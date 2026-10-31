# Patent Claims Checker — Design

**Date:** 2026-10-30
**Status:** Approved design, pending implementation plan
**Component:** LLDOCX Word add-in (VSTO, C#)
**Depends on:** the LL-Docx Sign-In Gate (`2026-10-31-lldocx-sign-in-gate-design.md`), which must be built first
**Revised 2026-10-31:** Deep Check now goes through an Azure gateway; no API key exists on users' machines

## Purpose

Add a tool to LLDOCX that identifies the claims in a patent document and checks them for:

1. Antecedent basis problems
2. Claim dependency problems
3. Typos and inconsistent terminology

Findings appear in a task pane. The document is never changed by the tool.

## Decisions

| Topic | Decision |
|---|---|
| Analysis engine | Hybrid: deterministic offline checks by default; optional "Deep Check" sends claims to Claude after explicit user confirmation |
| Output | Task pane list; clicking a finding selects the text in Word |
| Terminology scope | Claims only (no comparison against the specification) |
| Code structure | Pure-C# engine library + thin Word layer + unit-test project |
| Relationship to Scan | Separate ribbon buttons; existing Scan is unchanged |
| Access | Both buttons are gated by the Sign-In Gate (lozaip.com accounts only) |
| Deep Check backend | An Azure Function gateway owned by lozaip.com. It validates the user's Entra token and calls Claude. The model, prompt and credentials live only on the gateway |

## Architecture

```
Ribbon (Home > LL-Docx): [Scan] [Settings] [Check Claims] [Deep Check]
   │
ClaimsExtractor (LLDOCX add-in, Word interop)
   locate claims section → read paragraphs → List<RawClaim>
   │
LLDOCX.Claims (new class library, no Office references)
   ClaimParser → AntecedentChecker / DependencyChecker / TerminologyChecker
   │
List<Finding>
   │
ClaimsPaneControl (new CustomTaskPane) ← DeepCheckService merges AI findings
                                              │  HTTPS + user's Entra token (from the Sign-In Gate)
                                              ▼
                                   LLDOCX.Gateway (Azure Function) ──► Claude (Microsoft Foundry,
                                                                        or the Anthropic API)
```

### Projects

- **LLDOCX** (existing): adds `ClaimsExtractor`, `ClaimsPaneControl`, `DeepCheckService`, ribbon buttons and settings UI.
- **LLDOCX.Claims** (new netstandard2.0 class library): all parsing and checking. Input is plain text with offsets; output is findings. Testable without Word.
- **LLDOCX.DeepCheck** (new netstandard2.0 class library): the wire contract shared by the add-in and the gateway. It holds the request and response DTOs, the response parser and validator, and the mapping of AI findings back to document offsets. It has no Anthropic SDK reference.
- **LLDOCX.Gateway** (new Azure Functions app, .NET 8 isolated worker): validates tokens, enforces limits, builds the prompt, calls Claude and returns findings. It is the only component that references the Anthropic SDK.
- **LLDOCX.Claims.Tests** (new, MSTest on net8.0): tests of the engine, the Deep Check contract and the gateway's request handling, run with `dotnet test`.
- The add-in moves from .NET Framework 4.6.1 to **4.7.2**, as required by MSAL's broker package (Sign-In Gate) and the netstandard2.0 libraries.

### Core types (LLDOCX.Claims)

```csharp
class RawParagraph { string Text; int DocStart; }         // one document paragraph; DocStart = Word char offset of Text[0]
class Claim      { int Number; string Status;             // e.g. "Canceled", "Currently Amended", or null
                   string Preamble; string FullText; string PreambleNoun;
                   List<int> DependsOn; bool IsIndependent; bool IsMultipleDependent; bool HasCumulativeDependency;
                   ClaimCategory Category; int DocStart;
                   List<TextSegment> Segments; int ToDocOffset(int textIndex); }  // maps joined text back to the document
enum ClaimCategory { Method, Apparatus, Composition, ComputerReadableMedium, Unknown }
enum FindingType   { Antecedent, Dependency, Terminology, Formatting }
enum Severity      { Error, Warning, Info }
class Finding      { int ClaimNumber; FindingType Type; Severity Severity; string Message;
                     string Suggestion; int DocStart; int Length;
                     string Text;     // flagged text, used for stale detection
                     string Source;   // "Rules" | "AI"
                     bool DisputedByAI; }
class AnalysisResult { List<Claim> Claims; List<Finding> Findings; }
class ClaimsAnalyzer { AnalysisResult Analyze(IList<RawParagraph> paragraphs, ClaimsCheckOptions options); }
```

## Claim extraction (Word layer)

- If the user has a non-empty selection, only the selection is checked.
- Otherwise, find the claims heading. Matches are case-insensitive, on a paragraph of their own: `CLAIMS`, `WHAT IS CLAIMED IS:`, `WHAT IS CLAIMED:`, `I CLAIM:`, `WE CLAIM:`, `THE CLAIMS ARE:`, `CLAIMS:`. Use the last match before the `ABSTRACT` heading. If there is no `ABSTRACT` heading, use the last match in the document.
- The claims section runs from the heading to the next abstract heading (reuse the abstract heading list from `WordDocumentHelper.BuildDocumentMetaData`) or to the end of the document.
- Paragraph text is read with `Range.ListFormat.ListString` prepended when the paragraph is auto-numbered, so typed numbers and Word list numbers are handled the same way.
- Text is read from Word once. The engine runs on a background thread. All COM calls stay on the UI thread.
- If no claims section is found, show a message offering "Check selected text instead".

## Claim parsing

- A claim starts at a paragraph matching `^\s*(\d+)\s*[.)]\s+`. Following paragraphs that don't match are joined to the current claim, so claim elements on separate lines stay in their claim.
- Status markers right after the number are recognized: `(Canceled)`, `(Cancelled)`, `(Withdrawn)`, `(Original)`, `(Previously Presented)`, `(Currently Amended)`, `(New)`. Canceled claims are kept for numbering checks and skipped by content checks.
- The preamble is the text before the first transition (`comprising`, `consisting of`, `consisting essentially of`, `including`, `wherein`, `characterized in that`, `having`), ending at a colon or comma.
- Dependency references are matched with patterns covering:
  - `(of|in|according to|as (claimed|recited|defined|set forth) in) claim N`
  - `claims N or M`, `claims N, M, or P`
  - `any (one )?of claims N(-|–|to|through)M`
  - `any (one )?of the preceding claims`, which resolves to all earlier claims
- Category is determined from the first noun of the preamble:
  - `method` or `process` → Method
  - `system`, `apparatus`, `device`, `assembly`, `machine` (and similar) → Apparatus
  - `composition`, `compound`, `formulation` → Composition
  - `computer-readable`, `storage medium`, `storage media`, `program product` → ComputerReadableMedium
  - anything else → Unknown

## Checks

### Dependency checks

| Condition | Severity |
|---|---|
| Refers to a nonexistent claim | Error |
| Refers to itself | Error |
| Refers to a later-numbered claim | Error |
| Refers to a canceled claim | Error |
| Duplicate claim number | Error |
| Gap in claim numbering | Warning |
| Multiple dependent claim depends on another multiple dependent claim (37 CFR 1.75(c)) | Error |
| Multiple dependency written cumulatively ("claims 1 and 2") | Error |
| Multiple dependent claim present (fee/notice) | Info |
| Category mismatch (e.g. "The method of claim 3" where claim 3 is Apparatus) | Warning |
| Preamble noun mismatch (e.g. "The device of claim 1" where claim 1 recites "A system") | Warning |

### Antecedent basis checks

**Noun phrase extraction.** A noun phrase is:
- an optional determiner,
- then a run of modifier words,
- then a head noun.

The phrase ends at a boundary token: a preposition (`of`, `to`, `for`, `in`, `on`, `with`, `from`, `by`, `at`, `between`, `into`, `through`), a verb or participle cue (words ending in `-ing` or `-ed` that follow the head, `is`, `are`, `being`, `configured`, `adapted`, `operable`), a claim keyword (`wherein`, `whereby`, `comprising`, `including`, `having`, `and`, `or`), or punctuation. Ordinals (`first`, `second`, …) are kept as part of the phrase.

**Introducing determiners:** `a`, `an`, `one or more`, `at least one`, `a plurality of`, `plurality of`, `another`. (`each`, `any`, `one`, and bare plurals are not treated as introducers, which keeps false positives down.)

**Claim references** such as "The method of claim 1" or "the apparatus according to any of the preceding claims" are recognized and skipped by the antecedent check. They are handled by the dependency checks instead.

**Referring determiners:** `the`, `said`, `each of the`, `each of said`, `the at least one`, `the one or more`, `the plurality of`.

**Known terms for a claim.** The terms introduced in every ancestor along the dependency chain, plus the terms introduced earlier in the same claim. For a multiple dependent claim, a term has basis only if every referenced branch provides it.

**Matching order.**
1. Exact match on the normalized full phrase (lowercased, whitespace collapsed, simple plural stemming).
2. Match on head noun plus the same modifiers in a different order or with a subset of them. This is treated as a match.
3. Head noun only. This is a weak match and is reported as Info, not Error.

**Findings.**

| Condition | Severity |
|---|---|
| "the/said X" with no matching introduction | Error |
| "the X" when two or more distinct "X" terms were introduced (e.g. "a first X", "a second X") | Error |
| "the X" matched only on the head noun | Info |
| "a/an X" when X was already introduced in the chain | Warning |
| "the X" with no match but a close match exists | Error, with a suggestion ("did you mean 'the fastening member'?") |

**Allowlist** (no antecedent needed; editable in Settings): `the user`, `the art`, `the same`, `the following`, `the foregoing`, `the group`, `the presence`, `the absence`, `the range`, `the order`. An entry matches a phrase equal to it or starting with it ("the same temperature"). Phrases whose head noun is a unit of measure are also exempt.

**Singular/plural drift** is detected during this pass, because the match is already known at this point. It is reported with type Terminology and controlled by the terminology toggle.

### Terminology and typo checks (claims only)

| Condition | Severity |
|---|---|
| Two words (≥ 5 letters) within edit distance 1 (5–7 letters) or 2 (8+ letters), where one occurs ≥ 3× as often as the other; inflections such as "comprises"/"comprising" are excluded | Warning, suggesting the more frequent word |
| Singular/plural drift ("the sensors" when only "a sensor" was introduced) | Warning |
| Mixed use of "said" and "the" as referring determiners across the claim set | Info |
| Repeated word ("the the") | Error |
| Claim does not end with exactly one period | Warning |
| Period inside a claim body (not part of a number, abbreviation, or reference) | Warning |
| Unbalanced parentheses or brackets | Warning |

### Severity filtering

The pane shows findings at or above the minimum severity set in Settings. The default is Info, meaning everything is shown.

## Task pane (ClaimsPaneControl)

- A separate `CustomTaskPane` titled "LL-DOCX Claims Check". The existing Settings pane is not changed.
- Header: claim count, number of independent claims, and finding counts by severity. Controls: severity filter, group by Claim/Type, Re-run button.
- Each finding shows a severity icon, the claim number, the message, any suggestion, and its source ("Rules" or "AI").
- Clicking a finding selects `Document.Range(DocStart, DocStart + Length)` and scrolls to it. Before selecting, the range text is compared with the stored text. If they differ, the finding is marked "stale – re-run" and nothing is selected.
- AI findings have a "Dismiss" action. Rules findings that Deep Check disputes are shown struck through, with the AI's reason.

## Deep Check (Claude)

- Triggered by the "Deep Check" button. If no offline check has run yet, the offline check runs first.
- Before sending anything, a confirmation dialog appears every time. It shows the number of claims and an approximate character count, and offers "Don't ask again for this document" (remembered for the session, per document full name).
- The button is disabled when the gate is locked or offline, or when no gateway is configured.

### Add-in → gateway

- `POST {gatewayUrl}` with `Authorization: Bearer <token>`. The token comes from `ITokenSource.GetGatewayTokenAsync()` (Sign-In Gate), audience is the gateway API and the scope is `DeepCheck.Use`.
- **Request body (`DeepCheckRequest`):** `{ claims: [{ number, text }], offlineFindings: [{ index, claimNumber, type, message, text }] }`.
- **Response body (`DeepCheckResponse`):** `{ findings: [{ claimNumber, type, severity, message, suggestion, quotedText }], disputes: [{ index, reason }] }`.
- **Error responses** are `{ error: { code, message } }` with HTTP 400 (too large or invalid), 401/403 (token or role), 429 (per-user limit, with `Retry-After`), 422 (`refused`: the model declined) and 502/503 (upstream failure).
- The add-in sends no model name, prompt or credentials. It cannot choose or change them.
- AI findings are mapped back to document offsets. The AI returns the claim number and the quoted text, and the offset is found by searching within that claim's text. Findings that can't be located are shown without navigation.
- Network work runs off the UI thread, with a cancellable progress indicator. The HTTP timeout is 5 minutes.

### Gateway (LLDOCX.Gateway)

- **Authentication:** validates the JWT signature against the tenant's Entra signing keys. It checks:
  - the issuer is `https://login.microsoftonline.com/{tenantId}/v2.0`;
  - the audience is the gateway's app ID URI or client ID;
  - `tid` equals the configured tenant;
  - `scp` contains `DeepCheck.Use`;
  - if `RequiredRole` is set, `roles` contains it.

  Any failure returns 401 or 403.
- **Limits:**
  - total claims text at most `MaxInputChars` (default 200,000), otherwise 400;
  - at most `RequestsPerUserPerHour` per user `oid` (default 30), otherwise 429. The counter is held in memory per instance, which is acceptable at lozaip.com's scale.
- **Claude request:**
  - model: `Model` app setting, default `claude-opus-5-5` (on Foundry, the deployment name);
  - `output_config.effort`: `Effort` app setting, default `high`;
  - `output_config.format`: a JSON schema matching `DeepCheckResponse`;
  - system prompt: a US patent-claim reviewer focused only on antecedent basis, dependency and terminology;
  - user content: the numbered claims plus the offline findings with their indexes.
- **Provider** (`ClaudeProvider` app setting):
  - `Foundry` (preferred): `AnthropicFoundryClient` from the `Anthropic.Foundry` NuGet package, calling a Claude deployment in lozaip.com's Microsoft Foundry resource. It authenticates with the Function App's managed identity, so no key exists. If the SDK's Foundry client can't use a managed identity, the fallback is a Foundry key stored in Key Vault.
  - `Anthropic`: `AnthropicClient` with the API key read from Azure Key Vault through a Key Vault reference in app settings. It enables server-side refusal fallback (`fallbacks: "default"`, beta `server-side-fallback-2026-07-01`), which Foundry does not support.
- **Refusals:** the gateway checks `stop_reason` for `refusal` before reading content. A refusal returns 422 `refused`.
- **Logging:** Application Insights records the user's `oid`/UPN, the timestamp, claim count, input size, latency and outcome. It **never** records claim text or the model's output, since unfiled claims are confidential.
- **App settings** (Azure configuration, admin-controlled):
  - `TenantId`, `Audience`, `RequiredRole`
  - `ClaudeProvider`, `FoundryEndpoint`, `Model`, `Effort`
  - `AnthropicApiKey`: a Key Vault reference, used only by the `Anthropic` provider
  - `MaxInputChars`, `RequestsPerUserPerHour`
- **Hosting:** an Azure Functions Flex Consumption or Premium plan, HTTPS only, with the managed identity enabled. The function timeout is at least 5 minutes, because Deep Check calls can be long.
- **Still to verify** (the plan begins with a spike):
  - that `claude-opus-5-5` is deployable in Foundry in lozaip.com's region;
  - which authentication options `AnthropicFoundryClient` supports;
  - the exact SDK type names, using a compile-fix loop.

## Settings

- New fields on `SettingOption`, each with a default so existing saved settings still deserialize:
  - `ClaimsCheckAntecedent`, `ClaimsCheckDependency`, `ClaimsCheckTerminology` (bool, default true)
  - `ClaimsMinSeverity` (default Info)
  - `ClaimsAllowlist` (List<string>, defaults above)
- There is no model setting and no API key on the client. Both are gateway app settings controlled by the administrator.
- The Claims Check pane has a **Settings tab** with the toggles, the severity dropdown and an allowlist editor. They live there rather than in the existing Settings pane, whose fixed designer layout has no room. When the existing Settings pane saves, it keeps the claims fields that are on disk, so it doesn't overwrite them.

## Error handling

| Situation | Behavior |
|---|---|
| No claims section found | Message offering to check the selected text |
| Malformed or missing numbering | Parse what can be parsed; numbering problems become findings |
| Engine throws an exception | Log it, show "Claims check failed: <message>", and leave the pane unchanged |
| Deep Check: gate locked, offline or no gateway configured | Button disabled; tooltip explains why |
| Deep Check: token can't be acquired silently | Prompt to sign in again |
| Deep Check: 401/403 from gateway | "Your account isn't authorized for Deep Check. Contact your administrator." |
| Deep Check: 400 too large | "The claims are too long for Deep Check (limit N characters)." |
| Deep Check: 429 | "Deep Check limit reached; try again in N minutes." |
| Deep Check: 422 refused | "Claude declined to review these claims." Offline findings are kept |
| Deep Check: 502/503/timeout/no network | Message; offline findings are kept |
| Deep Check: response doesn't match the contract | Message; offline findings are kept |
| Gateway: Claude 429/5xx | The SDK retries, then the gateway returns 503 |
| Document edited after a check | Affected findings marked stale when clicked |

## Testing

- **LLDOCX.Claims.Tests** (MSTest, net8.0). The tests are inline claim sets with assertions on claim number, severity, and a substring of the message. They cover:
  - a clean claim set (zero findings)
  - missing antecedent, ambiguous first/second, re-introduction, weak head-noun match
  - multiple dependent chains, including the forbidden multiple-on-multiple case
  - method depending on apparatus, and preamble mismatch
  - missing/later/self/canceled references, numbering gaps and duplicates
  - typo variants, singular/plural drift, said/the mixing, repeated words, punctuation
  - status markers and canceled claims
- **False-positive guard:** at least three real claim sets from issued patents must produce zero Errors.
- Unit tests for the parser (numbering, joining paragraphs, dependency patterns, categories) and for noun-phrase extraction.
- **Gateway tests,** with the Claude client and token validation behind interfaces. They cover:
  - token rules (wrong tenant, missing scope, missing role);
  - the size limit;
  - the per-user rate limit;
  - mapping a refusal to 422 and an upstream failure to 503;
  - schema-valid output passed through;
  - a check that claim text never reaches the logger.
- **Deployed smoke test:** one Deep Check from Word against the real gateway.
- **Manual checklist** for the Word layer and Deep Check:
  - typed vs. auto-numbered claims
  - selection mode
  - navigation and stale detection
  - confirmation dialog
  - error messages with no network, an unauthorized account and the rate limit

## Out of scope

- Comparing claim terms against the specification (support/written description)
- Checking reference numerals
- Rules for non-US jurisdictions (EPO two-part form, and so on)
- Automatic fixes or rewriting claims
- Inserting findings as Word comments or highlights
