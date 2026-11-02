# Patent Claims Checker Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a "Check Claims" tool to the LLDOCX Word add-in. It finds the patent claims in the active document and reports antecedent basis, claim dependency, and terminology/typo problems in a task pane. An optional "Deep Check" sends the claims, after confirmation, to lozaip.com's Azure gateway, which has them reviewed by Claude.

**Architecture:**
- **`LLDOCX.Claims`** (netstandard2.0, no Office references) is a pure C# engine. It parses the claims and runs the checks on plain text with document offsets.
- **`LLDOCX.DeepCheck`** (netstandard2.0) holds the wire contract shared by the add-in and the gateway, plus the add-in's HTTP client.
- **`LLDOCX.Gateway`** (an Azure Function on .NET 10) validates the user's Entra token, enforces limits, and calls Claude through Microsoft Foundry or the Anthropic API. The model, prompt and credentials live only there.
- **The VSTO add-in** (`LLDOCX`, on .NET Framework 4.7.2 and gated by the sign-in gate) reads paragraphs from Word, runs the engine on a background thread, and shows the findings in a new task pane.
- **`LLDOCX.Claims.Tests` and `LLDOCX.Gateway.Tests`** (net10.0, MSTest) run with `dotnet test`.

**Tech Stack:**
- Add-in: C# / .NET Framework 4.7.2 VSTO (Word interop, WinForms).
- Libraries and tests: netstandard2.0 class libraries, `System.Text.Json`, MSTest on net10.0.
- Gateway: Azure Functions isolated worker on .NET 10, with the `Anthropic` and `Anthropic.Foundry` NuGet SDKs, `Azure.Identity` and `Microsoft.IdentityModel.*`.

**Spec:** `LLDOCX/docs/superpowers/specs/2026-10-30-patent-claims-checker-design.md`

**Prerequisite:** the sign-in gate plan, `LLDOCX/docs/superpowers/plans/2026-11-01-lldocx-sign-in-gate.md`, must be complete. It tracks the source, retargets the add-in to 4.7.2, installs the .NET 10 SDK, and adds `ThisAddIn.Gate`, `ThisAddIn.GateConfig`, `EnsureUnlockedAsync`, `HideAllPanes` and the per-machine installer.

## Global Constraints

- **Paths and shell:** all paths below are relative to the repository root `Tools/`. Shell commands run from `Tools/LLDOCX/` unless a step says otherwise.
- **Target frameworks:**
  - The add-in targets **.NET Framework 4.7.2**, set by the gate plan.
  - Libraries target **netstandard2.0**.
  - Test projects and the gateway target **net10.0**.
- **Dependencies:**
  - `LLDOCX.Claims` must not reference Office, WinForms, the Anthropic SDK, or `LLDOCX.DeepCheck`.
  - Only `LLDOCX.Gateway` references the Anthropic SDK. The add-in contains no Claude model setting, prompt, or credential.
- **Deep Check on the gateway:**
  - The default model is `claude-opus-5-5` (gateway app setting `Gateway__Model`).
  - Requests use `output_config.effort = high` and structured outputs (`output_config.format`).
  - `stop_reason` is checked for `refusal` (returned as 422) and `max_tokens` (returned as 502).
- **Gateway access:** the gateway accepts only Entra tokens from the lozaip.com tenant with the `DeepCheck.Use` scope (and the `RequiredRole`, when one is set). Claim text and Claude's output are never logged.
- **Consent:** no claims text leaves the PC without the confirmation dialog. The only exception is a document the user marked "Don't ask again" during the current session.
- **Gating:** every new ribbon button is gated by the sign-in gate (`getEnabled="GetEnabled"`), and each action starts with `EnsureUnlockedAsync()`.
- **The tool never modifies the document.**
- **Existing Scan behavior is unchanged.**
- **Severity order:** `Severity` enum values are `Error = 0`, `Warning = 1`, `Info = 2`. A severity filter keeps findings with `Severity <= minimum`.
- **Commits:** commit after every task. Each commit message ends with `Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>`.
- **Branch:** use `feature/patent-claims-checker`, which already exists.

## Build and test environment

- **Engine, contract and gateway with their tests (Tasks 2–10):** need the .NET 10 SDK, which the gate plan's Task 2 installs. They run with `dotnet test`.
- **VSTO add-in and installer (Tasks 11, 12, 13):** need Visual Studio 2022 with the **Office/SharePoint development** workload, the **.NET Framework 4.7.2 targeting pack**, and the WiX 3 extension. This machine has only VS Build Tools, so a human builds, runs and verifies those tasks in Visual Studio; F5 starts Word with the add-in loaded. Each ends with a manual verification checklist.
- **Azure deployment (Task 14):** needs lozaip.com's Azure subscription and an administrator. It follows the guide written in that task.

## File Structure

```
LLDOCX/
  LLDOCX.sln                                   (modify) add Claims, Claims.Tests, DeepCheck, Gateway, Gateway.Tests
  LLDOCX.Claims/                               (new) engine, netstandard2.0
    LLDOCX.Claims.csproj
    Models.cs              enums, RawParagraph, TextSegment, Claim, Finding
    TextUtil.cs            Token, Tokenize, Singularize, IsPlural, Levenshtein
    ClaimParser.cs         paragraphs -> claims (numbering, status, preamble, deps, category)
    DependencyChecker.cs   dependency/numbering findings
    NounPhraseExtractor.cs NounPhrase, DeterminerKind, Extract(text)
    AntecedentChecker.cs   antecedent findings (+ singular/plural drift)
    TerminologyChecker.cs  spelling variants, said/the mix, formatting
    ClaimsAnalyzer.cs      ClaimsCheckOptions, AnalysisResult, Analyze()
  LLDOCX.DeepCheck/                            (new) gateway wire contract + client, netstandard2.0
    LLDOCX.DeepCheck.csproj
    DeepCheckException.cs
    DeepCheckContract.cs   DeepCheckRequest, ClaimDto, RuleFindingDto, ErrorBody, ErrorCodes
    DeepCheckRequestBuilder.cs  claims + rule findings -> request
    DeepCheckResponseParser.cs  response JSON -> findings + disputes; IsWellFormed
    DeepCheckHttpClient.cs      POST to the gateway + error mapping
  LLDOCX.Claims.Tests/                         (new) MSTest, net10.0
    LLDOCX.Claims.Tests.csproj
    TestHelpers.cs
    TextUtilTests.cs, ClaimParserTests.cs, DependencyCheckerTests.cs,
    NounPhraseExtractorTests.cs, AntecedentCheckerTests.cs,
    TerminologyCheckerTests.cs, ClaimsAnalyzerTests.cs, DeepCheckProtocolTests.cs
  LLDOCX.Gateway/                              (new) Azure Functions isolated worker, net10.0
    LLDOCX.Gateway.csproj, host.json, local.settings.sample.json
    GatewayOptions.cs      app settings (Gateway__*)
    DeepCheckPrompt.cs     system prompt, JSON schema, user message builder
    TokenValidator.cs      Entra JWT validation (tenant, audience, scope, role)
    UserRateLimiter.cs     per-user hourly limit
    ClaudeReviewer.cs      Anthropic / Foundry SDK call
    DeepCheckHandler.cs    request pipeline (testable)
    DeepCheckFunction.cs   HTTP trigger POST /api/deepcheck
    Program.cs             DI + App Insights
  LLDOCX.Gateway.Tests/                        (new) MSTest, net10.0
    LLDOCX.Gateway.Tests.csproj, GatewayTests.cs
  LLDOCX/                                      (existing VSTO add-in)
    LLDOCX.csproj                              (modify) project refs, new files
    ThisAddIn.cs                               (modify) claims pane, controller
    RibbonExplorer.xml / RibbonExplorer.cs     (modify) two gated buttons + callbacks
    Options/SettingOption.cs                   (modify) claims settings fields
    CustomUserControl/SettingUserControl.cs    (modify) preserve claims settings on save
    Helpers/ClaimsExtractor.cs                 (new) Word -> RawParagraph list
    Helpers/ClaimsController.cs                (new) orchestrates check + deep check
    CustomUserControl/ClaimsPaneControl.cs     (new) findings tab + settings tab
    CustomUserControl/ClaimsSettingsPanel.cs   (new) claims settings UI
    CustomUserControl/DeepCheckConfirmDialog.cs (new) consent dialog
  LLDOCX.Auth/GateConfig.cs                    (modify, Task 14) compiled gateway URL + scope
  LLDOCXSetup/Product.wxs                      (modify) ship claims DLLs
  docs/admin/deep-check-setup.md               (new) Azure + Entra setup for Deep Check
```

---

### Task 1: Confirm the sign-in gate is in place

This task replaces the original Task 1. Tracking the source and retargeting to 4.7.2 moved to the sign-in gate plan.

**Files:** none

**Interfaces:**
- Consumes: the completed sign-in gate plan
- Produces: confirmation that the prerequisites exist

- [ ] **Step 1: Check the prerequisites**

Run, from `Tools/LLDOCX/`:

```bash
dotnet --list-sdks
git log --oneline -12
grep -n "EnsureUnlockedAsync\|HideAllPanes\|GateConfig" LLDOCX/ThisAddIn.cs
dotnet test LLDOCX.Auth.Tests
```

Expected:
- `dotnet --list-sdks` shows a `10.0.` SDK.
- The log shows the gate plan's commits, from "track add-in source" through "per-machine x64 MSI".
- `ThisAddIn.cs` contains all three members.
- The Auth tests pass.

If any check fails, finish the sign-in gate plan first. Nothing to commit.

---

### Task 2: Engine scaffold, core models, and text utilities

**Files:**
- Create: `LLDOCX/LLDOCX.Claims/LLDOCX.Claims.csproj`
- Create: `LLDOCX/LLDOCX.Claims/Models.cs`
- Create: `LLDOCX/LLDOCX.Claims/TextUtil.cs`
- Create: `LLDOCX/LLDOCX.Claims.Tests/LLDOCX.Claims.Tests.csproj`
- Create: `LLDOCX/LLDOCX.Claims.Tests/TestHelpers.cs`
- Test: `LLDOCX/LLDOCX.Claims.Tests/TextUtilTests.cs`
- Modify: `LLDOCX/LLDOCX.sln` (via `dotnet sln add`)

**Interfaces:**
- Consumes: nothing
- Produces (namespace `LLDOCX.Claims`):
  - `enum ClaimCategory { Unknown, Method, Apparatus, Composition, ComputerReadableMedium }`
  - `enum FindingType { Antecedent, Dependency, Terminology, Formatting }`
  - `enum Severity { Error = 0, Warning = 1, Info = 2 }`
  - `class RawParagraph(string text, int docStart)` with `Text`, `DocStart`
  - `class TextSegment(int textOffset, int docStart, int length)`
  - `class Claim`: `Number`, `Status`, `IsCanceled`, `FullText`, `Preamble`, `DependsOn` (List<int>), `IsIndependent`, `IsMultipleDependent`, `HasCumulativeDependency`, `DependencyRefStart`, `DependencyRefLength`, `Category`, `PreambleNoun`, `DocStart`, `Segments`, `int ToDocOffset(int textIndex)`
  - `class Finding`: `ClaimNumber`, `Type`, `Severity`, `Message`, `Suggestion`, `DocStart`, `Length`, `Text`, `Source` ("Rules"/"AI"), `DisputedByAI`, `DisputeReason`, `static Finding At(Claim, int start, int length, FindingType, Severity, string message, string suggestion = null)`
  - `struct Token(string value, int index)` with `Value`, `Index`, `End`, `IsWord`, `Lower`
  - `static class TextUtil`: `List<Token> Tokenize(string)`, `string Singularize(string lowerWord)`, `bool IsPlural(string lowerWord)`, `int Levenshtein(string, string)`
  - Test helpers `T.Paras(params string[])`, `T.Claims(params string[])`, `T.Joined(params string[])`, `T.Has(findings, claim, severity, text)`, `T.Dump(findings)`

- [ ] **Step 1: Confirm the .NET 10 SDK**

The gate plan's Task 2 installed it. Run: `dotnet --list-sdks`
Expected: a line starting with `10.0.` If it's missing, run `winget install --id Microsoft.DotNet.SDK.10 -e --accept-source-agreements --accept-package-agreements`, open a new shell, and check again.

- [ ] **Step 2: Create the engine project file `LLDOCX/LLDOCX.Claims/LLDOCX.Claims.csproj`**

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFramework>netstandard2.0</TargetFramework>
    <LangVersion>7.3</LangVersion>
    <RootNamespace>LLDOCX.Claims</RootNamespace>
  </PropertyGroup>
</Project>
```

- [ ] **Step 3: Create the test project `LLDOCX/LLDOCX.Claims.Tests/LLDOCX.Claims.Tests.csproj`**

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFramework>net10.0</TargetFramework>
    <ImplicitUsings>enable</ImplicitUsings>
    <Nullable>disable</Nullable>
    <IsPackable>false</IsPackable>
  </PropertyGroup>
  <ItemGroup>
    <PackageReference Include="Microsoft.NET.Test.Sdk" Version="17.11.1" />
    <PackageReference Include="MSTest.TestAdapter" Version="3.6.1" />
    <PackageReference Include="MSTest.TestFramework" Version="3.6.1" />
  </ItemGroup>
  <ItemGroup>
    <ProjectReference Include="..\LLDOCX.Claims\LLDOCX.Claims.csproj" />
  </ItemGroup>
</Project>
```

- [ ] **Step 4: Add both projects to the solution**

Run:

```bash
dotnet sln LLDOCX.sln add LLDOCX.Claims/LLDOCX.Claims.csproj LLDOCX.Claims.Tests/LLDOCX.Claims.Tests.csproj
```

Expected: `Project ... added to the solution.` twice.

- [ ] **Step 5: Write test helpers `LLDOCX/LLDOCX.Claims.Tests/TestHelpers.cs`**

```csharp
global using Microsoft.VisualStudio.TestTools.UnitTesting;
using LLDOCX.Claims;

namespace LLDOCX.Claims.Tests;

internal static class T
{
    /// <summary>One RawParagraph per line; DocStart is the line's offset in the "\n"-joined text.</summary>
    public static List<RawParagraph> Paras(params string[] lines)
    {
        var list = new List<RawParagraph>();
        int pos = 0;
        foreach (var line in lines)
        {
            list.Add(new RawParagraph(line, pos));
            pos += line.Length + 1;
        }
        return list;
    }

    public static string Joined(params string[] lines) => string.Join("\n", lines);

    public static List<Claim> Claims(params string[] lines) => throw new NotImplementedException(); // restored in Task 3

    public static bool Has(IEnumerable<Finding> findings, int claim, Severity severity, string text) =>
        findings.Any(f => f.ClaimNumber == claim && f.Severity == severity &&
                          f.Message.Contains(text, StringComparison.OrdinalIgnoreCase));

    public static string Dump(IEnumerable<Finding> findings) =>
        string.Join("\n", findings.Select(f => $"{f.ClaimNumber} {f.Severity} {f.Type}: {f.Message} | {f.Suggestion}"));
}
```

- [ ] **Step 6: Write the failing tests `LLDOCX/LLDOCX.Claims.Tests/TextUtilTests.cs`**

```csharp
using LLDOCX.Claims;

namespace LLDOCX.Claims.Tests;

[TestClass]
public class TextUtilTests
{
    [DataTestMethod]
    [DataRow("sensors", "sensor")]
    [DataRow("assemblies", "assembly")]
    [DataRow("processes", "process")]
    [DataRow("boxes", "box")]
    [DataRow("apparatus", "apparatus")]
    [DataRow("glass", "glass")]
    [DataRow("axis", "axis")]
    public void Singularize_handles_common_plural_forms(string word, string expected) =>
        Assert.AreEqual(expected, TextUtil.Singularize(word));

    [DataTestMethod]
    [DataRow("bracket", "braket", 1)]
    [DataRow("housing", "houseing", 1)]
    [DataRow("kitten", "sitting", 3)]
    [DataRow("", "abc", 3)]
    public void Levenshtein_computes_edit_distance(string a, string b, int expected) =>
        Assert.AreEqual(expected, TextUtil.Levenshtein(a, b));

    [TestMethod]
    public void Tokenize_splits_words_numbers_and_punctuation_with_offsets()
    {
        var tokens = TextUtil.Tokenize("a non-transitory medium 12, wherein");
        CollectionAssert.AreEqual(new[] { "a", "non-transitory", "medium", "12", ",", "wherein" },
            tokens.Select(t => t.Value).ToArray());
        Assert.AreEqual(2, tokens[1].Index);
        Assert.IsTrue(tokens[1].IsWord);
        Assert.IsFalse(tokens[3].IsWord);
    }

    [TestMethod]
    public void Claim_maps_text_offsets_across_segments()
    {
        var c = new Claim { FullText = "abcdefghij klmno" };
        c.Segments.Add(new TextSegment(0, 100, 10));
        c.Segments.Add(new TextSegment(11, 200, 5));
        Assert.AreEqual(103, c.ToDocOffset(3));
        Assert.AreEqual(110, c.ToDocOffset(10));
        Assert.AreEqual(200, c.ToDocOffset(11));
        Assert.AreEqual(201, c.ToDocOffset(12));
    }

    [TestMethod]
    public void Finding_At_computes_document_range_and_text()
    {
        var c = new Claim { Number = 4, FullText = "abcdefghij klmno" };
        c.Segments.Add(new TextSegment(0, 100, 10));
        c.Segments.Add(new TextSegment(11, 200, 5));
        var f = Finding.At(c, 12, 3, FindingType.Antecedent, Severity.Error, "msg");
        Assert.AreEqual(4, f.ClaimNumber);
        Assert.AreEqual(201, f.DocStart);
        Assert.AreEqual(3, f.Length);
        Assert.AreEqual("lmn", f.Text);
        Assert.AreEqual("Rules", f.Source);
    }
}
```

- [ ] **Step 7: Run tests to verify they fail**

Run: `dotnet test LLDOCX.Claims.Tests`
Expected: build FAILS with errors such as `The name 'TextUtil' does not exist` and `The type or namespace name 'Claim' could not be found`.

- [ ] **Step 8: Write `LLDOCX/LLDOCX.Claims/Models.cs`**

```csharp
using System;
using System.Collections.Generic;

namespace LLDOCX.Claims
{
    public enum ClaimCategory { Unknown, Method, Apparatus, Composition, ComputerReadableMedium }

    public enum FindingType { Antecedent, Dependency, Terminology, Formatting }

    /// <summary>Ordered so that a filter can keep findings with Severity &lt;= minimum.</summary>
    public enum Severity { Error = 0, Warning = 1, Info = 2 }

    /// <summary>One paragraph of document text. DocStart is the document offset of Text[0].</summary>
    public class RawParagraph
    {
        public RawParagraph(string text, int docStart)
        {
            Text = text ?? "";
            DocStart = docStart;
        }

        public string Text { get; }
        public int DocStart { get; }
    }

    /// <summary>Maps a run of Claim.FullText (starting at TextOffset) to the document (starting at DocStart).</summary>
    public class TextSegment
    {
        public TextSegment(int textOffset, int docStart, int length)
        {
            TextOffset = textOffset;
            DocStart = docStart;
            Length = length;
        }

        public int TextOffset { get; }
        public int DocStart { get; }
        public int Length { get; }
    }

    public class Claim
    {
        public int Number { get; set; }
        public string Status { get; set; }
        public bool IsCanceled => Status == "Canceled";
        public string FullText { get; set; } = "";
        public string Preamble { get; set; } = "";
        public List<int> DependsOn { get; } = new List<int>();
        public bool IsIndependent => DependsOn.Count == 0;
        public bool IsMultipleDependent { get; set; }
        public bool HasCumulativeDependency { get; set; }
        public int DependencyRefStart { get; set; } = -1;
        public int DependencyRefLength { get; set; }
        public ClaimCategory Category { get; set; }
        public string PreambleNoun { get; set; } = "";
        /// <summary>Document offset of the claim's content (after its number and status).</summary>
        public int DocStart { get; set; }
        public List<TextSegment> Segments { get; } = new List<TextSegment>();

        public int ToDocOffset(int textIndex)
        {
            if (Segments.Count == 0) return DocStart;
            foreach (var s in Segments)
            {
                if (textIndex < s.TextOffset) return s.DocStart;
                if (textIndex <= s.TextOffset + s.Length) return s.DocStart + (textIndex - s.TextOffset);
            }
            var last = Segments[Segments.Count - 1];
            return last.DocStart + last.Length;
        }
    }

    public class Finding
    {
        public int ClaimNumber { get; set; }
        public FindingType Type { get; set; }
        public Severity Severity { get; set; }
        public string Message { get; set; } = "";
        public string Suggestion { get; set; }
        /// <summary>Document offset of the flagged text, or -1 when it could not be located.</summary>
        public int DocStart { get; set; } = -1;
        public int Length { get; set; }
        /// <summary>The flagged text as it appeared when checked; used for stale detection.</summary>
        public string Text { get; set; } = "";
        public string Source { get; set; } = "Rules";
        public bool DisputedByAI { get; set; }
        public string DisputeReason { get; set; }

        public static Finding At(Claim claim, int start, int length, FindingType type, Severity severity,
            string message, string suggestion = null)
        {
            string text = claim.FullText ?? "";
            start = Math.Max(0, Math.Min(start, text.Length));
            length = Math.Max(0, Math.Min(length, text.Length - start));
            int docStart = claim.ToDocOffset(start);
            return new Finding
            {
                ClaimNumber = claim.Number,
                Type = type,
                Severity = severity,
                Message = message,
                Suggestion = suggestion,
                DocStart = docStart,
                Length = claim.ToDocOffset(start + length) - docStart,
                Text = text.Substring(start, length)
            };
        }
    }
}
```

- [ ] **Step 9: Write `LLDOCX/LLDOCX.Claims/TextUtil.cs`**

```csharp
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace LLDOCX.Claims
{
    public struct Token
    {
        public Token(string value, int index)
        {
            Value = value;
            Index = index;
        }

        public string Value { get; }
        public int Index { get; }
        public int End => Index + Value.Length;
        public bool IsWord => Value.Length > 0 && char.IsLetter(Value[0]);
        public string Lower => Value.ToLowerInvariant();
    }

    public static class TextUtil
    {
        static readonly Regex TokenRx = new Regex(@"[A-Za-z][A-Za-z0-9'\-]*|\d+(?:\.\d+)?|[^\sA-Za-z0-9]", RegexOptions.Compiled);

        public static List<Token> Tokenize(string text)
        {
            var list = new List<Token>();
            foreach (Match m in TokenRx.Matches(text ?? ""))
                list.Add(new Token(m.Value, m.Index));
            return list;
        }

        /// <summary>Rule-based singular form of a lowercase English word.</summary>
        public static string Singularize(string word)
        {
            if (string.IsNullOrEmpty(word) || word.Length <= 3) return word;
            if (word.EndsWith("ies") && word.Length > 4) return word.Substring(0, word.Length - 3) + "y";
            if (word.EndsWith("sses")) return word.Substring(0, word.Length - 2);
            if (word.EndsWith("xes") || word.EndsWith("ches") || word.EndsWith("shes")) return word.Substring(0, word.Length - 2);
            if (word.EndsWith("ss") || word.EndsWith("us") || word.EndsWith("is")) return word;
            if (word.EndsWith("s")) return word.Substring(0, word.Length - 1);
            return word;
        }

        public static bool IsPlural(string word) => Singularize(word) != word;

        public static int Levenshtein(string a, string b)
        {
            a = a ?? "";
            b = b ?? "";
            var prev = new int[b.Length + 1];
            var cur = new int[b.Length + 1];
            for (int j = 0; j <= b.Length; j++) prev[j] = j;
            for (int i = 1; i <= a.Length; i++)
            {
                cur[0] = i;
                for (int j = 1; j <= b.Length; j++)
                {
                    int cost = a[i - 1] == b[j - 1] ? 0 : 1;
                    cur[j] = Math.Min(Math.Min(cur[j - 1] + 1, prev[j] + 1), prev[j - 1] + cost);
                }
                var tmp = prev; prev = cur; cur = tmp;
            }
            return prev[b.Length];
        }
    }
}
```

- [ ] **Step 10: Run tests to verify they pass**

Run: `dotnet test LLDOCX.Claims.Tests`
Expected: `Passed!` with `Failed: 0`.

- [ ] **Step 11: Commit**

```bash
git add LLDOCX/LLDOCX.sln LLDOCX/LLDOCX.Claims LLDOCX/LLDOCX.Claims.Tests
git commit -m "feat(claims): add engine scaffold, core models, and text utilities

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 3: ClaimParser

**Files:**
- Create: `LLDOCX/LLDOCX.Claims/ClaimParser.cs`
- Modify: `LLDOCX/LLDOCX.Claims.Tests/TestHelpers.cs` (restore `T.Claims`)
- Test: `LLDOCX/LLDOCX.Claims.Tests/ClaimParserTests.cs`

**Interfaces:**
- Consumes: `RawParagraph`, `Claim`, `TextSegment`, `ClaimCategory`, `TextUtil.Singularize` (Task 2)
- Produces: `static class ClaimParser { static List<Claim> Parse(IList<RawParagraph> paragraphs); }`. Each `Claim` comes out with Number, Status, FullText, Segments, DocStart, Preamble, DependsOn, IsMultipleDependent, HasCumulativeDependency, DependencyRefStart/Length, Category, and PreambleNoun (lowercase, singular) populated.

- [ ] **Step 1: Restore `T.Claims` in `TestHelpers.cs`**

Replace the `Claims` line with:

```csharp
    public static List<Claim> Claims(params string[] lines) => ClaimParser.Parse(Paras(lines));
```

- [ ] **Step 2: Write the failing tests `LLDOCX/LLDOCX.Claims.Tests/ClaimParserTests.cs`**

```csharp
using LLDOCX.Claims;

namespace LLDOCX.Claims.Tests;

[TestClass]
public class ClaimParserTests
{
    static readonly string[] DeviceClaims =
    {
        "What is claimed is:",
        "1. A device comprising:",
        "a housing; and",
        "a lid attached to the housing.",
        "2. The device of claim 1, wherein the lid is hinged."
    };

    [TestMethod]
    public void Parses_typed_numbers_and_joins_continuation_paragraphs()
    {
        var claims = T.Claims(DeviceClaims);
        Assert.AreEqual(2, claims.Count);
        Assert.AreEqual(1, claims[0].Number);
        Assert.AreEqual("A device comprising: a housing; and a lid attached to the housing.", claims[0].FullText);
        Assert.IsTrue(claims[0].IsIndependent);
        CollectionAssert.AreEqual(new[] { 1 }, claims[1].DependsOn);
    }

    [TestMethod]
    public void Maps_claim_text_offsets_back_to_document_offsets()
    {
        var claims = T.Claims(DeviceClaims);
        string doc = T.Joined(DeviceClaims);
        int textIndex = claims[0].FullText.IndexOf("housing");
        Assert.AreEqual(doc.IndexOf("housing"), claims[0].ToDocOffset(textIndex));
        Assert.AreEqual(doc.IndexOf("A device"), claims[0].DocStart);
    }

    [TestMethod]
    public void Recognizes_status_markers()
    {
        var claims = T.Claims("1. (Canceled)", "2. (Currently Amended) A method comprising: heating a fluid.");
        Assert.IsTrue(claims[0].IsCanceled);
        Assert.AreEqual("", claims[0].FullText);
        Assert.AreEqual("Currently Amended", claims[1].Status);
        Assert.IsTrue(claims[1].FullText.StartsWith("A method"));
    }

    [DataTestMethod]
    [DataRow("2. The method of claim 1, wherein x.", "1", false, false)]
    [DataRow("3. The method according to claim 1 or 2, wherein x.", "1,2", true, false)]
    [DataRow("4. The method of any one of claims 1-3, wherein x.", "1,2,3", true, false)]
    [DataRow("4. The method of claims 1 and 2, wherein x.", "1,2", true, true)]
    [DataRow("5. The method of any of the preceding claims, wherein x.", "1,2,3,4", true, false)]
    [DataRow("6. The method as claimed in claims 1, 2, or 3, wherein x.", "1,2,3", true, false)]
    public void Parses_dependency_forms(string line, string deps, bool multiple, bool cumulative)
    {
        var c = T.Claims(line)[0];
        Assert.AreEqual(deps, string.Join(",", c.DependsOn));
        Assert.AreEqual(multiple, c.IsMultipleDependent);
        Assert.AreEqual(cumulative, c.HasCumulativeDependency);
        Assert.IsTrue(c.DependencyRefStart >= 0);
    }

    [DataTestMethod]
    [DataRow("1. A method of operating a device, comprising: x.", ClaimCategory.Method, "method")]
    [DataRow("1. A system for processing data, comprising: x.", ClaimCategory.Apparatus, "system")]
    [DataRow("1. A non-transitory computer-readable medium storing instructions that cause: x.", ClaimCategory.ComputerReadableMedium, "medium")]
    [DataRow("1. A composition comprising: x.", ClaimCategory.Composition, "composition")]
    [DataRow("2. The fluid dispensing apparatus of claim 1, wherein x.", ClaimCategory.Apparatus, "apparatus")]
    public void Detects_category_and_preamble_noun(string line, ClaimCategory category, string noun)
    {
        var c = T.Claims(line)[0];
        Assert.AreEqual(category, c.Category);
        Assert.AreEqual(noun, c.PreambleNoun);
    }

    [TestMethod]
    public void Ignores_text_before_first_claim_and_blank_paragraphs()
    {
        var claims = T.Claims("CLAIMS", "", "1. A cup comprising a handle.", "", "2. The cup of claim 1, wherein the handle is round.");
        Assert.AreEqual(2, claims.Count);
        Assert.AreEqual("A cup comprising a handle.", claims[0].FullText);
    }

    [TestMethod]
    public void Does_not_treat_decimal_numbers_as_claim_starts()
    {
        var claims = T.Claims("1. A rod comprising a diameter of", "2.5 mm.");
        Assert.AreEqual(1, claims.Count);
        Assert.AreEqual("A rod comprising a diameter of 2.5 mm.", claims[0].FullText);
    }
}
```

- [ ] **Step 3: Run tests to verify they fail**

Run: `dotnet test LLDOCX.Claims.Tests --filter ClaimParserTests`
Expected: build FAILS with `The name 'ClaimParser' does not exist in the current context`.

- [ ] **Step 4: Write `LLDOCX/LLDOCX.Claims/ClaimParser.cs`**

```csharp
using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace LLDOCX.Claims
{
    public static class ClaimParser
    {
        const RegexOptions Opts = RegexOptions.Compiled | RegexOptions.IgnoreCase;

        static readonly Regex ClaimStart = new Regex(@"^\s*(\d{1,3})\s*[.)](?!\d)\s*", RegexOptions.Compiled);
        static readonly Regex StatusRx = new Regex(@"^\((canceled|cancelled|withdrawn|original|previously presented|currently amended|new)\)\s*", Opts);
        static readonly Regex Transition = new Regex(@"\b(?:comprising|consisting essentially of|consisting of|including|wherein|characterized in that|having)\b|:", Opts);
        static readonly Regex Preceding = new Regex(@"\b(?:any|one)\s+(?:one\s+)?of\s+the\s+(?:preceding|previous|foregoing)\s+claims\b", Opts);
        static readonly Regex DepRef = new Regex(@"\bclaims?\s+(\d+(?:\s*(?:-|–|to|through|,|or|and|, or|, and)\s*\d+)*)", Opts);
        static readonly Regex NumberOrRange = new Regex(@"(\d+)(?:\s*(?:-|–|to|through)\s*(\d+))?", RegexOptions.Compiled);
        static readonly Regex AndRx = new Regex(@"\band\b", Opts);
        static readonly Regex LeadingDeterminer = new Regex(@"^\s*(?:a|an|the|one)\s+", Opts);
        static readonly Regex NounStop = new Regex(@"\b(?:for|of|according|as|in|comprising|consisting|including|having|wherein|that|which|configured|to|storing|with|on)\b|[,:;]", Opts);

        static readonly (Regex Pattern, ClaimCategory Category)[] CategoryKeywords =
        {
            (new Regex(@"computer[- ]readable|machine[- ]readable|processor[- ]readable|storage (?:medium|media)|program product", Opts), ClaimCategory.ComputerReadableMedium),
            (new Regex(@"\b(?:method|process)\b", Opts), ClaimCategory.Method),
            (new Regex(@"\b(?:composition|compound|formulation|mixture)\b", Opts), ClaimCategory.Composition),
            (new Regex(@"\b(?:system|apparatus|device|assembly|machine|circuit|kit|vehicle|tool|article|structure|instrument)\b", Opts), ClaimCategory.Apparatus),
        };

        public static List<Claim> Parse(IList<RawParagraph> paragraphs)
        {
            var claims = new List<Claim>();
            Claim current = null;
            StringBuilder sb = null;
            foreach (var p in paragraphs)
            {
                string text = p.Text.TrimEnd('\r', '\n', '\a', ' ', '\t');
                if (text.Trim().Length == 0) continue;

                var m = ClaimStart.Match(text);
                if (m.Success)
                {
                    if (current != null) Finish(current, sb);
                    current = new Claim { Number = int.Parse(m.Groups[1].Value) };
                    sb = new StringBuilder();
                    int contentStart = m.Length;
                    var sm = StatusRx.Match(text.Substring(contentStart));
                    if (sm.Success)
                    {
                        current.Status = CanonicalStatus(sm.Groups[1].Value);
                        contentStart += sm.Length;
                    }
                    current.DocStart = p.DocStart + contentStart;
                    Append(current, sb, text, contentStart, p.DocStart);
                    claims.Add(current);
                }
                else if (current != null)
                {
                    int lead = text.Length - text.TrimStart().Length;
                    Append(current, sb, text, lead, p.DocStart);
                }
            }
            if (current != null) Finish(current, sb);
            return claims;
        }

        static void Append(Claim claim, StringBuilder sb, string text, int start, int docStart)
        {
            string piece = text.Substring(start);
            if (piece.Length == 0) return;
            if (sb.Length > 0) sb.Append(' ');
            claim.Segments.Add(new TextSegment(sb.Length, docStart + start, piece.Length));
            sb.Append(piece);
        }

        static void Finish(Claim claim, StringBuilder sb)
        {
            claim.FullText = sb.ToString();
            var t = Transition.Match(claim.FullText);
            claim.Preamble = t.Success ? claim.FullText.Substring(0, t.Index) : claim.FullText;
            ParseDependencies(claim);
            claim.Category = DetectCategory(claim.Preamble);
            claim.PreambleNoun = DetectPreambleNoun(claim.Preamble);
        }

        static string CanonicalStatus(string status)
        {
            switch (status.ToLowerInvariant())
            {
                case "canceled":
                case "cancelled": return "Canceled";
                case "withdrawn": return "Withdrawn";
                case "original": return "Original";
                case "previously presented": return "Previously Presented";
                case "currently amended": return "Currently Amended";
                default: return "New";
            }
        }

        static void ParseDependencies(Claim claim)
        {
            string region = claim.Preamble;
            var pm = Preceding.Match(region);
            if (pm.Success)
            {
                for (int i = 1; i < claim.Number; i++) claim.DependsOn.Add(i);
                claim.IsMultipleDependent = true;
                claim.DependencyRefStart = pm.Index;
                claim.DependencyRefLength = pm.Length;
                return;
            }

            var matches = DepRef.Matches(region);
            if (matches.Count == 0) return;
            var first = matches[0];
            var last = matches[matches.Count - 1];
            claim.DependencyRefStart = first.Index;
            claim.DependencyRefLength = last.Index + last.Length - first.Index;

            bool hasRange = false;
            foreach (Match m in matches)
            {
                foreach (Match r in NumberOrRange.Matches(m.Groups[1].Value))
                {
                    int a = int.Parse(r.Groups[1].Value);
                    if (r.Groups[2].Success)
                    {
                        hasRange = true;
                        int b = int.Parse(r.Groups[2].Value);
                        for (int i = Math.Min(a, b); i <= Math.Max(a, b); i++) AddDependency(claim, i);
                    }
                    else AddDependency(claim, a);
                }
            }
            claim.IsMultipleDependent = hasRange || claim.DependsOn.Count > 1;
            string span = region.Substring(claim.DependencyRefStart, claim.DependencyRefLength);
            claim.HasCumulativeDependency = claim.IsMultipleDependent && AndRx.IsMatch(span);
        }

        static void AddDependency(Claim claim, int number)
        {
            if (!claim.DependsOn.Contains(number)) claim.DependsOn.Add(number);
        }

        static ClaimCategory DetectCategory(string preamble)
        {
            int best = int.MaxValue;
            var result = ClaimCategory.Unknown;
            foreach (var (pattern, category) in CategoryKeywords)
            {
                var m = pattern.Match(preamble);
                if (m.Success && m.Index < best)
                {
                    best = m.Index;
                    result = category;
                }
            }
            return result;
        }

        static string DetectPreambleNoun(string preamble)
        {
            string p = LeadingDeterminer.Replace(preamble, "", 1);
            var stop = NounStop.Match(p);
            string phrase = stop.Success ? p.Substring(0, stop.Index) : p;
            var words = phrase.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
            return words.Length == 0 ? "" : TextUtil.Singularize(words[words.Length - 1].ToLowerInvariant());
        }
    }
}
```

- [ ] **Step 5: Run tests to verify they pass**

Run: `dotnet test LLDOCX.Claims.Tests`
Expected: all tests pass, 0 failed.

- [ ] **Step 6: Commit**

```bash
git add LLDOCX/LLDOCX.Claims/ClaimParser.cs LLDOCX/LLDOCX.Claims.Tests
git commit -m "feat(claims): parse claim numbering, status, dependencies, and category

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 4: DependencyChecker

**Files:**
- Create: `LLDOCX/LLDOCX.Claims/DependencyChecker.cs`
- Test: `LLDOCX/LLDOCX.Claims.Tests/DependencyCheckerTests.cs`

**Interfaces:**
- Consumes: `ClaimParser.Parse`, `Claim`, `Finding.At` (Tasks 2–3)
- Produces: `static class DependencyChecker { static List<Finding> Check(IList<Claim> claims); }`. All findings have `Type = FindingType.Dependency`.

- [ ] **Step 1: Write the failing tests `LLDOCX/LLDOCX.Claims.Tests/DependencyCheckerTests.cs`**

```csharp
using LLDOCX.Claims;

namespace LLDOCX.Claims.Tests;

[TestClass]
public class DependencyCheckerTests
{
    [TestMethod]
    public void Flags_self_forward_nonexistent_references_and_numbering_gaps()
    {
        var f = DependencyChecker.Check(T.Claims(
            "1. A widget comprising a base.",
            "2. The widget of claim 2, wherein the base is flat.",
            "3. The widget of claim 5, wherein the base is round.",
            "5. The widget of claim 4, wherein the base is red."));
        string dump = T.Dump(f);
        Assert.IsTrue(T.Has(f, 2, Severity.Error, "depends on itself"), dump);
        Assert.IsTrue(T.Has(f, 3, Severity.Error, "comes after it"), dump);
        Assert.IsTrue(T.Has(f, 5, Severity.Error, "does not exist"), dump);
        Assert.IsTrue(T.Has(f, 5, Severity.Warning, "jumps from 3 to 5"), dump);
    }

    [TestMethod]
    public void Flags_duplicate_numbers_and_canceled_parents()
    {
        var f = DependencyChecker.Check(T.Claims(
            "1. A widget comprising a base.",
            "2. (Canceled)",
            "3. The widget of claim 2, wherein the base is flat.",
            "3. The widget of claim 1, wherein the base is round."));
        string dump = T.Dump(f);
        Assert.IsTrue(T.Has(f, 3, Severity.Error, "canceled claim 2"), dump);
        Assert.IsTrue(T.Has(f, 3, Severity.Error, "used more than once"), dump);
    }

    [TestMethod]
    public void Flags_improper_multiple_dependencies()
    {
        var f = DependencyChecker.Check(T.Claims(
            "1. A pump comprising a rotor.",
            "2. The pump of claim 1, wherein the rotor is steel.",
            "3. The pump of claim 1 or 2, wherein the rotor is coated.",
            "4. The pump of claim 2 or 3, further comprising a seal.",
            "5. The pump of claims 1 and 2, further comprising a valve."));
        string dump = T.Dump(f);
        Assert.IsTrue(T.Has(f, 4, Severity.Error, "not permitted"), dump);
        Assert.IsTrue(T.Has(f, 5, Severity.Error, "cumulatively"), dump);
        Assert.IsTrue(T.Has(f, 3, Severity.Info, "multiple dependent"), dump);
    }

    [TestMethod]
    public void Flags_category_and_preamble_mismatch()
    {
        var f = DependencyChecker.Check(T.Claims(
            "1. A system comprising a processor.",
            "2. The method of claim 1, wherein the processor is fast.",
            "3. The device of claim 1, wherein the processor is slow."));
        string dump = T.Dump(f);
        Assert.IsTrue(T.Has(f, 2, Severity.Warning, "different category"), dump);
        Assert.IsTrue(T.Has(f, 3, Severity.Warning, "'the device'"), dump);
    }

    [TestMethod]
    public void Clean_claim_set_has_no_findings()
    {
        var f = DependencyChecker.Check(T.Claims(
            "1. A lamp comprising a bulb.",
            "2. The lamp of claim 1, wherein the bulb is an LED."));
        Assert.AreEqual(0, f.Count, T.Dump(f));
    }
}
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `dotnet test LLDOCX.Claims.Tests --filter DependencyCheckerTests`
Expected: build FAILS with `The name 'DependencyChecker' does not exist in the current context`.

- [ ] **Step 3: Write `LLDOCX/LLDOCX.Claims/DependencyChecker.cs`**

```csharp
using System;
using System.Collections.Generic;
using System.Linq;

namespace LLDOCX.Claims
{
    public static class DependencyChecker
    {
        public static List<Finding> Check(IList<Claim> claims)
        {
            var findings = new List<Finding>();
            var byNumber = new Dictionary<int, Claim>();
            foreach (var c in claims)
                if (!byNumber.ContainsKey(c.Number)) byNumber[c.Number] = c;

            for (int i = 0; i < claims.Count; i++)
            {
                var c = claims[i];
                if (!ReferenceEquals(byNumber[c.Number], c))
                    findings.Add(AtStart(c, Severity.Error, $"Claim number {c.Number} is used more than once."));
                if (i > 0)
                {
                    int prev = claims[i - 1].Number;
                    if (c.Number != prev && c.Number != prev + 1)
                        findings.Add(AtStart(c, Severity.Warning, $"Claim numbering jumps from {prev} to {c.Number}."));
                }
            }

            foreach (var c in claims)
            {
                if (c.IsCanceled || c.IsIndependent) continue;

                foreach (var d in c.DependsOn)
                {
                    string message = null;
                    if (d == c.Number) message = $"Claim {c.Number} depends on itself.";
                    else if (d > c.Number) message = $"Claim {c.Number} depends on claim {d}, which comes after it. A claim may only refer to a preceding claim.";
                    else if (!byNumber.ContainsKey(d)) message = $"Claim {c.Number} refers to claim {d}, which does not exist.";
                    else if (byNumber[d].IsCanceled) message = $"Claim {c.Number} depends on canceled claim {d}.";
                    if (message != null) findings.Add(AtReference(c, Severity.Error, message));
                }

                if (c.HasCumulativeDependency)
                    findings.Add(AtReference(c, Severity.Error,
                        $"Claim {c.Number} refers to more than one claim cumulatively ('and'). A multiple dependent claim may refer to other claims only in the alternative ('or') (37 CFR 1.75(c))."));

                if (c.IsMultipleDependent)
                {
                    var multiParents = c.DependsOn
                        .Where(d => d < c.Number && byNumber.ContainsKey(d) && byNumber[d].IsMultipleDependent)
                        .ToList();
                    if (multiParents.Count > 0)
                        findings.Add(AtReference(c, Severity.Error,
                            $"Multiple dependent claim {c.Number} depends on multiple dependent claim(s) {string.Join(", ", multiParents)}, which is not permitted (37 CFR 1.75(c))."));
                    else
                        findings.Add(AtReference(c, Severity.Info,
                            $"Claim {c.Number} is a multiple dependent claim (an additional fee applies)."));
                }

                var parent = FirstValidParent(c, byNumber);
                if (parent == null) continue;
                var root = Root(parent, byNumber);
                if (c.Category != ClaimCategory.Unknown && parent.Category != ClaimCategory.Unknown && c.Category != parent.Category)
                {
                    findings.Add(AtReference(c, Severity.Warning,
                        $"Claim {c.Number} claims a different category ({Name(c.Category)}) than claim {parent.Number} ({Name(parent.Category)}), on which it depends."));
                }
                else if (c.Category == root.Category && c.PreambleNoun.Length > 0 && root.PreambleNoun.Length > 0 && c.PreambleNoun != root.PreambleNoun)
                {
                    findings.Add(AtReference(c, Severity.Warning,
                        $"Claim {c.Number} refers to 'the {c.PreambleNoun}' but claim {root.Number} claims 'a {root.PreambleNoun}'."));
                }
            }
            return findings;
        }

        static Claim FirstValidParent(Claim c, Dictionary<int, Claim> byNumber)
        {
            foreach (var d in c.DependsOn)
            {
                Claim p;
                if (d < c.Number && byNumber.TryGetValue(d, out p) && !p.IsCanceled) return p;
            }
            return null;
        }

        static Claim Root(Claim c, Dictionary<int, Claim> byNumber)
        {
            var visited = new HashSet<int>();
            var current = c;
            while (!current.IsIndependent && visited.Add(current.Number))
            {
                var next = FirstValidParent(current, byNumber);
                if (next == null) break;
                current = next;
            }
            return current;
        }

        static string Name(ClaimCategory category)
        {
            switch (category)
            {
                case ClaimCategory.Method: return "method";
                case ClaimCategory.Apparatus: return "apparatus/system";
                case ClaimCategory.Composition: return "composition";
                case ClaimCategory.ComputerReadableMedium: return "computer-readable medium";
                default: return "unknown";
            }
        }

        static Finding AtStart(Claim c, Severity severity, string message) =>
            Finding.At(c, 0, Math.Min(20, c.FullText.Length), FindingType.Dependency, severity, message);

        static Finding AtReference(Claim c, Severity severity, string message) =>
            c.DependencyRefStart >= 0
                ? Finding.At(c, c.DependencyRefStart, c.DependencyRefLength, FindingType.Dependency, severity, message)
                : AtStart(c, severity, message);
    }
}
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `dotnet test LLDOCX.Claims.Tests`
Expected: all tests pass, 0 failed.

- [ ] **Step 5: Commit**

```bash
git add LLDOCX/LLDOCX.Claims/DependencyChecker.cs LLDOCX/LLDOCX.Claims.Tests/DependencyCheckerTests.cs
git commit -m "feat(claims): check claim dependencies and numbering

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 5: NounPhraseExtractor

**Files:**
- Create: `LLDOCX/LLDOCX.Claims/NounPhraseExtractor.cs`
- Test: `LLDOCX/LLDOCX.Claims.Tests/NounPhraseExtractorTests.cs`

**Interfaces:**
- Consumes: `TextUtil.Tokenize`, `TextUtil.Singularize`, `TextUtil.IsPlural`, `Token` (Task 2)
- Produces:
  - `enum DeterminerKind { Introducing, Referring }`
  - `class NounPhrase`: `Kind`, `Determiner` (lowercase, e.g. `"the at least one"`), `Words` (lowercase list), `Start`, `Length`, `WordsStart`, `Text`, `WordsText`, `IsClaimReference`, `Head`, `HeadKey`, `Key`, `ModifierKeys` (List<string>), `IsPluralForm`
  - `static class NounPhraseExtractor { static List<NounPhrase> Extract(string text); }`

**Heuristics:** The extractor is deliberately simple. A phrase is a determiner, then up to 6 words, ending at the first boundary word or punctuation mark. After the first non-ordinal word, a word ending in `-ing`, `-ed`, `-ly`, or a plural `-s` also ends the phrase, unless it is in `NounExceptions`. Errors from these rules are consistent between the introduction and the later reference ("a plurality of temperature sensors" and "the temperature sensors" both reduce to "temperature"), so they rarely cause false antecedent errors.

- [ ] **Step 1: Write the failing tests `LLDOCX/LLDOCX.Claims.Tests/NounPhraseExtractorTests.cs`**

```csharp
using LLDOCX.Claims;

namespace LLDOCX.Claims.Tests;

[TestClass]
public class NounPhraseExtractorTests
{
    [TestMethod]
    public void Extracts_introductions_and_references()
    {
        var p = NounPhraseExtractor.Extract("a first arm coupled to a base, wherein the first arm rotates relative to the base.");
        Assert.AreEqual(4, p.Count, string.Join(" | ", p.Select(x => x.Text)));
        Assert.AreEqual(DeterminerKind.Introducing, p[0].Kind);
        Assert.AreEqual("first arm", p[0].Key);
        Assert.AreEqual("a first arm", p[0].Text);
        Assert.AreEqual("base", p[1].Key);
        Assert.AreEqual(DeterminerKind.Referring, p[2].Kind);
        Assert.AreEqual("first arm", p[2].Key);
        Assert.AreEqual("base", p[3].Key);
        Assert.AreEqual("arm", p[2].HeadKey);
        CollectionAssert.AreEqual(new[] { "first" }, p[2].ModifierKeys);
    }

    [TestMethod]
    public void Marks_preamble_claim_reference_and_stops_at_verbs()
    {
        var p = NounPhraseExtractor.Extract("The method of claim 1, wherein the valve opens.");
        Assert.AreEqual(2, p.Count, string.Join(" | ", p.Select(x => x.Text)));
        Assert.IsTrue(p[0].IsClaimReference);
        Assert.AreEqual("method", p[0].Key);
        Assert.IsFalse(p[1].IsClaimReference);
        Assert.AreEqual("valve", p[1].Key);
    }

    [TestMethod]
    public void Handles_multiword_determiners()
    {
        var p = NounPhraseExtractor.Extract(
            "a plurality of sensors and at least one controller, wherein each of the sensors and the at least one controller are connected.");
        CollectionAssert.AreEqual(
            new[] { "a plurality of", "at least one", "each of the", "the at least one" },
            p.Select(x => x.Determiner).ToArray());
        CollectionAssert.AreEqual(new[] { "sensor", "controller", "sensor", "controller" }, p.Select(x => x.Key).ToArray());
        Assert.IsTrue(p[0].IsPluralForm);
    }

    [TestMethod]
    public void Keeps_ing_nouns_and_stops_at_reference_numerals()
    {
        var p = NounPhraseExtractor.Extract("an outer housing 12 holding a spring.");
        Assert.AreEqual("outer housing", p[0].Key);
        Assert.AreEqual("an outer housing", p[0].Text);
        Assert.AreEqual("spring", p[1].Key);
    }
}
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `dotnet test LLDOCX.Claims.Tests --filter NounPhraseExtractorTests`
Expected: build FAILS with `The name 'NounPhraseExtractor' does not exist in the current context`.

- [ ] **Step 3: Write `LLDOCX/LLDOCX.Claims/NounPhraseExtractor.cs`**

```csharp
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace LLDOCX.Claims
{
    public enum DeterminerKind { Introducing, Referring }

    public class NounPhrase
    {
        public DeterminerKind Kind { get; set; }
        public string Determiner { get; set; } = "";
        public List<string> Words { get; } = new List<string>();
        /// <summary>Index of the determiner in the claim text.</summary>
        public int Start { get; set; }
        /// <summary>Length from the determiner through the last word.</summary>
        public int Length { get; set; }
        public int WordsStart { get; set; }
        public string Text { get; set; } = "";
        public string WordsText { get; set; } = "";
        /// <summary>True for "the method" in "The method of claim 1": a reference to another claim, not to an element.</summary>
        public bool IsClaimReference { get; set; }

        public string Head => Words[Words.Count - 1];
        public string HeadKey => TextUtil.Singularize(Head);
        public string Key => string.Join(" ", Words.Select(TextUtil.Singularize));
        public List<string> ModifierKeys => Words.Take(Words.Count - 1).Select(TextUtil.Singularize).ToList();
        public bool IsPluralForm => TextUtil.IsPlural(Head);
    }

    public static class NounPhraseExtractor
    {
        const int MaxWords = 6;

        static readonly (string[] Words, DeterminerKind Kind)[] Determiners = BuildDeterminers();

        static readonly HashSet<string> Boundary = new HashSet<string>
        {
            "of", "to", "for", "in", "on", "with", "from", "by", "at", "between", "into", "onto", "through", "via",
            "over", "under", "within", "along", "about", "against", "during", "until", "upon", "after", "before",
            "above", "below", "across", "around", "toward", "towards", "beyond", "near",
            "and", "or", "but", "nor",
            "is", "are", "was", "were", "be", "been", "being", "has", "have", "had", "having", "can", "may", "will",
            "must", "should", "shall", "would", "could", "does", "do", "did",
            "comprising", "comprises", "comprise", "including", "includes", "include", "containing", "contains",
            "contain", "consisting", "consists",
            "wherein", "whereby", "which", "that", "who", "whom", "whose", "when", "where", "while", "whether", "if",
            "than", "then", "thereby", "therein", "thereof", "thereto", "therefrom", "thereon", "such", "so", "as",
            "configured", "adapted", "operable", "arranged", "coupled", "connected", "disposed", "positioned",
            "located", "mounted", "attached", "defined", "provided", "formed",
            "a", "an", "the", "said", "another", "each", "every", "any", "all", "both", "either", "neither", "no",
            "not", "only", "also", "further", "respectively", "claim", "claims"
        };

        static readonly HashSet<string> Ordinals = new HashSet<string>
        {
            "first", "second", "third", "fourth", "fifth", "sixth", "seventh", "eighth", "ninth", "tenth"
        };

        /// <summary>Nouns that look like verbs/adverbs/plurals and must not end a phrase.</summary>
        static readonly HashSet<string> NounExceptions = new HashSet<string>
        {
            "housing", "bearing", "coating", "opening", "spring", "ring", "string", "fitting", "casing", "wiring",
            "lining", "setting", "building", "ceiling", "drawing", "mapping", "padding", "thing", "wing", "awning",
            "tubing", "piping", "molding", "moulding", "flooring", "siding", "packaging", "clothing", "railing",
            "fencing", "lighting",
            "bed", "red", "speed", "feed", "seed", "need", "shed", "reed",
            "assembly", "subassembly", "supply", "family", "anomaly", "reply", "poly",
            "means", "series", "species", "lens", "gas", "chassis", "bias", "canvas", "atlas"
        };

        static readonly Regex ClaimRefAfter = new Regex(
            @"\G\s*(?:,\s*)?(?:of|according\s+to|as\s+(?:claimed|recited|defined|set\s+forth|described)\s+in|in)\s+(?:any\s+(?:one\s+)?of\s+)?(?:the\s+(?:preceding|previous|foregoing)\s+)?claims?\b",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        static (string[] Words, DeterminerKind Kind)[] BuildDeterminers()
        {
            var list = new List<(string[] Words, DeterminerKind Kind)>();
            foreach (var d in new[] { "each of said", "each of the", "the at least one", "the one or more", "the plurality of",
                                      "said at least one", "said one or more", "said plurality of", "the", "said" })
                list.Add((d.Split(' '), DeterminerKind.Referring));
            foreach (var d in new[] { "a plurality of", "at least one", "one or more", "plurality of", "another", "a", "an" })
                list.Add((d.Split(' '), DeterminerKind.Introducing));
            return list.OrderByDescending(x => x.Words.Length).ToArray();
        }

        public static List<NounPhrase> Extract(string text)
        {
            var result = new List<NounPhrase>();
            if (string.IsNullOrEmpty(text)) return result;
            var tokens = TextUtil.Tokenize(text);
            int i = 0;
            while (i < tokens.Count)
            {
                int at = i;
                var det = Determiners.FirstOrDefault(d => Matches(tokens, at, d.Words));
                if (det.Words == null)
                {
                    i++;
                    continue;
                }

                var phrase = new NounPhrase
                {
                    Kind = det.Kind,
                    Determiner = string.Join(" ", det.Words),
                    Start = tokens[i].Index
                };
                int j = i + det.Words.Length;
                int nonOrdinal = 0;
                int lastEnd = -1;
                while (j < tokens.Count && phrase.Words.Count < MaxWords)
                {
                    var tok = tokens[j];
                    if (!tok.IsWord) break;
                    string w = tok.Lower;
                    if (Boundary.Contains(w)) break;
                    if (nonOrdinal > 0 && !NounExceptions.Contains(w) &&
                        (w.EndsWith("ing") || w.EndsWith("ed") || w.EndsWith("ly") || TextUtil.IsPlural(w)))
                        break;
                    if (phrase.Words.Count == 0) phrase.WordsStart = tok.Index;
                    phrase.Words.Add(w);
                    if (!Ordinals.Contains(w)) nonOrdinal++;
                    lastEnd = tok.End;
                    j++;
                }

                if (phrase.Words.Count > 0 && nonOrdinal > 0)
                {
                    phrase.Length = lastEnd - phrase.Start;
                    phrase.Text = text.Substring(phrase.Start, phrase.Length);
                    phrase.WordsText = text.Substring(phrase.WordsStart, lastEnd - phrase.WordsStart);
                    phrase.IsClaimReference = ClaimRefAfter.Match(text, lastEnd).Success;
                    result.Add(phrase);
                    i = j;
                }
                else
                {
                    i += det.Words.Length;
                }
            }
            return result;
        }

        static bool Matches(List<Token> tokens, int i, string[] words)
        {
            if (i + words.Length > tokens.Count) return false;
            for (int k = 0; k < words.Length; k++)
            {
                var t = tokens[i + k];
                if (!t.IsWord || t.Lower != words[k]) return false;
            }
            return true;
        }
    }
}
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `dotnet test LLDOCX.Claims.Tests`
Expected: all tests pass, 0 failed.

- [ ] **Step 5: Commit**

```bash
git add LLDOCX/LLDOCX.Claims/NounPhraseExtractor.cs LLDOCX/LLDOCX.Claims.Tests/NounPhraseExtractorTests.cs
git commit -m "feat(claims): extract introducing and referring noun phrases

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 6: AntecedentChecker

**Files:**
- Create: `LLDOCX/LLDOCX.Claims/AntecedentChecker.cs`
- Test: `LLDOCX/LLDOCX.Claims.Tests/AntecedentCheckerTests.cs`

**Interfaces:**
- Consumes: `Claim`, `Finding.At`, `NounPhraseExtractor.Extract`, `NounPhrase`, `DeterminerKind`, `TextUtil.Levenshtein`, `TextUtil.Singularize` (Tasks 2, 5)
- Produces: `static class AntecedentChecker { static List<Finding> Check(IList<Claim> claims, IEnumerable<string> allowlist, bool reportPlurality = true); }`. Findings have `Type = Antecedent`. The singular/plural drift findings have `Type = Terminology` and are emitted only when `reportPlurality` is true.

- [ ] **Step 1: Write the failing tests `LLDOCX/LLDOCX.Claims.Tests/AntecedentCheckerTests.cs`**

```csharp
using LLDOCX.Claims;

namespace LLDOCX.Claims.Tests;

[TestClass]
public class AntecedentCheckerTests
{
    static readonly string[] Allow = { "the user", "the same", "the group" };

    static List<Finding> Check(params string[] lines) => AntecedentChecker.Check(T.Claims(lines), Allow);

    [TestMethod]
    public void Clean_claims_have_no_findings()
    {
        var f = Check(
            "1. A lamp comprising a base and a bulb mounted on the base.",
            "2. The lamp of claim 1, wherein the bulb is removable.");
        Assert.AreEqual(0, f.Count, T.Dump(f));
    }

    [TestMethod]
    public void Flags_missing_basis()
    {
        var f = Check("1. A lamp comprising a base, wherein the shade is green.");
        Assert.IsTrue(T.Has(f, 1, Severity.Error, "No antecedent basis for 'the shade'"), T.Dump(f));
        Assert.AreEqual("the shade", f.Single().Text);
    }

    [TestMethod]
    public void Basis_comes_only_from_ancestor_claims()
    {
        var f = Check(
            "1. A lamp comprising a base.",
            "2. The lamp of claim 1, further comprising a shade.",
            "3. The lamp of claim 1, wherein the shade is green.");
        Assert.IsTrue(T.Has(f, 3, Severity.Error, "the shade"), T.Dump(f));
        Assert.IsFalse(f.Any(x => x.ClaimNumber == 2), T.Dump(f));
    }

    [TestMethod]
    public void Multiple_dependent_claim_needs_basis_in_every_branch()
    {
        var f = Check(
            "1. A cup comprising a handle.",
            "2. The cup of claim 1, further comprising a lid.",
            "3. The cup of claim 1 or 2, wherein the lid is clear.");
        Assert.IsTrue(T.Has(f, 3, Severity.Error, "the lid"), T.Dump(f));
    }

    [TestMethod]
    public void Flags_ambiguous_reference()
    {
        var f = Check("1. A clamp comprising a first jaw and a second jaw, wherein the jaw is serrated.");
        Assert.IsTrue(T.Has(f, 1, Severity.Error, "ambiguous"), T.Dump(f));
    }

    [TestMethod]
    public void Warns_on_reintroduction()
    {
        var f = Check("1. A lamp comprising a bulb and a socket holding a bulb.");
        Assert.IsTrue(T.Has(f, 1, Severity.Warning, "already has antecedent basis"), T.Dump(f));
    }

    [TestMethod]
    public void Reports_weak_head_noun_match_as_info()
    {
        var f = Check("1. A lamp comprising a housing, wherein the outer housing is black.");
        Assert.IsTrue(T.Has(f, 1, Severity.Info, "partly matches"), T.Dump(f));
        Assert.IsFalse(f.Any(x => x.Severity == Severity.Error), T.Dump(f));
    }

    [TestMethod]
    public void Suggests_close_match_for_typo()
    {
        var f = Check("1. A lamp comprising a reflector, wherein the reflecter is polished.");
        var finding = f.Single(x => x.Severity == Severity.Error);
        StringAssert.Contains(finding.Suggestion, "the reflector");
    }

    [TestMethod]
    public void Allowlisted_phrases_need_no_basis()
    {
        var f = Check("1. A method comprising heating a fluid to the same temperature.");
        Assert.AreEqual(0, f.Count, T.Dump(f));
    }

    [TestMethod]
    public void Flags_singular_plural_drift_as_terminology()
    {
        var f = Check("1. A board comprising a sensor, wherein the sensors are calibrated.");
        var finding = f.Single();
        Assert.AreEqual(FindingType.Terminology, finding.Type);
        StringAssert.Contains(finding.Message, "plural");
    }

    [TestMethod]
    public void Plural_reference_to_plurality_is_fine()
    {
        var f = Check("1. A board comprising a plurality of sensors, wherein the sensors are calibrated.");
        Assert.AreEqual(0, f.Count, T.Dump(f));
    }
}
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `dotnet test LLDOCX.Claims.Tests --filter AntecedentCheckerTests`
Expected: build FAILS with `The name 'AntecedentChecker' does not exist in the current context`.

- [ ] **Step 3: Write `LLDOCX/LLDOCX.Claims/AntecedentChecker.cs`**

```csharp
using System;
using System.Collections.Generic;
using System.Linq;

namespace LLDOCX.Claims
{
    public static class AntecedentChecker
    {
        /// <summary>Generic nouns that are routinely introduced more than once ("a portion of X ... a portion of Y").</summary>
        static readonly HashSet<string> GenericNouns = new HashSet<string>
        {
            "portion", "part", "end", "side", "surface", "number", "amount", "time", "set", "group", "pair",
            "value", "level", "degree", "piece", "section", "region", "area", "edge", "point", "step"
        };

        static readonly HashSet<string> Units = new HashSet<string>
        {
            "mm", "cm", "m", "um", "nm", "kg", "g", "mg", "ml", "l", "degree", "percent", "second", "minute",
            "hour", "hz", "khz", "mhz", "ghz", "volt", "v", "watt", "w", "celsius", "fahrenheit"
        };

        static readonly HashSet<string> PluralCapable = new HashSet<string>
        {
            "a plurality of", "plurality of", "one or more", "at least one"
        };

        public static List<Finding> Check(IList<Claim> claims, IEnumerable<string> allowlist, bool reportPlurality = true)
        {
            var findings = new List<Finding>();
            var allow = (allowlist ?? Enumerable.Empty<string>()).Select(NormalizeAllowEntry).Where(a => a.Length > 0).ToList();

            var byNumber = new Dictionary<int, Claim>();
            foreach (var c in claims)
                if (!c.IsCanceled && !byNumber.ContainsKey(c.Number)) byNumber[c.Number] = c;
            var phrases = byNumber.Values.ToDictionary(
                c => c.Number,
                c => NounPhraseExtractor.Extract(c.FullText).Where(p => !p.IsClaimReference).ToList());
            var context = new Context(byNumber, phrases);

            foreach (var claim in byNumber.Values.OrderBy(c => c.Number))
            {
                var known = new List<NounPhrase>(context.Inherited(claim));
                foreach (var p in phrases[claim.Number])
                {
                    if (p.Kind == DeterminerKind.Introducing)
                    {
                        if (p.Determiner != "another" && !GenericNouns.Contains(p.HeadKey) && known.Any(k => k.Key == p.Key))
                            findings.Add(Finding.At(claim, p.Start, p.Length, FindingType.Antecedent, Severity.Warning,
                                $"'{p.Text}' introduces a term that already has antecedent basis.",
                                $"If the same element is meant, use 'the {p.WordsText}'."));
                        known.Add(p);
                    }
                    else
                    {
                        CheckReference(claim, p, known, allow, reportPlurality, findings);
                    }
                }
            }
            return findings;
        }

        static void CheckReference(Claim claim, NounPhrase p, List<NounPhrase> known, List<string> allow,
            bool reportPlurality, List<Finding> findings)
        {
            if (IsAllowed(p, allow)) return;

            var exact = known.FirstOrDefault(k => k.Key == p.Key);
            if (exact != null)
            {
                if (reportPlurality) CheckPlurality(claim, p, exact, findings);
                return;
            }

            var mods = p.ModifierKeys;
            var candidates = known
                .Where(k => k.HeadKey == p.HeadKey && mods.All(m => k.ModifierKeys.Contains(m)))
                .GroupBy(k => k.Key).Select(g => g.First()).ToList();
            if (candidates.Count == 1)
            {
                if (reportPlurality) CheckPlurality(claim, p, candidates[0], findings);
                return;
            }
            if (candidates.Count > 1)
            {
                string options = string.Join(" or ", candidates.Select(k => $"'{k.Text}'"));
                findings.Add(Finding.At(claim, p.Start, p.Length, FindingType.Antecedent, Severity.Error,
                    $"'{p.Text}' is ambiguous: it could refer to {options}.",
                    $"Say which one, e.g. 'the {candidates[0].WordsText}'."));
                return;
            }

            var weak = known.FirstOrDefault(k => k.HeadKey == p.HeadKey);
            if (weak != null)
            {
                findings.Add(Finding.At(claim, p.Start, p.Length, FindingType.Antecedent, Severity.Info,
                    $"'{p.Text}' only partly matches '{weak.Text}' (same noun, different modifiers). Check its antecedent basis."));
                return;
            }

            var close = known
                .Where(k => k.Key.Length >= 5 && TextUtil.Levenshtein(k.Key, p.Key) <= 2)
                .OrderBy(k => TextUtil.Levenshtein(k.Key, p.Key))
                .FirstOrDefault();
            string suggestion = close != null
                ? $"Did you mean 'the {close.WordsText}'?"
                : $"Introduce it first (e.g. 'a {p.WordsText}') or correct the term.";
            findings.Add(Finding.At(claim, p.Start, p.Length, FindingType.Antecedent, Severity.Error,
                $"No antecedent basis for '{p.Text}'.", suggestion));
        }

        static void CheckPlurality(Claim claim, NounPhrase p, NounPhrase basis, List<Finding> findings)
        {
            if (PluralCapable.Contains(basis.Determiner) || p.Determiner.StartsWith("each")) return;
            if (p.IsPluralForm == basis.IsPluralForm) return;
            string pForm = p.IsPluralForm ? "plural" : "singular";
            string bForm = basis.IsPluralForm ? "plural" : "singular";
            findings.Add(Finding.At(claim, p.Start, p.Length, FindingType.Terminology, Severity.Warning,
                $"'{p.Text}' is {pForm} but its antecedent '{basis.Text}' is {bForm}.",
                $"Use 'the {basis.WordsText}'."));
        }

        static bool IsAllowed(NounPhrase p, List<string> allow)
        {
            if (Units.Contains(p.HeadKey)) return true;
            string key = p.Key;
            return allow.Any(a => key == a || key.StartsWith(a + " "));
        }

        static string NormalizeAllowEntry(string entry)
        {
            var words = (entry ?? "").ToLowerInvariant()
                .Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries).ToList();
            if (words.Count > 0 && (words[0] == "the" || words[0] == "said")) words.RemoveAt(0);
            return string.Join(" ", words.Select(TextUtil.Singularize));
        }

        sealed class Context
        {
            readonly Dictionary<int, Claim> _byNumber;
            readonly Dictionary<int, List<NounPhrase>> _phrases;
            readonly Dictionary<int, List<NounPhrase>> _closure = new Dictionary<int, List<NounPhrase>>();
            readonly HashSet<int> _visiting = new HashSet<int>();

            public Context(Dictionary<int, Claim> byNumber, Dictionary<int, List<NounPhrase>> phrases)
            {
                _byNumber = byNumber;
                _phrases = phrases;
            }

            /// <summary>Terms introduced by the claim's ancestors; for multiple dependency, only terms every branch provides.</summary>
            public List<NounPhrase> Inherited(Claim claim)
            {
                var parents = claim.DependsOn.Where(d => d < claim.Number && _byNumber.ContainsKey(d)).Distinct().ToList();
                if (parents.Count == 0) return new List<NounPhrase>();
                var sets = parents.Select(Closure).ToList();
                return sets[0].Where(p => sets.Skip(1).All(s => s.Any(q => q.Key == p.Key))).ToList();
            }

            List<NounPhrase> Closure(int number)
            {
                List<NounPhrase> cached;
                if (_closure.TryGetValue(number, out cached)) return cached;
                if (!_visiting.Add(number)) return new List<NounPhrase>();
                var result = Inherited(_byNumber[number]);
                result.AddRange(_phrases[number].Where(p => p.Kind == DeterminerKind.Introducing));
                _visiting.Remove(number);
                _closure[number] = result;
                return result;
            }
        }
    }
}
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `dotnet test LLDOCX.Claims.Tests`
Expected: all tests pass, 0 failed.

- [ ] **Step 5: Commit**

```bash
git add LLDOCX/LLDOCX.Claims/AntecedentChecker.cs LLDOCX/LLDOCX.Claims.Tests/AntecedentCheckerTests.cs
git commit -m "feat(claims): check antecedent basis across dependency chains

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 7: TerminologyChecker

**Files:**
- Create: `LLDOCX/LLDOCX.Claims/TerminologyChecker.cs`
- Test: `LLDOCX/LLDOCX.Claims.Tests/TerminologyCheckerTests.cs`

**Interfaces:**
- Consumes: `Claim`, `Finding.At`, `NounPhraseExtractor`, `TextUtil.Levenshtein` (Tasks 2, 5)
- Produces: `static class TerminologyChecker { static List<Finding> Check(IList<Claim> claims); }`. Spelling variants and the said/the mix are `Type = Terminology`. Repeated words and punctuation are `Type = Formatting`.

- [ ] **Step 1: Write the failing tests `LLDOCX/LLDOCX.Claims.Tests/TerminologyCheckerTests.cs`**

```csharp
using LLDOCX.Claims;

namespace LLDOCX.Claims.Tests;

[TestClass]
public class TerminologyCheckerTests
{
    [TestMethod]
    public void Flags_rare_spelling_variant_of_frequent_word()
    {
        var f = TerminologyChecker.Check(T.Claims(
            "1. A frame comprising a bracket.",
            "2. The frame of claim 1, wherein the bracket is steel.",
            "3. The frame of claim 1, wherein the bracket is welded.",
            "4. The frame of claim 1, wherein the braket is painted."));
        Assert.AreEqual(1, f.Count, T.Dump(f));
        Assert.IsTrue(T.Has(f, 4, Severity.Warning, "'bracket'"), T.Dump(f));
        Assert.AreEqual("braket", f[0].Text);
    }

    [TestMethod]
    public void Ignores_inflections()
    {
        var f = TerminologyChecker.Check(T.Claims(
            "1. A method comprising: connecting a wire.",
            "2. The method of claim 1, comprising connecting a cable.",
            "3. The method of claim 1, comprising connecting a pipe.",
            "4. The method of claim 1, wherein the wire comprises copper."));
        Assert.IsFalse(f.Any(x => x.Type == FindingType.Terminology), T.Dump(f));
    }

    [TestMethod]
    public void Notes_mixed_said_and_the()
    {
        var f = TerminologyChecker.Check(T.Claims(
            "1. A cup comprising a handle.",
            "2. The cup of claim 1, wherein said handle is round.",
            "3. The cup of claim 1, wherein the handle is square."));
        Assert.IsTrue(T.Has(f, 2, Severity.Info, "both 'said'"), T.Dump(f));
    }

    [TestMethod]
    public void Flags_formatting_problems()
    {
        var f = TerminologyChecker.Check(T.Claims(
            "1. A cup comprising a a handle",
            "2. The cup of claim 1. wherein the handle is round."));
        string dump = T.Dump(f);
        Assert.IsTrue(T.Has(f, 1, Severity.Error, "Repeated word"), dump);
        Assert.IsTrue(T.Has(f, 1, Severity.Warning, "does not end with a period"), dump);
        Assert.IsTrue(T.Has(f, 2, Severity.Warning, "period before its end"), dump);
    }

    [TestMethod]
    public void Flags_unbalanced_parentheses()
    {
        var f = TerminologyChecker.Check(T.Claims("1. A cup (10 comprising a handle."));
        Assert.IsTrue(T.Has(f, 1, Severity.Warning, "Unbalanced"), T.Dump(f));
    }

    [TestMethod]
    public void Abbreviation_periods_are_not_flagged()
    {
        var f = TerminologyChecker.Check(T.Claims("1. A cup comprising a handle, e.g. a loop, as in FIG. 2."));
        Assert.IsFalse(f.Any(x => x.Message.Contains("period before")), T.Dump(f));
    }
}
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `dotnet test LLDOCX.Claims.Tests --filter TerminologyCheckerTests`
Expected: build FAILS with `The name 'TerminologyChecker' does not exist in the current context`.

- [ ] **Step 3: Write `LLDOCX/LLDOCX.Claims/TerminologyChecker.cs`**

```csharp
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace LLDOCX.Claims
{
    public static class TerminologyChecker
    {
        static readonly Regex WordRx = new Regex(@"[A-Za-z][A-Za-z\-]*", RegexOptions.Compiled);
        static readonly Regex RepeatedRx = new Regex(@"\b([A-Za-z]+)\s+\1\b", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        static readonly Regex InnerPeriodRx = new Regex(
            @"(?<!\b(?:e\.g|i\.e|etc|approx|no|fig|figs|wt|vol|ca|vs|cf|al|inc|ltd|co))\.(?=\s+\S)",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);
        static readonly string[] InflectionSuffixes = { "ing", "ed", "es", "s", "er" };

        public static List<Finding> Check(IList<Claim> claims)
        {
            var active = claims.Where(c => !c.IsCanceled && c.FullText.Length > 0).ToList();
            var findings = new List<Finding>();
            CheckSpellingVariants(active, findings);
            CheckSaidTheMix(active, findings);
            foreach (var c in active) CheckFormatting(c, findings);
            return findings;
        }

        static void CheckSpellingVariants(List<Claim> claims, List<Finding> findings)
        {
            var occurrences = new Dictionary<string, List<KeyValuePair<Claim, Match>>>();
            foreach (var c in claims)
            {
                foreach (Match m in WordRx.Matches(c.FullText))
                {
                    string w = m.Value.ToLowerInvariant();
                    if (w.Length < 5) continue;
                    List<KeyValuePair<Claim, Match>> list;
                    if (!occurrences.TryGetValue(w, out list))
                        occurrences[w] = list = new List<KeyValuePair<Claim, Match>>();
                    list.Add(new KeyValuePair<Claim, Match>(c, m));
                }
            }

            foreach (var rare in occurrences)
            {
                string best = null;
                int bestCount = 0;
                foreach (var common in occurrences)
                {
                    int count = common.Value.Count;
                    if (common.Key == rare.Key || count < 3 * rare.Value.Count || count <= bestCount) continue;
                    if (IsInflection(rare.Key, common.Key)) continue;
                    int maxDistance = Math.Max(rare.Key.Length, common.Key.Length) >= 8 ? 2 : 1;
                    if (Math.Abs(rare.Key.Length - common.Key.Length) > maxDistance) continue;
                    if (TextUtil.Levenshtein(rare.Key, common.Key) > maxDistance) continue;
                    best = common.Key;
                    bestCount = count;
                }
                if (best == null) continue;
                foreach (var occ in rare.Value)
                    findings.Add(Finding.At(occ.Key, occ.Value.Index, occ.Value.Length, FindingType.Terminology, Severity.Warning,
                        $"'{occ.Value.Value}' may be a typo or variant of '{best}' (used {bestCount} times).",
                        $"Use '{best}' consistently."));
            }
        }

        static bool IsInflection(string a, string b) => a.StartsWith(b) || b.StartsWith(a) || Stem(a) == Stem(b);

        static string Stem(string word)
        {
            foreach (var s in InflectionSuffixes)
                if (word.EndsWith(s) && word.Length - s.Length >= 3) return word.Substring(0, word.Length - s.Length);
            return word;
        }

        static void CheckSaidTheMix(List<Claim> claims, List<Finding> findings)
        {
            var refs = new List<KeyValuePair<Claim, NounPhrase>>();
            foreach (var c in claims)
                foreach (var p in NounPhraseExtractor.Extract(c.FullText))
                    if (p.Kind == DeterminerKind.Referring && !p.IsClaimReference)
                        refs.Add(new KeyValuePair<Claim, NounPhrase>(c, p));

            int said = refs.Count(r => r.Value.Determiner.Contains("said"));
            int the = refs.Count - said;
            if (said == 0 || the == 0) return;
            bool saidIsMinority = said <= the;
            var first = refs.First(r => r.Value.Determiner.Contains("said") == saidIsMinority);
            findings.Add(Finding.At(first.Key, first.Value.Start, first.Value.Length, FindingType.Terminology, Severity.Info,
                $"The claims use both 'said' ({said}x) and 'the' ({the}x) to refer back to elements.",
                "Consider using one consistently."));
        }

        static void CheckFormatting(Claim c, List<Finding> findings)
        {
            string text = c.FullText;
            foreach (Match m in RepeatedRx.Matches(text))
                findings.Add(Finding.At(c, m.Index, m.Length, FindingType.Formatting, Severity.Error,
                    $"Repeated word '{m.Groups[1].Value}'.", $"Delete the extra '{m.Groups[1].Value}'."));

            string trimmed = text.TrimEnd();
            if (!trimmed.EndsWith("."))
                findings.Add(Finding.At(c, Math.Max(0, trimmed.Length - 1), 1, FindingType.Formatting, Severity.Warning,
                    $"Claim {c.Number} does not end with a period."));
            else if (trimmed.EndsWith(".."))
                findings.Add(Finding.At(c, trimmed.Length - 2, 2, FindingType.Formatting, Severity.Warning,
                    $"Claim {c.Number} ends with more than one period."));

            foreach (Match m in InnerPeriodRx.Matches(text))
                findings.Add(Finding.At(c, m.Index, 1, FindingType.Formatting, Severity.Warning,
                    $"Claim {c.Number} contains a period before its end. A claim must be a single sentence."));

            if (text.Count(ch => ch == '(') != text.Count(ch => ch == ')') ||
                text.Count(ch => ch == '[') != text.Count(ch => ch == ']'))
                findings.Add(Finding.At(c, 0, Math.Min(20, text.Length), FindingType.Formatting, Severity.Warning,
                    $"Unbalanced parentheses or brackets in claim {c.Number}."));
        }
    }
}
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `dotnet test LLDOCX.Claims.Tests`
Expected: all tests pass, 0 failed.

- [ ] **Step 5: Commit**

```bash
git add LLDOCX/LLDOCX.Claims/TerminologyChecker.cs LLDOCX/LLDOCX.Claims.Tests/TerminologyCheckerTests.cs
git commit -m "feat(claims): check terminology variants and claim formatting

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 8: ClaimsAnalyzer facade and false-positive guard

**Files:**
- Create: `LLDOCX/LLDOCX.Claims/ClaimsAnalyzer.cs`
- Test: `LLDOCX/LLDOCX.Claims.Tests/ClaimsAnalyzerTests.cs`
- Create: `LLDOCX/LLDOCX.Claims.Tests/Fixtures/real1.txt`, `real2.txt`, `real3.txt`

**Interfaces:**
- Consumes: `ClaimParser`, `DependencyChecker`, `AntecedentChecker`, `TerminologyChecker` (Tasks 3–7)
- Produces:
  - `class ClaimsCheckOptions { bool CheckAntecedent = true; bool CheckDependency = true; bool CheckTerminology = true; Severity MinSeverity = Info; List<string> Allowlist; static readonly string[] DefaultAllowlist; }`
  - `class AnalysisResult { List<Claim> Claims; List<Finding> Findings; }`
  - `static class ClaimsAnalyzer { static AnalysisResult Analyze(IList<RawParagraph> paragraphs, ClaimsCheckOptions options); }`. Findings are filtered by `MinSeverity` and sorted by claim number, then by document offset.

- [ ] **Step 1: Write the failing tests `LLDOCX/LLDOCX.Claims.Tests/ClaimsAnalyzerTests.cs`**

```csharp
using LLDOCX.Claims;

namespace LLDOCX.Claims.Tests;

[TestClass]
public class ClaimsAnalyzerTests
{
    /// <summary>A realistic, correctly drafted claim set. Must produce no errors or warnings.</summary>
    static readonly string[] CleanClaimSet =
    {
        "What is claimed is:",
        "1. A fluid dispensing apparatus, comprising:",
        "a housing defining a reservoir;",
        "a pump disposed within the housing and in fluid communication with the reservoir;",
        "a nozzle coupled to the pump; and",
        "a controller configured to actuate the pump in response to a signal from a sensor.",
        "2. The fluid dispensing apparatus of claim 1, wherein the sensor is an optical sensor.",
        "3. The fluid dispensing apparatus of claim 1, wherein the nozzle comprises a first outlet and a second outlet.",
        "4. The fluid dispensing apparatus of claim 3, wherein the first outlet has a diameter smaller than a diameter of the second outlet.",
        "5. A method of dispensing a fluid, the method comprising:",
        "receiving a signal from a sensor;",
        "actuating a pump in response to the signal; and",
        "dispensing the fluid through a nozzle.",
        "6. The method of claim 5, further comprising stopping the pump after a predetermined period."
    };

    [TestMethod]
    public void Clean_claim_set_produces_no_errors_or_warnings()
    {
        var result = ClaimsAnalyzer.Analyze(T.Paras(CleanClaimSet), new ClaimsCheckOptions());
        Assert.AreEqual(6, result.Claims.Count);
        var serious = result.Findings.Where(f => f.Severity != Severity.Info).ToList();
        Assert.AreEqual(0, serious.Count, T.Dump(serious));
    }

    [TestMethod]
    public void Applies_minimum_severity_filter()
    {
        var paras = T.Paras("1. A lamp comprising a housing, wherein the outer housing is black.");
        Assert.AreEqual(1, ClaimsAnalyzer.Analyze(paras, new ClaimsCheckOptions()).Findings.Count);
        var filtered = ClaimsAnalyzer.Analyze(paras, new ClaimsCheckOptions { MinSeverity = Severity.Warning });
        Assert.AreEqual(0, filtered.Findings.Count, T.Dump(filtered.Findings));
    }

    [TestMethod]
    public void Respects_check_toggles()
    {
        var paras = T.Paras("1. A lamp comprising a base, wherein the shade is green.");
        var on = ClaimsAnalyzer.Analyze(paras, new ClaimsCheckOptions());
        Assert.IsTrue(on.Findings.Any(f => f.Type == FindingType.Antecedent));
        var off = ClaimsAnalyzer.Analyze(paras, new ClaimsCheckOptions { CheckAntecedent = false });
        Assert.IsFalse(off.Findings.Any(f => f.Type == FindingType.Antecedent), T.Dump(off.Findings));
    }

    [TestMethod]
    public void Sorts_findings_by_claim_then_position()
    {
        var result = ClaimsAnalyzer.Analyze(T.Paras(
            "1. A lamp comprising a base, wherein the shade and the cord are green.",
            "2. The lamp of claim 1, wherein the plug is black."), new ClaimsCheckOptions());
        var numbers = result.Findings.Select(f => f.ClaimNumber).ToList();
        CollectionAssert.AreEqual(numbers.OrderBy(n => n).ToList(), numbers);
        var claim1 = result.Findings.Where(f => f.ClaimNumber == 1).ToList();
        Assert.IsTrue(claim1[0].DocStart < claim1[1].DocStart);
    }
}
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `dotnet test LLDOCX.Claims.Tests --filter ClaimsAnalyzerTests`
Expected: build FAILS with `The name 'ClaimsAnalyzer' does not exist in the current context`.

- [ ] **Step 3: Write `LLDOCX/LLDOCX.Claims/ClaimsAnalyzer.cs`**

```csharp
using System.Collections.Generic;
using System.Linq;

namespace LLDOCX.Claims
{
    public class ClaimsCheckOptions
    {
        public static readonly string[] DefaultAllowlist =
        {
            "the user", "the art", "the same", "the following", "the foregoing", "the group",
            "the presence", "the absence", "the range", "the order"
        };

        public bool CheckAntecedent { get; set; } = true;
        public bool CheckDependency { get; set; } = true;
        public bool CheckTerminology { get; set; } = true;
        public Severity MinSeverity { get; set; } = Severity.Info;
        public List<string> Allowlist { get; set; } = new List<string>(DefaultAllowlist);
    }

    public class AnalysisResult
    {
        public AnalysisResult(List<Claim> claims, List<Finding> findings)
        {
            Claims = claims;
            Findings = findings;
        }

        public List<Claim> Claims { get; }
        public List<Finding> Findings { get; }
    }

    public static class ClaimsAnalyzer
    {
        public static AnalysisResult Analyze(IList<RawParagraph> paragraphs, ClaimsCheckOptions options)
        {
            options = options ?? new ClaimsCheckOptions();
            var claims = ClaimParser.Parse(paragraphs);
            var findings = new List<Finding>();
            if (options.CheckDependency) findings.AddRange(DependencyChecker.Check(claims));
            if (options.CheckAntecedent) findings.AddRange(AntecedentChecker.Check(claims, options.Allowlist, options.CheckTerminology));
            if (options.CheckTerminology) findings.AddRange(TerminologyChecker.Check(claims));
            var visible = findings
                .Where(f => f.Severity <= options.MinSeverity)
                .OrderBy(f => f.ClaimNumber)
                .ThenBy(f => f.DocStart)
                .ToList();
            return new AnalysisResult(claims, visible);
        }
    }
}
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `dotnet test LLDOCX.Claims.Tests`
Expected: all tests pass, 0 failed.

If `Clean_claim_set_produces_no_errors_or_warnings` fails, the dump lists each false positive. Fix the rule that caused it, usually by adding a word to `Boundary` or `NounExceptions` in `NounPhraseExtractor.cs`. Do not weaken the test.

- [ ] **Step 5: Add three real claim sets as additional false-positive guards**

Copy the full claim text of three issued US patents in different fields (one mechanical, one software/method, one chemical composition) from Google Patents. Save them as `LLDOCX/LLDOCX.Claims.Tests/Fixtures/real1.txt`, `real2.txt`, and `real3.txt`, with one claim paragraph per line in the form `1. A ...`. Add to the test csproj:

```xml
  <ItemGroup>
    <None Update="Fixtures\*.txt" CopyToOutputDirectory="PreserveNewest" />
  </ItemGroup>
```

Add this test to `ClaimsAnalyzerTests`:

```csharp
    [DataTestMethod]
    [DataRow("real1.txt")]
    [DataRow("real2.txt")]
    [DataRow("real3.txt")]
    public void Real_issued_claims_produce_no_errors(string file)
    {
        var lines = File.ReadAllLines(Path.Combine(AppContext.BaseDirectory, "Fixtures", file));
        var result = ClaimsAnalyzer.Analyze(T.Paras(lines), new ClaimsCheckOptions());
        var errors = result.Findings.Where(f => f.Severity == Severity.Error).ToList();
        Assert.AreEqual(0, errors.Count, T.Dump(errors));
    }
```

Run: `dotnet test LLDOCX.Claims.Tests`
Expected: all pass. For each Error the dump shows, decide whether it is:
- **a real drafting problem in the issued patent:** acceptable, but replace that fixture with a different patent, since fixtures must be clean;
- **a heuristic false positive:** fix it in the extractor or checker word lists, add a focused unit test that reproduces it in the relevant `*Tests.cs`, and re-run.

- [ ] **Step 6: Commit**

```bash
git add LLDOCX/LLDOCX.Claims/ClaimsAnalyzer.cs LLDOCX/LLDOCX.Claims.Tests
git commit -m "feat(claims): add analyzer facade with options and false-positive guards

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 9: Deep Check contract and gateway client (LLDOCX.DeepCheck)

**Files:**
- Create: `LLDOCX/LLDOCX.DeepCheck/LLDOCX.DeepCheck.csproj`
- Create: `LLDOCX/LLDOCX.DeepCheck/DeepCheckException.cs`
- Create: `LLDOCX/LLDOCX.DeepCheck/DeepCheckContract.cs`
- Create: `LLDOCX/LLDOCX.DeepCheck/DeepCheckRequestBuilder.cs`
- Create: `LLDOCX/LLDOCX.DeepCheck/DeepCheckResponseParser.cs`
- Create: `LLDOCX/LLDOCX.DeepCheck/DeepCheckHttpClient.cs`
- Modify: `LLDOCX/LLDOCX.Claims.Tests/LLDOCX.Claims.Tests.csproj` (add project reference)
- Test: `LLDOCX/LLDOCX.Claims.Tests/DeepCheckProtocolTests.cs`

**Interfaces:**
- Consumes: `Claim`, `Finding`, `Finding.At`, `FindingType`, `Severity` (Task 2); `T.Claims` (Task 3)
- Produces (namespace `LLDOCX.DeepCheck`; used by the gateway in Task 10 and by the add-in in Task 12):
  - `class DeepCheckException : Exception` (its message is user-facing)
  - The wire DTOs:
    - `DeepCheckRequest { List<ClaimDto> Claims; List<RuleFindingDto> RuleFindings; }`
    - `ClaimDto { int Number; string Text; bool Canceled; }`
    - `RuleFindingDto { int Claim; string Type; string Severity; string Message; }`
    - `ErrorBody { ErrorDetail Error; }` and `ErrorDetail { string Code; string Message; }`
    - JSON names are camelCase: `claims`, `ruleFindings`, `number`, `text`, `canceled`, `claim`, `type`, `severity`, `message`, `error`, `code`.
  - `static class ErrorCodes` (string constants)
  - `static class DeepCheckRequestBuilder`:
    - `DeepCheckRequest Build(IList<Claim> claims, IList<Finding> ruleFindings)`;
    - `int TextLength(DeepCheckRequest request)`.
  - `static class DeepCheckResponseParser`:
    - `List<Finding> Parse(string json, IList<Claim> claims, IList<Finding> ruleFindings)` returns new AI findings (`Source = "AI"`, `DocStart = -1` when the quote isn't found), and sets `DisputedByAI`/`DisputeReason` on disputed `ruleFindings` in place;
    - `bool IsWellFormed(string json)`.
  - The response body is `{ findings: [{ claim, type, severity, quote, message, suggestion }], disputes: [{ index, reason }] }`.
  - `sealed class DeepCheckHttpClient(HttpClient http)`:
    - `static readonly TimeSpan DefaultTimeout` (240 s);
    - `Task<List<Finding>> RunAsync(string gatewayUrl, string bearerToken, IList<Claim> claims, IList<Finding> ruleFindings, CancellationToken ct)`, which throws `DeepCheckException` with the spec's user messages.

- [ ] **Step 1: Create `LLDOCX/LLDOCX.DeepCheck/LLDOCX.DeepCheck.csproj` and wire up references**

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFramework>netstandard2.0</TargetFramework>
    <LangVersion>7.3</LangVersion>
    <RootNamespace>LLDOCX.DeepCheck</RootNamespace>
  </PropertyGroup>
  <ItemGroup>
    <ProjectReference Include="..\LLDOCX.Claims\LLDOCX.Claims.csproj" />
  </ItemGroup>
</Project>
```

Run:

```bash
dotnet add LLDOCX.DeepCheck package System.Text.Json
dotnet sln LLDOCX.sln add LLDOCX.DeepCheck/LLDOCX.DeepCheck.csproj
```

Expected: `PackageReference for package 'System.Text.Json' ... added` and `Project ... added to the solution.`

Add to the test csproj's project-reference `ItemGroup`:

```xml
    <ProjectReference Include="..\LLDOCX.DeepCheck\LLDOCX.DeepCheck.csproj" />
```

- [ ] **Step 2: Write the failing tests `LLDOCX/LLDOCX.Claims.Tests/DeepCheckProtocolTests.cs`**

```csharp
using System.Net;
using System.Text.Json;
using LLDOCX.Claims;
using LLDOCX.DeepCheck;

namespace LLDOCX.Claims.Tests;

internal sealed class StubHandler : HttpMessageHandler
{
    public Func<HttpRequestMessage, HttpResponseMessage> Respond { get; set; }
    public HttpRequestMessage LastRequest { get; private set; }
    public string LastBody { get; private set; }

    protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken ct)
    {
        LastRequest = request;
        LastBody = request.Content == null ? null : await request.Content.ReadAsStringAsync(ct);
        return Respond(request);
    }
}

[TestClass]
public class DeepCheckProtocolTests
{
    const string Line = "1. A lamp comprising a base, wherein the shade is green.";
    const string Url = "https://gw.example.net/api/deepcheck";
    const string OkJson = """{"findings":[{"claim":1,"type":"Terminology","severity":"Warning","quote":"the shade","message":"m","suggestion":"s"}],"disputes":[{"index":0,"reason":"r"}]}""";

    static HttpResponseMessage Json(HttpStatusCode status, string body) =>
        new(status) { Content = new StringContent(body, System.Text.Encoding.UTF8, "application/json") };

    // ---- request ----

    [TestMethod]
    public void Request_lists_claims_canceled_claims_and_rule_findings_in_order()
    {
        var claims = T.Claims(Line, "2. (Canceled)");
        var rule = new List<Finding> { new Finding { ClaimNumber = 1, Type = FindingType.Antecedent, Severity = Severity.Error, Message = "No antecedent basis for 'the shade'." } };
        var r = DeepCheckRequestBuilder.Build(claims, rule);
        Assert.AreEqual(2, r.Claims.Count);
        StringAssert.StartsWith(r.Claims[0].Text, "A lamp comprising a base");
        Assert.IsTrue(r.Claims[1].Canceled);
        Assert.AreEqual("", r.Claims[1].Text);
        Assert.AreEqual("Antecedent", r.RuleFindings[0].Type);
        Assert.AreEqual("Error", r.RuleFindings[0].Severity);
        Assert.AreEqual(r.Claims[0].Text.Length, DeepCheckRequestBuilder.TextLength(r));
    }

    [TestMethod]
    public void Request_serializes_with_camel_case_names()
    {
        string json = JsonSerializer.Serialize(DeepCheckRequestBuilder.Build(T.Claims(Line), new List<Finding>()));
        StringAssert.Contains(json, "\"claims\":[{\"number\":1,");
        StringAssert.Contains(json, "\"ruleFindings\":[]");
    }

    // ---- response ----

    [TestMethod]
    public void Parser_maps_quotes_to_document_offsets_and_applies_disputes()
    {
        var claims = T.Claims(Line);
        var rule = new List<Finding> { new Finding { ClaimNumber = 1, Message = "x" } };
        var added = DeepCheckResponseParser.Parse(OkJson, claims, rule);
        Assert.AreEqual(1, added.Count);
        Assert.AreEqual(Line.IndexOf("the shade"), added[0].DocStart);
        Assert.AreEqual("the shade", added[0].Text);
        Assert.AreEqual("AI", added[0].Source);
        Assert.AreEqual(FindingType.Terminology, added[0].Type);
        Assert.AreEqual(Severity.Warning, added[0].Severity);
        Assert.IsTrue(rule[0].DisputedByAI);
        Assert.AreEqual("r", rule[0].DisputeReason);
    }

    [TestMethod]
    public void Parser_keeps_findings_whose_quote_is_not_found_without_location()
    {
        const string json = """{"findings":[{"claim":1,"type":"Antecedent","severity":"Error","quote":"not there","message":"m","suggestion":""}],"disputes":[]}""";
        var added = DeepCheckResponseParser.Parse(json, T.Claims(Line), new List<Finding>());
        Assert.AreEqual(-1, added[0].DocStart);
        Assert.IsNull(added[0].Suggestion);
    }

    [TestMethod]
    public void Parser_ignores_out_of_range_dispute_indexes()
    {
        var rule = new List<Finding> { new Finding() };
        DeepCheckResponseParser.Parse("""{"findings":[],"disputes":[{"index":5,"reason":"r"}]}""", T.Claims(Line), rule);
        Assert.IsFalse(rule[0].DisputedByAI);
    }

    [TestMethod]
    public void Parser_rejects_invalid_json()
    {
        Assert.ThrowsException<DeepCheckException>(() =>
            DeepCheckResponseParser.Parse("not json", T.Claims(Line), new List<Finding>()));
    }

    [DataTestMethod]
    [DataRow(OkJson, true)]
    [DataRow("""{"findings":[],"disputes":[]}""", true)]
    [DataRow("""{"findings":[]}""", false)]
    [DataRow("""{"findings":[{"claim":"1","type":"Antecedent","severity":"Error","quote":"q","message":"m","suggestion":""}],"disputes":[]}""", false)]
    [DataRow("""{"findings":[],"disputes":[{"index":0}]}""", false)]
    [DataRow("[]", false)]
    [DataRow("not json", false)]
    public void IsWellFormed_checks_the_response_shape(string json, bool expected) =>
        Assert.AreEqual(expected, DeepCheckResponseParser.IsWellFormed(json));

    // ---- HTTP client ----

    static (DeepCheckHttpClient client, StubHandler handler) NewClient(Func<HttpRequestMessage, HttpResponseMessage> respond)
    {
        var handler = new StubHandler { Respond = respond };
        return (new DeepCheckHttpClient(new HttpClient(handler)), handler);
    }

    [TestMethod]
    public async Task Client_posts_json_with_bearer_token_and_parses_findings()
    {
        var (client, handler) = NewClient(_ => Json(HttpStatusCode.OK, OkJson));
        var rule = new List<Finding> { new Finding { ClaimNumber = 1, Message = "x" } };
        var added = await client.RunAsync(Url, "tok123", T.Claims(Line), rule, CancellationToken.None);
        Assert.AreEqual(1, added.Count);
        Assert.AreEqual(HttpMethod.Post, handler.LastRequest.Method);
        Assert.AreEqual(Url, handler.LastRequest.RequestUri.ToString());
        Assert.AreEqual("Bearer", handler.LastRequest.Headers.Authorization.Scheme);
        Assert.AreEqual("tok123", handler.LastRequest.Headers.Authorization.Parameter);
        StringAssert.Contains(handler.LastBody, "\"ruleFindings\":[{\"claim\":1,");
    }

    static async Task<string> FailureMessage(HttpResponseMessage response)
    {
        var (client, _) = NewClient(_ => response);
        var ex = await Assert.ThrowsExceptionAsync<DeepCheckException>(() =>
            client.RunAsync(Url, "t", T.Claims(Line), new List<Finding>(), CancellationToken.None));
        return ex.Message;
    }

    [TestMethod]
    public async Task Client_maps_401_and_403_to_not_authorized()
    {
        Assert.AreEqual("Your account isn't authorized for Deep Check. Contact your administrator.",
            await FailureMessage(Json(HttpStatusCode.Unauthorized, "{}")));
        Assert.AreEqual("Your account isn't authorized for Deep Check. Contact your administrator.",
            await FailureMessage(Json(HttpStatusCode.Forbidden, "{}")));
    }

    [TestMethod]
    public async Task Client_maps_429_using_retry_after()
    {
        var response = Json((HttpStatusCode)429, """{"error":{"code":"rate_limited","message":"x"}}""");
        response.Headers.RetryAfter = new System.Net.Http.Headers.RetryConditionHeaderValue(TimeSpan.FromSeconds(90));
        Assert.AreEqual("Deep Check limit reached; try again in 2 minutes.", await FailureMessage(response));
    }

    [TestMethod]
    public async Task Client_maps_422_to_declined()
    {
        Assert.AreEqual("Claude declined to review these claims. The rule-based findings are unchanged.",
            await FailureMessage(Json((HttpStatusCode)422, """{"error":{"code":"refused","message":"x"}}""")));
    }

    [TestMethod]
    public async Task Client_shows_the_gateway_message_for_400()
    {
        Assert.AreEqual("The claims are too long for Deep Check (limit 200,000 characters). The rule-based findings are unchanged.",
            await FailureMessage(Json(HttpStatusCode.BadRequest,
                """{"error":{"code":"too_large","message":"The claims are too long for Deep Check (limit 200,000 characters)."}}""")));
    }

    [TestMethod]
    public async Task Client_uses_a_generic_message_for_5xx_without_body()
    {
        Assert.AreEqual("The Deep Check service had a problem. Try again shortly. The rule-based findings are unchanged.",
            await FailureMessage(new HttpResponseMessage(HttpStatusCode.ServiceUnavailable)));
    }

    [TestMethod]
    public async Task Client_maps_network_failure()
    {
        Assert.AreEqual("Could not reach the Deep Check service. Check your network connection.",
            await FailureFromThrow(new HttpRequestException("no route")));
    }

    [TestMethod]
    public async Task Client_maps_timeout()
    {
        Assert.AreEqual("Deep Check took too long and was stopped. Try checking a selection of claims.",
            await FailureFromThrow(new TaskCanceledException("timeout")));
    }

    [TestMethod]
    public async Task Client_rethrows_user_cancellation()
    {
        using var cts = new CancellationTokenSource();
        cts.Cancel();
        var (client, _) = NewClient(_ => throw new TaskCanceledException());
        try
        {
            await client.RunAsync(Url, "t", T.Claims(Line), new List<Finding>(), cts.Token);
            Assert.Fail("Expected cancellation.");
        }
        catch (OperationCanceledException)
        {
            // HttpClient may surface TaskCanceledException or OperationCanceledException; both mean the user canceled.
        }
    }

    static async Task<string> FailureFromThrow(Exception toThrow)
    {
        var (client, _) = NewClient(_ => throw toThrow);
        var ex = await Assert.ThrowsExceptionAsync<DeepCheckException>(() =>
            client.RunAsync(Url, "t", T.Claims(Line), new List<Finding>(), CancellationToken.None));
        return ex.Message;
    }
}
```

- [ ] **Step 3: Run the tests to verify they fail**

Run: `dotnet test LLDOCX.Claims.Tests --filter DeepCheckProtocolTests`
Expected: the build FAILS with `The type or namespace name 'DeepCheckRequestBuilder' could not be found`.

- [ ] **Step 4: Write `LLDOCX/LLDOCX.DeepCheck/DeepCheckException.cs`**

```csharp
using System;

namespace LLDOCX.DeepCheck
{
    /// <summary>A Deep Check failure whose Message is safe to show to the user.</summary>
    public class DeepCheckException : Exception
    {
        public DeepCheckException(string message, Exception inner = null) : base(message, inner) { }
    }
}
```

- [ ] **Step 5: Write `LLDOCX/LLDOCX.DeepCheck/DeepCheckContract.cs`**

```csharp
using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace LLDOCX.DeepCheck
{
    /// <summary>Body of POST {gatewayUrl}. Shared by the add-in (sender) and the gateway (receiver).</summary>
    public sealed class DeepCheckRequest
    {
        [JsonPropertyName("claims")] public List<ClaimDto> Claims { get; set; } = new List<ClaimDto>();
        [JsonPropertyName("ruleFindings")] public List<RuleFindingDto> RuleFindings { get; set; } = new List<RuleFindingDto>();
    }

    public sealed class ClaimDto
    {
        [JsonPropertyName("number")] public int Number { get; set; }
        [JsonPropertyName("text")] public string Text { get; set; }
        [JsonPropertyName("canceled")] public bool Canceled { get; set; }
    }

    /// <summary>An offline finding. Its position in the list is the index Claude uses for disputes.</summary>
    public sealed class RuleFindingDto
    {
        [JsonPropertyName("claim")] public int Claim { get; set; }
        [JsonPropertyName("type")] public string Type { get; set; }
        [JsonPropertyName("severity")] public string Severity { get; set; }
        [JsonPropertyName("message")] public string Message { get; set; }
    }

    /// <summary>Body of every non-200 gateway response.</summary>
    public sealed class ErrorBody
    {
        [JsonPropertyName("error")] public ErrorDetail Error { get; set; }
    }

    public sealed class ErrorDetail
    {
        [JsonPropertyName("code")] public string Code { get; set; }
        [JsonPropertyName("message")] public string Message { get; set; }
    }

    public static class ErrorCodes
    {
        public const string InvalidRequest = "invalid_request";
        public const string TooLarge = "too_large";
        public const string Unauthorized = "unauthorized";
        public const string Forbidden = "forbidden";
        public const string RateLimited = "rate_limited";
        public const string Refused = "refused";
        public const string Truncated = "truncated";
        public const string Timeout = "timeout";
        public const string UpstreamUnavailable = "upstream_unavailable";
        public const string BadUpstreamResponse = "bad_upstream_response";
    }
}
```

- [ ] **Step 6: Write `LLDOCX/LLDOCX.DeepCheck/DeepCheckRequestBuilder.cs`**

```csharp
using System.Collections.Generic;
using System.Linq;
using LLDOCX.Claims;

namespace LLDOCX.DeepCheck
{
    public static class DeepCheckRequestBuilder
    {
        public static DeepCheckRequest Build(IList<Claim> claims, IList<Finding> ruleFindings)
        {
            var request = new DeepCheckRequest();
            foreach (var c in claims)
                request.Claims.Add(new ClaimDto { Number = c.Number, Text = c.IsCanceled ? "" : c.FullText, Canceled = c.IsCanceled });
            foreach (var f in ruleFindings)
                request.RuleFindings.Add(new RuleFindingDto
                {
                    Claim = f.ClaimNumber,
                    Type = f.Type.ToString(),
                    Severity = f.Severity.ToString(),
                    Message = f.Message
                });
            return request;
        }

        public static int TextLength(DeepCheckRequest request) =>
            request.Claims == null ? 0 : request.Claims.Sum(c => c?.Text?.Length ?? 0);
    }
}
```

- [ ] **Step 7: Write `LLDOCX/LLDOCX.DeepCheck/DeepCheckResponseParser.cs`**

```csharp
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using LLDOCX.Claims;

namespace LLDOCX.DeepCheck
{
    public static class DeepCheckResponseParser
    {
        const string Unreadable = "Deep Check returned a response that could not be read. The rule-based findings are unchanged.";

        public static List<Finding> Parse(string json, IList<Claim> claims, IList<Finding> ruleFindings)
        {
            JsonDocument doc;
            try
            {
                doc = JsonDocument.Parse(json);
            }
            catch (JsonException ex)
            {
                throw new DeepCheckException(Unreadable, ex);
            }

            using (doc)
            {
                var root = doc.RootElement;
                if (root.ValueKind != JsonValueKind.Object) throw new DeepCheckException(Unreadable);

                var added = new List<Finding>();
                if (root.TryGetProperty("findings", out var findings) && findings.ValueKind == JsonValueKind.Array)
                    foreach (var f in findings.EnumerateArray())
                        added.Add(ToFinding(f, claims));

                if (root.TryGetProperty("disputes", out var disputes) && disputes.ValueKind == JsonValueKind.Array)
                {
                    foreach (var d in disputes.EnumerateArray())
                    {
                        int index = d.TryGetProperty("index", out var i) && i.ValueKind == JsonValueKind.Number ? i.GetInt32() : -1;
                        if (index < 0 || index >= ruleFindings.Count) continue;
                        ruleFindings[index].DisputedByAI = true;
                        ruleFindings[index].DisputeReason = Str(d, "reason");
                    }
                }
                return added;
            }
        }

        /// <summary>True when json has the response schema's shape. The gateway checks this before returning Claude's output.</summary>
        public static bool IsWellFormed(string json)
        {
            try
            {
                using (var doc = JsonDocument.Parse(json))
                {
                    var root = doc.RootElement;
                    if (root.ValueKind != JsonValueKind.Object) return false;
                    if (!root.TryGetProperty("findings", out var findings) || findings.ValueKind != JsonValueKind.Array) return false;
                    if (!root.TryGetProperty("disputes", out var disputes) || disputes.ValueKind != JsonValueKind.Array) return false;
                    foreach (var f in findings.EnumerateArray())
                    {
                        if (f.ValueKind != JsonValueKind.Object || !Is(f, "claim", JsonValueKind.Number)) return false;
                        foreach (var name in new[] { "type", "severity", "quote", "message", "suggestion" })
                            if (!Is(f, name, JsonValueKind.String)) return false;
                    }
                    foreach (var d in disputes.EnumerateArray())
                        if (d.ValueKind != JsonValueKind.Object || !Is(d, "index", JsonValueKind.Number) || !Is(d, "reason", JsonValueKind.String))
                            return false;
                    return true;
                }
            }
            catch (JsonException)
            {
                return false;
            }
        }

        static bool Is(JsonElement e, string name, JsonValueKind kind) => e.TryGetProperty(name, out var v) && v.ValueKind == kind;

        static Finding ToFinding(JsonElement f, IList<Claim> claims)
        {
            int number = f.TryGetProperty("claim", out var n) && n.ValueKind == JsonValueKind.Number ? n.GetInt32() : 0;
            var type = ParseEnum(Str(f, "type"), FindingType.Antecedent);
            var severity = ParseEnum(Str(f, "severity"), Severity.Warning);
            string quote = Str(f, "quote");
            string message = Str(f, "message");
            string suggestion = Str(f, "suggestion");
            if (suggestion.Length == 0) suggestion = null;

            var claim = claims.FirstOrDefault(c => c.Number == number);
            int idx = claim == null || quote.Length == 0 ? -1 : claim.FullText.IndexOf(quote, StringComparison.OrdinalIgnoreCase);
            var result = idx >= 0
                ? Finding.At(claim, idx, quote.Length, type, severity, message, suggestion)
                : new Finding { ClaimNumber = number, Type = type, Severity = severity, Message = message, Suggestion = suggestion, DocStart = -1, Text = quote };
            result.Source = "AI";
            return result;
        }

        static string Str(JsonElement e, string name) =>
            e.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.String ? v.GetString() ?? "" : "";

        static T ParseEnum<T>(string s, T fallback) where T : struct =>
            Enum.TryParse(s, true, out T value) ? value : fallback;
    }
}
```

- [ ] **Step 8: Write `LLDOCX/LLDOCX.DeepCheck/DeepCheckHttpClient.cs`**

```csharp
using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using LLDOCX.Claims;

namespace LLDOCX.DeepCheck
{
    /// <summary>Calls the lozaip.com Deep Check gateway. No model, prompt, or Claude credential lives on the client.</summary>
    public sealed class DeepCheckHttpClient
    {
        /// <summary>Just above the gateway's own 210-second Claude timeout (Azure stops HTTP requests at about 230 seconds).</summary>
        public static readonly TimeSpan DefaultTimeout = TimeSpan.FromSeconds(240);

        const string Kept = " The rule-based findings are unchanged.";
        readonly HttpClient _http;

        public DeepCheckHttpClient(HttpClient http)
        {
            _http = http ?? throw new ArgumentNullException(nameof(http));
        }

        public async Task<List<Finding>> RunAsync(string gatewayUrl, string bearerToken, IList<Claim> claims,
            IList<Finding> ruleFindings, CancellationToken ct)
        {
            string body = JsonSerializer.Serialize(DeepCheckRequestBuilder.Build(claims, ruleFindings));
            using (var request = new HttpRequestMessage(HttpMethod.Post, gatewayUrl))
            {
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", bearerToken);
                request.Content = new StringContent(body, Encoding.UTF8, "application/json");

                HttpResponseMessage response;
                try
                {
                    response = await _http.SendAsync(request, ct).ConfigureAwait(false);
                }
                catch (TaskCanceledException ex) when (!ct.IsCancellationRequested)
                {
                    throw new DeepCheckException("Deep Check took too long and was stopped. Try checking a selection of claims.", ex);
                }
                catch (HttpRequestException ex)
                {
                    throw new DeepCheckException("Could not reach the Deep Check service. Check your network connection.", ex);
                }

                using (response)
                {
                    string text = response.Content == null ? "" : await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                    if (response.IsSuccessStatusCode) return DeepCheckResponseParser.Parse(text, claims, ruleFindings);
                    throw new DeepCheckException(MessageFor((int)response.StatusCode, ServerMessage(text), response.Headers.RetryAfter?.Delta));
                }
            }
        }

        static string MessageFor(int status, string serverMessage, TimeSpan? retryAfter)
        {
            switch (status)
            {
                case 401:
                case 403:
                    return "Your account isn't authorized for Deep Check. Contact your administrator.";
                case 422:
                    return "Claude declined to review these claims." + Kept;
                case 429:
                    int minutes = retryAfter.HasValue ? Math.Max(1, (int)Math.Ceiling(retryAfter.Value.TotalMinutes)) : 60;
                    return "Deep Check limit reached; try again in " + minutes + (minutes == 1 ? " minute." : " minutes.");
                default:
                    if (!string.IsNullOrWhiteSpace(serverMessage)) return serverMessage + Kept;
                    return status >= 500
                        ? "The Deep Check service had a problem. Try again shortly." + Kept
                        : "Deep Check could not process these claims." + Kept;
            }
        }

        static string ServerMessage(string body)
        {
            if (string.IsNullOrWhiteSpace(body)) return null;
            try
            {
                return JsonSerializer.Deserialize<ErrorBody>(body)?.Error?.Message;
            }
            catch (JsonException)
            {
                return null;
            }
        }
    }
}
```

- [ ] **Step 9: Run the tests to verify they pass**

Run: `dotnet test LLDOCX.Claims.Tests`
Expected: all tests pass, 0 failed.

- [ ] **Step 10: Commit**

```bash
git add LLDOCX/LLDOCX.sln LLDOCX/LLDOCX.DeepCheck LLDOCX/LLDOCX.Claims.Tests
git commit -m "feat(deepcheck): add gateway wire contract, response parser, and HTTP client

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 10: Deep Check gateway (Azure Function)

**Files:**
- Create: `LLDOCX/LLDOCX.Gateway/LLDOCX.Gateway.csproj`, `host.json`, `local.settings.sample.json`
- Create: `LLDOCX/LLDOCX.Gateway/GatewayOptions.cs`
- Create: `LLDOCX/LLDOCX.Gateway/DeepCheckPrompt.cs`
- Create: `LLDOCX/LLDOCX.Gateway/TokenValidator.cs`
- Create: `LLDOCX/LLDOCX.Gateway/UserRateLimiter.cs`
- Create: `LLDOCX/LLDOCX.Gateway/ClaudeReviewer.cs`
- Create: `LLDOCX/LLDOCX.Gateway/DeepCheckHandler.cs`
- Create: `LLDOCX/LLDOCX.Gateway/DeepCheckFunction.cs`
- Create: `LLDOCX/LLDOCX.Gateway/Program.cs`
- Create: `LLDOCX/LLDOCX.Gateway.Tests/LLDOCX.Gateway.Tests.csproj`
- Test: `LLDOCX/LLDOCX.Gateway.Tests/GatewayTests.cs`

**Interfaces:**
- Consumes: `DeepCheckRequest`, `ClaimDto`, `RuleFindingDto`, `ErrorBody`, `ErrorDetail`, `ErrorCodes`, `DeepCheckRequestBuilder.TextLength`, `DeepCheckResponseParser.IsWellFormed` (Task 9)
- Produces (namespace `LLDOCX.Gateway`):
  - `sealed class GatewayOptions`:
    - settings: `TenantId`, `Audience` (comma-separated), `RequiredScope` (= "DeepCheck.Use"), `RequiredRole`, `ClaudeProvider` ("Foundry" | "Anthropic"), `FoundryResource`, `AnthropicApiKey`;
    - defaults: `Model` (= "claude-opus-5-5"), `Effort` (= "high"), `MaxOutputTokens` (= 16000), `MaxInputChars` (= 200000), `RequestsPerUserPerHour` (= 30), `ClaudeTimeoutSeconds` (= 210);
    - `IReadOnlyList<string> Audiences`, `void Validate()`.
  - `static class DeepCheckPrompt { const string SystemPrompt; const string ResponseSchemaJson; static string BuildUserMessage(DeepCheckRequest request); }`
  - Token validation:
    - `sealed class CallerIdentity(string objectId, string username)`;
    - `enum AuthFailureKind { Unauthenticated, Forbidden }` and `sealed class AuthFailure(AuthFailureKind kind, string message) : Exception`;
    - `interface ITokenValidator { Task<CallerIdentity> ValidateAsync(string authorizationHeader, CancellationToken ct); }`;
    - `sealed class EntraTokenValidator(GatewayOptions, IConfigurationManager<OpenIdConnectConfiguration>)`, with `static CreateConfigurationManager(string tenantId)`.
  - `sealed class UserRateLimiter(int perHour, Func<DateTime> utcNow) { bool TryAcquire(string userId, out TimeSpan retryAfter); }`
  - Claude call:
    - `enum ReviewOutcome { Ok, Refused, Truncated }` and `sealed class ReviewResult` (`Outcome`, `Json`; `static Ok(string)`, `Refused()`, `Truncated()`);
    - `sealed class UpstreamException(string message, Exception inner) : Exception`;
    - `interface IClaudeReviewer { Task<ReviewResult> ReviewAsync(string systemPrompt, string userMessage, string responseSchemaJson, CancellationToken ct); }`;
    - `sealed class SdkClaudeReviewer : IClaudeReviewer`, with `static Create(GatewayOptions)`.
  - Request handling:
    - `sealed class HandlerResult(int status, string body, int? retryAfterSeconds = null)`, with `static Error(int status, string code, string message, int? retryAfterSeconds = null)`;
    - `sealed class DeepCheckHandler(GatewayOptions, ITokenValidator, UserRateLimiter, IClaudeReviewer, ILogger<DeepCheckHandler>)` with `Task<HandlerResult> HandleAsync(string authorizationHeader, string body, CancellationToken ct)`.
  - HTTP endpoint: `POST /api/deepcheck`

- [ ] **Step 1: Create the gateway and test project files**

`LLDOCX/LLDOCX.Gateway/LLDOCX.Gateway.csproj`:

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFramework>net10.0</TargetFramework>
    <AzureFunctionsVersion>v4</AzureFunctionsVersion>
    <OutputType>Exe</OutputType>
    <ImplicitUsings>enable</ImplicitUsings>
    <Nullable>disable</Nullable>
    <RootNamespace>LLDOCX.Gateway</RootNamespace>
  </PropertyGroup>
  <ItemGroup>
    <ProjectReference Include="..\LLDOCX.DeepCheck\LLDOCX.DeepCheck.csproj" />
  </ItemGroup>
  <ItemGroup>
    <None Update="host.json" CopyToOutputDirectory="PreserveNewest" />
    <None Update="local.settings.json" CopyToOutputDirectory="PreserveNewest" CopyToPublishDirectory="Never" />
  </ItemGroup>
  <ItemGroup>
    <Using Include="System.Threading.ExecutionContext" Alias="ExecutionContext" />
  </ItemGroup>
</Project>
```

`LLDOCX/LLDOCX.Gateway/host.json`:

```json
{
  "version": "2.0",
  "functionTimeout": "00:05:00",
  "logging": {
    "applicationInsights": {
      "samplingSettings": { "isEnabled": true, "excludedTypes": "Request" }
    }
  }
}
```

`LLDOCX/LLDOCX.Gateway/local.settings.sample.json`. To run locally, copy it to `local.settings.json`; that file is gitignored by the gate plan's Task 1.

```json
{
  "IsEncrypted": false,
  "Values": {
    "AzureWebJobsStorage": "UseDevelopmentStorage=true",
    "FUNCTIONS_WORKER_RUNTIME": "dotnet-isolated",
    "Gateway__TenantId": "<lozaip.com tenant GUID>",
    "Gateway__Audience": "api://<gateway client id>,<gateway client id>",
    "Gateway__RequiredScope": "DeepCheck.Use",
    "Gateway__RequiredRole": "",
    "Gateway__ClaudeProvider": "Foundry",
    "Gateway__FoundryResource": "<Foundry resource name>",
    "Gateway__Model": "claude-opus-5-5",
    "Gateway__Effort": "high"
  }
}
```

`LLDOCX/LLDOCX.Gateway.Tests/LLDOCX.Gateway.Tests.csproj`:

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFramework>net10.0</TargetFramework>
    <ImplicitUsings>enable</ImplicitUsings>
    <Nullable>disable</Nullable>
    <IsPackable>false</IsPackable>
  </PropertyGroup>
  <ItemGroup>
    <PackageReference Include="Microsoft.NET.Test.Sdk" Version="17.11.1" />
    <PackageReference Include="MSTest.TestAdapter" Version="3.6.1" />
    <PackageReference Include="MSTest.TestFramework" Version="3.6.1" />
  </ItemGroup>
  <ItemGroup>
    <ProjectReference Include="..\LLDOCX.Gateway\LLDOCX.Gateway.csproj" />
  </ItemGroup>
</Project>
```

Run, from `Tools/LLDOCX/`:

```bash
dotnet add LLDOCX.Gateway package Microsoft.Azure.Functions.Worker
dotnet add LLDOCX.Gateway package Microsoft.Azure.Functions.Worker.Sdk
dotnet add LLDOCX.Gateway package Microsoft.Azure.Functions.Worker.Extensions.Http
dotnet add LLDOCX.Gateway package Microsoft.Azure.Functions.Worker.ApplicationInsights
dotnet add LLDOCX.Gateway package Microsoft.ApplicationInsights.WorkerService
dotnet add LLDOCX.Gateway package Microsoft.Extensions.Configuration.Binder
dotnet add LLDOCX.Gateway package Anthropic
dotnet add LLDOCX.Gateway package Anthropic.Foundry
dotnet add LLDOCX.Gateway package Azure.Identity
dotnet add LLDOCX.Gateway package Microsoft.IdentityModel.JsonWebTokens
dotnet add LLDOCX.Gateway package Microsoft.IdentityModel.Protocols.OpenIdConnect
dotnet sln LLDOCX.sln add LLDOCX.Gateway/LLDOCX.Gateway.csproj LLDOCX.Gateway.Tests/LLDOCX.Gateway.Tests.csproj
```

Expected: each command reports the package added, and the solution reports both projects added.

- [ ] **Step 2: Write the failing tests `LLDOCX/LLDOCX.Gateway.Tests/GatewayTests.cs`**

```csharp
global using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Security.Cryptography;
using System.Text.Json;
using LLDOCX.DeepCheck;
using LLDOCX.Gateway;
using Microsoft.Extensions.Logging;
using Microsoft.IdentityModel.JsonWebTokens;
using Microsoft.IdentityModel.Protocols;
using Microsoft.IdentityModel.Protocols.OpenIdConnect;
using Microsoft.IdentityModel.Tokens;

namespace LLDOCX.Gateway.Tests;

internal sealed class CapturingLogger : ILogger<DeepCheckHandler>
{
    public List<string> Lines { get; } = new();
    public IDisposable BeginScope<TState>(TState state) where TState : notnull => null;
    public bool IsEnabled(LogLevel logLevel) => true;
    public void Log<TState>(LogLevel logLevel, EventId eventId, TState state, Exception exception, Func<TState, Exception, string> formatter) =>
        Lines.Add(formatter(state, exception) + " | " + state + " | " + exception);
}

internal sealed class FakeValidator : ITokenValidator
{
    public AuthFailure Failure { get; set; }
    public Task<CallerIdentity> ValidateAsync(string authorizationHeader, CancellationToken ct) =>
        Failure != null ? throw Failure : Task.FromResult(new CallerIdentity("oid-1", "pat@lozaip.com"));
}

internal sealed class FakeReviewer : IClaudeReviewer
{
    public Func<ReviewResult> Result { get; set; } = () => ReviewResult.Ok("""{"findings":[],"disputes":[]}""");
    public string LastUserMessage { get; private set; }

    public Task<ReviewResult> ReviewAsync(string systemPrompt, string userMessage, string responseSchemaJson, CancellationToken ct)
    {
        LastUserMessage = userMessage;
        return Task.FromResult(Result());
    }
}

[TestClass]
public class PromptTests
{
    [TestMethod]
    public void User_message_lists_claims_canceled_claims_and_indexed_rule_findings()
    {
        var request = new DeepCheckRequest
        {
            Claims = { new ClaimDto { Number = 1, Text = "A lamp comprising a base." }, new ClaimDto { Number = 2, Canceled = true, Text = "" } },
            RuleFindings = { new RuleFindingDto { Claim = 1, Type = "Antecedent", Severity = "Error", Message = "No antecedent basis for 'the shade'." } }
        };
        string msg = DeepCheckPrompt.BuildUserMessage(request);
        StringAssert.Contains(msg, "1. A lamp comprising a base.");
        StringAssert.Contains(msg, "2. (Canceled)");
        StringAssert.Contains(msg, "[0] claim 1, Antecedent, Error: No antecedent basis for 'the shade'.");
    }

    [TestMethod]
    public void Schema_is_valid_json_with_required_fields()
    {
        using var doc = JsonDocument.Parse(DeepCheckPrompt.ResponseSchemaJson);
        var required = doc.RootElement.GetProperty("required").EnumerateArray().Select(e => e.GetString()).ToArray();
        CollectionAssert.AreEquivalent(new[] { "findings", "disputes" }, required);
    }
}

[TestClass]
public class TokenValidatorTests
{
    const string Tenant = "11111111-1111-1111-1111-111111111111";
    const string Audience = "api://gateway-app";
    const string ClientIdAudience = "22222222-2222-2222-2222-222222222222";
    static readonly RsaSecurityKey Key = new(RSA.Create(2048)) { KeyId = "test-key" };

    static GatewayOptions Options(string role = null) =>
        new() { TenantId = Tenant, Audience = Audience + "," + ClientIdAudience, RequiredRole = role };

    static EntraTokenValidator Validator(GatewayOptions o)
    {
        var config = new OpenIdConnectConfiguration();
        config.SigningKeys.Add(Key);
        return new EntraTokenValidator(o, new StaticConfigurationManager<OpenIdConnectConfiguration>(config));
    }

    static string Token(Action<Dictionary<string, object>> edit = null, string issuer = null, string audience = Audience, DateTime? expires = null)
    {
        var claims = new Dictionary<string, object>
        {
            ["tid"] = Tenant,
            ["oid"] = "oid-1",
            ["preferred_username"] = "pat@lozaip.com",
            ["scp"] = "User.Read DeepCheck.Use"
        };
        edit?.Invoke(claims);
        DateTime exp = expires ?? DateTime.UtcNow.AddMinutes(30);
        return new JsonWebTokenHandler().CreateToken(new SecurityTokenDescriptor
        {
            Issuer = issuer ?? $"https://login.microsoftonline.com/{Tenant}/v2.0",
            Audience = audience,
            Claims = claims,
            NotBefore = exp.AddHours(-1),
            IssuedAt = exp.AddHours(-1),
            Expires = exp,
            SigningCredentials = new SigningCredentials(Key, SecurityAlgorithms.RsaSha256)
        });
    }

    static async Task<AuthFailureKind> Fails(EntraTokenValidator v, string header)
    {
        var ex = await Assert.ThrowsExceptionAsync<AuthFailure>(() => v.ValidateAsync(header, CancellationToken.None));
        return ex.Kind;
    }

    [TestMethod]
    public async Task Valid_token_returns_the_caller()
    {
        var caller = await Validator(Options()).ValidateAsync("Bearer " + Token(), CancellationToken.None);
        Assert.AreEqual("oid-1", caller.ObjectId);
        Assert.AreEqual("pat@lozaip.com", caller.Username);
    }

    [TestMethod]
    public async Task Client_id_audience_is_also_accepted()
    {
        var caller = await Validator(Options()).ValidateAsync("Bearer " + Token(audience: ClientIdAudience), CancellationToken.None);
        Assert.AreEqual("oid-1", caller.ObjectId);
    }

    [TestMethod]
    public async Task Missing_or_malformed_header_is_unauthenticated()
    {
        Assert.AreEqual(AuthFailureKind.Unauthenticated, await Fails(Validator(Options()), null));
        Assert.AreEqual(AuthFailureKind.Unauthenticated, await Fails(Validator(Options()), "Basic abc"));
        Assert.AreEqual(AuthFailureKind.Unauthenticated, await Fails(Validator(Options()), "Bearer not-a-jwt"));
    }

    [TestMethod]
    public async Task Other_tenant_is_unauthenticated()
    {
        const string other = "99999999-9999-9999-9999-999999999999";
        string token = Token(c => c["tid"] = other, issuer: $"https://login.microsoftonline.com/{other}/v2.0");
        Assert.AreEqual(AuthFailureKind.Unauthenticated, await Fails(Validator(Options()), "Bearer " + token));
    }

    [TestMethod]
    public async Task Wrong_audience_is_unauthenticated()
    {
        Assert.AreEqual(AuthFailureKind.Unauthenticated, await Fails(Validator(Options()), "Bearer " + Token(audience: "api://other")));
    }

    [TestMethod]
    public async Task Expired_token_is_unauthenticated()
    {
        Assert.AreEqual(AuthFailureKind.Unauthenticated,
            await Fails(Validator(Options()), "Bearer " + Token(expires: DateTime.UtcNow.AddMinutes(-10))));
    }

    [TestMethod]
    public async Task Missing_scope_is_forbidden()
    {
        Assert.AreEqual(AuthFailureKind.Forbidden, await Fails(Validator(Options()), "Bearer " + Token(c => c["scp"] = "User.Read")));
    }

    [TestMethod]
    public async Task Required_role_is_enforced()
    {
        Assert.AreEqual(AuthFailureKind.Forbidden, await Fails(Validator(Options("DeepCheck.User")), "Bearer " + Token()));
        var caller = await Validator(Options("DeepCheck.User"))
            .ValidateAsync("Bearer " + Token(c => c["roles"] = new[] { "Other", "DeepCheck.User" }), CancellationToken.None);
        Assert.AreEqual("oid-1", caller.ObjectId);
    }
}

[TestClass]
public class RateLimiterTests
{
    [TestMethod]
    public void Allows_the_limit_then_blocks_with_retry_after()
    {
        var now = new DateTime(2026, 11, 1, 12, 0, 0, DateTimeKind.Utc);
        var limiter = new UserRateLimiter(2, () => now);
        Assert.IsTrue(limiter.TryAcquire("a", out _));
        now = now.AddMinutes(10);
        Assert.IsTrue(limiter.TryAcquire("a", out _));
        Assert.IsFalse(limiter.TryAcquire("a", out var retry));
        Assert.AreEqual(TimeSpan.FromMinutes(50), retry);
        Assert.IsTrue(limiter.TryAcquire("b", out _), "users are independent");
        now = now.AddMinutes(51);
        Assert.IsTrue(limiter.TryAcquire("a", out _), "the window slides");
    }
}

[TestClass]
public class DeepCheckHandlerTests
{
    const string Secret = "ZZ-CONFIDENTIAL-CLAIM-TEXT-ZZ";
    readonly FakeValidator _validator = new();
    readonly FakeReviewer _reviewer = new();
    readonly CapturingLogger _log = new();
    readonly GatewayOptions _options = new()
    {
        TenantId = "11111111-1111-1111-1111-111111111111",
        Audience = "api://gw",
        MaxInputChars = 1000,
        RequestsPerUserPerHour = 30
    };

    DeepCheckHandler NewHandler() =>
        new(_options, _validator, new UserRateLimiter(_options.RequestsPerUserPerHour, () => DateTime.UtcNow), _reviewer, _log);

    static string Body(string text = "A lamp comprising a base, wherein the shade is green. " + Secret) =>
        JsonSerializer.Serialize(new DeepCheckRequest { Claims = { new ClaimDto { Number = 1, Text = text } } });

    static ErrorDetail Error(HandlerResult r) => JsonSerializer.Deserialize<ErrorBody>(r.Body).Error;

    [TestMethod]
    public async Task Success_returns_claude_json_and_logs_metadata_only()
    {
        _reviewer.Result = () => ReviewResult.Ok("""{"findings":[{"claim":1,"type":"Antecedent","severity":"Error","quote":"the shade","message":"m","suggestion":""}],"disputes":[]}""");
        var r = await NewHandler().HandleAsync("Bearer x", Body(), CancellationToken.None);
        Assert.AreEqual(200, r.Status);
        Assert.IsTrue(DeepCheckResponseParser.IsWellFormed(r.Body));
        StringAssert.Contains(_reviewer.LastUserMessage, Secret);
        Assert.IsTrue(_log.Lines.Any(l => l.Contains("ok") && l.Contains("pat@lozaip.com")), string.Join("\n", _log.Lines));
        Assert.IsFalse(_log.Lines.Any(l => l.Contains(Secret) || l.Contains("the shade")), string.Join("\n", _log.Lines));
    }

    [TestMethod]
    public async Task Auth_failures_map_to_401_and_403()
    {
        _validator.Failure = new AuthFailure(AuthFailureKind.Unauthenticated, "x");
        Assert.AreEqual(401, (await NewHandler().HandleAsync(null, Body(), CancellationToken.None)).Status);
        _validator.Failure = new AuthFailure(AuthFailureKind.Forbidden, "x");
        var r = await NewHandler().HandleAsync("Bearer x", Body(), CancellationToken.None);
        Assert.AreEqual(403, r.Status);
        Assert.AreEqual(ErrorCodes.Forbidden, Error(r).Code);
    }

    [TestMethod]
    public async Task Invalid_or_empty_request_is_400()
    {
        Assert.AreEqual(ErrorCodes.InvalidRequest, Error(await NewHandler().HandleAsync("Bearer x", "not json", CancellationToken.None)).Code);
        Assert.AreEqual(ErrorCodes.InvalidRequest, Error(await NewHandler().HandleAsync("Bearer x", """{"claims":[]}""", CancellationToken.None)).Code);
    }

    [TestMethod]
    public async Task Too_large_is_400_with_the_limit_in_the_message()
    {
        var r = await NewHandler().HandleAsync("Bearer x", Body(new string('a', 1001)), CancellationToken.None);
        Assert.AreEqual(400, r.Status);
        Assert.AreEqual(ErrorCodes.TooLarge, Error(r).Code);
        StringAssert.Contains(Error(r).Message, "limit 1,000 characters");
    }

    [TestMethod]
    public async Task Rate_limit_is_429_with_retry_after()
    {
        _options.RequestsPerUserPerHour = 1;
        var handler = NewHandler();
        Assert.AreEqual(200, (await handler.HandleAsync("Bearer x", Body(), CancellationToken.None)).Status);
        var r = await handler.HandleAsync("Bearer x", Body(), CancellationToken.None);
        Assert.AreEqual(429, r.Status);
        Assert.IsTrue(r.RetryAfterSeconds > 3500);
    }

    [TestMethod]
    public async Task Refusal_is_422()
    {
        _reviewer.Result = ReviewResult.Refused;
        var r = await NewHandler().HandleAsync("Bearer x", Body(), CancellationToken.None);
        Assert.AreEqual(422, r.Status);
        Assert.AreEqual(ErrorCodes.Refused, Error(r).Code);
    }

    [TestMethod]
    public async Task Truncation_is_502()
    {
        _reviewer.Result = ReviewResult.Truncated;
        var r = await NewHandler().HandleAsync("Bearer x", Body(), CancellationToken.None);
        Assert.AreEqual(502, r.Status);
        Assert.AreEqual(ErrorCodes.Truncated, Error(r).Code);
    }

    [TestMethod]
    public async Task Upstream_failure_is_503_and_does_not_log_inner_messages()
    {
        _reviewer.Result = () => throw new UpstreamException("Claude request failed.", new InvalidOperationException(Secret));
        var r = await NewHandler().HandleAsync("Bearer x", Body(), CancellationToken.None);
        Assert.AreEqual(503, r.Status);
        Assert.IsFalse(_log.Lines.Any(l => l.Contains(Secret)), string.Join("\n", _log.Lines));
    }

    [TestMethod]
    public async Task Malformed_claude_output_is_502()
    {
        _reviewer.Result = () => ReviewResult.Ok("""{"findings":"nope"}""");
        Assert.AreEqual(ErrorCodes.BadUpstreamResponse, Error(await NewHandler().HandleAsync("Bearer x", Body(), CancellationToken.None)).Code);
    }

    [TestMethod]
    public void Options_validation_reports_missing_settings()
    {
        var ex = Assert.ThrowsException<InvalidOperationException>(() => new GatewayOptions().Validate());
        StringAssert.Contains(ex.Message, "TenantId");
        StringAssert.Contains(ex.Message, "Audience");
        StringAssert.Contains(ex.Message, "FoundryResource");
    }
}
```

- [ ] **Step 3: Run the tests to verify they fail**

Run: `dotnet test LLDOCX.Gateway.Tests`
Expected: the build FAILS with `The type or namespace name 'DeepCheckHandler' could not be found`.

- [ ] **Step 4: Write `LLDOCX/LLDOCX.Gateway/GatewayOptions.cs`**

```csharp
namespace LLDOCX.Gateway;

/// <summary>Gateway settings, bound from the "Gateway" configuration section (app settings named Gateway__Name).</summary>
public sealed class GatewayOptions
{
    public string TenantId { get; set; }

    /// <summary>Comma-separated accepted audiences: the gateway's App ID URI and/or client ID.</summary>
    public string Audience { get; set; }

    public string RequiredScope { get; set; } = "DeepCheck.Use";

    /// <summary>Optional app role the caller must hold (for example "DeepCheck.User").</summary>
    public string RequiredRole { get; set; }

    /// <summary>"Foundry" (managed identity, no key) or "Anthropic" (API key from Key Vault).</summary>
    public string ClaudeProvider { get; set; } = "Foundry";

    public string FoundryResource { get; set; }
    public string AnthropicApiKey { get; set; }

    /// <summary>Model ID, or the Foundry deployment name.</summary>
    public string Model { get; set; } = "claude-opus-5-5";

    public string Effort { get; set; } = "high";
    public int MaxOutputTokens { get; set; } = 16000;
    public int MaxInputChars { get; set; } = 200000;
    public int RequestsPerUserPerHour { get; set; } = 30;

    /// <summary>Kept under Azure's 230-second HTTP response limit.</summary>
    public int ClaudeTimeoutSeconds { get; set; } = 210;

    public IReadOnlyList<string> Audiences =>
        (Audience ?? "").Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

    public void Validate()
    {
        var errors = new List<string>();
        if (!Guid.TryParse(TenantId, out _)) errors.Add("Gateway:TenantId must be the lozaip.com tenant GUID.");
        if (Audiences.Count == 0) errors.Add("Gateway:Audience is required.");
        switch (ClaudeProvider)
        {
            case "Foundry":
                if (string.IsNullOrWhiteSpace(FoundryResource)) errors.Add("Gateway:FoundryResource is required for the Foundry provider.");
                break;
            case "Anthropic":
                if (string.IsNullOrWhiteSpace(AnthropicApiKey)) errors.Add("Gateway:AnthropicApiKey is required for the Anthropic provider.");
                break;
            default:
                errors.Add("Gateway:ClaudeProvider must be Foundry or Anthropic.");
                break;
        }
        if (errors.Count > 0) throw new InvalidOperationException(string.Join(" ", errors));
    }
}
```

- [ ] **Step 5: Write `LLDOCX/LLDOCX.Gateway/DeepCheckPrompt.cs`**

```csharp
using System.Text;
using LLDOCX.DeepCheck;

namespace LLDOCX.Gateway;

public static class DeepCheckPrompt
{
    public const string SystemPrompt =
        "You review the claims of a US patent application for three kinds of problems only:\n" +
        "1. Antecedent basis: every element referred to with 'the' or 'said' must be introduced earlier in the same claim " +
        "or in a claim it depends on (for a multiple dependent claim, in every claim it refers to); references must not be ambiguous.\n" +
        "2. Claim dependency: references to missing, later, or canceled claims; multiple dependent claims that refer to claims " +
        "cumulatively or depend on another multiple dependent claim (37 CFR 1.75(c)); dependent claims whose category or preamble " +
        "does not match the claim they depend on.\n" +
        "3. Terminology: typos and inconsistent names for the same element.\n\n" +
        "An automated rule checker has already reviewed the claims. Its findings are listed with indexes in <rule_findings>. " +
        "Put in `findings` only problems the rule checker missed. Put in `disputes` the index of each rule finding that is a " +
        "false positive, with a one-sentence reason. For every finding, `quote` must be an exact substring of that claim's text, " +
        "as short as possible while unique within the claim. `suggestion` is a concrete fix, or an empty string. " +
        "Do not comment on patentability, claim scope, or style. If there is nothing to report, return empty arrays.";

    public const string ResponseSchemaJson = """
        {
          "type": "object",
          "properties": {
            "findings": {
              "type": "array",
              "items": {
                "type": "object",
                "properties": {
                  "claim": { "type": "integer" },
                  "type": { "type": "string", "enum": ["Antecedent", "Dependency", "Terminology", "Formatting"] },
                  "severity": { "type": "string", "enum": ["Error", "Warning", "Info"] },
                  "quote": { "type": "string" },
                  "message": { "type": "string" },
                  "suggestion": { "type": "string" }
                },
                "required": ["claim", "type", "severity", "quote", "message", "suggestion"],
                "additionalProperties": false
              }
            },
            "disputes": {
              "type": "array",
              "items": {
                "type": "object",
                "properties": {
                  "index": { "type": "integer" },
                  "reason": { "type": "string" }
                },
                "required": ["index", "reason"],
                "additionalProperties": false
              }
            }
          },
          "required": ["findings", "disputes"],
          "additionalProperties": false
        }
        """;

    public static string BuildUserMessage(DeepCheckRequest request)
    {
        var sb = new StringBuilder();
        sb.AppendLine("<claims>");
        foreach (var c in request.Claims)
            sb.AppendLine(c.Canceled ? $"{c.Number}. (Canceled)" : $"{c.Number}. {c.Text}");
        sb.AppendLine("</claims>");
        sb.AppendLine("<rule_findings>");
        for (int i = 0; i < request.RuleFindings.Count; i++)
        {
            var f = request.RuleFindings[i];
            sb.AppendLine($"[{i}] claim {f.Claim}, {f.Type}, {f.Severity}: {f.Message}");
        }
        sb.AppendLine("</rule_findings>");
        return sb.ToString();
    }
}
```

- [ ] **Step 6: Write `LLDOCX/LLDOCX.Gateway/TokenValidator.cs`**

```csharp
using Microsoft.IdentityModel.JsonWebTokens;
using Microsoft.IdentityModel.Protocols;
using Microsoft.IdentityModel.Protocols.OpenIdConnect;
using Microsoft.IdentityModel.Tokens;

namespace LLDOCX.Gateway;

public sealed class CallerIdentity
{
    public CallerIdentity(string objectId, string username)
    {
        ObjectId = objectId;
        Username = username;
    }

    public string ObjectId { get; }
    public string Username { get; }
}

public enum AuthFailureKind { Unauthenticated, Forbidden }

public sealed class AuthFailure : Exception
{
    public AuthFailure(AuthFailureKind kind, string message) : base(message) => Kind = kind;
    public AuthFailureKind Kind { get; }
}

public interface ITokenValidator
{
    /// <summary>Returns the caller, or throws AuthFailure.</summary>
    Task<CallerIdentity> ValidateAsync(string authorizationHeader, CancellationToken ct);
}

/// <summary>Accepts only Entra access tokens issued by the lozaip.com tenant for this gateway, with the DeepCheck.Use scope.</summary>
public sealed class EntraTokenValidator : ITokenValidator
{
    readonly GatewayOptions _options;
    readonly IConfigurationManager<OpenIdConnectConfiguration> _config;
    readonly JsonWebTokenHandler _handler = new();

    public EntraTokenValidator(GatewayOptions options, IConfigurationManager<OpenIdConnectConfiguration> config)
    {
        _options = options;
        _config = config;
    }

    public static IConfigurationManager<OpenIdConnectConfiguration> CreateConfigurationManager(string tenantId) =>
        new ConfigurationManager<OpenIdConnectConfiguration>(
            $"https://login.microsoftonline.com/{tenantId}/v2.0/.well-known/openid-configuration",
            new OpenIdConnectConfigurationRetriever(),
            new HttpDocumentRetriever());

    public async Task<CallerIdentity> ValidateAsync(string authorizationHeader, CancellationToken ct)
    {
        const string prefix = "Bearer ";
        if (string.IsNullOrEmpty(authorizationHeader) || !authorizationHeader.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            throw new AuthFailure(AuthFailureKind.Unauthenticated, "Missing bearer token.");
        string token = authorizationHeader.Substring(prefix.Length).Trim();

        var config = await _config.GetConfigurationAsync(ct);
        var parameters = new TokenValidationParameters
        {
            ValidIssuers = new[]
            {
                $"https://login.microsoftonline.com/{_options.TenantId}/v2.0",
                $"https://sts.windows.net/{_options.TenantId}/"
            },
            ValidAudiences = _options.Audiences,
            IssuerSigningKeys = config.SigningKeys,
            ValidateLifetime = true,
            ClockSkew = TimeSpan.FromMinutes(2)
        };
        TokenValidationResult result = await _handler.ValidateTokenAsync(token, parameters);
        if (!result.IsValid) throw new AuthFailure(AuthFailureKind.Unauthenticated, "Invalid token.");

        var identity = result.ClaimsIdentity;
        if (!string.Equals(identity.FindFirst("tid")?.Value, _options.TenantId, StringComparison.OrdinalIgnoreCase))
            throw new AuthFailure(AuthFailureKind.Unauthenticated, "Token is not from the lozaip.com tenant.");

        var scopes = (identity.FindFirst("scp")?.Value ?? "").Split(' ', StringSplitOptions.RemoveEmptyEntries);
        if (!scopes.Contains(_options.RequiredScope, StringComparer.Ordinal))
            throw new AuthFailure(AuthFailureKind.Forbidden, "Missing scope " + _options.RequiredScope + ".");

        if (!string.IsNullOrWhiteSpace(_options.RequiredRole) && !identity.FindAll("roles").Any(c => c.Value == _options.RequiredRole))
            throw new AuthFailure(AuthFailureKind.Forbidden, "Missing role " + _options.RequiredRole + ".");

        string oid = identity.FindFirst("oid")?.Value;
        if (string.IsNullOrEmpty(oid)) throw new AuthFailure(AuthFailureKind.Unauthenticated, "Token has no oid.");
        string username = identity.FindFirst("preferred_username")?.Value ?? identity.FindFirst("upn")?.Value ?? oid;
        return new CallerIdentity(oid, username);
    }
}
```

- [ ] **Step 7: Write `LLDOCX/LLDOCX.Gateway/UserRateLimiter.cs`**

```csharp
namespace LLDOCX.Gateway;

/// <summary>Sliding one-hour window per user, held in memory per instance (acceptable at lozaip.com's scale).</summary>
public sealed class UserRateLimiter
{
    static readonly TimeSpan Window = TimeSpan.FromHours(1);
    readonly int _perHour;
    readonly Func<DateTime> _utcNow;
    readonly Dictionary<string, Queue<DateTime>> _hits = new();
    readonly object _gate = new();

    public UserRateLimiter(int perHour, Func<DateTime> utcNow)
    {
        _perHour = Math.Max(1, perHour);
        _utcNow = utcNow;
    }

    public bool TryAcquire(string userId, out TimeSpan retryAfter)
    {
        lock (_gate)
        {
            DateTime now = _utcNow();
            if (!_hits.TryGetValue(userId, out var queue)) _hits[userId] = queue = new Queue<DateTime>();
            while (queue.Count > 0 && queue.Peek() <= now - Window) queue.Dequeue();
            if (queue.Count >= _perHour)
            {
                retryAfter = queue.Peek() + Window - now;
                return false;
            }
            queue.Enqueue(now);
            retryAfter = TimeSpan.Zero;
            return true;
        }
    }
}
```

- [ ] **Step 8: Write `LLDOCX/LLDOCX.Gateway/ClaudeReviewer.cs`**

```csharp
using System.Text;
using System.Text.Json;
using Anthropic;
using Anthropic.Exceptions;
using Anthropic.Foundry;
using Anthropic.Models.Messages;
using Azure.Identity;

namespace LLDOCX.Gateway;

public enum ReviewOutcome { Ok, Refused, Truncated }

public sealed class ReviewResult
{
    ReviewResult(ReviewOutcome outcome, string json)
    {
        Outcome = outcome;
        Json = json;
    }

    public ReviewOutcome Outcome { get; }
    public string Json { get; }

    public static ReviewResult Ok(string json) => new(ReviewOutcome.Ok, json);
    public static ReviewResult Refused() => new(ReviewOutcome.Refused, null);
    public static ReviewResult Truncated() => new(ReviewOutcome.Truncated, null);
}

/// <summary>Claude could not be reached or returned an error. Message never contains claim text.</summary>
public sealed class UpstreamException : Exception
{
    public UpstreamException(string message, Exception inner) : base(message, inner) { }
}

public interface IClaudeReviewer
{
    Task<ReviewResult> ReviewAsync(string systemPrompt, string userMessage, string responseSchemaJson, CancellationToken ct);
}

public sealed class SdkClaudeReviewer : IClaudeReviewer
{
    readonly Func<MessageCreateParams, CancellationToken, Task<Message>> _create;
    readonly string _model;
    readonly string _effort;
    readonly int _maxTokens;

    public SdkClaudeReviewer(Func<MessageCreateParams, CancellationToken, Task<Message>> create, string model, string effort, int maxTokens)
    {
        _create = create;
        _model = model;
        _effort = effort;
        _maxTokens = maxTokens;
    }

    public static SdkClaudeReviewer Create(GatewayOptions o)
    {
        if (o.ClaudeProvider == "Anthropic")
        {
            var client = new AnthropicClient { ApiKey = o.AnthropicApiKey };
            return new SdkClaudeReviewer((p, ct) => client.Messages.Create(p, cancellationToken: ct), o.Model, o.Effort, o.MaxOutputTokens);
        }

        // Managed identity when running in Azure; developer credentials (az login) when running locally.
        var foundry = new AnthropicFoundryClient(new AnthropicFoundryIdentityTokenCredentials(new DefaultAzureCredential(), o.FoundryResource));
        return new SdkClaudeReviewer((p, ct) => foundry.Messages.Create(p, cancellationToken: ct), o.Model, o.Effort, o.MaxOutputTokens);
    }

    public async Task<ReviewResult> ReviewAsync(string systemPrompt, string userMessage, string responseSchemaJson, CancellationToken ct)
    {
        var schema = JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(responseSchemaJson);
        var parameters = new MessageCreateParams
        {
            Model = _model,
            MaxTokens = _maxTokens,
            System = systemPrompt,
            OutputConfig = new OutputConfig
            {
                Effort = ParseEffort(_effort),
                Format = new JsonOutputFormat { Schema = schema },
            },
            Messages = [new UserMessageParam { Content = userMessage }],
        };

        Message response;
        try
        {
            response = await _create(parameters, ct);
        }
        catch (AnthropicException ex)
        {
            throw new UpstreamException("Claude request failed (" + ex.GetType().Name + ").", ex);
        }

        if (response.StopReason == "refusal") return ReviewResult.Refused();
        if (response.StopReason == "max_tokens") return ReviewResult.Truncated();

        var text = new StringBuilder();
        foreach (var block in response.Content)
            if (block.TryPickText(out var t)) text.Append(t.Text);
        return ReviewResult.Ok(text.ToString());
    }

    static Effort ParseEffort(string value) => (value ?? "").Trim().ToLowerInvariant() switch
    {
        "low" => Effort.Low,
        "medium" => Effort.Medium,
        "max" => Effort.Max,
        _ => Effort.High,
    };
}
```

- [ ] **Step 9: Write `LLDOCX/LLDOCX.Gateway/DeepCheckHandler.cs`**

```csharp
using System.Diagnostics;
using System.Globalization;
using System.Text.Json;
using LLDOCX.DeepCheck;
using Microsoft.Extensions.Logging;

namespace LLDOCX.Gateway;

public sealed class HandlerResult
{
    public HandlerResult(int status, string body, int? retryAfterSeconds = null)
    {
        Status = status;
        Body = body;
        RetryAfterSeconds = retryAfterSeconds;
    }

    public int Status { get; }
    public string Body { get; }
    public int? RetryAfterSeconds { get; }

    public static HandlerResult Error(int status, string code, string message, int? retryAfterSeconds = null) =>
        new(status, JsonSerializer.Serialize(new ErrorBody { Error = new ErrorDetail { Code = code, Message = message } }), retryAfterSeconds);
}

/// <summary>Everything the HTTP function does, without the HTTP plumbing, so it can be tested.</summary>
public sealed class DeepCheckHandler
{
    readonly GatewayOptions _options;
    readonly ITokenValidator _validator;
    readonly UserRateLimiter _limiter;
    readonly IClaudeReviewer _reviewer;
    readonly ILogger<DeepCheckHandler> _log;

    public DeepCheckHandler(GatewayOptions options, ITokenValidator validator, UserRateLimiter limiter,
        IClaudeReviewer reviewer, ILogger<DeepCheckHandler> log)
    {
        _options = options;
        _validator = validator;
        _limiter = limiter;
        _reviewer = reviewer;
        _log = log;
    }

    public async Task<HandlerResult> HandleAsync(string authorizationHeader, string body, CancellationToken ct)
    {
        var started = Stopwatch.StartNew();

        CallerIdentity caller;
        try
        {
            caller = await _validator.ValidateAsync(authorizationHeader, ct);
        }
        catch (AuthFailure ex)
        {
            _log.LogWarning("Deep Check rejected: {Kind} ({Reason})", ex.Kind, ex.Message);
            return ex.Kind == AuthFailureKind.Forbidden
                ? HandlerResult.Error(403, ErrorCodes.Forbidden, "Your account isn't authorized for Deep Check.")
                : HandlerResult.Error(401, ErrorCodes.Unauthorized, "Sign-in required.");
        }

        DeepCheckRequest request;
        try
        {
            request = JsonSerializer.Deserialize<DeepCheckRequest>(body ?? "");
        }
        catch (JsonException)
        {
            request = null;
        }
        if (request?.Claims == null || request.Claims.Count == 0)
            return Done(caller, started, 0, 0, "invalid_request",
                HandlerResult.Error(400, ErrorCodes.InvalidRequest, "The request contained no claims."));
        request.RuleFindings ??= new List<RuleFindingDto>();

        int count = request.Claims.Count;
        int chars = DeepCheckRequestBuilder.TextLength(request);
        if (chars > _options.MaxInputChars)
            return Done(caller, started, count, chars, "too_large", HandlerResult.Error(400, ErrorCodes.TooLarge,
                string.Format(CultureInfo.InvariantCulture, "The claims are too long for Deep Check (limit {0:N0} characters).", _options.MaxInputChars)));

        if (!_limiter.TryAcquire(caller.ObjectId, out TimeSpan retryAfter))
            return Done(caller, started, count, chars, "rate_limited", HandlerResult.Error(429, ErrorCodes.RateLimited,
                "Deep Check limit reached.", Math.Max(1, (int)Math.Ceiling(retryAfter.TotalSeconds))));

        ReviewResult review;
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(ct);
        timeout.CancelAfter(TimeSpan.FromSeconds(_options.ClaudeTimeoutSeconds));
        try
        {
            review = await _reviewer.ReviewAsync(DeepCheckPrompt.SystemPrompt, DeepCheckPrompt.BuildUserMessage(request),
                DeepCheckPrompt.ResponseSchemaJson, timeout.Token);
        }
        catch (OperationCanceledException) when (!ct.IsCancellationRequested)
        {
            return Done(caller, started, count, chars, "timeout", HandlerResult.Error(504, ErrorCodes.Timeout,
                "Claude took too long to review these claims. Try checking a selection of claims."));
        }
        catch (UpstreamException ex)
        {
            // Log only exception types: inner messages could echo request content.
            _log.LogError("Deep Check upstream failure: {Error} / {Inner}", ex.Message, ex.InnerException?.GetType().Name);
            return Done(caller, started, count, chars, "upstream_unavailable", HandlerResult.Error(503, ErrorCodes.UpstreamUnavailable,
                "The Claude service is unavailable. Try again shortly."));
        }

        if (review.Outcome == ReviewOutcome.Refused)
            return Done(caller, started, count, chars, "refused",
                HandlerResult.Error(422, ErrorCodes.Refused, "Claude declined to review these claims."));
        if (review.Outcome == ReviewOutcome.Truncated)
            return Done(caller, started, count, chars, "truncated", HandlerResult.Error(502, ErrorCodes.Truncated,
                "Claude's review was cut off because the claim set is very large. Try checking a selection of claims."));
        if (!DeepCheckResponseParser.IsWellFormed(review.Json))
            return Done(caller, started, count, chars, "bad_upstream_response", HandlerResult.Error(502, ErrorCodes.BadUpstreamResponse,
                "Claude returned a response that could not be read."));

        return Done(caller, started, count, chars, "ok", new HandlerResult(200, review.Json));
    }

    HandlerResult Done(CallerIdentity caller, Stopwatch started, int claims, int chars, string outcome, HandlerResult result)
    {
        // Metadata only. Claim text and model output are confidential and are never logged.
        _log.LogInformation("Deep Check {Outcome} for {User} ({ObjectId}): {Claims} claims, {Chars} chars, {ElapsedMs} ms",
            outcome, caller.Username, caller.ObjectId, claims, chars, started.ElapsedMilliseconds);
        return result;
    }
}
```

- [ ] **Step 10: Write `LLDOCX/LLDOCX.Gateway/DeepCheckFunction.cs` and `LLDOCX/LLDOCX.Gateway/Program.cs`**

`DeepCheckFunction.cs`:

```csharp
using System.Net;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;

namespace LLDOCX.Gateway;

public sealed class DeepCheckFunction
{
    readonly DeepCheckHandler _handler;

    public DeepCheckFunction(DeepCheckHandler handler) => _handler = handler;

    // Anonymous at the Functions layer: DeepCheckHandler validates the caller's Entra token itself.
    [Function("deepcheck")]
    public async Task<HttpResponseData> Run(
        [HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = "deepcheck")] HttpRequestData req,
        CancellationToken ct)
    {
        string auth = req.Headers.TryGetValues("Authorization", out var values) ? values.FirstOrDefault() : null;
        string body;
        using (var reader = new StreamReader(req.Body))
            body = await reader.ReadToEndAsync(ct);

        HandlerResult result = await _handler.HandleAsync(auth, body, ct);
        var response = req.CreateResponse((HttpStatusCode)result.Status);
        response.Headers.Add("Content-Type", "application/json; charset=utf-8");
        if (result.RetryAfterSeconds.HasValue) response.Headers.Add("Retry-After", result.RetryAfterSeconds.Value.ToString());
        await response.WriteStringAsync(result.Body, ct);
        return response;
    }
}
```

`Program.cs`:

```csharp
using LLDOCX.Gateway;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;

var host = new HostBuilder()
    .ConfigureFunctionsWorkerDefaults()
    .ConfigureServices((context, services) =>
    {
        services.AddApplicationInsightsTelemetryWorkerService();
        services.ConfigureFunctionsApplicationInsights();

        var options = new GatewayOptions();
        context.Configuration.GetSection("Gateway").Bind(options);
        options.Validate();

        services.AddSingleton(options);
        services.AddSingleton<ITokenValidator>(new EntraTokenValidator(options, EntraTokenValidator.CreateConfigurationManager(options.TenantId)));
        services.AddSingleton(new UserRateLimiter(options.RequestsPerUserPerHour, () => DateTime.UtcNow));
        services.AddSingleton<IClaudeReviewer>(SdkClaudeReviewer.Create(options));
        services.AddSingleton<DeepCheckHandler>();
    })
    .Build();

host.Run();
```

- [ ] **Step 11: Build and fix SDK binding names**

Run: `dotnet build LLDOCX.Gateway`
Expected: `Build succeeded.`

These names come from the SDK documentation: `AnthropicFoundryClient`, `AnthropicFoundryIdentityTokenCredentials`, `MessageCreateParams`, `OutputConfig.Effort`, `Effort.High`, `UserMessageParam`, `TryPickText`, `StopReason == "refusal"`, and the `AnthropicException` hierarchy. If the compiler reports an unknown name, look it up in the package XML docs, for example `grep -n "OutputFormat" %USERPROFILE%\.nuget\packages\anthropic\<version>\lib\netstandard2.0\Anthropic.xml`. Likely adjustments:
- **`JsonOutputFormat`** (the non-beta structured-output format class). The beta name is `BetaJsonOutputFormat`.
- **The `cancellationToken:` parameter name** on `Messages.Create`.
- **`StopReason`:** if it is an `ApiEnum`, compare its raw string value.
- **`ILogger.BeginScope` constraint** in `CapturingLogger`: if the compiler warns about the `where TState : notnull` constraint, remove it.

Re-run until `Build succeeded.`

- [ ] **Step 12: Run the tests to verify they pass**

Run: `dotnet test LLDOCX.Gateway.Tests`
Expected: `Passed!` with 0 failed.

Then run all non-VSTO tests:

```bash
dotnet test LLDOCX.Claims.Tests && dotnet test LLDOCX.Auth.Tests && dotnet test LLDOCX.Gateway.Tests
```

Expected: all three pass.

- [ ] **Step 13: Commit**

```bash
git add LLDOCX/LLDOCX.sln LLDOCX/LLDOCX.Gateway LLDOCX/LLDOCX.Gateway.Tests
git commit -m "feat(gateway): add Deep Check Azure Function with Entra auth and Foundry client

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 11: Add-in claims settings

**Files:**
- Modify: `LLDOCX/LLDOCX/LLDOCX.csproj` (project references)
- Modify: `LLDOCX/LLDOCX/Options/SettingOption.cs`
- Modify: `LLDOCX/LLDOCX/CustomUserControl/SettingUserControl.cs:107` (`SaveAllSettng`)

**Interfaces:**
- Consumes: `Severity`, `ClaimsCheckOptions.DefaultAllowlist` (Task 8)
- Produces:
  - `SettingOption` properties: `bool ClaimsCheckAntecedent`, `bool ClaimsCheckDependency`, `bool ClaimsCheckTerminology`, `Severity ClaimsMinSeverity`, `List<string> ClaimsAllowlist`;
  - methods `void CopyClaimsSettingsFrom(SettingOption other)` and `ClaimsCheckOptions ToClaimsCheckOptions()`.
  - There is no model setting and no API key on the client.

- [ ] **Step 1: Add project references to `LLDOCX/LLDOCX/LLDOCX.csproj`**

In the `ItemGroup` that holds `<ProjectReference Include="..\LLDOCX.Auth\LLDOCX.Auth.csproj" />` (added by the gate plan), add:

```xml
    <ProjectReference Include="..\LLDOCX.Claims\LLDOCX.Claims.csproj" />
    <ProjectReference Include="..\LLDOCX.DeepCheck\LLDOCX.DeepCheck.csproj" />
```

- [ ] **Step 2: Replace `LLDOCX/LLDOCX/Options/SettingOption.cs`**

```csharp
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;
using LLDOCX.Claims;

namespace LLDOCX.Options
{
    [Serializable]
   public class SettingOption
    {
        public bool CheckPageSize { get; set; } = true;
        public bool CheckPageOrientation { get; set; } = true;
        public bool CheckMarkup { get; set; } = true;
        public bool CheckFont { get; set; } = true;
        public bool CheckHyperlink { get; set; } = true;
        public bool CheckComment { get; set; } = true;
        public bool CheckBookmarker { get; set; } = true;
        public bool RemoveDocumentProperty { get; set; } = true;
        public bool KeepTrackingOn { get; set; } = true;
        public bool HeadingValidation { get; set; } = true;
        public bool CheckedAll { get; set; } = true;
        public string DefaultFontName { get; set; } = "Times New Roman";

        // Claims Check settings (added in version 2; OptionalField keeps older settings files loadable)
        [field: OptionalField(VersionAdded = 2)]
        public bool ClaimsCheckAntecedent { get; set; } = true;
        [field: OptionalField(VersionAdded = 2)]
        public bool ClaimsCheckDependency { get; set; } = true;
        [field: OptionalField(VersionAdded = 2)]
        public bool ClaimsCheckTerminology { get; set; } = true;
        [field: OptionalField(VersionAdded = 2)]
        public Severity ClaimsMinSeverity { get; set; } = Severity.Info;
        [field: OptionalField(VersionAdded = 2)]
        public List<string> ClaimsAllowlist { get; set; } = new List<string>(ClaimsCheckOptions.DefaultAllowlist);

        // Property initializers do not run during BinaryFormatter deserialization; set defaults
        // here so fields missing from an older settings file get them.
        [OnDeserializing]
        private void SetClaimsDefaults(StreamingContext context)
        {
            ClaimsCheckAntecedent = true;
            ClaimsCheckDependency = true;
            ClaimsCheckTerminology = true;
            ClaimsMinSeverity = Severity.Info;
            ClaimsAllowlist = new List<string>(ClaimsCheckOptions.DefaultAllowlist);
        }

        public void CopyClaimsSettingsFrom(SettingOption other)
        {
            ClaimsCheckAntecedent = other.ClaimsCheckAntecedent;
            ClaimsCheckDependency = other.ClaimsCheckDependency;
            ClaimsCheckTerminology = other.ClaimsCheckTerminology;
            ClaimsMinSeverity = other.ClaimsMinSeverity;
            ClaimsAllowlist = new List<string>(other.ClaimsAllowlist ?? new List<string>(ClaimsCheckOptions.DefaultAllowlist));
        }

        public ClaimsCheckOptions ToClaimsCheckOptions()
        {
            return new ClaimsCheckOptions
            {
                CheckAntecedent = ClaimsCheckAntecedent,
                CheckDependency = ClaimsCheckDependency,
                CheckTerminology = ClaimsCheckTerminology,
                MinSeverity = ClaimsMinSeverity,
                Allowlist = new List<string>(ClaimsAllowlist ?? new List<string>(ClaimsCheckOptions.DefaultAllowlist))
            };
        }
    }
}
```

- [ ] **Step 3: Preserve claims settings when the existing Settings pane saves**

In `LLDOCX/LLDOCX/CustomUserControl/SettingUserControl.cs`, change the start of `SaveAllSettng()` from:

```csharp
        private void SaveAllSettng()
        {
            _option.CheckPageSize = ckbPageSize.Checked;
```

to:

```csharp
        private void SaveAllSettng()
        {
            // The Claims Check pane saves its own settings; don't overwrite them with this pane's stale copy.
            _option.CopyClaimsSettingsFrom(ProcessData.GetData());
            _option.CheckPageSize = ckbPageSize.Checked;
```

- [ ] **Step 4: Verify in Visual Studio (manual)**

1. Rebuild the solution. Expected: the build succeeds, and the add-in's `bin\Debug` contains `LLDOCX.Claims.dll`, `LLDOCX.DeepCheck.dll`, `System.Text.Json.dll` and `LLDOCX.dll.config`. There is no `Anthropic.dll`.
2. Press F5 and sign in.
3. Click **Settings**, toggle any checkbox and click Save. Expected: "Saved Setting successful" appears with no exception, which confirms an existing settings file in `%APPDATA%\LL-Docx\LLDocx.dll` still loads with the new fields.

- [ ] **Step 5: Commit**

```bash
git add LLDOCX/LLDOCX/LLDOCX.csproj LLDOCX/LLDOCX/Options LLDOCX/LLDOCX/CustomUserControl/SettingUserControl.cs
git commit -m "feat(lldocx): add claims check settings

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 12: Word integration (extractor, controller, pane, dialog, ribbon)

**Files:**
- Create: `LLDOCX/LLDOCX/Helpers/ClaimsExtractor.cs`
- Create: `LLDOCX/LLDOCX/Helpers/ClaimsController.cs`
- Create: `LLDOCX/LLDOCX/CustomUserControl/ClaimsPaneControl.cs`
- Create: `LLDOCX/LLDOCX/CustomUserControl/ClaimsSettingsPanel.cs`
- Create: `LLDOCX/LLDOCX/CustomUserControl/DeepCheckConfirmDialog.cs`
- Modify: `LLDOCX/LLDOCX/ThisAddIn.cs`
- Modify: `LLDOCX/LLDOCX/RibbonExplorer.xml`
- Modify: `LLDOCX/LLDOCX/RibbonExplorer.cs`
- Modify: `LLDOCX/LLDOCX/LLDOCX.csproj` (Compile items)

**Interfaces:**
- Consumes:
  - `ClaimsAnalyzer.Analyze`, `AnalysisResult`, `Finding`, `Severity`, `RawParagraph` (Task 8);
  - `DeepCheckHttpClient`, `DeepCheckException` (Task 9);
  - `SettingOption.ToClaimsCheckOptions`, `ProcessData` (Task 11);
  - from the gate plan: `ThisAddIn.Gate`, `ThisAddIn.GateConfig`, `ThisAddIn.EnsureUnlockedAsync()`, `ThisAddIn.HideAllPanes()`, `AuthGate.GetGatewayTokenAsync`, `SignInRequiredException`, `GateStatus`, and `RibbonExplorer.GetEnabled`.
- Produces:
  - `ClaimsExtractor.Extract(Word.Document doc, out bool fromSelection) : List<RawParagraph>`, which returns null when no claims section is found;
  - `ClaimsController`: `Task RunCheckAsync()`, `Task RunDeepCheckAsync()`, `AnalysisResult LastResult`, `string LastDocumentName`;
  - `ThisAddIn.ClaimsController` and `ThisAddIn.ShowClaimsPane() : ClaimsPaneControl`;
  - `ClaimsPaneControl`: `ShowResult(AnalysisResult, bool fromSelection)`, `ShowStatus(string)`, `ShowBusy(string, Action cancel)`, `HideBusy()`.

- [ ] **Step 1: Add the Compile items to `LLDOCX/LLDOCX/LLDOCX.csproj`**

In the `ItemGroup` with the other `Compile` items, add:

```xml
    <Compile Include="Helpers\ClaimsExtractor.cs" />
    <Compile Include="Helpers\ClaimsController.cs" />
    <Compile Include="CustomUserControl\ClaimsPaneControl.cs">
      <SubType>UserControl</SubType>
    </Compile>
    <Compile Include="CustomUserControl\ClaimsSettingsPanel.cs">
      <SubType>UserControl</SubType>
    </Compile>
    <Compile Include="CustomUserControl\DeepCheckConfirmDialog.cs">
      <SubType>Form</SubType>
    </Compile>
```

- [ ] **Step 2: Create `LLDOCX/LLDOCX/Helpers/ClaimsExtractor.cs`**

```csharp
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using LLDOCX.Claims;
using Word = Microsoft.Office.Interop.Word;

namespace LLDOCX.Helpers
{
    /// <summary>Reads the claims section (or the selection) of a Word document as RawParagraphs.</summary>
    static class ClaimsExtractor
    {
        static readonly Regex ClaimsHeading = new Regex(
            @"^\s*(?:claims|what is claimed is|what is claimed|i claim|we claim|the claims are)\s*:?\s*$",
            RegexOptions.IgnoreCase);

        static readonly string[] AbstractHeadings =
        {
            "ABSTRACT", "ABSTRACT OF THE INVENTION", "INVENTION ABSTRACT", "DISCLOSURE ABSTRACT",
            "ABSTRACT OF THE DISCLOSURE", "AMENDMENT TO THE ABSTRACT", "SUBSTITUTE ABSTRACT"
        };

        public static List<RawParagraph> Extract(Word.Document doc, out bool fromSelection)
        {
            fromSelection = false;
            Word.Range scope = null;
            Word.Selection selection = doc.Application.Selection;
            if (selection != null && selection.Range.End > selection.Range.Start)
            {
                scope = selection.Range;
                fromSelection = true;
            }
            else
            {
                scope = FindClaimsSection(doc);
            }
            if (scope == null) return null;

            var list = new List<RawParagraph>();
            foreach (Word.Paragraph p in scope.Paragraphs)
            {
                Word.Range r = p.Range;
                string text = r.Text ?? "";
                string prefix = "";
                var listType = r.ListFormat.ListType;
                if (listType != Word.WdListType.wdListNoNumbering && listType != Word.WdListType.wdListBullet &&
                    listType != Word.WdListType.wdListPictureBullet && !string.IsNullOrEmpty(r.ListFormat.ListString))
                {
                    // Auto-numbered paragraphs: the number is not part of Range.Text. Prepend it and
                    // shift DocStart so offsets into the paragraph text still match the document.
                    prefix = r.ListFormat.ListString + " ";
                }
                list.Add(new RawParagraph(prefix + text, r.Start - prefix.Length));
            }
            return list;
        }

        static Word.Range FindClaimsSection(Word.Document doc)
        {
            var paragraphs = new List<(string Text, int Start, int End)>();
            foreach (Word.Paragraph p in doc.Paragraphs)
            {
                Word.Range r = p.Range;
                paragraphs.Add(((r.Text ?? "").Trim('\r', '\n', '\a', ' ', '\t'), r.Start, r.End));
            }

            var headings = Enumerable.Range(0, paragraphs.Count).Where(i => ClaimsHeading.IsMatch(paragraphs[i].Text)).ToList();
            if (headings.Count == 0) return null;
            var abstracts = Enumerable.Range(0, paragraphs.Count)
                .Where(i => AbstractHeadings.Contains(paragraphs[i].Text.ToUpperInvariant()))
                .ToList();

            // Prefer the last claims heading that is followed by an abstract heading; otherwise the last one.
            var beforeAbstract = headings.Where(h => abstracts.Any(a => a > h)).ToList();
            int heading = beforeAbstract.Count > 0 ? beforeAbstract.Last() : headings.Last();
            if (heading + 1 >= paragraphs.Count) return null;

            int endIndex = abstracts.Where(a => a > heading).DefaultIfEmpty(-1).First();
            int start = paragraphs[heading + 1].Start;
            int end = endIndex >= 0 ? paragraphs[endIndex].Start : paragraphs[paragraphs.Count - 1].End;
            return end > start ? doc.Range(start, end) : null;
        }
    }
}
```

- [ ] **Step 3: Create `LLDOCX/LLDOCX/CustomUserControl/DeepCheckConfirmDialog.cs`**

```csharp
using System.Drawing;
using System.Windows.Forms;

namespace LLDOCX.CustomUserControl
{
    class DeepCheckConfirmDialog : Form
    {
        readonly CheckBox _dontAsk = new CheckBox { Text = "Don't ask again for this document", AutoSize = true };

        public DeepCheckConfirmDialog(int claimCount, int characterCount)
        {
            Text = LLDocxConstant.AppName + " – Deep Check";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MaximizeBox = false;
            MinimizeBox = false;
            AutoSize = true;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;
            Padding = new Padding(12);

            string message =
                $"Deep Check will send the text of {claimCount} claims (about {characterCount:N0} characters) " +
                "to lozaip.com's Deep Check service, which has them reviewed by Claude.\n\n" +
                "Only the claims are sent, not the rest of the document. The document is not changed.\n\n" +
                "Continue?";
            var send = new Button { Text = "Send", DialogResult = DialogResult.OK };
            var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel };

            var layout = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, AutoSize = true, WrapContents = false };
            layout.Controls.Add(new Label { Text = message, AutoSize = true, MaximumSize = new Size(380, 0) });
            layout.Controls.Add(_dontAsk);
            var buttons = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, AutoSize = true, Width = 380 };
            buttons.Controls.Add(cancel);
            buttons.Controls.Add(send);
            layout.Controls.Add(buttons);
            Controls.Add(layout);
            AcceptButton = send;
            CancelButton = cancel;
        }

        public bool DontAskAgain => _dontAsk.Checked;
    }
}
```

- [ ] **Step 4: Create `LLDOCX/LLDOCX/CustomUserControl/ClaimsSettingsPanel.cs`**

```csharp
using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using LLDOCX.Claims;
using LLDOCX.Options;

namespace LLDOCX.CustomUserControl
{
    class ClaimsSettingsPanel : UserControl
    {
        const int ControlWidth = 380;
        readonly CheckBox _antecedent = new CheckBox { Text = "Check antecedent basis", AutoSize = true };
        readonly CheckBox _dependency = new CheckBox { Text = "Check claim dependency", AutoSize = true };
        readonly CheckBox _terminology = new CheckBox { Text = "Check terminology and typos", AutoSize = true };
        readonly ComboBox _minSeverity = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 200 };
        readonly TextBox _allowlist = new TextBox { Multiline = true, ScrollBars = ScrollBars.Vertical, Width = ControlWidth, Height = 120 };

        public ClaimsSettingsPanel()
        {
            _minSeverity.Items.AddRange(new object[] { "Errors", "Errors + Warnings", "All" });

            var save = new Button { Text = "Save", Width = 110, Height = 34 };
            save.Click += OnSave;

            var layout = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoScroll = true,
                Padding = new Padding(8)
            };
            layout.Controls.Add(_antecedent);
            layout.Controls.Add(_dependency);
            layout.Controls.Add(_terminology);
            layout.Controls.Add(new Label { Text = "Show by default:", AutoSize = true, Margin = new Padding(3, 12, 3, 3) });
            layout.Controls.Add(_minSeverity);
            layout.Controls.Add(new Label { Text = "Phrases that need no antecedent (one per line):", AutoSize = true, Margin = new Padding(3, 12, 3, 3) });
            layout.Controls.Add(_allowlist);
            layout.Controls.Add(save);
            Controls.Add(layout);

            Load += (s, e) => LoadSettings();
        }

        void LoadSettings()
        {
            var o = ProcessData.GetData();
            _antecedent.Checked = o.ClaimsCheckAntecedent;
            _dependency.Checked = o.ClaimsCheckDependency;
            _terminology.Checked = o.ClaimsCheckTerminology;
            _minSeverity.SelectedIndex = (int)o.ClaimsMinSeverity;
            _allowlist.Text = string.Join(Environment.NewLine, o.ClaimsAllowlist ?? ClaimsCheckOptions.DefaultAllowlist.ToList());
        }

        void OnSave(object sender, EventArgs e)
        {
            var o = ProcessData.GetData();
            o.ClaimsCheckAntecedent = _antecedent.Checked;
            o.ClaimsCheckDependency = _dependency.Checked;
            o.ClaimsCheckTerminology = _terminology.Checked;
            o.ClaimsMinSeverity = (Severity)Math.Max(0, _minSeverity.SelectedIndex);
            o.ClaimsAllowlist = _allowlist.Lines.Select(l => l.Trim()).Where(l => l.Length > 0).ToList();
            ProcessData.SaveData(o);
            MessageBox.Show("Claims settings saved.", LLDocxConstant.AppName, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
```

- [ ] **Step 5: Create `LLDOCX/LLDOCX/CustomUserControl/ClaimsPaneControl.cs`**

```csharp
using System;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using LLDOCX.Claims;
using LLDOCX.Options;
using Word = Microsoft.Office.Interop.Word;

namespace LLDOCX.CustomUserControl
{
    public class ClaimsPaneControl : UserControl
    {
        readonly TabControl _tabs = new TabControl { Dock = DockStyle.Fill };
        readonly TabPage _findingsTab = new TabPage("Findings");
        readonly TabPage _settingsTab = new TabPage("Settings");
        readonly Label _summary = new Label { Dock = DockStyle.Top, Height = 40, Padding = new Padding(4) };
        readonly ComboBox _severityFilter = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 130 };
        readonly ComboBox _grouping = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 70 };
        readonly ListView _list = new ListView { Dock = DockStyle.Fill, View = View.Details, FullRowSelect = true, HideSelection = false, MultiSelect = false };
        readonly TextBox _details = new TextBox { Dock = DockStyle.Bottom, Height = 90, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical };
        readonly Panel _busyPanel = new Panel { Dock = DockStyle.Top, Height = 34, Visible = false };
        readonly Label _busyLabel = new Label { AutoSize = true, Location = new Point(6, 9) };
        readonly ContextMenuStrip _menu = new ContextMenuStrip();
        readonly ToolStripMenuItem _dismiss = new ToolStripMenuItem("Dismiss AI finding");
        Action _cancelAction;
        AnalysisResult _result;
        bool _fromSelection;

        public ClaimsPaneControl()
        {
            _severityFilter.Items.AddRange(new object[] { "Errors", "Errors + Warnings", "All" });
            _grouping.Items.AddRange(new object[] { "Claim", "Type" });
            _grouping.SelectedIndex = 0;

            var rerun = new Button { Text = "Re-run", AutoSize = true };
            rerun.Click += async (s, e) => await Globals.ThisAddIn.ClaimsController.RunCheckAsync();
            var toolbar = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 34, WrapContents = false };
            toolbar.Controls.Add(new Label { Text = "Show:", AutoSize = true, Margin = new Padding(3, 8, 0, 0) });
            toolbar.Controls.Add(_severityFilter);
            toolbar.Controls.Add(new Label { Text = "Group:", AutoSize = true, Margin = new Padding(6, 8, 0, 0) });
            toolbar.Controls.Add(_grouping);
            toolbar.Controls.Add(rerun);

            var cancel = new Button { Text = "Cancel", AutoSize = true, Location = new Point(250, 4) };
            cancel.Click += (s, e) => _cancelAction?.Invoke();
            _busyPanel.Controls.Add(_busyLabel);
            _busyPanel.Controls.Add(cancel);

            _list.Columns.Add("Severity", 65);
            _list.Columns.Add("Claim", 45);
            _list.Columns.Add("Type", 85);
            _list.Columns.Add("Message", 320);
            _list.Columns.Add("Source", 50);
            _list.SelectedIndexChanged += (s, e) => UpdateDetails();
            _list.MouseClick += (s, e) => { var f = SelectedFinding(); if (f != null) NavigateTo(f); };
            _menu.Items.Add(_dismiss);
            _menu.Opening += (s, e) => { var f = SelectedFinding(); _dismiss.Enabled = f != null && f.Source == "AI"; };
            _dismiss.Click += (s, e) => { var f = SelectedFinding(); if (f != null) { _result.Findings.Remove(f); Populate(); } };
            _list.ContextMenuStrip = _menu;

            _severityFilter.SelectedIndexChanged += (s, e) => Populate();
            _grouping.SelectedIndexChanged += (s, e) => Populate();

            // Dock order: the Fill control is added first; later edge-docked controls take space before it.
            _findingsTab.Controls.Add(_list);
            _findingsTab.Controls.Add(_details);
            _findingsTab.Controls.Add(_busyPanel);
            _findingsTab.Controls.Add(toolbar);
            _findingsTab.Controls.Add(_summary);
            _settingsTab.Controls.Add(new ClaimsSettingsPanel { Dock = DockStyle.Fill });
            _tabs.TabPages.Add(_findingsTab);
            _tabs.TabPages.Add(_settingsTab);
            Controls.Add(_tabs);

            Load += (s, e) => _severityFilter.SelectedIndex = (int)ProcessData.GetData().ClaimsMinSeverity;
        }

        public void ShowStatus(string text) => _summary.Text = text;

        public void ShowBusy(string text, Action cancel)
        {
            _busyLabel.Text = text;
            _cancelAction = cancel;
            _busyPanel.Visible = true;
        }

        public void HideBusy()
        {
            _busyPanel.Visible = false;
            _cancelAction = null;
        }

        public void SelectSettingsTab() => _tabs.SelectedTab = _settingsTab;

        public void ShowResult(AnalysisResult result, bool fromSelection)
        {
            _result = result;
            _fromSelection = fromSelection;
            _tabs.SelectedTab = _findingsTab;
            Populate();
        }

        void Populate()
        {
            _list.BeginUpdate();
            _list.Items.Clear();
            _list.Groups.Clear();
            _details.Text = "";
            if (_result == null)
            {
                _list.EndUpdate();
                return;
            }

            var max = (Severity)Math.Max(0, _severityFilter.SelectedIndex);
            int active = _result.Claims.Count(c => !c.IsCanceled);
            int independent = _result.Claims.Count(c => !c.IsCanceled && c.IsIndependent);
            Func<Severity, int> count = s => _result.Findings.Count(f => f.Severity == s && !f.DisputedByAI);
            _summary.Text = $"{active} claims ({independent} independent){(_fromSelection ? " in selection" : "")}: " +
                            $"{count(Severity.Error)} errors, {count(Severity.Warning)} warnings, {count(Severity.Info)} info";

            bool byClaim = _grouping.SelectedIndex != 1;
            var visible = _result.Findings.Where(f => f.Severity <= max)
                .OrderBy(f => byClaim ? f.ClaimNumber : (int)f.Type)
                .ThenBy(f => f.Severity)
                .ThenBy(f => f.DocStart);
            foreach (var f in visible)
            {
                string groupKey = byClaim ? "Claim " + f.ClaimNumber : f.Type.ToString();
                var group = _list.Groups[groupKey] ?? _list.Groups.Add(groupKey, groupKey);
                string message = f.Message + (string.IsNullOrEmpty(f.Suggestion) ? "" : "  → " + f.Suggestion);
                var item = new ListViewItem(new[] { f.Severity.ToString(), f.ClaimNumber.ToString(), f.Type.ToString(), message, f.Source })
                {
                    Group = group,
                    Tag = f,
                    ForeColor = f.Severity == Severity.Error ? Color.DarkRed
                              : f.Severity == Severity.Warning ? Color.DarkOrange
                              : SystemColors.GrayText
                };
                if (f.DisputedByAI) item.Font = new Font(_list.Font, FontStyle.Strikeout);
                _list.Items.Add(item);
            }
            _list.EndUpdate();
        }

        Finding SelectedFinding() => _list.SelectedItems.Count == 0 ? null : _list.SelectedItems[0].Tag as Finding;

        void UpdateDetails()
        {
            var f = SelectedFinding();
            if (f == null)
            {
                _details.Text = "";
                return;
            }
            string text = $"Claim {f.ClaimNumber} – {f.Type} – {f.Severity} ({f.Source})" + Environment.NewLine + f.Message;
            if (!string.IsNullOrEmpty(f.Suggestion)) text += Environment.NewLine + "Suggestion: " + f.Suggestion;
            if (f.DisputedByAI) text += Environment.NewLine + "Deep Check disputes this: " + f.DisputeReason;
            if (f.DocStart < 0) text += Environment.NewLine + "(This text could not be located in the document.)";
            _details.Text = text;
        }

        void NavigateTo(Finding f)
        {
            if (f.DocStart < 0) return;
            Word.Document doc;
            try { doc = Globals.ThisAddIn.Application.ActiveDocument; }
            catch (COMException) { return; }
            if (!string.Equals(doc.FullName, Globals.ThisAddIn.ClaimsController.LastDocumentName, StringComparison.OrdinalIgnoreCase))
            {
                MessageBox.Show("These findings belong to another document. Switch to that document or click Re-run.",
                    LLDocxConstant.AppName, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            Word.Range range = doc.Range(f.DocStart, f.DocStart + f.Length);
            if (Normalize(range.Text) != Normalize(f.Text))
            {
                var item = _list.SelectedItems[0];
                if (!item.SubItems[3].Text.StartsWith("(stale"))
                    item.SubItems[3].Text = "(stale – re-run) " + item.SubItems[3].Text;
                item.ForeColor = SystemColors.GrayText;
                return;
            }
            range.Select();
        }

        static string Normalize(string s) => Regex.Replace(s ?? "", @"\s+", " ").Trim();
    }
}
```

- [ ] **Step 6: Create `LLDOCX/LLDOCX/Helpers/ClaimsController.cs`**

```csharp
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Net.Http;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using LLDOCX.Auth;
using LLDOCX.Claims;
using LLDOCX.CustomUserControl;
using LLDOCX.DeepCheck;
using LLDOCX.Options;
using Word = Microsoft.Office.Interop.Word;

namespace LLDOCX.Helpers
{
    public class ClaimsController
    {
        static readonly DeepCheckHttpClient Gateway =
            new DeepCheckHttpClient(new HttpClient { Timeout = DeepCheckHttpClient.DefaultTimeout });

        readonly HashSet<string> _deepCheckApproved = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        CancellationTokenSource _deepCheckCts;

        public AnalysisResult LastResult { get; private set; }
        public string LastDocumentName { get; private set; }

        public async Task RunCheckAsync()
        {
            if (!await Globals.ThisAddIn.EnsureUnlockedAsync()) return;
            Word.Document doc = ActiveDocument();
            if (doc == null) return;
            var pane = Globals.ThisAddIn.ShowClaimsPane();

            bool fromSelection;
            List<RawParagraph> paragraphs = ClaimsExtractor.Extract(doc, out fromSelection);
            if (paragraphs == null)
            {
                MessageBox.Show("No claims section was found (for example, a \"CLAIMS\" or \"What is claimed is:\" heading).\n\n" +
                                "Select the claims and click Check Claims again to check the selected text.",
                    LLDocxConstant.AppName, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var options = ProcessData.GetData().ToClaimsCheckOptions();
            pane.ShowStatus("Checking claims…");
            try
            {
                var result = await Task.Run(() => ClaimsAnalyzer.Analyze(paragraphs, options));
                LastResult = result;
                LastDocumentName = doc.FullName;
                pane.ShowResult(result, fromSelection);
            }
            catch (Exception ex)
            {
                Trace.TraceError("LL-DOCX claims check failed: " + ex);
                pane.ShowStatus("");
                MessageBox.Show("Claims check failed: " + ex.Message, LLDocxConstant.AppName, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public async Task RunDeepCheckAsync()
        {
            if (_deepCheckCts != null) return; // already running
            var addIn = Globals.ThisAddIn;
            if (!await addIn.EnsureUnlockedAsync()) return;
            if (!addIn.GateConfig.DeepCheckAvailable)
            {
                MessageBox.Show("Deep Check is not configured. Contact your administrator.", LLDocxConstant.AppName,
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            if (addIn.Gate.State.Status != GateStatus.Unlocked)
            {
                MessageBox.Show("Deep Check needs an internet connection.", LLDocxConstant.AppName,
                    MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            Word.Document doc = ActiveDocument();
            if (doc == null) return;
            if (LastResult == null || !string.Equals(LastDocumentName, doc.FullName, StringComparison.OrdinalIgnoreCase))
            {
                await RunCheckAsync();
                if (LastResult == null || !string.Equals(LastDocumentName, doc.FullName, StringComparison.OrdinalIgnoreCase)) return;
            }

            var pane = addIn.ShowClaimsPane();
            var result = LastResult;
            if (!_deepCheckApproved.Contains(LastDocumentName))
            {
                int chars = result.Claims.Sum(c => c.FullText.Length);
                using (var dlg = new DeepCheckConfirmDialog(result.Claims.Count, chars))
                {
                    if (dlg.ShowDialog() != DialogResult.OK) return;
                    if (dlg.DontAskAgain) _deepCheckApproved.Add(LastDocumentName);
                }
            }

            // Replace any earlier AI results and disputes.
            result.Findings.RemoveAll(f => f.Source == "AI");
            foreach (var f in result.Findings) { f.DisputedByAI = false; f.DisputeReason = null; }
            var ruleFindings = result.Findings.ToList();

            _deepCheckCts = new CancellationTokenSource();
            var ct = _deepCheckCts.Token;
            pane.ShowBusy("Deep Check running…", () => _deepCheckCts?.Cancel());
            try
            {
                string accessToken = await addIn.Gate.GetGatewayTokenAsync(ct);
                var added = await Gateway.RunAsync(addIn.GateConfig.GatewayUrl, accessToken, result.Claims, ruleFindings, ct);
                result.Findings.AddRange(added);
                pane.ShowResult(result, false);
            }
            catch (OperationCanceledException)
            {
                pane.ShowResult(result, false);
                pane.ShowStatus("Deep Check canceled.");
            }
            catch (SignInRequiredException)
            {
                pane.ShowResult(result, false);
                MessageBox.Show("Your sign-in has expired. Click Sign in on the Home tab, then run Deep Check again.",
                    LLDocxConstant.AppName, MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (HttpRequestException)
            {
                pane.ShowResult(result, false);
                MessageBox.Show("Could not reach the sign-in service. Check your network connection.",
                    LLDocxConstant.AppName, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (DeepCheckException ex)
            {
                pane.ShowResult(result, false);
                MessageBox.Show(ex.Message, LLDocxConstant.AppName, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                Trace.TraceError("LL-DOCX Deep Check failed: " + ex);
                pane.ShowResult(result, false);
                MessageBox.Show("Deep Check failed: " + ex.Message, LLDocxConstant.AppName, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                pane.HideBusy();
                _deepCheckCts.Dispose();
                _deepCheckCts = null;
            }
        }

        static Word.Document ActiveDocument()
        {
            try
            {
                return Globals.ThisAddIn.Application.ActiveDocument;
            }
            catch (COMException)
            {
                MessageBox.Show("Open a document first.", LLDocxConstant.AppName, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return null;
            }
        }
    }
}
```

- [ ] **Step 7: Wire the pane and controller into `LLDOCX/LLDOCX/ThisAddIn.cs`**

Make three edits to the file as the gate plan left it:

(a) Replace:

```csharp
        public GateConfig GateConfig { get; private set; }
        public AuthGate Gate { get; private set; }
```

with:

```csharp
        public GateConfig GateConfig { get; private set; }
        public AuthGate Gate { get; private set; }
        public ClaimsController ClaimsController { get; private set; }
        private ClaimsPaneControl _claimsControl;
        private Microsoft.Office.Tools.CustomTaskPane _claimsTaskPane;
```

(b) In `ThisAddIn_Startup`, replace:

```csharp
            DocumentHeadingData = new List<DocumentHeadingData>();
            StartGate();
```

with:

```csharp
            DocumentHeadingData = new List<DocumentHeadingData>();
            ClaimsController = new ClaimsController();
            StartGate();
```

(c) Replace the whole `HideAllPanes` method:

```csharp
        public void HideAllPanes()
        {
            if (myCustomTaskPane != null) myCustomTaskPane.Visible = false;
            _ribbon?.ResetSettingsToggle();
        }
```

with:

```csharp
        public void HideAllPanes()
        {
            if (myCustomTaskPane != null) myCustomTaskPane.Visible = false;
            if (_claimsTaskPane != null) _claimsTaskPane.Visible = false;
            _ribbon?.ResetSettingsToggle();
        }

        public ClaimsPaneControl ShowClaimsPane()
        {
            if (_claimsControl == null)
            {
                _claimsControl = new ClaimsPaneControl();
            }
            if (_claimsTaskPane == null)
            {
                _claimsTaskPane = this.CustomTaskPanes.Add(_claimsControl, "LL-DOCX Claims Check", this.Application.ActiveWindow);
                _claimsTaskPane.Width = 450;
            }
            _claimsTaskPane.Visible = true;
            return _claimsControl;
        }
```

- [ ] **Step 8: Add the ribbon buttons to `LLDOCX/LLDOCX/RibbonExplorer.xml`**

After the `btnPanel` toggleButton line, and before `btnAccount`, add:

```xml
          <button id="btnCheckClaims" label="Check Claims" size="large" onAction="OnCheckClaims" getImage="GetImages" getEnabled="GetEnabled"
                  screentip="Check patent claims"
                  supertip="Checks the claims for antecedent basis, claim dependency, and terminology problems. The document is not changed."/>
          <button id="btnDeepCheck" label="Deep Check" size="large" onAction="OnDeepCheck" getImage="GetImages" getEnabled="GetEnabled"
                  screentip="Deep Check with Claude" getSupertip="GetDeepCheckSupertip"/>
```

- [ ] **Step 9: Update `LLDOCX/LLDOCX/RibbonExplorer.cs`**

Add `using LLDOCX.Auth;` to the usings. Replace `GetImages` and `GetEnabled` with:

```csharp
        public Bitmap GetImages(Office.IRibbonControl control)
        {
            if (control.Id == "btnScan" || control.Id == "btnAccount" || control.Id == "btnCheckClaims" || control.Id == "btnDeepCheck")
                return Properties.Resources.LL_logo;
            if (isPressed)
                return Properties.Resources.toggle_on;
            return Properties.Resources.toggle_off;
        }
        public bool GetEnabled(Office.IRibbonControl control)
        {
            var addIn = Globals.ThisAddIn;
            var gate = addIn.Gate;
            if (gate == null || !gate.State.IsUnlocked) return false;
            if (control.Id == "btnDeepCheck")
                return gate.State.Status == GateStatus.Unlocked && addIn.GateConfig.DeepCheckAvailable;
            return true;
        }
```

Add after `GetEnabled`:

```csharp
        public string GetDeepCheckSupertip(Office.IRibbonControl control)
        {
            var addIn = Globals.ThisAddIn;
            var gate = addIn.Gate;
            if (gate == null) return "";
            if (!gate.State.IsUnlocked) return gate.State.Message;
            if (!addIn.GateConfig.DeepCheckAvailable) return "Deep Check is not configured. Contact your administrator.";
            if (gate.State.Status == GateStatus.UnlockedOffline) return "Deep Check needs an internet connection.";
            return "Sends the claims to lozaip.com's Deep Check service for a second review by Claude. Asks before sending anything.";
        }
        public async void OnCheckClaims(Office.IRibbonControl control)
        {
            try { await Globals.ThisAddIn.ClaimsController.RunCheckAsync(); }
            catch (Exception ex) { System.Windows.Forms.MessageBox.Show("Claims check failed: " + ex.Message, LLDocxConstant.AppName); }
        }
        public async void OnDeepCheck(Office.IRibbonControl control)
        {
            try { await Globals.ThisAddIn.ClaimsController.RunDeepCheckAsync(); }
            catch (Exception ex) { System.Windows.Forms.MessageBox.Show("Deep Check failed: " + ex.Message, LLDocxConstant.AppName); }
        }
```

- [ ] **Step 10: Build and verify in Word (manual)**

Rebuild in Visual Studio. Expected: the build succeeds. Press F5 and sign in with a lozaip.com account (gate plan Task 6), then check each item:

1. **Ribbon buttons.** The Home tab shows **Check Claims** and **Deep Check** in the LL-Docx group. Deep Check is disabled while the gateway isn't configured (`GateDefaults.GatewayUrl` is empty), and hovering explains "Deep Check is not configured."
2. **Basic check.**
   - Create a document with a heading `What is claimed is:`, then these typed claims:
     - `1. A lamp comprising a base, wherein the shade is green.`
     - `2. The lamp of claim 3, wherein the base is round.`
     - `3. The lamp of claim 1, wherein the bulb is bright.`
   - Then add a heading `ABSTRACT` and one paragraph, and click **Check Claims**.
   - Expected: the "LL-DOCX Claims Check" pane opens. Claim 1 shows "No antecedent basis for 'the shade'", claim 2 shows "comes after it", and claim 3 shows "the bulb". The abstract is not checked.
3. **Navigation.** Click a finding: the flagged words are selected in the document.
4. **Stale findings.** Edit the flagged words, then click the same finding: the row shows "(stale – re-run)" and nothing is selected. Click **Re-run**: the findings refresh.
5. **Auto-numbered claims.** Convert the claims to a Word numbered list (Home → Numbering) with the typed numbers removed, and click **Check Claims**. Expected: the same findings, and clicking still selects the correct words.
6. **Selection mode.** Delete the "What is claimed is:" heading and click **Check Claims** with no selection. Expected: the "No claims section was found" message. Select the claims and click again: the findings appear, and the summary says "in selection".
7. **Filter and grouping.** Change "Show:" to "Errors" and "Group:" to "Type": the list filters and regroups.
8. **Settings tab.** Untick "Check antecedent basis", click Save, and click Re-run: the antecedent findings are gone. Re-tick the box and save.
9. **Existing Settings pane.** Click **Settings** in the ribbon and click its Save. Then reopen the Claims Settings tab: your claims settings are unchanged.
10. **Sign-out.** Click the Account button → sign out. The Claims Check pane hides, and Check Claims and Deep Check are disabled.
11. **Scan.** Click **Scan** (signed in): it behaves exactly as before.

Deep Check end-to-end checks happen in Task 14, after the gateway is deployed.

- [ ] **Step 11: Commit**

```bash
git add LLDOCX/LLDOCX
git commit -m "feat(lldocx): add Check Claims and Deep Check to the Word ribbon

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 13: Installer components for the claims checker

**Files:**
- Modify: `LLDOCX/LLDOCXSetup/Product.wxs` (`ProductComponents`, `Version`)

**Interfaces:**
- Consumes: the Release build output of the add-in (Tasks 1–12) and the gate plan's `Product.wxs`
- Produces: an MSI that also installs `LLDOCX.Claims.dll`, `LLDOCX.DeepCheck.dll` and their dependencies

- [ ] **Step 1: List the assemblies the add-in now ships**

In Visual Studio, switch to **Release** and rebuild. Then run:

```bash
ls LLDOCX/bin/Release/*.dll
```

Expected new entries since the gate plan: `LLDOCX.Claims.dll`, `LLDOCX.DeepCheck.dll`, `System.Text.Json.dll`, and its dependencies (for example `System.Text.Encodings.Web.dll`, `System.Memory.dll`, `System.Buffers.dll`, `System.Runtime.CompilerServices.Unsafe.dll`, `System.Threading.Tasks.Extensions.dll`, `System.Numerics.Vectors.dll`, `Microsoft.Bcl.AsyncInterfaces.dll`). Some of these may already be listed from the gate plan. The command's output is authoritative.

- [ ] **Step 2: Add a component for each new file**

Inside `<ComponentGroup Id="ProductComponents" ...>`, add one block per DLL from Step 1 that isn't already listed. In the Ids, replace `.` and `-` with `_`:

```xml
			<Component Id="LLDOCX_Claims_dll_Component" Win64="no">
				<File Id="LLDOCX_Claims_dll" KeyPath="yes" Name="LLDOCX.Claims.dll" Source="$(var.AddinFiles)" />
			</Component>
			<Component Id="LLDOCX_DeepCheck_dll_Component" Win64="no">
				<File Id="LLDOCX_DeepCheck_dll" KeyPath="yes" Name="LLDOCX.DeepCheck.dll" Source="$(var.AddinFiles)" />
			</Component>
			<Component Id="System_Text_Json_dll_Component" Win64="no">
				<File Id="System_Text_Json_dll" KeyPath="yes" Name="System.Text.Json.dll" Source="$(var.AddinFiles)" />
			</Component>
```

Repeat the same pattern for every remaining new DLL. Then change the `Version` attribute of `<Product>` to `1.2.0.0`.

- [ ] **Step 3: Build and test the installer (manual)**

1. Build `LLDOCXSetup` in Release | x64 and sign the MSI (gate plan Task 7, Step 4).
2. Install it over the gate-only version on a test VM. Expected: the upgrade succeeds, and the user stays signed in.
3. Run checklist items 2 and 3 from Task 12. Expected: Check Claims works from the installed add-in.

- [ ] **Step 4: Commit**

```bash
git add LLDOCX/LLDOCXSetup/Product.wxs
git commit -m "build(lldocx): ship claims checker assemblies in the installer

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 14: Deploy the gateway and enable Deep Check

**Files:**
- Create: `LLDOCX/docs/admin/deep-check-setup.md`
- Modify: `LLDOCX/LLDOCX.Auth/GateConfig.cs` (`GateDefaults.GatewayUrl`, `GateDefaults.GatewayScope`)

**Interfaces:**
- Consumes: the gateway (Task 10), the add-in (Task 12), the installer (Task 13), and the gate plan's `docs/admin/sign-in-setup.md`
- Produces:
  - a running gateway at `https://<function app>.azurewebsites.net/api/deepcheck`;
  - an add-in build with the gateway values compiled in;
  - the admin guide.

lozaip.com's Azure administrator runs this task together with the developer, using the guide written in Step 1.

- [ ] **Step 1: Write `LLDOCX/docs/admin/deep-check-setup.md`**

````markdown
# LL-Docx Deep Check: Azure setup (administrators)

Deep Check sends a document's claims, after the user confirms, to an Azure Function owned by lozaip.com. The function checks that
the caller is signed in with a lozaip.com account, then asks Claude for a review. The model, prompt and credentials exist only in
Azure: nothing secret is on users' PCs. Claim text is never logged.

Prerequisites: the add-in registration from `sign-in-setup.md`, an Azure subscription, the Azure CLI (`az login`), and the
.NET 10 SDK.

## 1. Claude in Microsoft Foundry

1. In the Foundry portal (https://ai.azure.com), create or choose a Foundry resource. Note its **resource name**.
2. Deploy **claude-opus-5-5**, choosing **Hosted on Azure** (Global, or Data Zone for US-only processing). Keep the deployment
   name `claude-opus-5-5`, or note your custom name.
3. Make sure Microsoft Entra ID authentication is enabled for the resource.

## 2. Register the gateway API in Entra ID

1. **App registrations** → **New registration**: `LL-Docx Gateway`, single tenant. Note its **Application (client) ID**.
2. **Expose an API** → set the Application ID URI to `api://<gateway client id>`. Then **Add a scope**:
   - Scope name: `DeepCheck.Use`
   - Who can consent: **Admins only**
   - Display name: "Use LL-Docx Deep Check"
3. **Manifest:** set `"requestedAccessTokenVersion": 2` (in older manifests, `"accessTokenAcceptedVersion": 2`). Save.
4. Optional, to choose who may use Deep Check:
   - Under **App roles**, create `DeepCheck.User` (allowed member types: Users/Groups).
   - Then under **Enterprise applications** → **LL-Docx Gateway** → **Users and groups**, assign the security group.
5. Go back to the **LL-Docx Add-in** registration:
   - **API permissions** → **Add** → **APIs my organization uses** → **LL-Docx Gateway** → **DeepCheck.Use** (delegated).
   - Then **Grant admin consent**.

## 3. Create the Function App

Replace the angle-bracket values, then run:

```bash
az group create -n rg-lldocx -l eastus2
az storage account create -n <storageaccountname> -g rg-lldocx -l eastus2 --sku Standard_LRS
az functionapp create -n <functionappname> -g rg-lldocx --storage-account <storageaccountname> \
  --flexconsumption-location eastus2 --runtime dotnet-isolated --runtime-version 10.0
az functionapp update -n <functionappname> -g rg-lldocx --set httpsOnly=true
az functionapp identity assign -n <functionappname> -g rg-lldocx
```

If the CLI rejects runtime version 10.0, list what's available with
`az functionapp list-flexconsumption-runtimes -l eastus2 --runtime dotnet-isolated`, and choose the newest supported .NET LTS. Also
change `TargetFramework` in `LLDOCX.Gateway.csproj` to match.

Give the function's managed identity access to Foundry:

```bash
az role assignment create --assignee <principalId from identity assign> \
  --role "Cognitive Services User" --scope <Foundry resource ID>
```

## 4. Configure the gateway

```bash
az functionapp config appsettings set -n <functionappname> -g rg-lldocx --settings \
  Gateway__TenantId=<tenant id> \
  "Gateway__Audience=api://<gateway client id>,<gateway client id>" \
  Gateway__RequiredScope=DeepCheck.Use \
  Gateway__RequiredRole=<DeepCheck.User, or leave empty> \
  Gateway__ClaudeProvider=Foundry \
  Gateway__FoundryResource=<Foundry resource name> \
  Gateway__Model=<deployment name, default claude-opus-5-5> \
  Gateway__Effort=high \
  Gateway__MaxInputChars=200000 \
  Gateway__RequestsPerUserPerHour=30
```

**To use the Anthropic API instead of Foundry:**
1. Store the key in Key Vault, and give the function's identity the "Key Vault Secrets User" role.
2. Set `Gateway__ClaudeProvider=Anthropic` and `Gateway__AnthropicApiKey=@Microsoft.KeyVault(SecretUri=https://<vault>.vault.azure.net/secrets/<name>/)`.

**To change the model later:** change `Gateway__Model`. No add-in update is needed.

## 5. Deploy

From `Tools/LLDOCX/`:

```bash
dotnet publish LLDOCX.Gateway -c Release -o publish/gateway
cd publish/gateway && zip -r ../gateway.zip . && cd ../..
az functionapp deployment source config-zip -g rg-lldocx -n <functionappname> --src publish/gateway.zip
```

On Windows without `zip`, use PowerShell instead: `Compress-Archive -Path publish/gateway/* -DestinationPath publish/gateway.zip`.

## 6. Tell the developer the add-in values

| Value | Example |
|---|---|
| Gateway URL | `https://<functionappname>.azurewebsites.net/api/deepcheck` |
| Gateway scope | `api://<gateway client id>/DeepCheck.Use` |

## Limits to know

- **Request size:** Azure returns an error if a request runs longer than about 230 seconds. The gateway stops Claude after 210 seconds
  and tells the user to check a selection of claims. If this happens often, set `Gateway__Effort=medium`.
- **Rate limit:** the per-user limit is kept in memory on each instance, so it is approximate when Azure runs several instances.
- **What is logged:** Application Insights records who ran Deep Check, when, the claim count and size, the latency and the outcome. It
  never records claim text or Claude's output.
````

- [ ] **Step 2: Deploy (manual, administrator)**

Follow sections 1–5 of the guide.
Expected:
- `az functionapp function list -g rg-lldocx -n <functionappname> --query "[].name"` lists `deepcheck`.
- In the portal, the Function App's **Log stream** shows the host starting with no `Gateway:` configuration errors.

- [ ] **Step 3: Smoke-test the gateway without Word (manual, optional)**

1. **Allow the Azure CLI to request a gateway token** (temporarily). In the gateway's **Expose an API** page, add the client application `04b07795-8ddb-461a-bbee-02f9e1bf7b46`, which is the Azure CLI, for the `DeepCheck.Use` scope.
2. **Send a test request:**

```bash
TOKEN=$(az account get-access-token --scope "api://<gateway client id>/DeepCheck.Use" --query accessToken -o tsv)
curl -s -X POST "https://<functionappname>.azurewebsites.net/api/deepcheck" \
  -H "Authorization: Bearer $TOKEN" -H "Content-Type: application/json" \
  -d '{"claims":[{"number":1,"text":"A lamp comprising a base, wherein the shade is green.","canceled":false}],"ruleFindings":[]}'
```

Expected: a JSON body with `findings` and `disputes` arrays.

3. **Check the rejection without a token:**

```bash
curl -s -o /dev/null -w "%{http_code}\n" -X POST "https://<functionappname>.azurewebsites.net/api/deepcheck" -d '{}'
```

Expected: `401`.

4. **Remove the Azure CLI client authorization** again.

- [ ] **Step 4: Compile the gateway values into the add-in**

In `LLDOCX/LLDOCX.Auth/GateConfig.cs`, class `GateDefaults`, set the two constants to the real values from the guide's section 6:

```csharp
        public const string GatewayUrl = "https://<functionappname>.azurewebsites.net/api/deepcheck";
        public const string GatewayScope = "api://<gateway client id>/DeepCheck.Use";
```

Run: `dotnet test LLDOCX.Auth.Tests`
Expected: all pass; the tests don't depend on these constants.

- [ ] **Step 5: Verify Deep Check end-to-end in Word (manual)**

Rebuild the add-in and press F5. Sign out and back in, so the new gateway scope is consented and cached. Then check:

1. **Enabled when online.** Deep Check is enabled. Hovering shows "Sends the claims to lozaip.com's Deep Check service…"
2. **Confirmation and results.**
   - On the Task 12 test document, click **Deep Check**. The confirmation dialog shows the claim count.
   - Click Send: "Deep Check running…" appears with Cancel.
   - When it finishes, any AI findings show Source "AI", and any disputed rule findings are struck through.
   - Right-click an AI finding → **Dismiss AI finding**: it disappears.
3. **"Don't ask again".** Run Deep Check again with "Don't ask again" ticked, then once more. The second time there is no dialog.
4. **Unauthorized user.** With `Gateway__RequiredRole=DeepCheck.User` set and a test user not in the group, click Deep Check. Expected: "Your account isn't authorized for Deep Check. Contact your administrator." The rule findings stay.
5. **Rate limit.** Set `Gateway__RequestsPerUserPerHour=1` and run Deep Check twice. The second run shows "Deep Check limit reached; try again in N minutes." Restore the setting.
6. **Offline.** Disconnect the network. Deep Check becomes disabled on the next ribbon refresh, with the tooltip "Deep Check needs an internet connection." If you click it during the transition, "Could not reach the Deep Check service…" appears and the findings remain.
7. **Logging.** In Application Insights → Logs, run `traces | where message startswith "Deep Check" | take 20`. Each entry shows the outcome, user, claim count, size and milliseconds, and no claim text.

- [ ] **Step 6: Ship the release**

1. Rebuild the add-in and installer in Release | x64.
2. Sign the build (gate plan Task 7, Step 4), bump the MSI `Version`, and distribute it.

- [ ] **Step 7: Commit**

```bash
git add LLDOCX/docs/admin/deep-check-setup.md LLDOCX/LLDOCX.Auth/GateConfig.cs
git commit -m "docs(lldocx): add Deep Check Azure setup guide and enable the gateway

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

## Spec updates made with this plan

These changes were made to the spec while writing this plan, and the spec file has been updated to match:

- The engine's input type is `RawParagraph`, one per document paragraph. `ClaimsAnalyzer.Analyze` returns an `AnalysisResult` that holds both the parsed claims and the findings, so the pane can show claim counts.
- Introducing determiners: `each`, `any`, `one`, and bare plurals were dropped to cut false positives. Phrases like "The method of claim 1" are skipped automatically, so the allowlist doesn't need "the method of claim" entries, and "the group consisting of" becomes "the group".
- Typo rule: edit distance 1 for words of 5–7 letters and 2 for 8+ letters; inflections are excluded instead of being flagged.
- Singular/plural drift is detected during the antecedent pass and controlled by the terminology toggle.
- The Claims settings live on a Settings tab inside the Claims Check pane rather than in the fixed-layout existing Settings pane.
- **Revised 2026-11-01 (Azure gateway):**
  - **Framework.** Test projects and the gateway target **.NET 10** (LTS), because .NET 8 reaches end of support on 2026-11-10.
  - **Wire contract.** The response body reuses Claude's structured-output schema (`claim`, `quote` and so on) instead of new field names. Error bodies are `{ error: { code, message } }`, with codes listed in `ErrorCodes`.
  - **Gateway settings.** The gateway setting is `FoundryResource` (the resource name), not an endpoint URL, because the C# Foundry client builds the URL from the resource name.
  - **No server-side fallback.** Both providers use the standard (non-beta) Messages call with no server-side refusal fallback, which Foundry doesn't support. A refusal returns 422 with the same message for either provider.
  - **Timeouts.** Azure stops HTTP requests after about 230 seconds, so the gateway stops the Claude call after 210 seconds (504 `timeout`), and the add-in's HTTP timeout is 240 seconds, not 5 minutes.
  - **Gateway tests.** These live in a separate `LLDOCX.Gateway.Tests` project, so the claims engine tests don't depend on the Azure and Anthropic packages.
  - **Source of the old Task 1.** Tracking the source and retargeting to 4.7.2 is now Task 1 of the sign-in gate plan.
