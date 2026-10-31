# LL-Docx Sign-In Gate — Design

**Date:** 2026-10-31
**Status:** Approved design, pending implementation plan
**Component:** LLDOCX Word add-in (VSTO, C#)
**Built before:** the Patent Claims Checker (`2026-10-30-patent-claims-checker-design.md`), which depends on it

## Purpose

Only people with a lozaip.com Microsoft 365 account may use the LL-Docx tools. The gate must work in every edition of Word, including standalone (non-subscription) Office, where Word may not be signed in to any account.

## Decisions

| Topic | Decision |
|---|---|
| Identity provider | Microsoft Entra ID, the lozaip.com Microsoft 365 tenant |
| When the user signs in | On first use, inside the add-in. Never at install time |
| Dependence on Office sign-in | None. The add-in signs in on its own, so standalone Office works |
| What is stored | A renewable token (MSAL token cache), encrypted per Windows user. Passwords are never seen or stored |
| Offline use | Allowed for 30 days after the last successful online check. Admins can lower this |
| Scope | Gates every LL-Docx ribbon button, existing and new |
| Configuration | Admin-controlled file plus an optional Group Policy registry override. Contains no secrets |

## How it works

1. When Word starts, the add-in loads its admin configuration. It then tries to renew the user's token silently on a background thread, so Word's startup is not delayed.
2. If no account is cached or renewal needs the user, the tools stay locked and the ribbon shows **Sign in**.
3. **Sign in** opens the Microsoft sign-in window. That is the Windows account broker (WAM) where available, and the system browser otherwise. The window is restricted to the lozaip.com tenant and enforces whatever MFA and Conditional Access policies the tenant has.
4. On success, the add-in checks that the account's tenant ID equals the configured tenant ID, records the time of the check, and unlocks the tools.
5. Later Word sessions renew silently. Users only see the sign-in window again when Entra requires it, for example after a password reset, a revoked session or expired tokens.

## Access policy

`AccessPolicy.Evaluate(...)` is a pure function. Its inputs are the result of the silent renewal, the tenant ID of the cached account, the configured tenant ID, the stored last-verified time, the current time and the grace period. It returns a `GateState`.

| Renewal result | Condition | State | Ribbon message |
|---|---|---|---|
| Success | Tenant matches | **Unlocked** (last-verified time updated) | Signed-in account's UPN |
| Success | Tenant differs | **Locked** (account signed out) | "This account is not a lozaip.com account." |
| Network unavailable | Last verified ≤ grace period ago | **Unlocked (offline)** | "Offline — verified N days ago" |
| Network unavailable | Last verified > grace period ago, none recorded, or in the future | **Locked** | "Connect to the internet to verify your account." |
| Interaction required (expired, revoked, password changed, account disabled) | — | **Locked** | "Sign in with your lozaip.com account." |
| No cached account | — | **Locked** | "Sign in with your lozaip.com account." |
| Configuration missing or invalid | — | **Locked** | "LL-Docx is not configured. Contact your administrator." |

- **Clock rollback:** a last-verified time later than the current time is treated as expired.
- **Re-checks:** the gate re-checks silently whenever a gated button is clicked and the last online check is more than 24 hours old, so a Word session left open for weeks still notices a disabled account. Deep Check always requests a fresh token, so it fails immediately for a disabled account whatever the gate state.
- **When the state changes to Locked:** the ribbon is invalidated, every gated button is disabled, and open LL-Docx task panes are hidden.

## Configuration

**File:** `%ProgramData%\LL-Docx\lldocx.config.json`. The installer writes it from MSI properties. Users can read it; only administrators can change it.

```json
{
  "tenantId": "00000000-0000-0000-0000-000000000000",
  "clientId": "00000000-0000-0000-0000-000000000000",
  "gatewayUrl": "https://lldocx-gateway.azurewebsites.net/api/deepcheck",
  "gatewayScope": "api://00000000-0000-0000-0000-000000000000/DeepCheck.Use",
  "offlineGraceDays": 30
}
```

| Field | Meaning |
|---|---|
| `tenantId` | The lozaip.com Entra tenant ID. Used for the sign-in authority and the tenant check |
| `clientId` | The add-in's Entra app registration (a public client) |
| `gatewayUrl` | The Deep Check gateway endpoint (see the claims checker spec) |
| `gatewayScope` | The gateway API scope. It is requested at sign-in, so a single consent covers Deep Check |
| `offlineGraceDays` | 1–30. Values outside that range are clamped |

- **Group Policy override:** values under `HKLM\SOFTWARE\Policies\LL-Docx` with the same names (`TenantId`, `ClientId`, `GatewayUrl`, `GatewayScope`, `OfflineGraceDays`) override the file, one field at a time.
- **Validation:** missing or invalid `tenantId` or `clientId` (they must be GUIDs) makes the configuration invalid, and the tools lock. A missing gateway setting only disables Deep Check.
- **No secrets:** nothing in the configuration is secret. A user who reads it learns IDs that are useless without a lozaip.com sign-in.

## Components

| Component | Project | Responsibility |
|---|---|---|
| `GateConfig`, `GateConfigLoader` | LLDOCX.Auth (new, netstandard2.0) | Read the file, apply registry overrides, validate. File and registry access go through small interfaces so they can be tested |
| `AccessPolicy`, `GateState`, `RenewalResult` | LLDOCX.Auth | The pure decision table above |
| `LastVerifiedStore` | LLDOCX.Auth | Stores the last-verified time and account ID in `%LOCALAPPDATA%\LL-Docx\gate.dat`, encrypted with DPAPI (CurrentUser) |
| `ITokenSource` | LLDOCX.Auth | `Task<RenewalOutcome> RenewSilentAsync(ct)`, `Task<RenewalOutcome> SignInInteractiveAsync(IntPtr parentHwnd, ct)`, `Task<string> GetGatewayTokenAsync(ct)`, `Task SignOutAsync()` |
| `AuthGate` | LLDOCX.Auth | Combines the token source, policy and store. Exposes `State` and a `StateChanged` event. Implements the 24-hour re-check |
| `MsalTokenSource` | LLDOCX add-in | The `ITokenSource` implementation: MSAL.NET `PublicClientApplication` with authority `https://login.microsoftonline.com/{tenantId}`, the WAM broker (`Microsoft.Identity.Client.Broker`) and a system-browser fallback. The token cache is persisted with `Microsoft.Identity.Client.Extensions.Msal` (DPAPI) at `%LOCALAPPDATA%\LL-Docx\msal.cache` |
| Ribbon changes | LLDOCX add-in | A `getEnabled` callback on all LL-Docx buttons. An **Account** button that shows Sign in or the signed-in UPN, with Sign out |
| `ThisAddIn` changes | LLDOCX add-in | Creates the gate at startup, starts the silent check, invalidates the ribbon and hides panes on state changes |

- **Classifying errors:** `MsalUiRequiredException` means interaction is required. Network failures (`HttpRequestException`, or `MsalServiceException`/`MsalClientException` caused by connectivity) mean the network is unavailable. Other exceptions are logged and treated as interaction required.
- **Signing out** removes the account from MSAL, deletes `gate.dat` and locks the tools.

## Admin setup (documented as a checklist in the plan)

1. **Register the add-in in Entra ID**
   - Name it "LL-Docx Add-in".
   - Make it single-tenant, as a public client.
   - Redirect URIs: `ms-appx-web://microsoft.aad.brokerplugin/{clientId}` and `http://localhost`.
2. **Grant it access to the gateway API.** Give it the delegated permission `DeepCheck.Use` on the gateway API, and grant admin consent so users aren't prompted for consent.
3. **Optionally restrict who can use LL-Docx.** On the enterprise application, set "Assignment required" and assign a security group. Then only members of that group can sign in, which lets you control access per person.
4. **Install with the configuration values:**
   ```
   msiexec /i LLDOCXSetup.msi TENANTID=… CLIENTID=… GATEWAYURL=… GATEWAYSCOPE=…
   ```
   Alternatively, push the registry values through Group Policy.

## Error handling

| Situation | Behavior |
|---|---|
| The user cancels the sign-in window | Stay locked; no message |
| Sign-in fails (wrong tenant, blocked by Conditional Access, not assigned) | Stay locked; show Entra's message |
| `gate.dat` is unreadable or corrupt | Treated as never verified (online renewal still works) |
| MSAL cache is unreadable | Treated as no cached account |
| An unexpected exception in the gate | Log it and stay locked. Never unlock on an error |

## Limits

- The gate is enforced inside the add-in. Someone who modifies the add-in's DLLs could bypass it for the offline tools, which hold no secrets. The 30-day offline timestamp can likewise be tampered with by a determined local user.
- Deep Check is protected on the server: the gateway accepts only valid lozaip.com tokens, so a bypassed client gains nothing there.

## Testing

**LLDOCX.Auth.Tests** (MSTest, net8.0), covering:
- the `AccessPolicy` table, every row, including the boundaries (exactly 30 days, 30 days plus one second), clock rollback and the clamping of `offlineGraceDays`;
- `GateConfigLoader`: a missing file, malformed JSON, invalid GUIDs, a registry override for each field, and the clamping;
- `AuthGate` with a fake `ITokenSource` and fake clock: state transitions, the `StateChanged` event, the 24-hour re-check, sign-out clearing state, and an exception leaving it locked;
- `LastVerifiedStore` round trip and corruption handling.

**Manual checklist** (Visual Studio, Word):
- standalone Office not signed in to any account;
- Microsoft 365 Office signed in to a different account;
- a non-lozaip account rejected;
- offline within and beyond the grace period (set `OfflineGraceDays` to 1 through policy and roll back `gate.dat`);
- sign-out;
- a disabled test account locking at the next re-check;
- the missing-configuration message.

## Out of scope

- Identity providers other than Entra ID
- Per-seat licensing or usage billing
- Word for Mac and Word on the web
