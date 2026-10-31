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
| Configuration | lozaip.com values compiled into the signed add-in. The tenant ID can't be overridden; other values can be overridden only through admin-only Group Policy registry values. Contains no secrets |
| Installation | One per-machine MSI with no sign-in at install time. Distributed by any means (file share, email, Group Policy, scripts). No dependency on Intune or any device-management tool |

## How it works

1. When Word starts, the add-in loads its compiled configuration and any Group Policy overrides. It then tries to renew the user's token silently on a background thread, so Word's startup is not delayed.
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

The values are compiled into the add-in as constants in `GateDefaults`. The add-in assembly is code-signed, so they can't be changed without breaking the signature.

| Setting | Compiled default | Group Policy override (`HKLM\SOFTWARE\Policies\LL-Docx`) |
|---|---|---|
| `TenantId` | The lozaip.com Entra tenant ID | **Not overridable.** This stops the add-in from being pointed at another tenant |
| `ClientId` | The add-in's Entra app registration (a public client) | `ClientId` (REG_SZ, must be a GUID) |
| `GatewayUrl` | The Deep Check gateway endpoint (see the claims checker spec) | `GatewayUrl` (REG_SZ, must be `https://`) |
| `GatewayScope` | `api://<gateway app ID>/DeepCheck.Use`. Requested at sign-in, so a single consent covers Deep Check | `GatewayScope` (REG_SZ) |
| `OfflineGraceDays` | 30 | `OfflineGraceDays` (REG_DWORD). It can only **lower** the value, and is clamped to 1–30 |

- **Invalid overrides** are ignored and logged, and the compiled default is used.
- **Changing the tenant** would need a rebuild, which is acceptable because lozaip.com's tenant doesn't change.
- **The IDs and the tenant ID are filled in** during the build from the admin setup values (see below). Until real values exist, the build uses placeholder GUIDs, and the gate shows "LL-Docx is not configured".
- **No secrets:** nothing here is secret. Someone who decompiles the add-in learns IDs that are useless without a lozaip.com sign-in.

## Installation and deployment

The same installer serves Microsoft 365 and standalone Office. The gate is independent of how Office is licensed, and installing requires no sign-in.

- **One per-machine MSI** (WiX, `InstallScope="perMachine"`, needs administrator rights). Installing is as simple as double-clicking it. It can also be installed silently (`msiexec /i LLDOCXSetup.msi /qn`) through Group Policy software installation, a login script or any deployment tool. No tool is required.
- **Registration for all users.** The add-in is registered under `HKLM\Software\Microsoft\Office\Word\Addins\LLDOCX`, so every user of the PC gets it. On 64-bit Windows it is registered in both the 64-bit and the `WOW6432Node` views, covering 64-bit and 32-bit Office. This fixes the current installer, which writes the registration to `HKCU` from a per-machine package, so only the user who ran it sees the add-in.
- **Prerequisite checks:** the installer refuses to install, with a clear message, when any of these is missing:
  - .NET Framework 4.7.2 or later;
  - the VSTO 2010 runtime (`HKLM\SOFTWARE\Microsoft\VSTO Runtime Setup\v4R`, either view);
  - Word 2016 or later.
- **Code signing:** the add-in manifests and assemblies are signed with lozaip.com's production code-signing certificate, as is the MSI. The temporary development certificate (`LLDOCX_TemporaryKey.pfx`) is never used for a release and is never committed.
- **Updates:** run a newer MSI. The existing `MajorUpgrade` element replaces the old version in place. Sign-in state (in the user profile) survives upgrades.
- **Uninstalling** removes the files and the registration. Per-user sign-in data in `%LOCALAPPDATA%\LL-Docx` is left in place, and Sign out clears it.
- **What users see:**
  - Microsoft 365 users who are signed in to Windows or Office with their lozaip.com account usually finish the first sign-in silently or with one click, through the Windows account broker.
  - Standalone Office users see the Microsoft sign-in window once.
  - After that, both renew silently.

## Components

| Component | Project | Responsibility |
|---|---|---|
| `GateDefaults`, `GateConfig`, `GateConfigLoader` | LLDOCX.Auth (new, netstandard2.0) | Hold the compiled defaults, apply the allowed policy overrides and validate them. Registry access goes through a small interface so it can be tested |
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
4. **Give the developer the tenant ID, the add-in client ID and the gateway values** to compile into `GateDefaults`.
5. **Obtain a code-signing certificate** for lozaip.com, used to sign the release build and the MSI.
6. **Distribute the signed MSI** by any means. Group Policy overrides are optional.

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
- `GateConfigLoader`:
  - compiled defaults only;
  - each allowed override;
  - an attempted `TenantId` override being ignored;
  - invalid values being ignored;
  - `OfflineGraceDays` able only to lower the value;
  - placeholder IDs reported as "not configured";
- `AuthGate` with a fake `ITokenSource` and fake clock: state transitions, the `StateChanged` event, the 24-hour re-check, sign-out clearing state, and an exception leaving it locked;
- `LastVerifiedStore` round trip and corruption handling.

**Manual checklist** (Visual Studio, Word):
- standalone Office not signed in to any account;
- Microsoft 365 Office signed in to a different account;
- a non-lozaip account rejected;
- offline within and beyond the grace period (set `OfflineGraceDays` to 1 through policy and roll back `gate.dat`);
- sign-out;
- a disabled test account locking at the next re-check;
- the missing-configuration message;
- installer, on a clean VM with standalone Office 2021 and on one with Microsoft 365 Apps, with both 32-bit and 64-bit Office:
  - the add-in is visible to a second Windows user;
  - the prerequisite failure messages appear;
  - a major upgrade keeps the user signed in;
  - uninstalling works.

## Out of scope

- Identity providers other than Entra ID
- Per-seat licensing or usage billing
- Word for Mac and Word on the web
- Automatic deployment (Intune, Microsoft 365 Centralized Deployment) and a built-in updater
- A sign-in-protected download page
