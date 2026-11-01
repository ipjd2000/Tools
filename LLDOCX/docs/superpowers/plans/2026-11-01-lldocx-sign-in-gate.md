# LL-Docx Sign-In Gate Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Lock every LL-Docx ribbon button until the user has signed in with a lozaip.com Microsoft 365 (Entra ID) account. This works in standalone and subscription Office, with 30 days of offline use. The add-in ships as one per-machine MSI.

**Architecture:**
- **`LLDOCX.Auth`** (netstandard2.0, no Office or MSAL references) holds the pure parts: compiled configuration with Group Policy overrides, the access-policy decision table, the encrypted last-verified store, and the `AuthGate` state machine. It talks to an `ITokenSource` interface.
- **The VSTO add-in** implements `ITokenSource` with MSAL.NET and the Windows broker (WAM). It wires the gate into `ThisAddIn` and the ribbon (`getEnabled` and an Account button).
- **The WiX installer** becomes a per-machine x64 MSI that registers the add-in under HKLM in both registry views and checks prerequisites.

**Tech Stack:** C# / .NET Framework 4.7.2 VSTO, netstandard2.0, MSTest on net10.0, MSAL.NET (`Microsoft.Identity.Client`, `.Broker`, `.Extensions.Msal`), DPAPI (`System.Security.Cryptography.ProtectedData`), WiX 3.

**Spec:** `LLDOCX/docs/superpowers/specs/2026-10-31-lldocx-sign-in-gate-design.md`

**Followed by:** `LLDOCX/docs/superpowers/plans/2026-10-30-patent-claims-checker.md`. Its Task 1 is replaced by Task 1 of this plan.

## Global Constraints

- **Paths and shell:** all paths are relative to the repository root `Tools/`. Shell commands run from `Tools/LLDOCX/` unless a step says otherwise.
- **Branch and commits:**
  - Work on the existing branch `feature/patent-claims-checker`. The gate and the claims checker ship together.
  - Commit after every task. Each commit message ends with `Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>`.
- **Target frameworks:**
  - The add-in targets **.NET Framework 4.7.2**.
  - `LLDOCX.Auth` targets **netstandard2.0** with `LangVersion` 7.3, and must not reference Office, WinForms or MSAL.
  - Test projects target **net10.0** (.NET 10 LTS; .NET 8 reaches end of support on 2026-11-10).
- **Identity:** the identity provider is Microsoft Entra ID, tenant lozaip.com. The sign-in authority is `https://login.microsoftonline.com/{TenantId}`.
- **Never stored:** passwords are never seen or stored. Only the MSAL token cache (DPAPI, CurrentUser) at `%LOCALAPPDATA%\LL-Docx\msal.cache` and the last-verified record at `%LOCALAPPDATA%\LL-Docx\gate.dat` (DPAPI, CurrentUser).
- **Offline use** is allowed for at most **30 days** after the last successful online check. Group Policy `OfflineGraceDays` can only lower the value (clamped to 1–30).
- **`TenantId` is compiled in** and can never be overridden. `ClientId`, `GatewayUrl`, `GatewayScope` and `OfflineGraceDays` can be overridden only from `HKLM\SOFTWARE\Policies\LL-Docx`.
- **Never unlock on an error:** any unexpected exception leaves the tools locked.
- **Re-checks:** the gate re-checks silently when a gated button is clicked and the last online check is more than 24 hours old.
- **Signing key:** `LLDOCX_TemporaryKey.pfx` is never committed and never used for a release build.
- **Installer:** one per-machine MSI, no sign-in at install time, no dependency on Intune or any device-management tool.

## Build and test environment

- **`LLDOCX.Auth` and its tests (Tasks 2–5):** need the .NET 10 SDK, which Task 2 installs. They run with `dotnet test`.
- **VSTO add-in and installer (Tasks 1, 6, 7):** need Visual Studio 2022 with the **Office/SharePoint development** workload, the **.NET Framework 4.7.2 targeting pack**, and the **WiX Toolset 3 extension**. This machine has only VS Build Tools, so a human builds and verifies these tasks in Visual Studio. Each ends with a manual checklist.
- **Real sign-in testing (Task 6):** needs the Entra app registration from Task 7's admin guide. At minimum, a tenant ID and a client ID must exist. Until they do, the add-in correctly shows "LL-Docx is not configured".

## File Structure

```
LLDOCX/
  .gitignore                                   (new) ignore .vs/, bin/, obj/, *.user, *.pfx
  LLDOCX.sln                                   (modify) add LLDOCX.Auth + LLDOCX.Auth.Tests
  LLDOCX.Auth/                                 (new) netstandard2.0, pure gate logic
    LLDOCX.Auth.csproj
    GateState.cs           GateStatus, LockReason, GateState (messages)
    RenewalOutcome.cs      RenewalStatus, RenewalOutcome, VerifiedRecord
    AccessPolicy.cs        the decision table
    GateConfig.cs          GateConfig, GateDefaults (compiled values), IPolicySource, GateConfigLoader
    LastVerifiedStore.cs   ILastVerifiedStore, DpapiLastVerifiedStore
    AuthGate.cs            ITokenSource, IClock, SystemClock, SignInRequiredException, AuthGate
  LLDOCX.Auth.Tests/                           (new) MSTest, net10.0
    LLDOCX.Auth.Tests.csproj
    Fakes.cs               FakeTokenSource, FakeClock, MemoryStore, DictionaryPolicySource
    AccessPolicyTests.cs, GateConfigLoaderTests.cs, LastVerifiedStoreTests.cs, AuthGateTests.cs
  LLDOCX/                                      (existing VSTO add-in)
    LLDOCX.csproj                              (modify) 4.7.2, LangVersion, MSAL packages, Auth reference
    Auth/MsalTokenSource.cs                    (new) ITokenSource via MSAL + WAM
    Auth/RegistryPolicySource.cs               (new) IPolicySource over HKLM\SOFTWARE\Policies\LL-Docx
    ThisAddIn.cs                               (modify) gate startup, state handling, pane null-check fix
    RibbonExplorer.xml / RibbonExplorer.cs     (modify) getEnabled on all buttons, Account button
  LLDOCXSetup/
    LLDOCXSetup.wixproj                        (modify) x64 platform
    Product.wxs                                (modify) HKLM registration (both views), prerequisites, new DLLs
  docs/admin/sign-in-setup.md                  (new) Entra registration, values to compile in, signing, distribution
```

---

### Task 1: Track the LLDOCX source and retarget to .NET Framework 4.7.2

**Files:**
- Create: `LLDOCX/.gitignore`
- Modify: `LLDOCX/LLDOCX/LLDOCX.csproj` (line 30: TargetFrameworkVersion, plus LangVersion and binding redirects)

**Interfaces:**
- Consumes: nothing
- Produces: the add-in source is tracked in git, and the add-in builds against .NET Framework 4.7.2 with `LangVersion` latest and PackageReference restore. Later tasks add references to this csproj.

- [ ] **Step 1: Create `LLDOCX/.gitignore`**

```gitignore
.vs/
bin/
obj/
*.user
*.suo
*.pfx
packages/
TestResults/
local.settings.json
```

- [ ] **Step 2: Retarget and enable binding redirects in `LLDOCX/LLDOCX/LLDOCX.csproj`**

Change line 30 from:

```xml
    <TargetFrameworkVersion>v4.6.1</TargetFrameworkVersion>
```

to:

```xml
    <TargetFrameworkVersion>v4.7.2</TargetFrameworkVersion>
    <LangVersion>latest</LangVersion>
    <AutoGenerateBindingRedirects>true</AutoGenerateBindingRedirects>
    <GenerateBindingRedirectsOutputType>true</GenerateBindingRedirectsOutputType>
    <RestoreProjectStyle>PackageReference</RestoreProjectStyle>
```

- [ ] **Step 3: Verify in Visual Studio (manual)**

1. Open `LLDOCX/LLDOCX.sln` in Visual Studio 2022 and choose Build → Rebuild Solution. Expected: the build succeeds.
2. Press F5. Expected: Word opens, and the Home tab shows the LL-Docx group with **Scan** and **Settings**.
3. Click **Scan** on any document. Expected: it behaves as before.

- [ ] **Step 4: Commit**

Run from `Tools/`:

```bash
git add LLDOCX/.gitignore LLDOCX/LLDOCX.sln LLDOCX/LLDOCX LLDOCX/LLDOCXSetup
git status --short   # confirm LLDOCX_TemporaryKey.pfx and .vs/ are NOT listed
git commit -m "chore(lldocx): track add-in source, retarget to .NET Framework 4.7.2

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 2: Auth library scaffold and access policy

**Files:**
- Create: `LLDOCX/LLDOCX.Auth/LLDOCX.Auth.csproj`
- Create: `LLDOCX/LLDOCX.Auth/GateState.cs`
- Create: `LLDOCX/LLDOCX.Auth/RenewalOutcome.cs`
- Create: `LLDOCX/LLDOCX.Auth/AccessPolicy.cs`
- Create: `LLDOCX/LLDOCX.Auth.Tests/LLDOCX.Auth.Tests.csproj`
- Test: `LLDOCX/LLDOCX.Auth.Tests/AccessPolicyTests.cs`
- Modify: `LLDOCX/LLDOCX.sln` (via `dotnet sln add`)

**Interfaces:**
- Consumes: nothing
- Produces (namespace `LLDOCX.Auth`):
  - `enum GateStatus { Locked, Unlocked, UnlockedOffline }`
  - `enum LockReason { None, Checking, NotConfigured, SignInRequired, WrongTenant, OfflineTooLong }`
  - `sealed class GateState`:
    - members: `Status`, `Reason`, `Username`, `OfflineDays` (int?), `Detail` (string), `bool IsUnlocked`, `string Message`, `bool SameAs(GateState)`;
    - factories: `static GateState Online(string username)`, `static GateState Offline(string username, int days)`, `static GateState Locked(LockReason reason, string detail = null)`.
  - `enum RenewalStatus { Success, NetworkUnavailable, InteractionRequired, NoAccount }`
  - `sealed class RenewalOutcome`:
    - members: `Status`, `TenantId`, `Username`, `AccountId`, `ErrorMessage`;
    - factories: `static Success(string tenantId, string username, string accountId)`, `static Network()`, `static Interaction(string errorMessage)`, `static NoAccount()`.
  - `sealed class VerifiedRecord(DateTime verifiedUtc, string accountId, string username)`
  - `static class AccessPolicy`: `const int MaxGraceDays = 30`, `static GateState Evaluate(RenewalOutcome outcome, string configuredTenantId, VerifiedRecord last, DateTime nowUtc, int graceDays)`

- [ ] **Step 1: Install the .NET 10 SDK**

Run: `winget install --id Microsoft.DotNet.SDK.10 -e --accept-source-agreements --accept-package-agreements`
Then, in a new shell, run `dotnet --list-sdks`.
Expected: a line starting with `10.0.`

- [ ] **Step 2: Create `LLDOCX/LLDOCX.Auth/LLDOCX.Auth.csproj`**

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFramework>netstandard2.0</TargetFramework>
    <LangVersion>7.3</LangVersion>
    <RootNamespace>LLDOCX.Auth</RootNamespace>
  </PropertyGroup>
</Project>
```

- [ ] **Step 3: Create `LLDOCX/LLDOCX.Auth.Tests/LLDOCX.Auth.Tests.csproj`**

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
    <ProjectReference Include="..\LLDOCX.Auth\LLDOCX.Auth.csproj" />
  </ItemGroup>
</Project>
```

- [ ] **Step 4: Add both projects to the solution**

Run: `dotnet sln LLDOCX.sln add LLDOCX.Auth/LLDOCX.Auth.csproj LLDOCX.Auth.Tests/LLDOCX.Auth.Tests.csproj`
Expected: `Project ... added to the solution.` twice.

- [ ] **Step 5: Write the failing tests `LLDOCX/LLDOCX.Auth.Tests/AccessPolicyTests.cs`**

```csharp
global using Microsoft.VisualStudio.TestTools.UnitTesting;
using LLDOCX.Auth;

namespace LLDOCX.Auth.Tests;

[TestClass]
public class AccessPolicyTests
{
    const string Tenant = "11111111-1111-1111-1111-111111111111";
    static readonly DateTime Now = new DateTime(2026, 11, 1, 12, 0, 0, DateTimeKind.Utc);

    static VerifiedRecord VerifiedAgo(TimeSpan age) => new VerifiedRecord(Now - age, "acct", "pat@lozaip.com");

    [TestMethod]
    public void Success_in_configured_tenant_unlocks_online()
    {
        var s = AccessPolicy.Evaluate(RenewalOutcome.Success(Tenant, "pat@lozaip.com", "acct"), Tenant, null, Now, 30);
        Assert.AreEqual(GateStatus.Unlocked, s.Status);
        Assert.AreEqual("pat@lozaip.com", s.Username);
        Assert.AreEqual("Signed in as pat@lozaip.com", s.Message);
        Assert.IsTrue(s.IsUnlocked);
    }

    [TestMethod]
    public void Tenant_comparison_ignores_case()
    {
        var s = AccessPolicy.Evaluate(RenewalOutcome.Success(Tenant.ToUpperInvariant(), "pat@lozaip.com", "acct"), Tenant, null, Now, 30);
        Assert.AreEqual(GateStatus.Unlocked, s.Status);
    }

    [TestMethod]
    public void Success_in_other_tenant_locks_as_wrong_tenant()
    {
        var s = AccessPolicy.Evaluate(RenewalOutcome.Success("22222222-2222-2222-2222-222222222222", "x@other.com", "a"), Tenant, null, Now, 30);
        Assert.AreEqual(GateStatus.Locked, s.Status);
        Assert.AreEqual(LockReason.WrongTenant, s.Reason);
        Assert.AreEqual("This account is not a lozaip.com account.", s.Message);
    }

    [TestMethod]
    public void Offline_within_grace_unlocks_offline_with_day_count()
    {
        var s = AccessPolicy.Evaluate(RenewalOutcome.Network(), Tenant, VerifiedAgo(TimeSpan.FromDays(3.5)), Now, 30);
        Assert.AreEqual(GateStatus.UnlockedOffline, s.Status);
        Assert.AreEqual(3, s.OfflineDays);
        Assert.AreEqual("pat@lozaip.com", s.Username);
        Assert.AreEqual("Offline — verified 3 days ago", s.Message);
        Assert.IsTrue(s.IsUnlocked);
    }

    [TestMethod]
    public void Offline_message_uses_singular_day()
    {
        var s = AccessPolicy.Evaluate(RenewalOutcome.Network(), Tenant, VerifiedAgo(TimeSpan.FromDays(1.2)), Now, 30);
        Assert.AreEqual("Offline — verified 1 day ago", s.Message);
    }

    [TestMethod]
    public void Offline_at_exactly_the_grace_period_still_unlocks()
    {
        var s = AccessPolicy.Evaluate(RenewalOutcome.Network(), Tenant, VerifiedAgo(TimeSpan.FromDays(30)), Now, 30);
        Assert.AreEqual(GateStatus.UnlockedOffline, s.Status);
    }

    [TestMethod]
    public void Offline_one_second_past_the_grace_period_locks()
    {
        var s = AccessPolicy.Evaluate(RenewalOutcome.Network(), Tenant, VerifiedAgo(TimeSpan.FromDays(30) + TimeSpan.FromSeconds(1)), Now, 30);
        Assert.AreEqual(LockReason.OfflineTooLong, s.Reason);
        Assert.AreEqual("Connect to the internet to verify your account.", s.Message);
    }

    [TestMethod]
    public void Offline_with_no_record_locks()
    {
        var s = AccessPolicy.Evaluate(RenewalOutcome.Network(), Tenant, null, Now, 30);
        Assert.AreEqual(LockReason.OfflineTooLong, s.Reason);
    }

    [TestMethod]
    public void Offline_with_record_in_the_future_locks_clock_rollback()
    {
        var s = AccessPolicy.Evaluate(RenewalOutcome.Network(), Tenant, VerifiedAgo(TimeSpan.FromMinutes(-5)), Now, 30);
        Assert.AreEqual(LockReason.OfflineTooLong, s.Reason);
    }

    [TestMethod]
    public void Grace_days_above_maximum_are_clamped_to_30()
    {
        var s = AccessPolicy.Evaluate(RenewalOutcome.Network(), Tenant, VerifiedAgo(TimeSpan.FromDays(31)), Now, 90);
        Assert.AreEqual(LockReason.OfflineTooLong, s.Reason);
    }

    [TestMethod]
    public void Grace_days_below_one_are_clamped_to_1()
    {
        var s = AccessPolicy.Evaluate(RenewalOutcome.Network(), Tenant, VerifiedAgo(TimeSpan.FromHours(20)), Now, 0);
        Assert.AreEqual(GateStatus.UnlockedOffline, s.Status);
    }

    [TestMethod]
    public void Interaction_required_locks_and_keeps_the_error_detail()
    {
        var s = AccessPolicy.Evaluate(RenewalOutcome.Interaction("AADSTS50057: The user account is disabled."), Tenant, VerifiedAgo(TimeSpan.FromDays(1)), Now, 30);
        Assert.AreEqual(LockReason.SignInRequired, s.Reason);
        Assert.AreEqual("Sign in with your lozaip.com account.", s.Message);
        Assert.AreEqual("AADSTS50057: The user account is disabled.", s.Detail);
    }

    [TestMethod]
    public void No_account_locks_with_sign_in_message()
    {
        var s = AccessPolicy.Evaluate(RenewalOutcome.NoAccount(), Tenant, null, Now, 30);
        Assert.AreEqual(LockReason.SignInRequired, s.Reason);
        Assert.IsFalse(s.IsUnlocked);
    }

    [TestMethod]
    public void Locked_messages_cover_checking_and_not_configured()
    {
        Assert.AreEqual("Verifying your lozaip.com account…", GateState.Locked(LockReason.Checking).Message);
        Assert.AreEqual("LL-Docx is not configured. Contact your administrator.", GateState.Locked(LockReason.NotConfigured).Message);
    }

    [TestMethod]
    public void SameAs_compares_status_reason_user_and_days()
    {
        Assert.IsTrue(GateState.Online("a").SameAs(GateState.Online("a")));
        Assert.IsFalse(GateState.Online("a").SameAs(GateState.Online("b")));
        Assert.IsFalse(GateState.Offline("a", 1).SameAs(GateState.Offline("a", 2)));
        Assert.IsFalse(GateState.Locked(LockReason.SignInRequired).SameAs(GateState.Locked(LockReason.WrongTenant)));
    }
}
```

- [ ] **Step 6: Run the tests to verify they fail**

Run: `dotnet test LLDOCX.Auth.Tests`
Expected: the build FAILS with `The type or namespace name 'AccessPolicy' could not be found`.

- [ ] **Step 7: Write `LLDOCX/LLDOCX.Auth/GateState.cs`**

```csharp
namespace LLDOCX.Auth
{
    public enum GateStatus { Locked, Unlocked, UnlockedOffline }

    public enum LockReason { None, Checking, NotConfigured, SignInRequired, WrongTenant, OfflineTooLong }

    /// <summary>Whether the LL-Docx tools may be used right now, and what to tell the user.</summary>
    public sealed class GateState
    {
        GateState(GateStatus status, LockReason reason, string username, int? offlineDays, string detail)
        {
            Status = status;
            Reason = reason;
            Username = username;
            OfflineDays = offlineDays;
            Detail = detail;
        }

        public GateStatus Status { get; }
        public LockReason Reason { get; }
        public string Username { get; }
        public int? OfflineDays { get; }

        /// <summary>Extra error text from Entra (for example why a sign-in failed). May be null.</summary>
        public string Detail { get; }

        public bool IsUnlocked => Status != GateStatus.Locked;

        public string Message
        {
            get
            {
                if (Status == GateStatus.Unlocked) return "Signed in as " + Username;
                if (Status == GateStatus.UnlockedOffline)
                    return "Offline — verified " + OfflineDays + (OfflineDays == 1 ? " day ago" : " days ago");
                switch (Reason)
                {
                    case LockReason.Checking: return "Verifying your lozaip.com account…";
                    case LockReason.NotConfigured: return "LL-Docx is not configured. Contact your administrator.";
                    case LockReason.WrongTenant: return "This account is not a lozaip.com account.";
                    case LockReason.OfflineTooLong: return "Connect to the internet to verify your account.";
                    default: return "Sign in with your lozaip.com account.";
                }
            }
        }

        public static GateState Online(string username) => new GateState(GateStatus.Unlocked, LockReason.None, username, null, null);

        public static GateState Offline(string username, int days) => new GateState(GateStatus.UnlockedOffline, LockReason.None, username, days, null);

        public static GateState Locked(LockReason reason, string detail = null) => new GateState(GateStatus.Locked, reason, null, null, detail);

        public bool SameAs(GateState other) =>
            other != null && Status == other.Status && Reason == other.Reason &&
            Username == other.Username && OfflineDays == other.OfflineDays;
    }
}
```

- [ ] **Step 8: Write `LLDOCX/LLDOCX.Auth/RenewalOutcome.cs`**

```csharp
using System;

namespace LLDOCX.Auth
{
    public enum RenewalStatus { Success, NetworkUnavailable, InteractionRequired, NoAccount }

    /// <summary>The result of asking the identity provider for a token, silently or interactively.</summary>
    public sealed class RenewalOutcome
    {
        RenewalOutcome(RenewalStatus status, string tenantId, string username, string accountId, string errorMessage)
        {
            Status = status;
            TenantId = tenantId;
            Username = username;
            AccountId = accountId;
            ErrorMessage = errorMessage;
        }

        public RenewalStatus Status { get; }
        public string TenantId { get; }
        public string Username { get; }
        public string AccountId { get; }
        public string ErrorMessage { get; }

        public static RenewalOutcome Success(string tenantId, string username, string accountId) =>
            new RenewalOutcome(RenewalStatus.Success, tenantId, username, accountId, null);

        public static RenewalOutcome Network() => new RenewalOutcome(RenewalStatus.NetworkUnavailable, null, null, null, null);

        public static RenewalOutcome Interaction(string errorMessage) =>
            new RenewalOutcome(RenewalStatus.InteractionRequired, null, null, null, errorMessage);

        public static RenewalOutcome NoAccount() => new RenewalOutcome(RenewalStatus.NoAccount, null, null, null, null);
    }

    /// <summary>When the user's account was last confirmed online.</summary>
    public sealed class VerifiedRecord
    {
        public VerifiedRecord(DateTime verifiedUtc, string accountId, string username)
        {
            VerifiedUtc = verifiedUtc;
            AccountId = accountId;
            Username = username;
        }

        public DateTime VerifiedUtc { get; }
        public string AccountId { get; }
        public string Username { get; }
    }
}
```

- [ ] **Step 9: Write `LLDOCX/LLDOCX.Auth/AccessPolicy.cs`**

```csharp
using System;

namespace LLDOCX.Auth
{
    /// <summary>The access decision table from the sign-in gate spec. Pure: no I/O, no clock.</summary>
    public static class AccessPolicy
    {
        public const int MaxGraceDays = 30;

        public static GateState Evaluate(RenewalOutcome outcome, string configuredTenantId, VerifiedRecord last,
            DateTime nowUtc, int graceDays)
        {
            int grace = Math.Max(1, Math.Min(MaxGraceDays, graceDays));
            switch (outcome.Status)
            {
                case RenewalStatus.Success:
                    return string.Equals(outcome.TenantId, configuredTenantId, StringComparison.OrdinalIgnoreCase)
                        ? GateState.Online(outcome.Username)
                        : GateState.Locked(LockReason.WrongTenant);

                case RenewalStatus.NetworkUnavailable:
                    if (last == null) return GateState.Locked(LockReason.OfflineTooLong);
                    TimeSpan age = nowUtc - last.VerifiedUtc;
                    if (age < TimeSpan.Zero || age > TimeSpan.FromDays(grace)) return GateState.Locked(LockReason.OfflineTooLong);
                    return GateState.Offline(last.Username, (int)age.TotalDays);

                case RenewalStatus.InteractionRequired:
                    return GateState.Locked(LockReason.SignInRequired, outcome.ErrorMessage);

                default:
                    return GateState.Locked(LockReason.SignInRequired);
            }
        }
    }
}
```

- [ ] **Step 10: Run the tests to verify they pass**

Run: `dotnet test LLDOCX.Auth.Tests`
Expected: `Passed!` with 0 failed.

- [ ] **Step 11: Commit**

```bash
git add LLDOCX/LLDOCX.sln LLDOCX/LLDOCX.Auth LLDOCX/LLDOCX.Auth.Tests
git commit -m "feat(auth): add gate state and access policy decision table

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 3: Compiled configuration and Group Policy overrides

**Files:**
- Create: `LLDOCX/LLDOCX.Auth/GateConfig.cs`
- Create: `LLDOCX/LLDOCX.Auth.Tests/Fakes.cs`
- Test: `LLDOCX/LLDOCX.Auth.Tests/GateConfigLoaderTests.cs`

**Interfaces:**
- Consumes: `AccessPolicy.MaxGraceDays` (Task 2)
- Produces (namespace `LLDOCX.Auth`):
  - `sealed class GateConfig(string tenantId, string clientId, string gatewayUrl, string gatewayScope, int offlineGraceDays)`:
    - properties: `TenantId`, `ClientId`, `GatewayUrl`, `GatewayScope`, `OfflineGraceDays`;
    - `bool IsConfigured` (tenant and client are non-empty GUIDs);
    - `bool DeepCheckAvailable` (configured, plus an https gateway URL and a non-empty scope).
  - `static class GateDefaults`: constants `TenantId`, `ClientId`, `GatewayUrl`, `GatewayScope`, `OfflineGraceDays`; `static GateConfig Create()`
  - `interface IPolicySource { string GetString(string name); int? GetInt(string name); }` (returns null when the value is absent)
  - `static class GateConfigLoader { static GateConfig Load(IPolicySource policy, GateConfig defaults, Action<string> log = null); }`
  - Test fake `DictionaryPolicySource`

- [ ] **Step 1: Write `LLDOCX/LLDOCX.Auth.Tests/Fakes.cs`** (with only the policy fake for now; Task 5 adds more)

```csharp
using LLDOCX.Auth;

namespace LLDOCX.Auth.Tests;

internal sealed class DictionaryPolicySource : IPolicySource
{
    public Dictionary<string, object> Values { get; } = new(StringComparer.OrdinalIgnoreCase);

    public string GetString(string name) => Values.TryGetValue(name, out var v) ? v as string : null;

    public int? GetInt(string name) => Values.TryGetValue(name, out var v) && v is int i ? i : null;
}
```

- [ ] **Step 2: Write the failing tests `LLDOCX/LLDOCX.Auth.Tests/GateConfigLoaderTests.cs`**

```csharp
using LLDOCX.Auth;

namespace LLDOCX.Auth.Tests;

[TestClass]
public class GateConfigLoaderTests
{
    const string Tenant = "11111111-1111-1111-1111-111111111111";
    const string Client = "22222222-2222-2222-2222-222222222222";
    const string OtherClient = "33333333-3333-3333-3333-333333333333";
    static readonly GateConfig Defaults = new(Tenant, Client, "https://gw.example.net/api/deepcheck", "api://gw/DeepCheck.Use", 30);

    [TestMethod]
    public void Uses_compiled_defaults_when_no_policy_is_set()
    {
        var c = GateConfigLoader.Load(new DictionaryPolicySource(), Defaults);
        Assert.AreEqual(Tenant, c.TenantId);
        Assert.AreEqual(Client, c.ClientId);
        Assert.AreEqual("https://gw.example.net/api/deepcheck", c.GatewayUrl);
        Assert.AreEqual("api://gw/DeepCheck.Use", c.GatewayScope);
        Assert.AreEqual(30, c.OfflineGraceDays);
        Assert.IsTrue(c.IsConfigured);
        Assert.IsTrue(c.DeepCheckAvailable);
    }

    [TestMethod]
    public void Null_policy_source_is_treated_as_empty()
    {
        Assert.AreEqual(Client, GateConfigLoader.Load(null, Defaults).ClientId);
    }

    [TestMethod]
    public void Applies_each_allowed_override()
    {
        var p = new DictionaryPolicySource();
        p.Values["ClientId"] = OtherClient;
        p.Values["GatewayUrl"] = "https://other.example.net/api/deepcheck";
        p.Values["GatewayScope"] = "api://other/DeepCheck.Use";
        p.Values["OfflineGraceDays"] = 7;
        var c = GateConfigLoader.Load(p, Defaults);
        Assert.AreEqual(OtherClient, c.ClientId);
        Assert.AreEqual("https://other.example.net/api/deepcheck", c.GatewayUrl);
        Assert.AreEqual("api://other/DeepCheck.Use", c.GatewayScope);
        Assert.AreEqual(7, c.OfflineGraceDays);
    }

    [TestMethod]
    public void Ignores_tenant_override_and_logs_it()
    {
        var p = new DictionaryPolicySource();
        p.Values["TenantId"] = "44444444-4444-4444-4444-444444444444";
        var log = new List<string>();
        var c = GateConfigLoader.Load(p, Defaults, log.Add);
        Assert.AreEqual(Tenant, c.TenantId);
        Assert.IsTrue(log.Any(m => m.Contains("TenantId")));
    }

    [TestMethod]
    public void Ignores_invalid_override_values()
    {
        var p = new DictionaryPolicySource();
        p.Values["ClientId"] = "not-a-guid";
        p.Values["GatewayUrl"] = "http://insecure.example.net/api/deepcheck";
        p.Values["GatewayScope"] = "   ";
        var log = new List<string>();
        var c = GateConfigLoader.Load(p, Defaults, log.Add);
        Assert.AreEqual(Client, c.ClientId);
        Assert.AreEqual("https://gw.example.net/api/deepcheck", c.GatewayUrl);
        Assert.AreEqual("api://gw/DeepCheck.Use", c.GatewayScope);
        Assert.AreEqual(3, log.Count);
    }

    [TestMethod]
    public void Grace_override_can_only_lower_the_value()
    {
        var p = new DictionaryPolicySource();
        p.Values["OfflineGraceDays"] = 45;
        Assert.AreEqual(30, GateConfigLoader.Load(p, Defaults).OfflineGraceDays);
    }

    [TestMethod]
    public void Grace_override_below_one_is_clamped_to_one()
    {
        var p = new DictionaryPolicySource();
        p.Values["OfflineGraceDays"] = 0;
        Assert.AreEqual(1, GateConfigLoader.Load(p, Defaults).OfflineGraceDays);
    }

    [TestMethod]
    public void Placeholder_ids_mean_not_configured()
    {
        var c = GateConfigLoader.Load(new DictionaryPolicySource(), GateDefaults.Create());
        Assert.IsFalse(c.IsConfigured);
        Assert.IsFalse(c.DeepCheckAvailable);
    }

    [TestMethod]
    public void Missing_gateway_only_disables_deep_check()
    {
        var c = GateConfigLoader.Load(new DictionaryPolicySource(), new GateConfig(Tenant, Client, "", "", 30));
        Assert.IsTrue(c.IsConfigured);
        Assert.IsFalse(c.DeepCheckAvailable);
    }
}
```

- [ ] **Step 3: Run the tests to verify they fail**

Run: `dotnet test LLDOCX.Auth.Tests --filter GateConfigLoaderTests`
Expected: the build FAILS with `The type or namespace name 'GateConfig' could not be found`.

- [ ] **Step 4: Write `LLDOCX/LLDOCX.Auth/GateConfig.cs`**

```csharp
using System;

namespace LLDOCX.Auth
{
    public sealed class GateConfig
    {
        public GateConfig(string tenantId, string clientId, string gatewayUrl, string gatewayScope, int offlineGraceDays)
        {
            TenantId = tenantId ?? "";
            ClientId = clientId ?? "";
            GatewayUrl = gatewayUrl ?? "";
            GatewayScope = gatewayScope ?? "";
            OfflineGraceDays = offlineGraceDays;
        }

        public string TenantId { get; }
        public string ClientId { get; }
        public string GatewayUrl { get; }
        public string GatewayScope { get; }
        public int OfflineGraceDays { get; }

        public bool IsConfigured => IsRealGuid(TenantId) && IsRealGuid(ClientId);

        public bool DeepCheckAvailable => IsConfigured && IsHttpsUrl(GatewayUrl) && !string.IsNullOrWhiteSpace(GatewayScope);

        internal static bool IsRealGuid(string s)
        {
            Guid g;
            return Guid.TryParse(s, out g) && g != Guid.Empty;
        }

        internal static bool IsHttpsUrl(string s)
        {
            Uri u;
            return Uri.TryCreate(s, UriKind.Absolute, out u) && u.Scheme == Uri.UriSchemeHttps;
        }
    }

    /// <summary>
    /// lozaip.com values compiled into the signed add-in. Fill these in from the admin setup values
    /// (docs/admin/sign-in-setup.md). The all-zero placeholders mean "not configured".
    /// </summary>
    public static class GateDefaults
    {
        public const string TenantId = "00000000-0000-0000-0000-000000000000";
        public const string ClientId = "00000000-0000-0000-0000-000000000000";
        public const string GatewayUrl = "";
        public const string GatewayScope = "";
        public const int OfflineGraceDays = 30;

        public static GateConfig Create() => new GateConfig(TenantId, ClientId, GatewayUrl, GatewayScope, OfflineGraceDays);
    }

    /// <summary>Admin-only policy values (HKLM\SOFTWARE\Policies\LL-Docx in the add-in).</summary>
    public interface IPolicySource
    {
        /// <summary>Returns null when the value is absent.</summary>
        string GetString(string name);

        /// <summary>Returns null when the value is absent.</summary>
        int? GetInt(string name);
    }

    public static class GateConfigLoader
    {
        public static GateConfig Load(IPolicySource policy, GateConfig defaults, Action<string> log = null)
        {
            string clientId = defaults.ClientId;
            string gatewayUrl = defaults.GatewayUrl;
            string gatewayScope = defaults.GatewayScope;
            int grace = Math.Max(1, Math.Min(AccessPolicy.MaxGraceDays, defaults.OfflineGraceDays));

            if (policy != null)
            {
                if (policy.GetString("TenantId") != null)
                    log?.Invoke("TenantId cannot be overridden by policy; ignoring the policy value.");

                string value = policy.GetString("ClientId");
                if (value != null)
                {
                    if (GateConfig.IsRealGuid(value.Trim())) clientId = value.Trim();
                    else log?.Invoke("Ignoring invalid ClientId policy value.");
                }

                value = policy.GetString("GatewayUrl");
                if (value != null)
                {
                    if (GateConfig.IsHttpsUrl(value.Trim())) gatewayUrl = value.Trim();
                    else log?.Invoke("Ignoring invalid GatewayUrl policy value (must be https).");
                }

                value = policy.GetString("GatewayScope");
                if (value != null)
                {
                    if (value.Trim().Length > 0) gatewayScope = value.Trim();
                    else log?.Invoke("Ignoring empty GatewayScope policy value.");
                }

                int? days = policy.GetInt("OfflineGraceDays");
                if (days.HasValue) grace = Math.Min(grace, Math.Max(1, days.Value));
            }

            return new GateConfig(defaults.TenantId, clientId, gatewayUrl, gatewayScope, grace);
        }
    }
}
```

- [ ] **Step 5: Run the tests to verify they pass**

Run: `dotnet test LLDOCX.Auth.Tests`
Expected: `Passed!` with 0 failed.

- [ ] **Step 6: Commit**

```bash
git add LLDOCX/LLDOCX.Auth LLDOCX/LLDOCX.Auth.Tests
git commit -m "feat(auth): add compiled gate config with admin policy overrides

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 4: Encrypted last-verified store

**Files:**
- Create: `LLDOCX/LLDOCX.Auth/LastVerifiedStore.cs`
- Modify: `LLDOCX/LLDOCX.Auth/LLDOCX.Auth.csproj` (ProtectedData package)
- Test: `LLDOCX/LLDOCX.Auth.Tests/LastVerifiedStoreTests.cs`

**Interfaces:**
- Consumes: `VerifiedRecord` (Task 2)
- Produces (namespace `LLDOCX.Auth`):
  - `interface ILastVerifiedStore { VerifiedRecord Load(); void Save(VerifiedRecord record); void Clear(); }`. `Load` returns null when the record is missing or unreadable.
  - `sealed class DpapiLastVerifiedStore(string filePath) : ILastVerifiedStore`

- [ ] **Step 1: Add the DPAPI package**

Run: `dotnet add LLDOCX.Auth package System.Security.Cryptography.ProtectedData`
Expected: `PackageReference for package 'System.Security.Cryptography.ProtectedData' version '...' added`.

- [ ] **Step 2: Write the failing tests `LLDOCX/LLDOCX.Auth.Tests/LastVerifiedStoreTests.cs`**

```csharp
using LLDOCX.Auth;

namespace LLDOCX.Auth.Tests;

[TestClass]
public class LastVerifiedStoreTests
{
    string _dir;
    string _file;

    [TestInitialize]
    public void Init()
    {
        _dir = Path.Combine(Path.GetTempPath(), "lldocx-auth-tests-" + Guid.NewGuid().ToString("N"));
        _file = Path.Combine(_dir, "gate.dat");
    }

    [TestCleanup]
    public void Cleanup()
    {
        if (Directory.Exists(_dir)) Directory.Delete(_dir, true);
    }

    [TestMethod]
    public void Round_trips_a_record_and_creates_the_folder()
    {
        var store = new DpapiLastVerifiedStore(_file);
        var when = new DateTime(2026, 11, 1, 8, 30, 0, DateTimeKind.Utc);
        store.Save(new VerifiedRecord(when, "acct-1", "pat@lozaip.com"));
        var loaded = new DpapiLastVerifiedStore(_file).Load();
        Assert.AreEqual(when, loaded.VerifiedUtc);
        Assert.AreEqual(DateTimeKind.Utc, loaded.VerifiedUtc.Kind);
        Assert.AreEqual("acct-1", loaded.AccountId);
        Assert.AreEqual("pat@lozaip.com", loaded.Username);
    }

    [TestMethod]
    public void File_is_not_plain_text()
    {
        new DpapiLastVerifiedStore(_file).Save(new VerifiedRecord(DateTime.UtcNow, "acct-1", "pat@lozaip.com"));
        StringAssert.DoesNotMatch(File.ReadAllText(_file), new System.Text.RegularExpressions.Regex("lozaip"));
    }

    [TestMethod]
    public void Missing_file_loads_as_null()
    {
        Assert.IsNull(new DpapiLastVerifiedStore(_file).Load());
    }

    [TestMethod]
    public void Corrupt_file_loads_as_null()
    {
        Directory.CreateDirectory(_dir);
        File.WriteAllBytes(_file, new byte[] { 1, 2, 3, 4, 5 });
        Assert.IsNull(new DpapiLastVerifiedStore(_file).Load());
    }

    [TestMethod]
    public void Clear_deletes_the_record()
    {
        var store = new DpapiLastVerifiedStore(_file);
        store.Save(new VerifiedRecord(DateTime.UtcNow, "a", "u"));
        store.Clear();
        Assert.IsNull(store.Load());
        store.Clear(); // clearing twice is harmless
    }
}
```

- [ ] **Step 3: Run the tests to verify they fail**

Run: `dotnet test LLDOCX.Auth.Tests --filter LastVerifiedStoreTests`
Expected: the build FAILS with `The type or namespace name 'DpapiLastVerifiedStore' could not be found`.

- [ ] **Step 4: Write `LLDOCX/LLDOCX.Auth/LastVerifiedStore.cs`**

```csharp
using System;
using System.Globalization;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace LLDOCX.Auth
{
    public interface ILastVerifiedStore
    {
        /// <summary>Returns null when there is no record or it can't be read.</summary>
        VerifiedRecord Load();

        void Save(VerifiedRecord record);

        void Clear();
    }

    /// <summary>Stores the last-verified record encrypted with DPAPI for the current Windows user.</summary>
    public sealed class DpapiLastVerifiedStore : ILastVerifiedStore
    {
        static readonly byte[] Entropy = Encoding.UTF8.GetBytes("LL-Docx gate record v1");
        readonly string _filePath;

        public DpapiLastVerifiedStore(string filePath)
        {
            _filePath = filePath;
        }

        public VerifiedRecord Load()
        {
            try
            {
                if (!File.Exists(_filePath)) return null;
                byte[] plain = ProtectedData.Unprotect(File.ReadAllBytes(_filePath), Entropy, DataProtectionScope.CurrentUser);
                string[] lines = Encoding.UTF8.GetString(plain).Split('\n');
                if (lines.Length != 4 || lines[0] != "v1") return null;
                long ticks = long.Parse(lines[1], CultureInfo.InvariantCulture);
                return new VerifiedRecord(new DateTime(ticks, DateTimeKind.Utc), lines[2], lines[3]);
            }
            catch (Exception ex) when (ex is CryptographicException || ex is FormatException || ex is IOException ||
                                       ex is UnauthorizedAccessException || ex is OverflowException || ex is ArgumentException)
            {
                return null;
            }
        }

        public void Save(VerifiedRecord record)
        {
            string text = string.Join("\n", "v1",
                record.VerifiedUtc.ToUniversalTime().Ticks.ToString(CultureInfo.InvariantCulture),
                Clean(record.AccountId), Clean(record.Username));
            byte[] cipher = ProtectedData.Protect(Encoding.UTF8.GetBytes(text), Entropy, DataProtectionScope.CurrentUser);
            Directory.CreateDirectory(Path.GetDirectoryName(_filePath));
            File.WriteAllBytes(_filePath, cipher);
        }

        public void Clear()
        {
            if (File.Exists(_filePath)) File.Delete(_filePath);
        }

        static string Clean(string s) => (s ?? "").Replace("\n", " ").Replace("\r", " ");
    }
}
```

- [ ] **Step 5: Run the tests to verify they pass**

Run: `dotnet test LLDOCX.Auth.Tests`
Expected: `Passed!` with 0 failed.

- [ ] **Step 6: Commit**

```bash
git add LLDOCX/LLDOCX.Auth LLDOCX/LLDOCX.Auth.Tests
git commit -m "feat(auth): add DPAPI-encrypted last-verified store

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 5: AuthGate state machine

**Files:**
- Create: `LLDOCX/LLDOCX.Auth/AuthGate.cs`
- Modify: `LLDOCX/LLDOCX.Auth.Tests/Fakes.cs` (add fakes)
- Test: `LLDOCX/LLDOCX.Auth.Tests/AuthGateTests.cs`

**Interfaces:**
- Consumes: `GateState`, `RenewalOutcome`, `VerifiedRecord`, `AccessPolicy` (Task 2); `GateConfig` (Task 3); `ILastVerifiedStore` (Task 4)
- Produces (namespace `LLDOCX.Auth`):
  - `interface ITokenSource`:
    - `Task<RenewalOutcome> RenewSilentAsync(CancellationToken ct)`;
    - `Task<RenewalOutcome> SignInInteractiveAsync(IntPtr parentWindow, CancellationToken ct)`, which throws `OperationCanceledException` when the user cancels;
    - `Task<string> GetGatewayTokenAsync(CancellationToken ct)`, which throws `SignInRequiredException`;
    - `Task SignOutAsync()`.
  - `interface IClock { DateTime UtcNow { get; } }` and `sealed class SystemClock : IClock` with `static readonly SystemClock Instance`
  - `class SignInRequiredException : Exception`
  - `sealed class AuthGate(GateConfig config, ITokenSource tokens, ILastVerifiedStore store, IClock clock, Action<string> log = null)`:
    - state: `GateState State`, `event EventHandler StateChanged`;
    - `Task<GateState> CheckAsync(CancellationToken ct = default)`
    - `Task<GateState> SignInAsync(IntPtr parentWindow, CancellationToken ct = default)`
    - `Task SignOutAsync()`
    - `Task<GateState> EnsureFreshAsync(CancellationToken ct = default)`
    - `Task<string> GetGatewayTokenAsync(CancellationToken ct)`
    - constants: `RecheckInterval` = 24 h, `OfflineRetryInterval` = 5 min.
  - `tokens` may be null when `config.IsConfigured` is false.

- [ ] **Step 1: Add the fakes to `LLDOCX/LLDOCX.Auth.Tests/Fakes.cs`**

Append to the file:

```csharp
internal sealed class FakeClock : IClock
{
    public DateTime UtcNow { get; set; } = new DateTime(2026, 11, 1, 12, 0, 0, DateTimeKind.Utc);
}

internal sealed class MemoryStore : ILastVerifiedStore
{
    public VerifiedRecord Record { get; set; }
    public VerifiedRecord Load() => Record;
    public void Save(VerifiedRecord record) => Record = record;
    public void Clear() => Record = null;
}

internal sealed class FakeTokenSource : ITokenSource
{
    public Queue<Func<RenewalOutcome>> Silent { get; } = new();
    public Func<RenewalOutcome> Interactive { get; set; }
    public Func<string> GatewayToken { get; set; } = () => "token";
    public int SilentCalls { get; private set; }
    public int SignOutCalls { get; private set; }

    public Task<RenewalOutcome> RenewSilentAsync(CancellationToken ct)
    {
        SilentCalls++;
        return Task.FromResult(Silent.Dequeue()());
    }

    public Task<RenewalOutcome> SignInInteractiveAsync(IntPtr parentWindow, CancellationToken ct) => Task.FromResult(Interactive());

    public Task<string> GetGatewayTokenAsync(CancellationToken ct) => Task.FromResult(GatewayToken());

    public Task SignOutAsync()
    {
        SignOutCalls++;
        return Task.CompletedTask;
    }
}
```

- [ ] **Step 2: Write the failing tests `LLDOCX/LLDOCX.Auth.Tests/AuthGateTests.cs`**

```csharp
using LLDOCX.Auth;

namespace LLDOCX.Auth.Tests;

[TestClass]
public class AuthGateTests
{
    const string Tenant = "11111111-1111-1111-1111-111111111111";
    const string Client = "22222222-2222-2222-2222-222222222222";
    static readonly GateConfig Config = new(Tenant, Client, "https://gw.example.net/api/deepcheck", "api://gw/DeepCheck.Use", 30);

    readonly FakeTokenSource _tokens = new();
    readonly MemoryStore _store = new();
    readonly FakeClock _clock = new();

    AuthGate NewGate(GateConfig config = null) => new(config ?? Config, _tokens, _store, _clock);

    static RenewalOutcome Ok() => RenewalOutcome.Success(Tenant, "pat@lozaip.com", "acct");

    [TestMethod]
    public void Starts_locked_while_checking()
    {
        Assert.AreEqual(LockReason.Checking, NewGate().State.Reason);
    }

    [TestMethod]
    public async Task Not_configured_locks_without_calling_the_token_source()
    {
        var gate = new AuthGate(GateDefaults.Create(), null, _store, _clock);
        var s = await gate.CheckAsync();
        Assert.AreEqual(LockReason.NotConfigured, s.Reason);
        Assert.AreEqual(0, _tokens.SilentCalls);
    }

    [TestMethod]
    public async Task Successful_check_unlocks_records_verification_and_raises_event()
    {
        var gate = NewGate();
        int events = 0;
        gate.StateChanged += (s, e) => events++;
        _tokens.Silent.Enqueue(Ok);
        var state = await gate.CheckAsync();
        Assert.AreEqual(GateStatus.Unlocked, state.Status);
        Assert.AreEqual(1, events);
        Assert.AreEqual(_clock.UtcNow, _store.Record.VerifiedUtc);
        Assert.AreEqual("pat@lozaip.com", _store.Record.Username);
    }

    [TestMethod]
    public async Task Same_state_twice_raises_event_once()
    {
        var gate = NewGate();
        int events = 0;
        gate.StateChanged += (s, e) => events++;
        _tokens.Silent.Enqueue(Ok);
        _tokens.Silent.Enqueue(Ok);
        await gate.CheckAsync();
        await gate.CheckAsync();
        Assert.AreEqual(1, events);
    }

    [TestMethod]
    public async Task Offline_uses_the_stored_record()
    {
        _store.Record = new VerifiedRecord(_clock.UtcNow.AddDays(-2), "acct", "pat@lozaip.com");
        _tokens.Silent.Enqueue(RenewalOutcome.Network);
        var s = await NewGate().CheckAsync();
        Assert.AreEqual(GateStatus.UnlockedOffline, s.Status);
        Assert.AreEqual(2, s.OfflineDays);
    }

    [TestMethod]
    public async Task Wrong_tenant_signs_out_and_clears_the_record()
    {
        _store.Record = new VerifiedRecord(_clock.UtcNow, "acct", "pat@lozaip.com");
        _tokens.Silent.Enqueue(() => RenewalOutcome.Success("99999999-9999-9999-9999-999999999999", "x@other.com", "b"));
        var s = await NewGate().CheckAsync();
        Assert.AreEqual(LockReason.WrongTenant, s.Reason);
        Assert.AreEqual(1, _tokens.SignOutCalls);
        Assert.IsNull(_store.Record);
    }

    [TestMethod]
    public async Task Exception_from_token_source_stays_locked()
    {
        _tokens.Silent.Enqueue(() => throw new InvalidOperationException("boom"));
        var s = await NewGate().CheckAsync();
        Assert.IsFalse(s.IsUnlocked);
        Assert.AreEqual(LockReason.SignInRequired, s.Reason);
    }

    [TestMethod]
    public async Task Interactive_sign_in_unlocks()
    {
        _tokens.Interactive = Ok;
        var s = await NewGate().SignInAsync(IntPtr.Zero);
        Assert.AreEqual(GateStatus.Unlocked, s.Status);
    }

    [TestMethod]
    public async Task Cancelled_sign_in_keeps_the_previous_state()
    {
        var gate = NewGate();
        _tokens.Silent.Enqueue(RenewalOutcome.NoAccount);
        await gate.CheckAsync();
        _tokens.Interactive = () => throw new OperationCanceledException();
        var s = await gate.SignInAsync(IntPtr.Zero);
        Assert.AreEqual(LockReason.SignInRequired, s.Reason);
        Assert.IsNull(s.Detail);
    }

    [TestMethod]
    public async Task Failed_sign_in_keeps_the_entra_message_as_detail()
    {
        _tokens.Interactive = () => RenewalOutcome.Interaction("AADSTS50105: not assigned");
        var s = await NewGate().SignInAsync(IntPtr.Zero);
        Assert.AreEqual("AADSTS50105: not assigned", s.Detail);
    }

    [TestMethod]
    public async Task Sign_out_locks_and_clears()
    {
        var gate = NewGate();
        _tokens.Silent.Enqueue(Ok);
        await gate.CheckAsync();
        await gate.SignOutAsync();
        Assert.AreEqual(LockReason.SignInRequired, gate.State.Reason);
        Assert.IsNull(_store.Record);
        Assert.AreEqual(1, _tokens.SignOutCalls);
    }

    [TestMethod]
    public async Task EnsureFresh_does_not_recheck_within_24_hours()
    {
        var gate = NewGate();
        _tokens.Silent.Enqueue(Ok);
        await gate.CheckAsync();
        _clock.UtcNow = _clock.UtcNow.AddHours(23);
        await gate.EnsureFreshAsync();
        Assert.AreEqual(1, _tokens.SilentCalls);
    }

    [TestMethod]
    public async Task EnsureFresh_rechecks_after_24_hours_and_detects_disabled_account()
    {
        var gate = NewGate();
        _tokens.Silent.Enqueue(Ok);
        await gate.CheckAsync();
        _clock.UtcNow = _clock.UtcNow.AddHours(25);
        _tokens.Silent.Enqueue(() => RenewalOutcome.Interaction("AADSTS50057: disabled"));
        var s = await gate.EnsureFreshAsync();
        Assert.AreEqual(2, _tokens.SilentCalls);
        Assert.AreEqual(LockReason.SignInRequired, s.Reason);
    }

    [TestMethod]
    public async Task EnsureFresh_retries_offline_state_at_most_every_5_minutes()
    {
        _store.Record = new VerifiedRecord(_clock.UtcNow.AddDays(-1), "acct", "pat@lozaip.com");
        var gate = NewGate();
        _tokens.Silent.Enqueue(RenewalOutcome.Network);
        await gate.CheckAsync();
        _clock.UtcNow = _clock.UtcNow.AddMinutes(2);
        await gate.EnsureFreshAsync();
        Assert.AreEqual(1, _tokens.SilentCalls);
        _clock.UtcNow = _clock.UtcNow.AddMinutes(4);
        _tokens.Silent.Enqueue(Ok);
        var s = await gate.EnsureFreshAsync();
        Assert.AreEqual(2, _tokens.SilentCalls);
        Assert.AreEqual(GateStatus.Unlocked, s.Status);
    }

    [TestMethod]
    public async Task EnsureFresh_rechecks_when_locked_for_being_offline_too_long()
    {
        var gate = NewGate();
        _tokens.Silent.Enqueue(RenewalOutcome.Network);
        await gate.CheckAsync();
        Assert.AreEqual(LockReason.OfflineTooLong, gate.State.Reason);
        _clock.UtcNow = _clock.UtcNow.AddMinutes(6);
        _tokens.Silent.Enqueue(Ok);
        Assert.AreEqual(GateStatus.Unlocked, (await gate.EnsureFreshAsync()).Status);
    }

    [TestMethod]
    public async Task EnsureFresh_does_not_silently_retry_when_sign_in_is_required()
    {
        var gate = NewGate();
        _tokens.Silent.Enqueue(RenewalOutcome.NoAccount);
        await gate.CheckAsync();
        _clock.UtcNow = _clock.UtcNow.AddDays(2);
        await gate.EnsureFreshAsync();
        Assert.AreEqual(1, _tokens.SilentCalls);
    }

    [TestMethod]
    public async Task Gateway_token_sign_in_required_locks_the_gate()
    {
        var gate = NewGate();
        _tokens.Silent.Enqueue(Ok);
        await gate.CheckAsync();
        _tokens.GatewayToken = () => throw new SignInRequiredException("expired");
        await Assert.ThrowsExceptionAsync<SignInRequiredException>(() => gate.GetGatewayTokenAsync(CancellationToken.None));
        Assert.AreEqual(LockReason.SignInRequired, gate.State.Reason);
    }

    [TestMethod]
    public async Task Gateway_token_requires_a_configured_gateway()
    {
        var gate = NewGate(new GateConfig(Tenant, Client, "", "", 30));
        await Assert.ThrowsExceptionAsync<InvalidOperationException>(() => gate.GetGatewayTokenAsync(CancellationToken.None));
    }
}
```

- [ ] **Step 3: Run the tests to verify they fail**

Run: `dotnet test LLDOCX.Auth.Tests --filter AuthGateTests`
Expected: the build FAILS with `The type or namespace name 'ITokenSource' could not be found`.

- [ ] **Step 4: Write `LLDOCX/LLDOCX.Auth/AuthGate.cs`**

```csharp
using System;
using System.Threading;
using System.Threading.Tasks;

namespace LLDOCX.Auth
{
    public interface ITokenSource
    {
        Task<RenewalOutcome> RenewSilentAsync(CancellationToken ct);

        /// <summary>Shows the Microsoft sign-in window. Throws OperationCanceledException if the user cancels.</summary>
        Task<RenewalOutcome> SignInInteractiveAsync(IntPtr parentWindow, CancellationToken ct);

        /// <summary>Returns an access token for the Deep Check gateway. Throws SignInRequiredException.</summary>
        Task<string> GetGatewayTokenAsync(CancellationToken ct);

        Task SignOutAsync();
    }

    public interface IClock
    {
        DateTime UtcNow { get; }
    }

    public sealed class SystemClock : IClock
    {
        public static readonly SystemClock Instance = new SystemClock();
        public DateTime UtcNow => DateTime.UtcNow;
    }

    public class SignInRequiredException : Exception
    {
        public SignInRequiredException(string message, Exception inner = null) : base(message, inner) { }
    }

    /// <summary>Decides whether the LL-Docx tools are unlocked. Never unlocks on an error.</summary>
    public sealed class AuthGate
    {
        public static readonly TimeSpan RecheckInterval = TimeSpan.FromHours(24);
        public static readonly TimeSpan OfflineRetryInterval = TimeSpan.FromMinutes(5);

        readonly GateConfig _config;
        readonly ITokenSource _tokens;
        readonly ILastVerifiedStore _store;
        readonly IClock _clock;
        readonly Action<string> _log;
        readonly SemaphoreSlim _lock = new SemaphoreSlim(1, 1);
        DateTime? _lastOnlineCheckUtc;
        DateTime? _lastAttemptUtc;
        GateState _state = GateState.Locked(LockReason.Checking);

        public AuthGate(GateConfig config, ITokenSource tokens, ILastVerifiedStore store, IClock clock, Action<string> log = null)
        {
            _config = config ?? throw new ArgumentNullException(nameof(config));
            if (config.IsConfigured && tokens == null) throw new ArgumentNullException(nameof(tokens));
            _tokens = tokens;
            _store = store ?? throw new ArgumentNullException(nameof(store));
            _clock = clock ?? throw new ArgumentNullException(nameof(clock));
            _log = log;
        }

        public GateState State => _state;

        /// <summary>Raised on a thread-pool thread when the state changes. Marshal to the UI thread.</summary>
        public event EventHandler StateChanged;

        public async Task<GateState> CheckAsync(CancellationToken ct = default(CancellationToken))
        {
            if (!_config.IsConfigured) return SetState(GateState.Locked(LockReason.NotConfigured));
            await _lock.WaitAsync(ct).ConfigureAwait(false);
            try
            {
                RenewalOutcome outcome;
                try
                {
                    outcome = await _tokens.RenewSilentAsync(ct).ConfigureAwait(false);
                }
                catch (Exception ex)
                {
                    _log?.Invoke("Silent token renewal failed: " + ex);
                    outcome = RenewalOutcome.Interaction(null);
                }
                return await ApplyAsync(outcome).ConfigureAwait(false);
            }
            finally
            {
                _lock.Release();
            }
        }

        public async Task<GateState> SignInAsync(IntPtr parentWindow, CancellationToken ct = default(CancellationToken))
        {
            if (!_config.IsConfigured) return SetState(GateState.Locked(LockReason.NotConfigured));
            await _lock.WaitAsync(ct).ConfigureAwait(false);
            try
            {
                RenewalOutcome outcome;
                try
                {
                    outcome = await _tokens.SignInInteractiveAsync(parentWindow, ct).ConfigureAwait(false);
                }
                catch (OperationCanceledException)
                {
                    return _state;
                }
                catch (Exception ex)
                {
                    _log?.Invoke("Interactive sign-in failed: " + ex);
                    outcome = RenewalOutcome.Interaction(ex.Message);
                }
                return await ApplyAsync(outcome).ConfigureAwait(false);
            }
            finally
            {
                _lock.Release();
            }
        }

        public async Task SignOutAsync()
        {
            if (_tokens != null)
            {
                try { await _tokens.SignOutAsync().ConfigureAwait(false); }
                catch (Exception ex) { _log?.Invoke("Sign-out failed: " + ex); }
            }
            SafeClear();
            _lastOnlineCheckUtc = null;
            SetState(GateState.Locked(_config.IsConfigured ? LockReason.SignInRequired : LockReason.NotConfigured));
        }

        /// <summary>Call before every gated action: re-checks silently when the last online check is stale.</summary>
        public Task<GateState> EnsureFreshAsync(CancellationToken ct = default(CancellationToken))
        {
            DateTime now = _clock.UtcNow;
            bool retryAllowed = _lastAttemptUtc == null || now - _lastAttemptUtc.Value >= OfflineRetryInterval;
            switch (_state.Status)
            {
                case GateStatus.Unlocked:
                    if (_lastOnlineCheckUtc == null || now - _lastOnlineCheckUtc.Value > RecheckInterval) return CheckAsync(ct);
                    return Task.FromResult(_state);
                case GateStatus.UnlockedOffline:
                    return retryAllowed ? CheckAsync(ct) : Task.FromResult(_state);
                default:
                    bool retryable = _state.Reason == LockReason.OfflineTooLong || _state.Reason == LockReason.Checking;
                    return retryable && retryAllowed ? CheckAsync(ct) : Task.FromResult(_state);
            }
        }

        public async Task<string> GetGatewayTokenAsync(CancellationToken ct)
        {
            if (!_config.DeepCheckAvailable) throw new InvalidOperationException("No Deep Check gateway is configured.");
            try
            {
                return await _tokens.GetGatewayTokenAsync(ct).ConfigureAwait(false);
            }
            catch (SignInRequiredException)
            {
                SetState(GateState.Locked(LockReason.SignInRequired));
                throw;
            }
        }

        async Task<GateState> ApplyAsync(RenewalOutcome outcome)
        {
            DateTime now = _clock.UtcNow;
            _lastAttemptUtc = now;
            if (outcome.Status != RenewalStatus.NetworkUnavailable) _lastOnlineCheckUtc = now;

            var state = AccessPolicy.Evaluate(outcome, _config.TenantId, SafeLoad(), now, _config.OfflineGraceDays);
            if (state.Status == GateStatus.Unlocked)
            {
                SafeSave(new VerifiedRecord(now, outcome.AccountId, outcome.Username));
            }
            else if (state.Reason == LockReason.WrongTenant)
            {
                try { await _tokens.SignOutAsync().ConfigureAwait(false); }
                catch (Exception ex) { _log?.Invoke("Sign-out after wrong tenant failed: " + ex); }
                SafeClear();
            }
            return SetState(state);
        }

        GateState SetState(GateState state)
        {
            bool changed = !state.SameAs(_state);
            _state = state;
            if (changed) StateChanged?.Invoke(this, EventArgs.Empty);
            return state;
        }

        VerifiedRecord SafeLoad()
        {
            try { return _store.Load(); }
            catch (Exception ex) { _log?.Invoke("Reading gate record failed: " + ex); return null; }
        }

        void SafeSave(VerifiedRecord record)
        {
            try { _store.Save(record); }
            catch (Exception ex) { _log?.Invoke("Saving gate record failed: " + ex); }
        }

        void SafeClear()
        {
            try { _store.Clear(); }
            catch (Exception ex) { _log?.Invoke("Clearing gate record failed: " + ex); }
        }
    }
}
```

- [ ] **Step 5: Run the tests to verify they pass**

Run: `dotnet test LLDOCX.Auth.Tests`
Expected: `Passed!` with 0 failed.

- [ ] **Step 6: Commit**

```bash
git add LLDOCX/LLDOCX.Auth LLDOCX/LLDOCX.Auth.Tests
git commit -m "feat(auth): add AuthGate state machine with 24h re-check

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 6: Wire the gate into the Word add-in

**Files:**
- Modify: `LLDOCX/LLDOCX/LLDOCX.csproj` (Auth project reference, MSAL packages, Compile items)
- Create: `LLDOCX/LLDOCX/Auth/MsalTokenSource.cs`
- Create: `LLDOCX/LLDOCX/Auth/RegistryPolicySource.cs`
- Modify: `LLDOCX/LLDOCX/ThisAddIn.cs`
- Modify: `LLDOCX/LLDOCX/RibbonExplorer.xml`
- Modify: `LLDOCX/LLDOCX/RibbonExplorer.cs`

**Interfaces:**
- Consumes: everything in `LLDOCX.Auth` (Tasks 2–5)
- Produces:
  - `ThisAddIn` members:
    - `GateConfig GateConfig`, `AuthGate Gate`;
    - `Task<bool> EnsureUnlockedAsync()`, which re-checks when stale and shows the lock message when locked;
    - `Task SignInAsync()`, `void HideAllPanes()`.
  - `RibbonExplorer` members:
    - `void InvalidateAll()`, `void ResetSettingsToggle()`;
    - ribbon callbacks `GetEnabled`, `GetAccountLabel`, `GetAccountScreentip`, `OnAccount`.
  - The claims-checker plan adds its buttons, and extends `HideAllPanes` and `GetEnabled`, on top of this.
  - The add-in-side classes live in namespace `LLDOCX.AddinAuth`, which avoids a clash with the `LLDOCX.Auth` library namespace.

- [ ] **Step 1: Add references to `LLDOCX/LLDOCX/LLDOCX.csproj`**

1. **Add the NuGet packages.** In Visual Studio, right-click the **LLDOCX** project → **Manage NuGet Packages** → Browse, and install the latest stable versions of:
   - `Microsoft.Identity.Client.Broker` (this brings `Microsoft.Identity.Client`)
   - `Microsoft.Identity.Client.Extensions.Msal`

   Because Task 1 set `RestoreProjectStyle` to PackageReference, this adds `<PackageReference>` items to the csproj.
2. **Add the project reference.** After the `ItemGroup` that contains `<Compile Include="CustomEvents\MissingHeadingHandler.cs" />`, add:

```xml
  <ItemGroup>
    <ProjectReference Include="..\LLDOCX.Auth\LLDOCX.Auth.csproj" />
  </ItemGroup>
```

3. **Add the new source files.** In the `ItemGroup` with the other `Compile` items, add:

```xml
    <Compile Include="Auth\MsalTokenSource.cs" />
    <Compile Include="Auth\RegistryPolicySource.cs" />
```

- [ ] **Step 2: Create `LLDOCX/LLDOCX/Auth/RegistryPolicySource.cs`**

```csharp
using System;
using LLDOCX.Auth;
using Microsoft.Win32;

namespace LLDOCX.AddinAuth
{
    /// <summary>Reads admin-only overrides from HKLM\SOFTWARE\Policies\LL-Docx (set by Group Policy or an admin).</summary>
    sealed class RegistryPolicySource : IPolicySource
    {
        const string KeyPath = @"SOFTWARE\Policies\LL-Docx";

        public string GetString(string name) => Read(name) as string;

        public int? GetInt(string name)
        {
            object value = Read(name);
            if (value is int i) return i;
            return null;
        }

        static object Read(string name)
        {
            try
            {
                using (var hklm = RegistryKey.OpenBaseKey(RegistryHive.LocalMachine, RegistryView.Registry64))
                using (var key = hklm.OpenSubKey(KeyPath))
                {
                    return key?.GetValue(name);
                }
            }
            catch (Exception ex) when (ex is System.Security.SecurityException || ex is UnauthorizedAccessException)
            {
                return null;
            }
        }
    }
}
```

- [ ] **Step 3: Create `LLDOCX/LLDOCX/Auth/MsalTokenSource.cs`**

```csharp
using System;
using System.IO;
using System.Linq;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using LLDOCX.Auth;
using Microsoft.Identity.Client;
using Microsoft.Identity.Client.Broker;
using Microsoft.Identity.Client.Extensions.Msal;

namespace LLDOCX.AddinAuth
{
    /// <summary>ITokenSource backed by MSAL.NET, the Windows account broker (WAM) and a DPAPI token cache.</summary>
    sealed class MsalTokenSource : ITokenSource
    {
        readonly IPublicClientApplication _app;
        readonly string[] _signInScopes;
        readonly string[] _gatewayScopes;
        readonly Lazy<Task> _cacheReady;

        public MsalTokenSource(GateConfig config, string dataFolder)
        {
            _gatewayScopes = config.DeepCheckAvailable ? new[] { config.GatewayScope } : null;
            // Requesting the gateway scope at sign-in means one consent covers Deep Check later.
            _signInScopes = _gatewayScopes ?? new[] { "User.Read" };
            _app = PublicClientApplicationBuilder.Create(config.ClientId)
                .WithAuthority(AzureCloudInstance.AzurePublic, config.TenantId)
                .WithRedirectUri("http://localhost")
                .WithBroker(new BrokerOptions(BrokerOptions.OperatingSystems.Windows) { Title = "LL-Docx" })
                .Build();
            _cacheReady = new Lazy<Task>(async () =>
            {
                Directory.CreateDirectory(dataFolder);
                var props = new StorageCreationPropertiesBuilder("msal.cache", dataFolder).Build();
                var helper = await MsalCacheHelper.CreateAsync(props).ConfigureAwait(false);
                helper.RegisterCache(_app.UserTokenCache);
            });
        }

        public async Task<RenewalOutcome> RenewSilentAsync(CancellationToken ct)
        {
            try
            {
                await _cacheReady.Value.ConfigureAwait(false);
                var account = (await _app.GetAccountsAsync().ConfigureAwait(false)).FirstOrDefault();
                if (account == null) return RenewalOutcome.NoAccount();
                var result = await _app.AcquireTokenSilent(_signInScopes, account).ExecuteAsync(ct).ConfigureAwait(false);
                return ToOutcome(result);
            }
            catch (MsalUiRequiredException ex)
            {
                return RenewalOutcome.Interaction(ex.Message);
            }
            catch (Exception ex) when (IsNetwork(ex))
            {
                return RenewalOutcome.Network();
            }
        }

        public async Task<RenewalOutcome> SignInInteractiveAsync(IntPtr parentWindow, CancellationToken ct)
        {
            await _cacheReady.Value.ConfigureAwait(false);
            try
            {
                var result = await _app.AcquireTokenInteractive(_signInScopes)
                    .WithParentActivityOrWindow(parentWindow)
                    .WithPrompt(Prompt.SelectAccount)
                    .WithUseEmbeddedWebView(false)
                    .ExecuteAsync(ct).ConfigureAwait(false);
                return ToOutcome(result);
            }
            catch (MsalClientException ex) when (ex.ErrorCode == MsalError.AuthenticationCanceledError)
            {
                throw new OperationCanceledException("Sign-in canceled.", ex);
            }
            catch (Exception ex) when (IsNetwork(ex))
            {
                return RenewalOutcome.Network();
            }
            catch (MsalException ex)
            {
                return RenewalOutcome.Interaction(ex.Message);
            }
        }

        public async Task<string> GetGatewayTokenAsync(CancellationToken ct)
        {
            if (_gatewayScopes == null) throw new InvalidOperationException("No Deep Check gateway is configured.");
            await _cacheReady.Value.ConfigureAwait(false);
            var account = (await _app.GetAccountsAsync().ConfigureAwait(false)).FirstOrDefault();
            if (account == null) throw new SignInRequiredException("Not signed in.");
            try
            {
                var result = await _app.AcquireTokenSilent(_gatewayScopes, account).ExecuteAsync(ct).ConfigureAwait(false);
                return result.AccessToken;
            }
            catch (MsalUiRequiredException ex)
            {
                throw new SignInRequiredException(ex.Message, ex);
            }
        }

        public async Task SignOutAsync()
        {
            await _cacheReady.Value.ConfigureAwait(false);
            foreach (var account in await _app.GetAccountsAsync().ConfigureAwait(false))
                await _app.RemoveAsync(account).ConfigureAwait(false);
        }

        static RenewalOutcome ToOutcome(AuthenticationResult result) =>
            RenewalOutcome.Success(result.TenantId, result.Account?.Username, result.Account?.HomeAccountId?.Identifier);

        static bool IsNetwork(Exception ex)
        {
            for (var e = ex; e != null; e = e.InnerException)
            {
                if (e is HttpRequestException) return true;
                if (e is MsalException m && (m.ErrorCode == MsalError.RequestTimeout || m.ErrorCode == "service_not_available"))
                    return true;
            }
            return false;
        }
    }
}
```

- [ ] **Step 4: Replace the class body of `LLDOCX/LLDOCX/ThisAddIn.cs` above `#region VSTO generated code`**

Replace the usings at the top of the file with:

```csharp
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;
using Microsoft.Office.Tools.Word;
using LLDOCX.AddinAuth;
using LLDOCX.Auth;
using LLDOCX.CustomUserControl;
using LLDOCX.CustomEvents;
using LLDOCX.Helpers;
```

Replace everything from `private SettingUserControl _settingControl;` through the end of `ThisAddIn_Shutdown` with:

```csharp
        private SettingUserControl _settingControl;
        private Microsoft.Office.Tools.CustomTaskPane myCustomTaskPane;
        private RibbonExplorer _ribbon;
        private SynchronizationContext _ui;
        public MissingHeadingHandler MissingHeadingHandler { get; set; }
        public List<DocumentHeadingData> DocumentHeadingData { get; set; }
        public GateConfig GateConfig { get; private set; }
        public AuthGate Gate { get; private set; }

        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // Lets async ribbon callbacks resume on the Office UI thread after awaiting background work.
            if (!(SynchronizationContext.Current is WindowsFormsSynchronizationContext))
                SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());
            _ui = SynchronizationContext.Current;
            MissingHeadingHandler = new MissingHeadingHandler();
            DocumentHeadingData = new List<DocumentHeadingData>();
            StartGate();
        }

        private void StartGate()
        {
            GateConfig = GateConfigLoader.Load(new RegistryPolicySource(), GateDefaults.Create(), Log);
            string dataFolder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "LL-Docx");
            ITokenSource tokens = GateConfig.IsConfigured ? new MsalTokenSource(GateConfig, dataFolder) : null;
            Gate = new AuthGate(GateConfig, tokens, new DpapiLastVerifiedStore(Path.Combine(dataFolder, "gate.dat")), SystemClock.Instance, Log);
            Gate.StateChanged += (s, args) => _ui.Post(_ => OnGateStateChanged(), null);
            // CheckAsync never throws for token errors; it runs in the background so Word starts immediately.
            Gate.CheckAsync().ContinueWith(t => Log("Gate check faulted: " + t.Exception), TaskContinuationOptions.OnlyOnFaulted);
        }

        private void OnGateStateChanged()
        {
            _ribbon?.InvalidateAll();
            if (!Gate.State.IsUnlocked) HideAllPanes();
        }

        public void HideAllPanes()
        {
            if (myCustomTaskPane != null) myCustomTaskPane.Visible = false;
            _ribbon?.ResetSettingsToggle();
        }

        /// <summary>Call at the start of every gated action. Re-verifies when stale; explains why when locked.</summary>
        public async Task<bool> EnsureUnlockedAsync()
        {
            var state = await Gate.EnsureFreshAsync();
            if (state.IsUnlocked) return true;
            MessageBox.Show(state.Message, LLDocxConstant.AppName, MessageBoxButtons.OK, MessageBoxIcon.Information);
            return false;
        }

        public async Task SignInAsync()
        {
            if (!GateConfig.IsConfigured)
            {
                MessageBox.Show(Gate.State.Message, LLDocxConstant.AppName, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            var state = await Gate.SignInAsync(WordWindowHandle());
            if (!state.IsUnlocked && (state.Detail != null || state.Reason == LockReason.WrongTenant))
            {
                string text = state.Message + (string.IsNullOrEmpty(state.Detail) ? "" : Environment.NewLine + Environment.NewLine + state.Detail);
                MessageBox.Show(text, LLDocxConstant.AppName, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
        }

        private IntPtr WordWindowHandle()
        {
            try { return new IntPtr(Application.ActiveWindow.Hwnd); }
            catch (COMException) { return Process.GetCurrentProcess().MainWindowHandle; }
        }

        private static void Log(string message) => Trace.TraceInformation("LL-DOCX: " + message);

        public void ProcessSideBarPanel(bool isVisible)
        {
            if (_settingControl == null)
            {
                _settingControl = new SettingUserControl();
            }
            if (myCustomTaskPane == null)
            {
                myCustomTaskPane = this.CustomTaskPanes.Add(_settingControl, "LL-DOCX Setting", this.Application.ActiveWindow);
                myCustomTaskPane.Width = 450;
            }
            myCustomTaskPane.Visible = isVisible;
        }

        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            _ribbon = new RibbonExplorer();
            return _ribbon;
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }
```

The old `if (CustomTaskPanes.Count == 0)` check becomes `if (myCustomTaskPane == null)`. Without that change, the Settings pane would never be created once another LL-Docx pane exists.

- [ ] **Step 5: Gate the ribbon in `LLDOCX/LLDOCX/RibbonExplorer.xml`**

Replace the `<group>` element with:

```xml
        <group id="MyGroup"
               label="LL-Docx">
          <button id="btnScan" label="Scan" size="large" onAction="OnScanning" getImage="GetImages" getEnabled="GetEnabled"/>
          <toggleButton id="btnPanel" label="Settings" getPressed="GetPressed" getImage="GetImages" size="large" onAction="ShowSetting" getEnabled="GetEnabled"/>
          <button id="btnAccount" getLabel="GetAccountLabel" getScreentip="GetAccountScreentip" size="large" onAction="OnAccount" getImage="GetImages"
                  supertip="Sign in with your lozaip.com account to use LL-Docx, or sign out."/>
        </group>
```

- [ ] **Step 6: Update `LLDOCX/LLDOCX/RibbonExplorer.cs`**

1. Add `using System.Threading.Tasks;` to the usings.
2. Replace `GetImages`, `OnScanning`, `ShowSetting` and `InvalidateControl` with the versions below.
3. Add the new methods (`GetEnabled` through `ResetSettingsToggle`) after `ShowSetting`.

```csharp
        public Bitmap GetImages(Office.IRibbonControl control)
        {
            if (control.Id == "btnScan" || control.Id == "btnAccount")
                return Properties.Resources.LL_logo;
            if (isPressed)
                return Properties.Resources.toggle_on;
            return Properties.Resources.toggle_off;
        }
        public async void OnScanning(Office.IRibbonControl control)
        {
            if (!await Globals.ThisAddIn.EnsureUnlockedAsync()) return;
            var option = ProcessData.GetData();
            var missingheadings = Helpers.WordDocumentHelper.ValidateDocument(option);
            if (missingheadings.Count > 0 || Globals.ThisAddIn.DocumentHeadingData.Count > 0)
            {
                if (!this.isPressed)
                {
                    this.isPressed = true;
                    InvalidateControl(control.Id);
                }
                Globals.ThisAddIn.ProcessSideBarPanel(this.isPressed);
                Globals.ThisAddIn.MissingHeadingHandler.OnEmailItemChanged(missingheadings);
            }
        }
        public void InvalidateControl(string controlId)
        {
            ribbon?.InvalidateControl(controlId);
        }
        public async void ShowSetting(Office.IRibbonControl control, bool isPressed)
        {
            if (isPressed && !await Globals.ThisAddIn.EnsureUnlockedAsync())
            {
                this.isPressed = false;
                InvalidateControl(control.Id);
                return;
            }
            this.isPressed = isPressed;
            InvalidateControl(control.Id);
            Globals.ThisAddIn.ProcessSideBarPanel(this.isPressed);
        }
        public bool GetEnabled(Office.IRibbonControl control)
        {
            var gate = Globals.ThisAddIn.Gate;
            return gate != null && gate.State.IsUnlocked;
        }
        public string GetAccountLabel(Office.IRibbonControl control)
        {
            var gate = Globals.ThisAddIn.Gate;
            if (gate == null || !gate.State.IsUnlocked) return "Sign in";
            string user = gate.State.Username ?? "Account";
            int at = user.IndexOf('@');
            return at > 0 ? user.Substring(0, at) : user;
        }
        public string GetAccountScreentip(Office.IRibbonControl control)
        {
            var gate = Globals.ThisAddIn.Gate;
            return gate == null ? LLDocxConstant.AppName : gate.State.Message;
        }
        public async void OnAccount(Office.IRibbonControl control)
        {
            try
            {
                var gate = Globals.ThisAddIn.Gate;
                if (gate.State.IsUnlocked)
                {
                    var answer = System.Windows.Forms.MessageBox.Show(gate.State.Message + "\n\nSign out of LL-Docx?",
                        LLDocxConstant.AppName, System.Windows.Forms.MessageBoxButtons.YesNo, System.Windows.Forms.MessageBoxIcon.Question);
                    if (answer == System.Windows.Forms.DialogResult.Yes) await gate.SignOutAsync();
                }
                else
                {
                    await Globals.ThisAddIn.SignInAsync();
                }
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show("Sign-in failed: " + ex.Message, LLDocxConstant.AppName);
            }
        }
        public void InvalidateAll()
        {
            ribbon?.Invalidate();
        }
        public void ResetSettingsToggle()
        {
            if (!isPressed) return;
            isPressed = false;
            ribbon?.InvalidateControl("btnPanel");
        }
```

- [ ] **Step 7: Build and verify in Word with the placeholder configuration (manual)**

Rebuild in Visual Studio. Expected: the build succeeds. If MSAL member names differ in the installed version, for example `MsalError.RequestTimeout` or `BrokerOptions.Title`, fix them using IntelliSense, keeping the same intent, and rebuild.

Press F5, then check:

1. The LL-Docx group shows **Scan**, **Settings** and **Sign in**. Scan and Settings are disabled.
2. Hover over **Sign in**. The tooltip reads "LL-Docx is not configured. Contact your administrator."
3. Click **Sign in**. A message shows the same text.

- [ ] **Step 8: Verify real sign-in (manual, after the admin registers the app)**

Put the tenant ID and add-in client ID from the admin into `LLDOCX/LLDOCX.Auth/GateConfig.cs` (`GateDefaults.TenantId`, `GateDefaults.ClientId`), rebuild and press F5. Then check:

1. **Startup:** Scan and Settings are disabled and the Account button shows **Sign in**. If the account broker already holds the account silently, the buttons unlock within a few seconds.
2. **Sign-in:** click **Sign in** and sign in with a lozaip.com account. The buttons enable and the Account button shows the user name. Click **Scan**: it works as before.
3. **Restart:** close and reopen Word. The buttons unlock with no sign-in window.
4. **Sign-out:** click the Account button → **Yes**. The buttons disable and any open Settings pane hides.
5. **Wrong account:** sign in with a non-lozaip account, such as a personal Microsoft account or another organization's. Entra refuses it, or the add-in shows "This account is not a lozaip.com account." The buttons stay disabled.
6. **Standalone Office:** on a machine with standalone Office 2021 that is not signed in to Office, install the build, sign in once, then restart. It stays unlocked.
7. **Microsoft 365 signed into another account:** on Microsoft 365 Office signed in to a different Microsoft 365 account, LL-Docx still requires the lozaip.com account.
8. **Offline grace period:**
   - Sign in, then disconnect the network and restart Word. The Account tooltip reads "Offline — verified 0 days ago", and the tools work.
   - As an admin, set `HKLM\SOFTWARE\Policies\LL-Docx\OfflineGraceDays` (DWORD) = 1. Close Word, and set the PC clock forward two days. Start Word, still offline: the buttons are disabled, and clicking Scan says "Connect to the internet to verify your account."
   - Restore the clock, reconnect and delete the policy value.
9. **Disabled account:** with a test lozaip.com account, sign in, then have the admin disable the account. After Word restarts, or 24 hours later on the next button click, the tools lock with "Sign in with your lozaip.com account."

- [ ] **Step 9: Commit**

```bash
git add LLDOCX/LLDOCX
git commit -m "feat(lldocx): require lozaip.com sign-in for all LL-Docx tools

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 7: Per-machine installer and admin setup guide

**Files:**
- Modify: `LLDOCX/LLDOCXSetup/LLDOCXSetup.wixproj` (x64 platform)
- Modify: `LLDOCX/LLDOCXSetup/Product.wxs`
- Create: `LLDOCX/docs/admin/sign-in-setup.md`

**Interfaces:**
- Consumes: the Release build output of the add-in (Tasks 1–6)
- Produces:
  - a signed, per-machine x64 MSI that registers the add-in for all users in both registry views and refuses to install without its prerequisites;
  - the admin guide that the claims-checker plan's gateway task extends.

- [ ] **Step 1: Make the installer an x64 package**

In Visual Studio, open **Build → Configuration Manager**. For the `LLDOCXSetup` project, choose platform **x64** (create it from x86 if needed) for both Debug and Release. An x64 MSI can write both the 64-bit registry view and the 32-bit one (`WOW6432Node`), which is needed to support both 64-bit and 32-bit Office.

- [ ] **Step 2: Replace `LLDOCX/LLDOCXSetup/Product.wxs`**

The registration moves from HKCU to HKLM, written to both registry views. The prerequisite checks are added. The `LLDOCX.dll.config` component is uncommented, and the gate assembly is added.

```xml
<?xml version="1.0" encoding="UTF-8"?>
<Wix xmlns="http://schemas.microsoft.com/wix/2006/wi">
	<Product Id="*" Name="LL-Docx" Language="1033" Version="1.1.0.0" Manufacturer="LLDOCX" UpgradeCode="ba7c8330-a0d4-4900-b79c-0ef68ac8db75">
		<Package InstallerVersion="200" Compressed="yes" InstallScope="perMachine" Platform="x64" />

		<MajorUpgrade DowngradeErrorMessage="A newer version of [ProductName] is already installed." AllowSameVersionUpgrades="yes" />

		<!-- .NET Framework 4.7.2 or later: release key 461808 -->
		<PropertyRef Id="NETFRAMEWORK45"/>
		<Condition Message="LL-Docx requires .NET Framework 4.7.2 or later. Install it from https://dotnet.microsoft.com/download/dotnet-framework and run this installer again.">
			<![CDATA[Installed OR (NETFRAMEWORK45 AND NETFRAMEWORK45 >= "#461808")]]>
		</Condition>

		<!-- VSTO 2010 runtime (installed with Office or separately), in either registry view -->
		<Property Id="VSTORUNTIME64">
			<RegistrySearch Id="VstoRuntime64" Root="HKLM" Key="SOFTWARE\Microsoft\VSTO Runtime Setup\v4R" Name="Version" Type="raw" Win64="yes" />
		</Property>
		<Property Id="VSTORUNTIME32">
			<RegistrySearch Id="VstoRuntime32" Root="HKLM" Key="SOFTWARE\Microsoft\VSTO Runtime Setup\v4R" Name="Version" Type="raw" Win64="no" />
		</Property>
		<Condition Message="LL-Docx requires the Visual Studio 2010 Tools for Office Runtime. Install it from https://aka.ms/VSTORuntime and run this installer again.">
			<![CDATA[Installed OR VSTORUNTIME64 OR VSTORUNTIME32]]>
		</Condition>

		<!-- Word 2016 or later (all current Office versions, including Microsoft 365, report 16) -->
		<Property Id="WORDCURVER">
			<RegistrySearch Id="WordCurVer" Root="HKCR" Key="Word.Application\CurVer" Type="raw" />
		</Property>
		<Condition Message="LL-Docx requires Microsoft Word 2016 or later.">
			<![CDATA[Installed OR WORDCURVER = "Word.Application.16"]]>
		</Condition>

		<Condition Message="LL-Docx requires 64-bit Windows.">
			<![CDATA[Installed OR VersionNT64]]>
		</Condition>

		<Media Id="1" Cabinet="LLDOCX.cab" EmbedCab="yes"/>

		<Feature Id="ProductFeature" Title="LL-Docx" Level="1">
			<ComponentGroupRef Id="ProductComponents" />
			<ComponentRef Id="Registry_Addin64" />
			<ComponentRef Id="Registry_Addin32" />
		</Feature>
		<UIRef Id="WixUI_Mondo"></UIRef>
		<UIRef Id="WixUI_ErrorProgressText" />
		<WixVariable Id="WixUILicenseRtf" Value="EULA.rtf" />
		<Property Id="WIXUI_EXITDIALOGOPTIONALTEXT" Value="Thank you for installing LL-Docx. Open Word and click Sign in on the Home tab." />
	</Product>

	<Fragment>
		<Directory Id="TARGETDIR" Name="SourceDir">
			<Directory Id="ProgramFilesFolder">
				<Directory Id="INSTALLFOLDER" Name="LLDOCX" />
				<!-- Per-machine registration, read by 64-bit Office -->
				<Component Id="Registry_Addin64" Guid="7B0B7F5E-2C55-4C47-9E0E-6F1F2B6D6A01" Win64="yes">
					<RegistryKey Root="HKLM" Key="Software\Microsoft\Office\Word\Addins\LLDOCX">
						<RegistryValue Name="FriendlyName" Value="LL-Docx" Type="string" KeyPath="yes" />
						<RegistryValue Name="Description" Value="LL-Docx tools for Word" Type="string" />
						<RegistryValue Name="Manifest" Value="file:///[INSTALLFOLDER]LLDOCX.vsto|vstolocal" Type="string" />
						<RegistryValue Name="LoadBehavior" Value="3" Type="integer" />
					</RegistryKey>
				</Component>
				<!-- Per-machine registration, read by 32-bit Office on 64-bit Windows (WOW6432Node) -->
				<Component Id="Registry_Addin32" Guid="7B0B7F5E-2C55-4C47-9E0E-6F1F2B6D6A02" Win64="no">
					<RegistryKey Root="HKLM" Key="Software\Microsoft\Office\Word\Addins\LLDOCX">
						<RegistryValue Name="FriendlyName" Value="LL-Docx" Type="string" KeyPath="yes" />
						<RegistryValue Name="Description" Value="LL-Docx tools for Word" Type="string" />
						<RegistryValue Name="Manifest" Value="file:///[INSTALLFOLDER]LLDOCX.vsto|vstolocal" Type="string" />
						<RegistryValue Name="LoadBehavior" Value="3" Type="integer" />
					</RegistryKey>
				</Component>
			</Directory>
		</Directory>
	</Fragment>

	<Fragment>
		<ComponentGroup Id="ProductComponents" Directory="INSTALLFOLDER">
			<Component Id="LLDOCX_vsto_Component" Win64="no">
				<File Id="LLDOCX_vsto" KeyPath="yes" Name="LLDOCX.vsto" Source="$(var.AddinFiles)" />
			</Component>
			<Component Id="LLDOCX_dll_manifest_Component" Win64="no">
				<File Id="LLDOCX.dll.manifest" KeyPath="yes" Name="LLDOCX.dll.manifest" Source="$(var.AddinFiles)" />
			</Component>
			<Component Id="LLDOCX_dll_Component" Win64="no">
				<File Id="LLDOCX_dll" KeyPath="yes" Name="LLDOCX.dll" Source="$(var.AddinFiles)" />
			</Component>
			<Component Id="LLDOCX_dll_config_Component" Win64="no">
				<File Id="LLDOCX_dll_config" KeyPath="yes" Name="LLDOCX.dll.config" Source="$(var.AddinFiles)" />
			</Component>
			<Component Id="MSOfficeToolsCommon_dll_Component" Win64="no">
				<File Id="MSOfficeToolsCommon_dll" KeyPath="yes" Name="Microsoft.Office.Tools.Common.v4.0.Utilities.dll" Source="$(var.AddinFiles)" />
			</Component>
			<Component Id="Microsoft.VisualStudio_Tools_Applications_Runtime_dll_component" Win64="no">
				<File Id="Microsoft.VisualStudio.Tools_Applications_Runtime.dll" KeyPath="yes" Name="Microsoft.VisualStudio.Tools.Applications.Runtime.dll" Source="$(var.AddinFiles)" />
			</Component>
			<Component Id="LLDOCX_Auth_dll_Component" Win64="no">
				<File Id="LLDOCX_Auth_dll" KeyPath="yes" Name="LLDOCX.Auth.dll" Source="$(var.AddinFiles)" />
			</Component>
			<!-- Step 3 adds one component per remaining dependency DLL here. -->
		</ComponentGroup>
	</Fragment>
</Wix>
```

**Note:** the old HKCU registration from version 1.0.0.x is removed automatically by `MajorUpgrade`, because those components belonged to the old product.

- [ ] **Step 3: Add the MSAL dependency assemblies**

In Visual Studio, switch to **Release** and rebuild the add-in. Then run:

```bash
ls LLDOCX/bin/Release/*.dll
ls -R LLDOCX/bin/Release/runtimes 2>/dev/null
```

For every DLL in the first listing that isn't already in `ProductComponents`, add one component inside `<ComponentGroup Id="ProductComponents" ...>`. In the Ids, replace `.` and `-` in the file name with `_`. For example:

```xml
			<Component Id="Microsoft_Identity_Client_dll_Component" Win64="no">
				<File Id="Microsoft_Identity_Client_dll" KeyPath="yes" Name="Microsoft.Identity.Client.dll" Source="$(var.AddinFiles)" />
			</Component>
```

Expected entries include `Microsoft.Identity.Client.dll`, `Microsoft.Identity.Client.Broker.dll`, `Microsoft.Identity.Client.NativeInterop.dll`, `Microsoft.Identity.Client.Extensions.Msal.dll`, `Microsoft.IdentityModel.Abstractions.dll`, `System.Security.Cryptography.ProtectedData.dll`, and their `System.*` dependencies; the output of the command above is authoritative.

If the second listing shows native WAM binaries (for example `runtimes/win-x64/native/msalruntime.dll` and `runtimes/win-x86/native/msalruntime_x86.dll`), install them in the same relative folders. Add this inside `<Directory Id="INSTALLFOLDER" Name="LLDOCX">`, which then becomes an element with children instead of self-closing:

```xml
				<Directory Id="RUNTIMES" Name="runtimes">
					<Directory Id="RT_WINX64" Name="win-x64"><Directory Id="RT_WINX64_NATIVE" Name="native" /></Directory>
					<Directory Id="RT_WINX86" Name="win-x86"><Directory Id="RT_WINX86_NATIVE" Name="native" /></Directory>
				</Directory>
```

and add these components after the `ProductComponents` group, in the same fragment:

```xml
		<ComponentGroup Id="NativeRuntimeComponents">
			<Component Id="msalruntime_x64_Component" Directory="RT_WINX64_NATIVE" Win64="no">
				<File Id="msalruntime_x64" KeyPath="yes" Name="msalruntime.dll" Source="$(var.AddinFiles)runtimes\win-x64\native\msalruntime.dll" />
			</Component>
			<Component Id="msalruntime_x86_Component" Directory="RT_WINX86_NATIVE" Win64="no">
				<File Id="msalruntime_x86" KeyPath="yes" Name="msalruntime_x86.dll" Source="$(var.AddinFiles)runtimes\win-x86\native\msalruntime_x86.dll" />
			</Component>
		</ComponentGroup>
```

Then reference the group from the Feature with `<ComponentGroupRef Id="NativeRuntimeComponents" />`. Use the exact file names from the listing. If the native DLLs are instead placed directly next to `LLDOCX.dll`, add them as plain components in `ProductComponents`.

- [ ] **Step 4: Sign the release build (manual, needs lozaip.com's code-signing certificate)**

1. **Sign the add-in's manifests.** In the LLDOCX project properties → **Signing**, check "Sign the ClickOnce manifests" and choose **Select from Store…** → the lozaip.com code-signing certificate. Never select `LLDOCX_TemporaryKey.pfx` for a release. This signs `LLDOCX.vsto` and `LLDOCX.dll.manifest`, whose hashes cover `LLDOCX.dll` and `LLDOCX.Auth.dll`, so a modified DLL won't load.
2. **Build the installer.** Rebuild the add-in and `LLDOCXSetup` in Release | x64.
3. **Sign the MSI.** Run, from a Developer Command Prompt:

```bash
signtool sign /a /fd SHA256 /tr http://timestamp.digicert.com /td SHA256 LLDOCXSetup/bin/x64/Release/LLDOCXSetup.msi
```

Expected: `Successfully signed`. Use the timestamp URL your certificate vendor recommends.

- [ ] **Step 5: Test the installer (manual)**

Use clean VMs: one with standalone Office 2021 (32-bit) and one with Microsoft 365 Apps (64-bit). On each:

1. **Install.** Double-click the MSI and install. Expected: no sign-in is asked for during installation.
2. **First user.** Open Word as the installing user: the LL-Docx group appears, with **Sign in**.
3. **Second user.** Sign in to Windows as a second user and open Word: the LL-Docx group appears for them too, confirming per-machine registration.
4. **Upgrade.** Sign in to LL-Docx, then install a newer MSI build (bump `Version`). Reopen Word: the add-in is still there, and the user is still signed in.
5. **Uninstall.** Uninstall from Settings → Apps: the LL-Docx group is gone from Word.
6. **Prerequisite checks.** On a VM without Office, run the MSI: it refuses with "LL-Docx requires Microsoft Word 2016 or later."
7. **Silent install.** As an admin, run `msiexec /i LLDOCXSetup.msi /qn`: it installs with no UI.

- [ ] **Step 6: Write `LLDOCX/docs/admin/sign-in-setup.md`**

```markdown
# LL-Docx: sign-in and deployment setup (administrators)

LL-Docx only works for people signed in with a lozaip.com Microsoft 365 account. It works the same way with
standalone Office and Microsoft 365 Office. Installation needs no sign-in, and no device-management tool is required.

## 1. Register the add-in in Microsoft Entra ID

1. In the Microsoft Entra admin center, go to **App registrations** → **New registration**.
   - Name: `LL-Docx Add-in`
   - Supported account types: **Accounts in this organizational directory only (single tenant)**
2. Open the new app → **Authentication** → **Add a platform** → **Mobile and desktop applications**. Add both redirect URIs:
   - `ms-appx-web://microsoft.aad.brokerplugin/<Application (client) ID>`
   - `http://localhost`
3. Still under **Authentication**, set **Allow public client flows** to **Yes**.
4. Under **API permissions**, keep **Microsoft Graph → User.Read (delegated)**. After the Deep Check gateway is registered (see
   `deep-check-setup.md`), add **APIs my organization uses → LL-Docx Gateway → DeepCheck.Use (delegated)**.
5. Click **Grant admin consent for lozaip.com**, so users aren't asked to consent.
6. Optional, to control access per person: go to **Enterprise applications** → **LL-Docx Add-in** → **Properties**, set
   **Assignment required?** to **Yes**, then under **Users and groups** assign the security group of people allowed to use LL-Docx.

## 2. Send these values to the developer

They are compiled into the signed add-in (`LLDOCX.Auth/GateConfig.cs`, class `GateDefaults`):

| Value | Where to find it |
|---|---|
| Directory (tenant) ID | The app's **Overview** page |
| Application (client) ID | The app's **Overview** page |
| Gateway URL and scope | From `deep-check-setup.md`, when Deep Check is set up |

None of these are secrets.

## 3. Code signing

Provide a code-signing certificate issued to lozaip.com (an organization-validated or EV certificate from a public certificate
authority). Release builds of the add-in and the MSI are signed with it.

## 4. Distribute the installer

- **Interactive:** users with administrator rights double-click `LLDOCXSetup.msi`.
- **Silent:** `msiexec /i LLDOCXSetup.msi /qn`. This works from a script, Group Policy software installation, or any deployment tool.
- **Updates:** run a newer MSI. It replaces the old version in place, and users stay signed in.
- **Requirements:** 64-bit Windows 10 or 11, Word 2016 or later (standalone or Microsoft 365), .NET Framework 4.7.2+, and the VSTO
  runtime (normally installed with Office).

## 5. Optional Group Policy overrides

These are admin-only values under `HKLM\SOFTWARE\Policies\LL-Docx`. The tenant ID can never be overridden.

| Name | Type | Effect |
|---|---|---|
| `ClientId` | REG_SZ | A different add-in app registration (GUID) |
| `GatewayUrl` | REG_SZ | A different Deep Check gateway (must be `https://`) |
| `GatewayScope` | REG_SZ | A different gateway scope |
| `OfflineGraceDays` | REG_DWORD | Shortens the offline allowance (1–30 days; it can't be lengthened) |

## What users see

- **First use:** Word shows a **Sign in** button on the Home tab. If the person's lozaip.com account is already signed in to Windows,
  signing in usually takes one click. Otherwise they see the Microsoft sign-in page once, including MFA.
- **Offline:** the tools keep working for up to 30 days after the last successful check.
- **Leavers:** disabling the person's account locks LL-Docx at their next Word start, or within 24 hours. Deep Check stops at once.
```

- [ ] **Step 7: Commit**

```bash
git add LLDOCX/LLDOCXSetup LLDOCX/docs/admin/sign-in-setup.md
git commit -m "build(lldocx): per-machine x64 MSI with HKLM registration and prerequisite checks

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

## Spec updates made with this plan

These changes were made to the spec while writing this plan, and the spec file has been updated to match:

- **New `Checking` state.** The gate starts as Locked/`Checking` ("Verifying your lozaip.com account…") until the first silent check finishes. While locked for being offline too long, or while still checking, clicking a gated button retries the check. While unlocked offline, the gate retries the online check on a button click, at most every 5 minutes.
- **Sign-in scope.** Sign-in requests the gateway scope when a gateway is configured, and `User.Read` otherwise.
- **x64 installer.** The installer is an x64 package so it can register in both registry views. It therefore requires 64-bit Windows.
- **Manifest signing.** The add-in's tamper protection comes from signing the VSTO manifests, whose hashes cover the add-in DLLs.
