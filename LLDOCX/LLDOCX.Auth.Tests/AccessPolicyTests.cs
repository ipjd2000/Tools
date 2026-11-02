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
