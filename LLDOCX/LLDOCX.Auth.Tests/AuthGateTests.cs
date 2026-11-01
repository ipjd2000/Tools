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
