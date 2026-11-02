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
                catch (OperationCanceledException) when (ct.IsCancellationRequested)
                {
                    return _state;
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
            await _lock.WaitAsync().ConfigureAwait(false);
            try
            {
                if (_tokens != null)
                {
                    try { await _tokens.SignOutAsync().ConfigureAwait(false); }
                    catch (Exception ex) { _log?.Invoke("Sign-out failed: " + ex); }
                }
                SafeClear();
                _lastOnlineCheckUtc = null;
                _lastAttemptUtc = null;
                SetState(GateState.Locked(_config.IsConfigured ? LockReason.SignInRequired : LockReason.NotConfigured));
            }
            finally
            {
                _lock.Release();
            }
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
