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
