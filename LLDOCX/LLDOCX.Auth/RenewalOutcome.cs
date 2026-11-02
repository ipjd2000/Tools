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
