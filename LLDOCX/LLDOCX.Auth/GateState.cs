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
