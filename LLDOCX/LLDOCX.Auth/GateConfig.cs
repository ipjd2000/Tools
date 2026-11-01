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
