namespace LLDOCX.Gateway;

/// <summary>Gateway settings, bound from the "Gateway" configuration section (app settings named Gateway__Name).</summary>
public sealed class GatewayOptions
{
    public string TenantId { get; set; }

    /// <summary>Comma-separated accepted audiences: the gateway's App ID URI and/or client ID.</summary>
    public string Audience { get; set; }

    public string RequiredScope { get; set; } = "DeepCheck.Use";

    /// <summary>Optional app role the caller must hold (for example "DeepCheck.User").</summary>
    public string RequiredRole { get; set; }

    /// <summary>"Foundry" (managed identity, no key) or "Anthropic" (API key from Key Vault).</summary>
    public string ClaudeProvider { get; set; } = "Foundry";

    public string FoundryResource { get; set; }
    public string AnthropicApiKey { get; set; }

    /// <summary>Model ID, or the Foundry deployment name.</summary>
    public string Model { get; set; } = "claude-opus-5-5";

    public string Effort { get; set; } = "high";
    public int MaxOutputTokens { get; set; } = 16000;
    public int MaxInputChars { get; set; } = 200000;
    public int RequestsPerUserPerHour { get; set; } = 30;

    /// <summary>Kept under Azure's 230-second HTTP response limit.</summary>
    public int ClaudeTimeoutSeconds { get; set; } = 210;

    public IReadOnlyList<string> Audiences =>
        (Audience ?? "").Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries);

    public void Validate()
    {
        var errors = new List<string>();
        if (!Guid.TryParse(TenantId, out _)) errors.Add("Gateway:TenantId must be the lozaip.com tenant GUID.");
        if (Audiences.Count == 0) errors.Add("Gateway:Audience is required.");
        switch (ClaudeProvider)
        {
            case "Foundry":
                if (string.IsNullOrWhiteSpace(FoundryResource)) errors.Add("Gateway:FoundryResource is required for the Foundry provider.");
                break;
            case "Anthropic":
                if (string.IsNullOrWhiteSpace(AnthropicApiKey)) errors.Add("Gateway:AnthropicApiKey is required for the Anthropic provider.");
                break;
            default:
                errors.Add("Gateway:ClaudeProvider must be Foundry or Anthropic.");
                break;
        }
        if (errors.Count > 0) throw new InvalidOperationException(string.Join(" ", errors));
    }
}
