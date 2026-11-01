using Microsoft.IdentityModel.JsonWebTokens;
using Microsoft.IdentityModel.Protocols;
using Microsoft.IdentityModel.Protocols.OpenIdConnect;
using Microsoft.IdentityModel.Tokens;

namespace LLDOCX.Gateway;

public sealed class CallerIdentity
{
    public CallerIdentity(string objectId, string username)
    {
        ObjectId = objectId;
        Username = username;
    }

    public string ObjectId { get; }
    public string Username { get; }
}

public enum AuthFailureKind { Unauthenticated, Forbidden }

public sealed class AuthFailure : Exception
{
    public AuthFailure(AuthFailureKind kind, string message) : base(message) => Kind = kind;
    public AuthFailureKind Kind { get; }
}

public interface ITokenValidator
{
    /// <summary>Returns the caller, or throws AuthFailure.</summary>
    Task<CallerIdentity> ValidateAsync(string authorizationHeader, CancellationToken ct);
}

/// <summary>Accepts only Entra access tokens issued by the lozaip.com tenant for this gateway, with the DeepCheck.Use scope.</summary>
public sealed class EntraTokenValidator : ITokenValidator
{
    readonly GatewayOptions _options;
    readonly IConfigurationManager<OpenIdConnectConfiguration> _config;
    readonly JsonWebTokenHandler _handler = new();

    public EntraTokenValidator(GatewayOptions options, IConfigurationManager<OpenIdConnectConfiguration> config)
    {
        _options = options;
        _config = config;
    }

    public static IConfigurationManager<OpenIdConnectConfiguration> CreateConfigurationManager(string tenantId) =>
        new ConfigurationManager<OpenIdConnectConfiguration>(
            $"https://login.microsoftonline.com/{tenantId}/v2.0/.well-known/openid-configuration",
            new OpenIdConnectConfigurationRetriever(),
            new HttpDocumentRetriever());

    public async Task<CallerIdentity> ValidateAsync(string authorizationHeader, CancellationToken ct)
    {
        const string prefix = "Bearer ";
        if (string.IsNullOrEmpty(authorizationHeader) || !authorizationHeader.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            throw new AuthFailure(AuthFailureKind.Unauthenticated, "Missing bearer token.");
        string token = authorizationHeader.Substring(prefix.Length).Trim();

        var config = await _config.GetConfigurationAsync(ct);
        var parameters = new TokenValidationParameters
        {
            ValidIssuers = new[]
            {
                $"https://login.microsoftonline.com/{_options.TenantId}/v2.0",
                $"https://sts.windows.net/{_options.TenantId}/"
            },
            ValidAudiences = _options.Audiences,
            IssuerSigningKeys = config.SigningKeys,
            ValidateLifetime = true,
            ClockSkew = TimeSpan.FromMinutes(2)
        };
        TokenValidationResult result = await _handler.ValidateTokenAsync(token, parameters);
        if (!result.IsValid) throw new AuthFailure(AuthFailureKind.Unauthenticated, "Invalid token.");

        var identity = result.ClaimsIdentity;
        if (!string.Equals(identity.FindFirst("tid")?.Value, _options.TenantId, StringComparison.OrdinalIgnoreCase))
            throw new AuthFailure(AuthFailureKind.Unauthenticated, "Token is not from the lozaip.com tenant.");

        var scopes = (identity.FindFirst("scp")?.Value ?? "").Split(' ', StringSplitOptions.RemoveEmptyEntries);
        if (!scopes.Contains(_options.RequiredScope, StringComparer.Ordinal))
            throw new AuthFailure(AuthFailureKind.Forbidden, "Missing scope " + _options.RequiredScope + ".");

        if (!string.IsNullOrWhiteSpace(_options.RequiredRole) && !identity.FindAll("roles").Any(c => c.Value == _options.RequiredRole))
            throw new AuthFailure(AuthFailureKind.Forbidden, "Missing role " + _options.RequiredRole + ".");

        string oid = identity.FindFirst("oid")?.Value;
        if (string.IsNullOrEmpty(oid)) throw new AuthFailure(AuthFailureKind.Unauthenticated, "Token has no oid.");
        string username = identity.FindFirst("preferred_username")?.Value ?? identity.FindFirst("upn")?.Value ?? oid;
        return new CallerIdentity(oid, username);
    }
}
