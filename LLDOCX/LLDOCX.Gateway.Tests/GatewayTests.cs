global using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Security.Cryptography;
using System.Text.Json;
using LLDOCX.DeepCheck;
using LLDOCX.Gateway;
using Microsoft.Extensions.Logging;
using Microsoft.IdentityModel.JsonWebTokens;
using Microsoft.IdentityModel.Protocols;
using Microsoft.IdentityModel.Protocols.OpenIdConnect;
using Microsoft.IdentityModel.Tokens;

namespace LLDOCX.Gateway.Tests;

internal sealed class CapturingLogger : ILogger<DeepCheckHandler>
{
    public List<string> Lines { get; } = new();
    public IDisposable BeginScope<TState>(TState state) where TState : notnull => null;
    public bool IsEnabled(LogLevel logLevel) => true;
    public void Log<TState>(LogLevel logLevel, EventId eventId, TState state, Exception exception, Func<TState, Exception, string> formatter) =>
        Lines.Add(formatter(state, exception) + " | " + state + " | " + exception);
}

internal sealed class FakeValidator : ITokenValidator
{
    public AuthFailure Failure { get; set; }
    public Task<CallerIdentity> ValidateAsync(string authorizationHeader, CancellationToken ct) =>
        Failure != null ? throw Failure : Task.FromResult(new CallerIdentity("oid-1", "pat@lozaip.com"));
}

internal sealed class FakeReviewer : IClaudeReviewer
{
    public Func<ReviewResult> Result { get; set; } = () => ReviewResult.Ok("""{"findings":[],"disputes":[]}""");
    public string LastUserMessage { get; private set; }

    public Task<ReviewResult> ReviewAsync(string systemPrompt, string userMessage, string responseSchemaJson, CancellationToken ct)
    {
        LastUserMessage = userMessage;
        return Task.FromResult(Result());
    }
}

[TestClass]
public class PromptTests
{
    [TestMethod]
    public void User_message_lists_claims_canceled_claims_and_indexed_rule_findings()
    {
        var request = new DeepCheckRequest
        {
            Claims = { new ClaimDto { Number = 1, Text = "A lamp comprising a base." }, new ClaimDto { Number = 2, Canceled = true, Text = "" } },
            RuleFindings = { new RuleFindingDto { Claim = 1, Type = "Antecedent", Severity = "Error", Message = "No antecedent basis for 'the shade'." } }
        };
        string msg = DeepCheckPrompt.BuildUserMessage(request);
        StringAssert.Contains(msg, "1. A lamp comprising a base.");
        StringAssert.Contains(msg, "2. (Canceled)");
        StringAssert.Contains(msg, "[0] claim 1, Antecedent, Error: No antecedent basis for 'the shade'.");
    }

    [TestMethod]
    public void Schema_is_valid_json_with_required_fields()
    {
        using var doc = JsonDocument.Parse(DeepCheckPrompt.ResponseSchemaJson);
        var required = doc.RootElement.GetProperty("required").EnumerateArray().Select(e => e.GetString()).ToArray();
        CollectionAssert.AreEquivalent(new[] { "findings", "disputes" }, required);
    }
}

[TestClass]
public class TokenValidatorTests
{
    const string Tenant = "11111111-1111-1111-1111-111111111111";
    const string Audience = "api://gateway-app";
    const string ClientIdAudience = "22222222-2222-2222-2222-222222222222";
    static readonly RsaSecurityKey Key = new(RSA.Create(2048)) { KeyId = "test-key" };

    static GatewayOptions Options(string role = null) =>
        new() { TenantId = Tenant, Audience = Audience + "," + ClientIdAudience, RequiredRole = role };

    static EntraTokenValidator Validator(GatewayOptions o)
    {
        var config = new OpenIdConnectConfiguration();
        config.SigningKeys.Add(Key);
        return new EntraTokenValidator(o, new StaticConfigurationManager<OpenIdConnectConfiguration>(config));
    }

    static string Token(Action<Dictionary<string, object>> edit = null, string issuer = null, string audience = Audience, DateTime? expires = null)
    {
        var claims = new Dictionary<string, object>
        {
            ["tid"] = Tenant,
            ["oid"] = "oid-1",
            ["preferred_username"] = "pat@lozaip.com",
            ["scp"] = "User.Read DeepCheck.Use"
        };
        edit?.Invoke(claims);
        DateTime exp = expires ?? DateTime.UtcNow.AddMinutes(30);
        return new JsonWebTokenHandler().CreateToken(new SecurityTokenDescriptor
        {
            Issuer = issuer ?? $"https://login.microsoftonline.com/{Tenant}/v2.0",
            Audience = audience,
            Claims = claims,
            NotBefore = exp.AddHours(-1),
            IssuedAt = exp.AddHours(-1),
            Expires = exp,
            SigningCredentials = new SigningCredentials(Key, SecurityAlgorithms.RsaSha256)
        });
    }

    static async Task<AuthFailureKind> Fails(EntraTokenValidator v, string header)
    {
        var ex = await Assert.ThrowsExceptionAsync<AuthFailure>(() => v.ValidateAsync(header, CancellationToken.None));
        return ex.Kind;
    }

    [TestMethod]
    public async Task Valid_token_returns_the_caller()
    {
        var caller = await Validator(Options()).ValidateAsync("Bearer " + Token(), CancellationToken.None);
        Assert.AreEqual("oid-1", caller.ObjectId);
        Assert.AreEqual("pat@lozaip.com", caller.Username);
    }

    [TestMethod]
    public async Task Client_id_audience_is_also_accepted()
    {
        var caller = await Validator(Options()).ValidateAsync("Bearer " + Token(audience: ClientIdAudience), CancellationToken.None);
        Assert.AreEqual("oid-1", caller.ObjectId);
    }

    [TestMethod]
    public async Task Missing_or_malformed_header_is_unauthenticated()
    {
        Assert.AreEqual(AuthFailureKind.Unauthenticated, await Fails(Validator(Options()), null));
        Assert.AreEqual(AuthFailureKind.Unauthenticated, await Fails(Validator(Options()), "Basic abc"));
        Assert.AreEqual(AuthFailureKind.Unauthenticated, await Fails(Validator(Options()), "Bearer not-a-jwt"));
    }

    [TestMethod]
    public async Task Other_tenant_is_unauthenticated()
    {
        const string other = "99999999-9999-9999-9999-999999999999";
        string token = Token(c => c["tid"] = other, issuer: $"https://login.microsoftonline.com/{other}/v2.0");
        Assert.AreEqual(AuthFailureKind.Unauthenticated, await Fails(Validator(Options()), "Bearer " + token));
    }

    [TestMethod]
    public async Task Wrong_audience_is_unauthenticated()
    {
        Assert.AreEqual(AuthFailureKind.Unauthenticated, await Fails(Validator(Options()), "Bearer " + Token(audience: "api://other")));
    }

    [TestMethod]
    public async Task Expired_token_is_unauthenticated()
    {
        Assert.AreEqual(AuthFailureKind.Unauthenticated,
            await Fails(Validator(Options()), "Bearer " + Token(expires: DateTime.UtcNow.AddMinutes(-10))));
    }

    [TestMethod]
    public async Task Token_signed_with_an_unknown_key_is_unauthenticated()
    {
        var otherKey = new RsaSecurityKey(RSA.Create(2048)) { KeyId = "other-key" };
        var claims = new Dictionary<string, object> { ["tid"] = Tenant, ["oid"] = "oid-1", ["scp"] = "DeepCheck.Use" };
        string token = new JsonWebTokenHandler().CreateToken(new SecurityTokenDescriptor
        {
            Issuer = $"https://login.microsoftonline.com/{Tenant}/v2.0",
            Audience = Audience,
            Claims = claims,
            Expires = DateTime.UtcNow.AddMinutes(30),
            SigningCredentials = new SigningCredentials(otherKey, SecurityAlgorithms.RsaSha256)
        });
        Assert.AreEqual(AuthFailureKind.Unauthenticated, await Fails(Validator(Options()), "Bearer " + token));
    }

    [TestMethod]
    public async Task Missing_scope_is_forbidden()
    {
        Assert.AreEqual(AuthFailureKind.Forbidden, await Fails(Validator(Options()), "Bearer " + Token(c => c["scp"] = "User.Read")));
    }

    [TestMethod]
    public async Task Required_role_is_enforced()
    {
        Assert.AreEqual(AuthFailureKind.Forbidden, await Fails(Validator(Options("DeepCheck.User")), "Bearer " + Token()));
        var caller = await Validator(Options("DeepCheck.User"))
            .ValidateAsync("Bearer " + Token(c => c["roles"] = new[] { "Other", "DeepCheck.User" }), CancellationToken.None);
        Assert.AreEqual("oid-1", caller.ObjectId);
    }
}

[TestClass]
public class RateLimiterTests
{
    [TestMethod]
    public void Allows_the_limit_then_blocks_with_retry_after()
    {
        var now = new DateTime(2026, 11, 1, 12, 0, 0, DateTimeKind.Utc);
        var limiter = new UserRateLimiter(2, () => now);
        Assert.IsTrue(limiter.TryAcquire("a", out _));
        now = now.AddMinutes(10);
        Assert.IsTrue(limiter.TryAcquire("a", out _));
        Assert.IsFalse(limiter.TryAcquire("a", out var retry));
        Assert.AreEqual(TimeSpan.FromMinutes(50), retry);
        Assert.IsTrue(limiter.TryAcquire("b", out _), "users are independent");
        now = now.AddMinutes(51);
        Assert.IsTrue(limiter.TryAcquire("a", out _), "the window slides");
    }
}

[TestClass]
public class DeepCheckHandlerTests
{
    const string Secret = "ZZ-CONFIDENTIAL-CLAIM-TEXT-ZZ";
    readonly FakeValidator _validator = new();
    readonly FakeReviewer _reviewer = new();
    readonly CapturingLogger _log = new();
    readonly GatewayOptions _options = new()
    {
        TenantId = "11111111-1111-1111-1111-111111111111",
        Audience = "api://gw",
        MaxInputChars = 1000,
        RequestsPerUserPerHour = 30
    };

    DeepCheckHandler NewHandler() =>
        new(_options, _validator, new UserRateLimiter(_options.RequestsPerUserPerHour, () => DateTime.UtcNow), _reviewer, _log);

    static string Body(string text = "A lamp comprising a base, wherein the shade is green. " + Secret) =>
        JsonSerializer.Serialize(new DeepCheckRequest { Claims = { new ClaimDto { Number = 1, Text = text } } });

    static ErrorDetail Error(HandlerResult r) => JsonSerializer.Deserialize<ErrorBody>(r.Body).Error;

    [TestMethod]
    public async Task Success_returns_claude_json_and_logs_metadata_only()
    {
        _reviewer.Result = () => ReviewResult.Ok("""{"findings":[{"claim":1,"type":"Antecedent","severity":"Error","quote":"the shade","message":"m","suggestion":""}],"disputes":[]}""");
        var r = await NewHandler().HandleAsync("Bearer x", Body(), CancellationToken.None);
        Assert.AreEqual(200, r.Status);
        Assert.IsTrue(DeepCheckResponseParser.IsWellFormed(r.Body));
        StringAssert.Contains(_reviewer.LastUserMessage, Secret);
        Assert.IsTrue(_log.Lines.Any(l => l.Contains("ok") && l.Contains("pat@lozaip.com")), string.Join("\n", _log.Lines));
        Assert.IsFalse(_log.Lines.Any(l => l.Contains(Secret) || l.Contains("the shade")), string.Join("\n", _log.Lines));
    }

    [TestMethod]
    public async Task Auth_failures_map_to_401_and_403()
    {
        _validator.Failure = new AuthFailure(AuthFailureKind.Unauthenticated, "x");
        Assert.AreEqual(401, (await NewHandler().HandleAsync(null, Body(), CancellationToken.None)).Status);
        _validator.Failure = new AuthFailure(AuthFailureKind.Forbidden, "x");
        var r = await NewHandler().HandleAsync("Bearer x", Body(), CancellationToken.None);
        Assert.AreEqual(403, r.Status);
        Assert.AreEqual(ErrorCodes.Forbidden, Error(r).Code);
    }

    [TestMethod]
    public async Task Invalid_or_empty_request_is_400()
    {
        Assert.AreEqual(ErrorCodes.InvalidRequest, Error(await NewHandler().HandleAsync("Bearer x", "not json", CancellationToken.None)).Code);
        Assert.AreEqual(ErrorCodes.InvalidRequest, Error(await NewHandler().HandleAsync("Bearer x", """{"claims":[]}""", CancellationToken.None)).Code);
    }

    [TestMethod]
    public async Task Too_large_is_400_with_the_limit_in_the_message()
    {
        var r = await NewHandler().HandleAsync("Bearer x", Body(new string('a', 1001)), CancellationToken.None);
        Assert.AreEqual(400, r.Status);
        Assert.AreEqual(ErrorCodes.TooLarge, Error(r).Code);
        StringAssert.Contains(Error(r).Message, "limit 1,000 characters");
    }

    [TestMethod]
    public async Task Rate_limit_is_429_with_retry_after()
    {
        _options.RequestsPerUserPerHour = 1;
        var handler = NewHandler();
        Assert.AreEqual(200, (await handler.HandleAsync("Bearer x", Body(), CancellationToken.None)).Status);
        var r = await handler.HandleAsync("Bearer x", Body(), CancellationToken.None);
        Assert.AreEqual(429, r.Status);
        Assert.IsTrue(r.RetryAfterSeconds > 3500);
    }

    [TestMethod]
    public async Task Refusal_is_422()
    {
        _reviewer.Result = ReviewResult.Refused;
        var r = await NewHandler().HandleAsync("Bearer x", Body(), CancellationToken.None);
        Assert.AreEqual(422, r.Status);
        Assert.AreEqual(ErrorCodes.Refused, Error(r).Code);
    }

    [TestMethod]
    public async Task Truncation_is_502()
    {
        _reviewer.Result = ReviewResult.Truncated;
        var r = await NewHandler().HandleAsync("Bearer x", Body(), CancellationToken.None);
        Assert.AreEqual(502, r.Status);
        Assert.AreEqual(ErrorCodes.Truncated, Error(r).Code);
    }

    [TestMethod]
    public async Task Upstream_failure_is_503_and_does_not_log_inner_messages()
    {
        _reviewer.Result = () => throw new UpstreamException("Claude request failed.", new InvalidOperationException(Secret));
        var r = await NewHandler().HandleAsync("Bearer x", Body(), CancellationToken.None);
        Assert.AreEqual(503, r.Status);
        Assert.IsFalse(_log.Lines.Any(l => l.Contains(Secret)), string.Join("\n", _log.Lines));
    }

    [TestMethod]
    public async Task Malformed_claude_output_is_502()
    {
        _reviewer.Result = () => ReviewResult.Ok("""{"findings":"nope"}""");
        Assert.AreEqual(ErrorCodes.BadUpstreamResponse, Error(await NewHandler().HandleAsync("Bearer x", Body(), CancellationToken.None)).Code);
    }

    [TestMethod]
    public async Task Unexpected_reviewer_exception_is_503_and_does_not_log_its_message()
    {
        _reviewer.Result = () => throw new InvalidOperationException(Secret);
        var r = await NewHandler().HandleAsync("Bearer x", Body(), CancellationToken.None);
        Assert.AreEqual(503, r.Status);
        Assert.AreEqual(ErrorCodes.UpstreamUnavailable, Error(r).Code);
        Assert.IsFalse(_log.Lines.Any(l => l.Contains(Secret)), string.Join("\n", _log.Lines));
    }

    sealed class ThrowingValidator : ITokenValidator
    {
        public Task<CallerIdentity> ValidateAsync(string authorizationHeader, CancellationToken ct) =>
            throw new HttpRequestException("metadata fetch failed " + Secret);
    }

    [TestMethod]
    public async Task Unexpected_validator_exception_is_503()
    {
        var handler = new DeepCheckHandler(_options, new ThrowingValidator(), new UserRateLimiter(30, () => DateTime.UtcNow), _reviewer, _log);
        var r = await handler.HandleAsync("Bearer x", Body(), CancellationToken.None);
        Assert.AreEqual(503, r.Status);
        Assert.IsFalse(_log.Lines.Any(l => l.Contains(Secret)), string.Join("\n", _log.Lines));
    }

    [TestMethod]
    public void Options_validation_reports_missing_settings()
    {
        var ex = Assert.ThrowsException<InvalidOperationException>(() => new GatewayOptions().Validate());
        StringAssert.Contains(ex.Message, "TenantId");
        StringAssert.Contains(ex.Message, "Audience");
        StringAssert.Contains(ex.Message, "FoundryResource");
    }
}
