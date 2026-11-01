using System.Diagnostics;
using System.Globalization;
using System.Text.Json;
using LLDOCX.DeepCheck;
using Microsoft.Extensions.Logging;

namespace LLDOCX.Gateway;

public sealed class HandlerResult
{
    public HandlerResult(int status, string body, int? retryAfterSeconds = null)
    {
        Status = status;
        Body = body;
        RetryAfterSeconds = retryAfterSeconds;
    }

    public int Status { get; }
    public string Body { get; }
    public int? RetryAfterSeconds { get; }

    public static HandlerResult Error(int status, string code, string message, int? retryAfterSeconds = null) =>
        new(status, JsonSerializer.Serialize(new ErrorBody { Error = new ErrorDetail { Code = code, Message = message } }), retryAfterSeconds);
}

/// <summary>Everything the HTTP function does, without the HTTP plumbing, so it can be tested.</summary>
public sealed class DeepCheckHandler
{
    readonly GatewayOptions _options;
    readonly ITokenValidator _validator;
    readonly UserRateLimiter _limiter;
    readonly IClaudeReviewer _reviewer;
    readonly ILogger<DeepCheckHandler> _log;

    public DeepCheckHandler(GatewayOptions options, ITokenValidator validator, UserRateLimiter limiter,
        IClaudeReviewer reviewer, ILogger<DeepCheckHandler> log)
    {
        _options = options;
        _validator = validator;
        _limiter = limiter;
        _reviewer = reviewer;
        _log = log;
    }

    public async Task<HandlerResult> HandleAsync(string authorizationHeader, string body, CancellationToken ct)
    {
        var started = Stopwatch.StartNew();

        CallerIdentity caller;
        try
        {
            caller = await _validator.ValidateAsync(authorizationHeader, ct);
        }
        catch (AuthFailure ex)
        {
            _log.LogWarning("Deep Check rejected: {Kind} ({Reason})", ex.Kind, ex.Message);
            return ex.Kind == AuthFailureKind.Forbidden
                ? HandlerResult.Error(403, ErrorCodes.Forbidden, "Your account isn't authorized for Deep Check.")
                : HandlerResult.Error(401, ErrorCodes.Unauthorized, "Sign-in required.");
        }
        catch (Exception ex) when (!ct.IsCancellationRequested)
        {
            _log.LogError("Deep Check token validation failed unexpectedly: {ErrorType}", ex.GetType().Name);
            return HandlerResult.Error(503, ErrorCodes.UpstreamUnavailable, "Sign-in could not be verified. Try again shortly.");
        }

        DeepCheckRequest request;
        try
        {
            request = JsonSerializer.Deserialize<DeepCheckRequest>(body ?? "");
        }
        catch (JsonException)
        {
            request = null;
        }
        if (request?.Claims == null || request.Claims.Count == 0)
            return Done(caller, started, 0, 0, "invalid_request",
                HandlerResult.Error(400, ErrorCodes.InvalidRequest, "The request contained no claims."));
        request.RuleFindings ??= new List<RuleFindingDto>();

        int count = request.Claims.Count;
        int chars = DeepCheckRequestBuilder.TextLength(request);
        if (chars > _options.MaxInputChars)
            return Done(caller, started, count, chars, "too_large", HandlerResult.Error(400, ErrorCodes.TooLarge,
                string.Format(CultureInfo.InvariantCulture, "The claims are too long for Deep Check (limit {0:N0} characters).", _options.MaxInputChars)));

        if (!_limiter.TryAcquire(caller.ObjectId, out TimeSpan retryAfter))
            return Done(caller, started, count, chars, "rate_limited", HandlerResult.Error(429, ErrorCodes.RateLimited,
                "Deep Check limit reached.", Math.Max(1, (int)Math.Ceiling(retryAfter.TotalSeconds))));

        ReviewResult review;
        using var timeout = CancellationTokenSource.CreateLinkedTokenSource(ct);
        timeout.CancelAfter(TimeSpan.FromSeconds(_options.ClaudeTimeoutSeconds));
        try
        {
            review = await _reviewer.ReviewAsync(DeepCheckPrompt.SystemPrompt, DeepCheckPrompt.BuildUserMessage(request),
                DeepCheckPrompt.ResponseSchemaJson, timeout.Token);
        }
        catch (OperationCanceledException) when (!ct.IsCancellationRequested)
        {
            return Done(caller, started, count, chars, "timeout", HandlerResult.Error(504, ErrorCodes.Timeout,
                "Claude took too long to review these claims. Try checking a selection of claims."));
        }
        catch (UpstreamException ex)
        {
            // Log only exception types: inner messages could echo request content.
            _log.LogError("Deep Check upstream failure: {Error} / {Inner}", ex.Message, ex.InnerException?.GetType().Name);
            return Done(caller, started, count, chars, "upstream_unavailable", HandlerResult.Error(503, ErrorCodes.UpstreamUnavailable,
                "The Claude service is unavailable. Try again shortly."));
        }
        catch (Exception ex) when (!ct.IsCancellationRequested)
        {
            // Type name only: messages could echo request content.
            _log.LogError("Deep Check unexpected failure: {ErrorType}", ex.GetType().Name);
            return Done(caller, started, count, chars, "upstream_unavailable", HandlerResult.Error(503, ErrorCodes.UpstreamUnavailable,
                "The Claude service is unavailable. Try again shortly."));
        }

        if (review.Outcome == ReviewOutcome.Refused)
            return Done(caller, started, count, chars, "refused",
                HandlerResult.Error(422, ErrorCodes.Refused, "Claude declined to review these claims."));
        if (review.Outcome == ReviewOutcome.Truncated)
            return Done(caller, started, count, chars, "truncated", HandlerResult.Error(502, ErrorCodes.Truncated,
                "Claude's review was cut off because the claim set is very large. Try checking a selection of claims."));
        if (!DeepCheckResponseParser.IsWellFormed(review.Json))
            return Done(caller, started, count, chars, "bad_upstream_response", HandlerResult.Error(502, ErrorCodes.BadUpstreamResponse,
                "Claude returned a response that could not be read."));

        return Done(caller, started, count, chars, "ok", new HandlerResult(200, review.Json));
    }

    HandlerResult Done(CallerIdentity caller, Stopwatch started, int claims, int chars, string outcome, HandlerResult result)
    {
        // Metadata only. Claim text and model output are confidential and are never logged.
        _log.LogInformation("Deep Check {Outcome} for {User} ({ObjectId}): {Claims} claims, {Chars} chars, {ElapsedMs} ms",
            outcome, caller.Username, caller.ObjectId, claims, chars, started.ElapsedMilliseconds);
        return result;
    }
}
