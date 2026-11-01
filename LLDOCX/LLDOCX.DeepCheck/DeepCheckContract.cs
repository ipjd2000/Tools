using System.Collections.Generic;
using System.Text.Json.Serialization;

namespace LLDOCX.DeepCheck
{
    /// <summary>Body of POST {gatewayUrl}. Shared by the add-in (sender) and the gateway (receiver).</summary>
    public sealed class DeepCheckRequest
    {
        [JsonPropertyName("claims")] public List<ClaimDto> Claims { get; set; } = new List<ClaimDto>();
        [JsonPropertyName("ruleFindings")] public List<RuleFindingDto> RuleFindings { get; set; } = new List<RuleFindingDto>();
    }

    public sealed class ClaimDto
    {
        [JsonPropertyName("number")] public int Number { get; set; }
        [JsonPropertyName("text")] public string Text { get; set; }
        [JsonPropertyName("canceled")] public bool Canceled { get; set; }
    }

    /// <summary>An offline finding. Its position in the list is the index Claude uses for disputes.</summary>
    public sealed class RuleFindingDto
    {
        [JsonPropertyName("claim")] public int Claim { get; set; }
        [JsonPropertyName("type")] public string Type { get; set; }
        [JsonPropertyName("severity")] public string Severity { get; set; }
        [JsonPropertyName("message")] public string Message { get; set; }
    }

    /// <summary>Body of every non-200 gateway response.</summary>
    public sealed class ErrorBody
    {
        [JsonPropertyName("error")] public ErrorDetail Error { get; set; }
    }

    public sealed class ErrorDetail
    {
        [JsonPropertyName("code")] public string Code { get; set; }
        [JsonPropertyName("message")] public string Message { get; set; }
    }

    public static class ErrorCodes
    {
        public const string InvalidRequest = "invalid_request";
        public const string TooLarge = "too_large";
        public const string Unauthorized = "unauthorized";
        public const string Forbidden = "forbidden";
        public const string RateLimited = "rate_limited";
        public const string Refused = "refused";
        public const string Truncated = "truncated";
        public const string Timeout = "timeout";
        public const string UpstreamUnavailable = "upstream_unavailable";
        public const string BadUpstreamResponse = "bad_upstream_response";
    }
}
