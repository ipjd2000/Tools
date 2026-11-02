using System.Text;
using System.Text.Json;
using Anthropic;
using Anthropic.Exceptions;
using Anthropic.Foundry;
using Anthropic.Models.Messages;
using Azure.Identity;

namespace LLDOCX.Gateway;

public enum ReviewOutcome { Ok, Refused, Truncated }

public sealed class ReviewResult
{
    ReviewResult(ReviewOutcome outcome, string json)
    {
        Outcome = outcome;
        Json = json;
    }

    public ReviewOutcome Outcome { get; }
    public string Json { get; }

    public static ReviewResult Ok(string json) => new(ReviewOutcome.Ok, json);
    public static ReviewResult Refused() => new(ReviewOutcome.Refused, null);
    public static ReviewResult Truncated() => new(ReviewOutcome.Truncated, null);
}

/// <summary>Claude could not be reached or returned an error. Message never contains claim text.</summary>
public sealed class UpstreamException : Exception
{
    public UpstreamException(string message, Exception inner) : base(message, inner) { }
}

public interface IClaudeReviewer
{
    Task<ReviewResult> ReviewAsync(string systemPrompt, string userMessage, string responseSchemaJson, CancellationToken ct);
}

public sealed class SdkClaudeReviewer : IClaudeReviewer
{
    readonly Func<MessageCreateParams, CancellationToken, Task<Message>> _create;
    readonly string _model;
    readonly string _effort;
    readonly int _maxTokens;

    public SdkClaudeReviewer(Func<MessageCreateParams, CancellationToken, Task<Message>> create, string model, string effort, int maxTokens)
    {
        _create = create;
        _model = model;
        _effort = effort;
        _maxTokens = maxTokens;
    }

    public static SdkClaudeReviewer Create(GatewayOptions o)
    {
        if (o.ClaudeProvider == "Anthropic")
        {
            var client = new AnthropicClient { ApiKey = o.AnthropicApiKey };
            return new SdkClaudeReviewer((p, ct) => client.Messages.Create(p, cancellationToken: ct), o.Model, o.Effort, o.MaxOutputTokens);
        }

        // Managed identity when running in Azure; developer credentials (az login) when running locally.
        var foundry = new AnthropicFoundryClient(new AnthropicFoundryIdentityTokenCredentials(new DefaultAzureCredential(), o.FoundryResource));
        return new SdkClaudeReviewer((p, ct) => foundry.Messages.Create(p, cancellationToken: ct), o.Model, o.Effort, o.MaxOutputTokens);
    }

    public async Task<ReviewResult> ReviewAsync(string systemPrompt, string userMessage, string responseSchemaJson, CancellationToken ct)
    {
        var schema = JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(responseSchemaJson);
        var parameters = new MessageCreateParams
        {
            Model = _model,
            MaxTokens = _maxTokens,
            System = systemPrompt,
            OutputConfig = new OutputConfig
            {
                Effort = ParseEffort(_effort),
                Format = new JsonOutputFormat { Schema = schema },
            },
            Messages = [new UserMessageParam { Content = userMessage }],
        };

        Message response;
        try
        {
            response = await _create(parameters, ct);
        }
        catch (AnthropicException ex)
        {
            throw new UpstreamException("Claude request failed (" + ex.GetType().Name + ").", ex);
        }

        if (response.StopReason == "refusal") return ReviewResult.Refused();
        if (response.StopReason == "max_tokens") return ReviewResult.Truncated();

        var text = new StringBuilder();
        foreach (var block in response.Content)
            if (block.TryPickText(out var t)) text.Append(t.Text);
        return ReviewResult.Ok(text.ToString());
    }

    static Effort ParseEffort(string value) => (value ?? "").Trim().ToLowerInvariant() switch
    {
        "low" => Effort.Low,
        "medium" => Effort.Medium,
        "max" => Effort.Max,
        _ => Effort.High,
    };
}
