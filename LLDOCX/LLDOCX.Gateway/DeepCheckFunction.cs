using System.Net;
using Microsoft.Azure.Functions.Worker;
using Microsoft.Azure.Functions.Worker.Http;

namespace LLDOCX.Gateway;

public sealed class DeepCheckFunction
{
    readonly DeepCheckHandler _handler;

    public DeepCheckFunction(DeepCheckHandler handler) => _handler = handler;

    // Anonymous at the Functions layer: DeepCheckHandler validates the caller's Entra token itself.
    [Function("deepcheck")]
    public async Task<HttpResponseData> Run(
        [HttpTrigger(AuthorizationLevel.Anonymous, "post", Route = "deepcheck")] HttpRequestData req,
        CancellationToken ct)
    {
        string auth = req.Headers.TryGetValues("Authorization", out var values) ? values.FirstOrDefault() : null;
        string body;
        using (var reader = new StreamReader(req.Body))
            body = await reader.ReadToEndAsync(ct);

        HandlerResult result = await _handler.HandleAsync(auth, body, ct);
        var response = req.CreateResponse((HttpStatusCode)result.Status);
        response.Headers.Add("Content-Type", "application/json; charset=utf-8");
        if (result.RetryAfterSeconds.HasValue) response.Headers.Add("Retry-After", result.RetryAfterSeconds.Value.ToString());
        await response.WriteStringAsync(result.Body, ct);
        return response;
    }
}
