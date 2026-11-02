using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using LLDOCX.Claims;

namespace LLDOCX.DeepCheck
{
    /// <summary>Calls the lozaip.com Deep Check gateway. No model, prompt, or Claude credential lives on the client.</summary>
    public sealed class DeepCheckHttpClient
    {
        /// <summary>Just above the gateway's own 210-second Claude timeout (Azure stops HTTP requests at about 230 seconds).</summary>
        public static readonly TimeSpan DefaultTimeout = TimeSpan.FromSeconds(240);

        const string Kept = " The rule-based findings are unchanged.";
        readonly HttpClient _http;

        public DeepCheckHttpClient(HttpClient http)
        {
            _http = http ?? throw new ArgumentNullException(nameof(http));
        }

        public async Task<List<Finding>> RunAsync(string gatewayUrl, string bearerToken, IList<Claim> claims,
            IList<Finding> ruleFindings, CancellationToken ct)
        {
            string body = JsonSerializer.Serialize(DeepCheckRequestBuilder.Build(claims, ruleFindings));
            using (var request = new HttpRequestMessage(HttpMethod.Post, gatewayUrl))
            {
                request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", bearerToken);
                request.Content = new StringContent(body, Encoding.UTF8, "application/json");

                HttpResponseMessage response;
                try
                {
                    response = await _http.SendAsync(request, ct).ConfigureAwait(false);
                }
                catch (TaskCanceledException ex) when (!ct.IsCancellationRequested)
                {
                    throw new DeepCheckException("Deep Check took too long and was stopped. Try checking a selection of claims.", ex);
                }
                catch (HttpRequestException ex)
                {
                    throw new DeepCheckException("Could not reach the Deep Check service. Check your network connection.", ex);
                }

                using (response)
                {
                    string text = response.Content == null ? "" : await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                    if (response.IsSuccessStatusCode) return DeepCheckResponseParser.Parse(text, claims, ruleFindings);
                    throw new DeepCheckException(MessageFor((int)response.StatusCode, ServerMessage(text), response.Headers.RetryAfter?.Delta));
                }
            }
        }

        static string MessageFor(int status, string serverMessage, TimeSpan? retryAfter)
        {
            switch (status)
            {
                case 401:
                case 403:
                    return "Your account isn't authorized for Deep Check. Contact your administrator.";
                case 422:
                    return "Claude declined to review these claims." + Kept;
                case 429:
                    int minutes = retryAfter.HasValue ? Math.Max(1, (int)Math.Ceiling(retryAfter.Value.TotalMinutes)) : 60;
                    return "Deep Check limit reached; try again in " + minutes + (minutes == 1 ? " minute." : " minutes.");
                default:
                    if (!string.IsNullOrWhiteSpace(serverMessage)) return serverMessage + Kept;
                    return status >= 500
                        ? "The Deep Check service had a problem. Try again shortly." + Kept
                        : "Deep Check could not process these claims." + Kept;
            }
        }

        static string ServerMessage(string body)
        {
            if (string.IsNullOrWhiteSpace(body)) return null;
            try
            {
                return JsonSerializer.Deserialize<ErrorBody>(body)?.Error?.Message;
            }
            catch (JsonException)
            {
                return null;
            }
        }
    }
}
