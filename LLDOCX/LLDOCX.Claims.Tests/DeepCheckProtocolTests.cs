using System.Net;
using System.Text.Json;
using LLDOCX.Claims;
using LLDOCX.DeepCheck;

namespace LLDOCX.Claims.Tests;

internal sealed class StubHandler : HttpMessageHandler
{
    public Func<HttpRequestMessage, HttpResponseMessage> Respond { get; set; }
    public HttpRequestMessage LastRequest { get; private set; }
    public string LastBody { get; private set; }

    protected override async Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken ct)
    {
        LastRequest = request;
        LastBody = request.Content == null ? null : await request.Content.ReadAsStringAsync(ct);
        return Respond(request);
    }
}

[TestClass]
public class DeepCheckProtocolTests
{
    const string Line = "1. A lamp comprising a base, wherein the shade is green.";
    const string Url = "https://gw.example.net/api/deepcheck";
    const string OkJson = """{"findings":[{"claim":1,"type":"Terminology","severity":"Warning","quote":"the shade","message":"m","suggestion":"s"}],"disputes":[{"index":0,"reason":"r"}]}""";

    static HttpResponseMessage Json(HttpStatusCode status, string body) =>
        new(status) { Content = new StringContent(body, System.Text.Encoding.UTF8, "application/json") };

    // ---- request ----

    [TestMethod]
    public void Request_lists_claims_canceled_claims_and_rule_findings_in_order()
    {
        var claims = T.Claims(Line, "2. (Canceled)");
        var rule = new List<Finding> { new Finding { ClaimNumber = 1, Type = FindingType.Antecedent, Severity = Severity.Error, Message = "No antecedent basis for 'the shade'." } };
        var r = DeepCheckRequestBuilder.Build(claims, rule);
        Assert.AreEqual(2, r.Claims.Count);
        StringAssert.StartsWith(r.Claims[0].Text, "A lamp comprising a base");
        Assert.IsTrue(r.Claims[1].Canceled);
        Assert.AreEqual("", r.Claims[1].Text);
        Assert.AreEqual("Antecedent", r.RuleFindings[0].Type);
        Assert.AreEqual("Error", r.RuleFindings[0].Severity);
        Assert.AreEqual(r.Claims[0].Text.Length, DeepCheckRequestBuilder.TextLength(r));
    }

    [TestMethod]
    public void Request_serializes_with_camel_case_names()
    {
        string json = JsonSerializer.Serialize(DeepCheckRequestBuilder.Build(T.Claims(Line), new List<Finding>()));
        StringAssert.Contains(json, "\"claims\":[{\"number\":1,");
        StringAssert.Contains(json, "\"ruleFindings\":[]");
    }

    // ---- response ----

    [TestMethod]
    public void Parser_maps_quotes_to_document_offsets_and_applies_disputes()
    {
        var claims = T.Claims(Line);
        var rule = new List<Finding> { new Finding { ClaimNumber = 1, Message = "x" } };
        var added = DeepCheckResponseParser.Parse(OkJson, claims, rule);
        Assert.AreEqual(1, added.Count);
        Assert.AreEqual(Line.IndexOf("the shade"), added[0].DocStart);
        Assert.AreEqual("the shade", added[0].Text);
        Assert.AreEqual("AI", added[0].Source);
        Assert.AreEqual(FindingType.Terminology, added[0].Type);
        Assert.AreEqual(Severity.Warning, added[0].Severity);
        Assert.IsTrue(rule[0].DisputedByAI);
        Assert.AreEqual("r", rule[0].DisputeReason);
    }

    [TestMethod]
    public void Parser_keeps_findings_whose_quote_is_not_found_without_location()
    {
        const string json = """{"findings":[{"claim":1,"type":"Antecedent","severity":"Error","quote":"not there","message":"m","suggestion":""}],"disputes":[]}""";
        var added = DeepCheckResponseParser.Parse(json, T.Claims(Line), new List<Finding>());
        Assert.AreEqual(-1, added[0].DocStart);
        Assert.IsNull(added[0].Suggestion);
    }

    [TestMethod]
    public void Parser_ignores_out_of_range_dispute_indexes()
    {
        var rule = new List<Finding> { new Finding() };
        DeepCheckResponseParser.Parse("""{"findings":[],"disputes":[{"index":5,"reason":"r"}]}""", T.Claims(Line), rule);
        Assert.IsFalse(rule[0].DisputedByAI);
    }

    [TestMethod]
    public void Parser_rejects_invalid_json()
    {
        Assert.ThrowsException<DeepCheckException>(() =>
            DeepCheckResponseParser.Parse("not json", T.Claims(Line), new List<Finding>()));
    }

    [DataTestMethod]
    [DataRow(OkJson, true)]
    [DataRow("""{"findings":[],"disputes":[]}""", true)]
    [DataRow("""{"findings":[]}""", false)]
    [DataRow("""{"findings":[{"claim":"1","type":"Antecedent","severity":"Error","quote":"q","message":"m","suggestion":""}],"disputes":[]}""", false)]
    [DataRow("""{"findings":[],"disputes":[{"index":0}]}""", false)]
    [DataRow("[]", false)]
    [DataRow("not json", false)]
    public void IsWellFormed_checks_the_response_shape(string json, bool expected) =>
        Assert.AreEqual(expected, DeepCheckResponseParser.IsWellFormed(json));


    [TestMethod]
    public void Non_integer_claim_is_not_well_formed_and_parse_throws_deepcheck_exception()
    {
        const string json = """{"findings":[{"claim":1.5,"type":"Antecedent","severity":"Error","quote":"q","message":"m","suggestion":""}],"disputes":[]}""";
        Assert.IsFalse(DeepCheckResponseParser.IsWellFormed(json));
        Assert.ThrowsException<DeepCheckException>(() =>
            DeepCheckResponseParser.Parse(json, T.Claims(Line), new List<Finding>()));
    }

    [TestMethod]
    public void Non_integer_dispute_index_is_not_well_formed()
    {
        const string json = """{"findings":[],"disputes":[{"index":0.5,"reason":"r"}]}""";
        Assert.IsFalse(DeepCheckResponseParser.IsWellFormed(json));
        Assert.ThrowsException<DeepCheckException>(() =>
            DeepCheckResponseParser.Parse(json, T.Claims(Line), new List<Finding> { new Finding { ClaimNumber = 1, Message = "x" } }));
    }

    [TestMethod]
    public void Non_object_elements_are_skipped_by_parse_and_rejected_by_is_well_formed()
    {
        const string json = """{"findings":[5,{"claim":1,"type":"Terminology","severity":"Warning","quote":"the shade","message":"m","suggestion":"s"}],"disputes":["x"]}""";
        Assert.IsFalse(DeepCheckResponseParser.IsWellFormed(json));
        var added = DeepCheckResponseParser.Parse(json, T.Claims(Line), new List<Finding>());
        Assert.AreEqual(1, added.Count);
    }
    // ---- HTTP client ----

    static (DeepCheckHttpClient client, StubHandler handler) NewClient(Func<HttpRequestMessage, HttpResponseMessage> respond)
    {
        var handler = new StubHandler { Respond = respond };
        return (new DeepCheckHttpClient(new HttpClient(handler)), handler);
    }

    [TestMethod]
    public async Task Client_posts_json_with_bearer_token_and_parses_findings()
    {
        var (client, handler) = NewClient(_ => Json(HttpStatusCode.OK, OkJson));
        var rule = new List<Finding> { new Finding { ClaimNumber = 1, Message = "x" } };
        var added = await client.RunAsync(Url, "tok123", T.Claims(Line), rule, CancellationToken.None);
        Assert.AreEqual(1, added.Count);
        Assert.AreEqual(HttpMethod.Post, handler.LastRequest.Method);
        Assert.AreEqual(Url, handler.LastRequest.RequestUri.ToString());
        Assert.AreEqual("Bearer", handler.LastRequest.Headers.Authorization.Scheme);
        Assert.AreEqual("tok123", handler.LastRequest.Headers.Authorization.Parameter);
        StringAssert.Contains(handler.LastBody, "\"ruleFindings\":[{\"claim\":1,");
    }

    static async Task<string> FailureMessage(HttpResponseMessage response)
    {
        var (client, _) = NewClient(_ => response);
        var ex = await Assert.ThrowsExceptionAsync<DeepCheckException>(() =>
            client.RunAsync(Url, "t", T.Claims(Line), new List<Finding>(), CancellationToken.None));
        return ex.Message;
    }

    [TestMethod]
    public async Task Client_maps_401_and_403_to_not_authorized()
    {
        Assert.AreEqual("Your account isn't authorized for Deep Check. Contact your administrator.",
            await FailureMessage(Json(HttpStatusCode.Unauthorized, "{}")));
        Assert.AreEqual("Your account isn't authorized for Deep Check. Contact your administrator.",
            await FailureMessage(Json(HttpStatusCode.Forbidden, "{}")));
    }

    [TestMethod]
    public async Task Client_maps_429_using_retry_after()
    {
        var response = Json((HttpStatusCode)429, """{"error":{"code":"rate_limited","message":"x"}}""");
        response.Headers.RetryAfter = new System.Net.Http.Headers.RetryConditionHeaderValue(TimeSpan.FromSeconds(90));
        Assert.AreEqual("Deep Check limit reached; try again in 2 minutes.", await FailureMessage(response));
    }

    [TestMethod]
    public async Task Client_maps_422_to_declined()
    {
        Assert.AreEqual("Claude declined to review these claims. The rule-based findings are unchanged.",
            await FailureMessage(Json((HttpStatusCode)422, """{"error":{"code":"refused","message":"x"}}""")));
    }

    [TestMethod]
    public async Task Client_shows_the_gateway_message_for_400()
    {
        Assert.AreEqual("The claims are too long for Deep Check (limit 200,000 characters). The rule-based findings are unchanged.",
            await FailureMessage(Json(HttpStatusCode.BadRequest,
                """{"error":{"code":"too_large","message":"The claims are too long for Deep Check (limit 200,000 characters)."}}""")));
    }

    [TestMethod]
    public async Task Client_uses_a_generic_message_for_5xx_without_body()
    {
        Assert.AreEqual("The Deep Check service had a problem. Try again shortly. The rule-based findings are unchanged.",
            await FailureMessage(new HttpResponseMessage(HttpStatusCode.ServiceUnavailable)));
    }

    [TestMethod]
    public async Task Client_maps_network_failure()
    {
        Assert.AreEqual("Could not reach the Deep Check service. Check your network connection.",
            await FailureFromThrow(new HttpRequestException("no route")));
    }

    [TestMethod]
    public async Task Client_maps_timeout()
    {
        Assert.AreEqual("Deep Check took too long and was stopped. Try checking a selection of claims.",
            await FailureFromThrow(new TaskCanceledException("timeout")));
    }

    [TestMethod]
    public async Task Client_rethrows_user_cancellation()
    {
        using var cts = new CancellationTokenSource();
        cts.Cancel();
        var (client, _) = NewClient(_ => throw new TaskCanceledException());
        try
        {
            await client.RunAsync(Url, "t", T.Claims(Line), new List<Finding>(), cts.Token);
            Assert.Fail("Expected cancellation.");
        }
        catch (OperationCanceledException)
        {
            // HttpClient may surface TaskCanceledException or OperationCanceledException; both mean the user canceled.
        }
    }

    static async Task<string> FailureFromThrow(Exception toThrow)
    {
        var (client, _) = NewClient(_ => throw toThrow);
        var ex = await Assert.ThrowsExceptionAsync<DeepCheckException>(() =>
            client.RunAsync(Url, "t", T.Claims(Line), new List<Finding>(), CancellationToken.None));
        return ex.Message;
    }
}
