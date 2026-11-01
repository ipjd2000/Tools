using LLDOCX.Auth;

namespace LLDOCX.Auth.Tests;

[TestClass]
public class GateConfigLoaderTests
{
    const string Tenant = "11111111-1111-1111-1111-111111111111";
    const string Client = "22222222-2222-2222-2222-222222222222";
    const string OtherClient = "33333333-3333-3333-3333-333333333333";
    static readonly GateConfig Defaults = new(Tenant, Client, "https://gw.example.net/api/deepcheck", "api://gw/DeepCheck.Use", 30);

    [TestMethod]
    public void Uses_compiled_defaults_when_no_policy_is_set()
    {
        var c = GateConfigLoader.Load(new DictionaryPolicySource(), Defaults);
        Assert.AreEqual(Tenant, c.TenantId);
        Assert.AreEqual(Client, c.ClientId);
        Assert.AreEqual("https://gw.example.net/api/deepcheck", c.GatewayUrl);
        Assert.AreEqual("api://gw/DeepCheck.Use", c.GatewayScope);
        Assert.AreEqual(30, c.OfflineGraceDays);
        Assert.IsTrue(c.IsConfigured);
        Assert.IsTrue(c.DeepCheckAvailable);
    }

    [TestMethod]
    public void Null_policy_source_is_treated_as_empty()
    {
        Assert.AreEqual(Client, GateConfigLoader.Load(null, Defaults).ClientId);
    }

    [TestMethod]
    public void Applies_each_allowed_override()
    {
        var p = new DictionaryPolicySource();
        p.Values["ClientId"] = OtherClient;
        p.Values["GatewayUrl"] = "https://other.example.net/api/deepcheck";
        p.Values["GatewayScope"] = "api://other/DeepCheck.Use";
        p.Values["OfflineGraceDays"] = 7;
        var c = GateConfigLoader.Load(p, Defaults);
        Assert.AreEqual(OtherClient, c.ClientId);
        Assert.AreEqual("https://other.example.net/api/deepcheck", c.GatewayUrl);
        Assert.AreEqual("api://other/DeepCheck.Use", c.GatewayScope);
        Assert.AreEqual(7, c.OfflineGraceDays);
    }

    [TestMethod]
    public void Ignores_tenant_override_and_logs_it()
    {
        var p = new DictionaryPolicySource();
        p.Values["TenantId"] = "44444444-4444-4444-4444-444444444444";
        var log = new List<string>();
        var c = GateConfigLoader.Load(p, Defaults, log.Add);
        Assert.AreEqual(Tenant, c.TenantId);
        Assert.IsTrue(log.Any(m => m.Contains("TenantId")));
    }

    [TestMethod]
    public void Ignores_invalid_override_values()
    {
        var p = new DictionaryPolicySource();
        p.Values["ClientId"] = "not-a-guid";
        p.Values["GatewayUrl"] = "http://insecure.example.net/api/deepcheck";
        p.Values["GatewayScope"] = "   ";
        var log = new List<string>();
        var c = GateConfigLoader.Load(p, Defaults, log.Add);
        Assert.AreEqual(Client, c.ClientId);
        Assert.AreEqual("https://gw.example.net/api/deepcheck", c.GatewayUrl);
        Assert.AreEqual("api://gw/DeepCheck.Use", c.GatewayScope);
        Assert.AreEqual(3, log.Count);
    }

    [TestMethod]
    public void Grace_override_can_only_lower_the_value()
    {
        var p = new DictionaryPolicySource();
        p.Values["OfflineGraceDays"] = 45;
        Assert.AreEqual(30, GateConfigLoader.Load(p, Defaults).OfflineGraceDays);
    }

    [TestMethod]
    public void Grace_override_below_one_is_clamped_to_one()
    {
        var p = new DictionaryPolicySource();
        p.Values["OfflineGraceDays"] = 0;
        Assert.AreEqual(1, GateConfigLoader.Load(p, Defaults).OfflineGraceDays);
    }

    [TestMethod]
    public void Placeholder_ids_mean_not_configured()
    {
        var c = GateConfigLoader.Load(new DictionaryPolicySource(), GateDefaults.Create());
        Assert.IsFalse(c.IsConfigured);
        Assert.IsFalse(c.DeepCheckAvailable);
    }

    [TestMethod]
    public void Missing_gateway_only_disables_deep_check()
    {
        var c = GateConfigLoader.Load(new DictionaryPolicySource(), new GateConfig(Tenant, Client, "", "", 30));
        Assert.IsTrue(c.IsConfigured);
        Assert.IsFalse(c.DeepCheckAvailable);
    }
}
