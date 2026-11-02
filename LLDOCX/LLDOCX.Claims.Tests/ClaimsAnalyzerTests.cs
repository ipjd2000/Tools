using LLDOCX.Claims;

namespace LLDOCX.Claims.Tests;

[TestClass]
public class ClaimsAnalyzerTests
{
    /// <summary>A realistic, correctly drafted claim set. Must produce no errors or warnings.</summary>
    static readonly string[] CleanClaimSet =
    {
        "What is claimed is:",
        "1. A fluid dispensing apparatus, comprising:",
        "a housing defining a reservoir;",
        "a pump disposed within the housing and in fluid communication with the reservoir;",
        "a nozzle coupled to the pump; and",
        "a controller configured to actuate the pump in response to a signal from a sensor.",
        "2. The fluid dispensing apparatus of claim 1, wherein the sensor is an optical sensor.",
        "3. The fluid dispensing apparatus of claim 1, wherein the nozzle comprises a first outlet and a second outlet.",
        "4. The fluid dispensing apparatus of claim 3, wherein the first outlet has a diameter smaller than a diameter of the second outlet.",
        "5. A method of dispensing a fluid, the method comprising:",
        "receiving a signal from a sensor;",
        "actuating a pump in response to the signal; and",
        "dispensing the fluid through a nozzle.",
        "6. The method of claim 5, further comprising stopping the pump after a predetermined period."
    };

    [TestMethod]
    public void Clean_claim_set_produces_no_errors_or_warnings()
    {
        var result = ClaimsAnalyzer.Analyze(T.Paras(CleanClaimSet), new ClaimsCheckOptions());
        Assert.AreEqual(6, result.Claims.Count);
        var serious = result.Findings.Where(f => f.Severity != Severity.Info).ToList();
        Assert.AreEqual(0, serious.Count, T.Dump(serious));
    }

    [TestMethod]
    public void Applies_minimum_severity_filter()
    {
        var paras = T.Paras("1. A lamp comprising a housing, wherein the outer housing is black.");
        Assert.AreEqual(1, ClaimsAnalyzer.Analyze(paras, new ClaimsCheckOptions()).Findings.Count);
        var filtered = ClaimsAnalyzer.Analyze(paras, new ClaimsCheckOptions { MinSeverity = Severity.Warning });
        Assert.AreEqual(0, filtered.Findings.Count, T.Dump(filtered.Findings));
    }

    [TestMethod]
    public void Respects_check_toggles()
    {
        var paras = T.Paras("1. A lamp comprising a base, wherein the shade is green.");
        var on = ClaimsAnalyzer.Analyze(paras, new ClaimsCheckOptions());
        Assert.IsTrue(on.Findings.Any(f => f.Type == FindingType.Antecedent));
        var off = ClaimsAnalyzer.Analyze(paras, new ClaimsCheckOptions { CheckAntecedent = false });
        Assert.IsFalse(off.Findings.Any(f => f.Type == FindingType.Antecedent), T.Dump(off.Findings));
    }

    [TestMethod]
    public void Sorts_findings_by_claim_then_position()
    {
        var result = ClaimsAnalyzer.Analyze(T.Paras(
            "1. A lamp comprising a base, wherein the shade and the cord are green.",
            "2. The lamp of claim 1, wherein the plug is black."), new ClaimsCheckOptions());
        var numbers = result.Findings.Select(f => f.ClaimNumber).ToList();
        CollectionAssert.AreEqual(numbers.OrderBy(n => n).ToList(), numbers);
        var claim1 = result.Findings.Where(f => f.ClaimNumber == 1).ToList();
        Assert.IsTrue(claim1[0].DocStart < claim1[1].DocStart);
    }

    [DataTestMethod]
    [DataRow("real1.txt")]
    [DataRow("real2.txt")]
    [DataRow("real3.txt")]
    public void Real_issued_claims_produce_no_errors(string file)
    {
        var lines = File.ReadAllLines(Path.Combine(AppContext.BaseDirectory, "Fixtures", file));
        var result = ClaimsAnalyzer.Analyze(T.Paras(lines), new ClaimsCheckOptions());
        var errors = result.Findings.Where(f => f.Severity == Severity.Error).ToList();
        Assert.AreEqual(0, errors.Count, T.Dump(errors));
    }
}
