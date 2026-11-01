using LLDOCX.Claims;

namespace LLDOCX.Claims.Tests;

[TestClass]
public class DependencyCheckerTests
{
    [TestMethod]
    public void Flags_self_forward_nonexistent_references_and_numbering_gaps()
    {
        var f = DependencyChecker.Check(T.Claims(
            "1. A widget comprising a base.",
            "2. The widget of claim 2, wherein the base is flat.",
            "3. The widget of claim 5, wherein the base is round.",
            "5. The widget of claim 4, wherein the base is red."));
        string dump = T.Dump(f);
        Assert.IsTrue(T.Has(f, 2, Severity.Error, "depends on itself"), dump);
        Assert.IsTrue(T.Has(f, 3, Severity.Error, "comes after it"), dump);
        Assert.IsTrue(T.Has(f, 5, Severity.Error, "does not exist"), dump);
        Assert.IsTrue(T.Has(f, 5, Severity.Warning, "jumps from 3 to 5"), dump);
    }

    [TestMethod]
    public void Flags_duplicate_numbers_and_canceled_parents()
    {
        var f = DependencyChecker.Check(T.Claims(
            "1. A widget comprising a base.",
            "2. (Canceled)",
            "3. The widget of claim 2, wherein the base is flat.",
            "3. The widget of claim 1, wherein the base is round."));
        string dump = T.Dump(f);
        Assert.IsTrue(T.Has(f, 3, Severity.Error, "canceled claim 2"), dump);
        Assert.IsTrue(T.Has(f, 3, Severity.Error, "used more than once"), dump);
    }

    [TestMethod]
    public void Flags_improper_multiple_dependencies()
    {
        var f = DependencyChecker.Check(T.Claims(
            "1. A pump comprising a rotor.",
            "2. The pump of claim 1, wherein the rotor is steel.",
            "3. The pump of claim 1 or 2, wherein the rotor is coated.",
            "4. The pump of claim 2 or 3, further comprising a seal.",
            "5. The pump of claims 1 and 2, further comprising a valve."));
        string dump = T.Dump(f);
        Assert.IsTrue(T.Has(f, 4, Severity.Error, "not permitted"), dump);
        Assert.IsTrue(T.Has(f, 5, Severity.Error, "cumulatively"), dump);
        Assert.IsTrue(T.Has(f, 3, Severity.Info, "multiple dependent"), dump);
    }

    [TestMethod]
    public void Flags_category_and_preamble_mismatch()
    {
        var f = DependencyChecker.Check(T.Claims(
            "1. A system comprising a processor.",
            "2. The method of claim 1, wherein the processor is fast.",
            "3. The device of claim 1, wherein the processor is slow."));
        string dump = T.Dump(f);
        Assert.IsTrue(T.Has(f, 2, Severity.Warning, "different category"), dump);
        Assert.IsTrue(T.Has(f, 3, Severity.Warning, "'the device'"), dump);
    }

    [TestMethod]
    public void Clean_claim_set_has_no_findings()
    {
        var f = DependencyChecker.Check(T.Claims(
            "1. A lamp comprising a bulb.",
            "2. The lamp of claim 1, wherein the bulb is an LED."));
        Assert.AreEqual(0, f.Count, T.Dump(f));
    }
}
