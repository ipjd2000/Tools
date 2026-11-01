using LLDOCX.Claims;

namespace LLDOCX.Claims.Tests;

[TestClass]
public class TerminologyCheckerTests
{
    [TestMethod]
    public void Flags_rare_spelling_variant_of_frequent_word()
    {
        var f = TerminologyChecker.Check(T.Claims(
            "1. A frame comprising a bracket.",
            "2. The frame of claim 1, wherein the bracket is steel.",
            "3. The frame of claim 1, wherein the bracket is welded.",
            "4. The frame of claim 1, wherein the braket is painted."));
        Assert.AreEqual(1, f.Count, T.Dump(f));
        Assert.IsTrue(T.Has(f, 4, Severity.Warning, "'bracket'"), T.Dump(f));
        Assert.AreEqual("braket", f[0].Text);
    }

    [TestMethod]
    public void Ignores_inflections()
    {
        var f = TerminologyChecker.Check(T.Claims(
            "1. A method comprising: connecting a wire.",
            "2. The method of claim 1, comprising connecting a cable.",
            "3. The method of claim 1, comprising connecting a pipe.",
            "4. The method of claim 1, wherein the wire comprises copper."));
        Assert.IsFalse(f.Any(x => x.Type == FindingType.Terminology), T.Dump(f));
    }

    [TestMethod]
    public void Notes_mixed_said_and_the()
    {
        var f = TerminologyChecker.Check(T.Claims(
            "1. A cup comprising a handle.",
            "2. The cup of claim 1, wherein said handle is round.",
            "3. The cup of claim 1, wherein the handle is square."));
        Assert.IsTrue(T.Has(f, 2, Severity.Info, "both 'said'"), T.Dump(f));
    }

    [TestMethod]
    public void Flags_formatting_problems()
    {
        var f = TerminologyChecker.Check(T.Claims(
            "1. A cup comprising a a handle",
            "2. The cup of claim 1. wherein the handle is round."));
        string dump = T.Dump(f);
        Assert.IsTrue(T.Has(f, 1, Severity.Error, "Repeated word"), dump);
        Assert.IsTrue(T.Has(f, 1, Severity.Warning, "does not end with a period"), dump);
        Assert.IsTrue(T.Has(f, 2, Severity.Warning, "period before its end"), dump);
    }

    [TestMethod]
    public void Flags_unbalanced_parentheses()
    {
        var f = TerminologyChecker.Check(T.Claims("1. A cup (10 comprising a handle."));
        Assert.IsTrue(T.Has(f, 1, Severity.Warning, "Unbalanced"), T.Dump(f));
    }

    [TestMethod]
    public void Abbreviation_periods_are_not_flagged()
    {
        var f = TerminologyChecker.Check(T.Claims("1. A cup comprising a handle, e.g. a loop, as in FIG. 2."));
        Assert.IsFalse(f.Any(x => x.Message.Contains("period before")), T.Dump(f));
    }
}
