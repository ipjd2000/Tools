using LLDOCX.Claims;

namespace LLDOCX.Claims.Tests;

[TestClass]
public class AntecedentCheckerTests
{
    static readonly string[] Allow = { "the user", "the same", "the group" };

    static List<Finding> Check(params string[] lines) => AntecedentChecker.Check(T.Claims(lines), Allow);

    [TestMethod]
    public void Clean_claims_have_no_findings()
    {
        var f = Check(
            "1. A lamp comprising a base and a bulb mounted on the base.",
            "2. The lamp of claim 1, wherein the bulb is removable.");
        Assert.AreEqual(0, f.Count, T.Dump(f));
    }

    [TestMethod]
    public void Flags_missing_basis()
    {
        var f = Check("1. A lamp comprising a base, wherein the shade is green.");
        Assert.IsTrue(T.Has(f, 1, Severity.Error, "No antecedent basis for 'the shade'"), T.Dump(f));
        Assert.AreEqual("the shade", f.Single().Text);
    }

    [TestMethod]
    public void Basis_comes_only_from_ancestor_claims()
    {
        var f = Check(
            "1. A lamp comprising a base.",
            "2. The lamp of claim 1, further comprising a shade.",
            "3. The lamp of claim 1, wherein the shade is green.");
        Assert.IsTrue(T.Has(f, 3, Severity.Error, "the shade"), T.Dump(f));
        Assert.IsFalse(f.Any(x => x.ClaimNumber == 2), T.Dump(f));
    }

    [TestMethod]
    public void Multiple_dependent_claim_needs_basis_in_every_branch()
    {
        var f = Check(
            "1. A cup comprising a handle.",
            "2. The cup of claim 1, further comprising a lid.",
            "3. The cup of claim 1 or 2, wherein the lid is clear.");
        Assert.IsTrue(T.Has(f, 3, Severity.Error, "the lid"), T.Dump(f));
    }

    [TestMethod]
    public void Flags_ambiguous_reference()
    {
        var f = Check("1. A clamp comprising a first jaw and a second jaw, wherein the jaw is serrated.");
        Assert.IsTrue(T.Has(f, 1, Severity.Error, "ambiguous"), T.Dump(f));
    }

    [TestMethod]
    public void Warns_on_reintroduction()
    {
        var f = Check("1. A lamp comprising a bulb and a socket holding a bulb.");
        Assert.IsTrue(T.Has(f, 1, Severity.Warning, "already has antecedent basis"), T.Dump(f));
    }

    [TestMethod]
    public void Reports_weak_head_noun_match_as_info()
    {
        var f = Check("1. A lamp comprising a housing, wherein the outer housing is black.");
        Assert.IsTrue(T.Has(f, 1, Severity.Info, "partly matches"), T.Dump(f));
        Assert.IsFalse(f.Any(x => x.Severity == Severity.Error), T.Dump(f));
    }

    [TestMethod]
    public void Suggests_close_match_for_typo()
    {
        var f = Check("1. A lamp comprising a reflector, wherein the reflecter is polished.");
        var finding = f.Single(x => x.Severity == Severity.Error);
        StringAssert.Contains(finding.Suggestion, "the reflector");
    }

    [TestMethod]
    public void Allowlisted_phrases_need_no_basis()
    {
        var f = Check("1. A method comprising heating a fluid to the same temperature.");
        Assert.AreEqual(0, f.Count, T.Dump(f));
    }

    [TestMethod]
    public void Flags_singular_plural_drift_as_terminology()
    {
        var f = Check("1. A board comprising a sensor, wherein the sensors are calibrated.");
        var finding = f.Single();
        Assert.AreEqual(FindingType.Terminology, finding.Type);
        StringAssert.Contains(finding.Message, "plural");
    }

    [TestMethod]
    public void Plural_reference_to_plurality_is_fine()
    {
        var f = Check("1. A board comprising a plurality of sensors, wherein the sensors are calibrated.");
        Assert.AreEqual(0, f.Count, T.Dump(f));
    }

    [TestMethod]
    public void Bare_plural_or_mass_noun_counts_as_introduction()
    {
        var f = Check("1. A method comprising receiving data from a server, wherein the data is encrypted.");
        Assert.AreEqual(0, f.Count(x => x.Severity == Severity.Error), T.Dump(f));
    }

    [TestMethod]
    public void Claim_subject_starting_with_the_is_not_flagged()
    {
        var f = Check(
            "1. A compound of formula I.",
            "2. The sodium salt of the compound of claim 1.");
        Assert.AreEqual(0, f.Count(x => x.Severity == Severity.Error), T.Dump(f));
    }

    [TestMethod]
    public void Elliptical_coordinated_modifier_matches_compound_term()
    {
        var f = Check("1. A garment comprising a front section and a rear section, wherein said front and rear sections are padded.");
        Assert.AreEqual(0, f.Count(x => x.Severity == Severity.Error), T.Dump(f));
    }

    [TestMethod]
    public void The_latter_is_not_flagged()
    {
        var f = Check("1. A lamp comprising a base and a shade, wherein the latter is green.");
        Assert.AreEqual(0, f.Count(x => x.Severity == Severity.Error), T.Dump(f));
    }
}
