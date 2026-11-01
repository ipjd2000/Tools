using LLDOCX.Claims;

namespace LLDOCX.Claims.Tests;

[TestClass]
public class NounPhraseExtractorTests
{
    [TestMethod]
    public void Extracts_introductions_and_references()
    {
        var p = NounPhraseExtractor.Extract("a first arm coupled to a base, wherein the first arm rotates relative to the base.");
        Assert.AreEqual(4, p.Count, string.Join(" | ", p.Select(x => x.Text)));
        Assert.AreEqual(DeterminerKind.Introducing, p[0].Kind);
        Assert.AreEqual("first arm", p[0].Key);
        Assert.AreEqual("a first arm", p[0].Text);
        Assert.AreEqual("base", p[1].Key);
        Assert.AreEqual(DeterminerKind.Referring, p[2].Kind);
        Assert.AreEqual("first arm", p[2].Key);
        Assert.AreEqual("base", p[3].Key);
        Assert.AreEqual("arm", p[2].HeadKey);
        CollectionAssert.AreEqual(new[] { "first" }, p[2].ModifierKeys);
    }

    [TestMethod]
    public void Marks_preamble_claim_reference_and_stops_at_verbs()
    {
        var p = NounPhraseExtractor.Extract("The method of claim 1, wherein the valve opens.");
        Assert.AreEqual(2, p.Count, string.Join(" | ", p.Select(x => x.Text)));
        Assert.IsTrue(p[0].IsClaimReference);
        Assert.AreEqual("method", p[0].Key);
        Assert.IsFalse(p[1].IsClaimReference);
        Assert.AreEqual("valve", p[1].Key);
    }

    [TestMethod]
    public void Handles_multiword_determiners()
    {
        var p = NounPhraseExtractor.Extract(
            "a plurality of sensors and at least one controller, wherein each of the sensors and the at least one controller are connected.");
        CollectionAssert.AreEqual(
            new[] { "a plurality of", "at least one", "each of the", "the at least one" },
            p.Select(x => x.Determiner).ToArray());
        CollectionAssert.AreEqual(new[] { "sensor", "controller", "sensor", "controller" }, p.Select(x => x.Key).ToArray());
        Assert.IsTrue(p[0].IsPluralForm);
    }

    [TestMethod]
    public void Keeps_ing_nouns_and_stops_at_reference_numerals()
    {
        var p = NounPhraseExtractor.Extract("an outer housing 12 holding a spring.");
        Assert.AreEqual("outer housing", p[0].Key);
        Assert.AreEqual("an outer housing", p[0].Text);
        Assert.AreEqual("spring", p[1].Key);
    }

    [TestMethod]
    public void Adjective_heads_keep_following_plural_noun()
    {
        var p = NounPhraseExtractor.Extract("wherein the upper ends are joined and the corresponding underlying padding section flexes");
        Assert.AreEqual("upper end", p[0].Key, string.Join(" | ", p.Select(x => x.Key)));
        Assert.AreEqual("corresponding underlying padding section", p[1].Key, string.Join(" | ", p.Select(x => x.Key)));
    }

    [TestMethod]
    public void Leading_participle_is_a_modifier_not_a_phrase_end()
    {
        var p = NounPhraseExtractor.Extract("based on the stored partitioning information, and the selected sensors");
        Assert.AreEqual("stored partitioning information", p[0].Key, string.Join(" | ", p.Select(x => x.Key)));
        Assert.AreEqual("selected sensor", p[1].Key, string.Join(" | ", p.Select(x => x.Key)));
    }
}
