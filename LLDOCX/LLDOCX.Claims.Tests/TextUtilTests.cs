using LLDOCX.Claims;

namespace LLDOCX.Claims.Tests;

[TestClass]
public class TextUtilTests
{
    [DataTestMethod]
    [DataRow("sensors", "sensor")]
    [DataRow("assemblies", "assembly")]
    [DataRow("processes", "process")]
    [DataRow("boxes", "box")]
    [DataRow("apparatus", "apparatus")]
    [DataRow("glass", "glass")]
    [DataRow("axis", "axis")]
    public void Singularize_handles_common_plural_forms(string word, string expected) =>
        Assert.AreEqual(expected, TextUtil.Singularize(word));

    [DataTestMethod]
    [DataRow("bracket", "braket", 1)]
    [DataRow("housing", "houseing", 1)]
    [DataRow("kitten", "sitting", 3)]
    [DataRow("", "abc", 3)]
    public void Levenshtein_computes_edit_distance(string a, string b, int expected) =>
        Assert.AreEqual(expected, TextUtil.Levenshtein(a, b));

    [TestMethod]
    public void Tokenize_splits_words_numbers_and_punctuation_with_offsets()
    {
        var tokens = TextUtil.Tokenize("a non-transitory medium 12, wherein");
        CollectionAssert.AreEqual(new[] { "a", "non-transitory", "medium", "12", ",", "wherein" },
            tokens.Select(t => t.Value).ToArray());
        Assert.AreEqual(2, tokens[1].Index);
        Assert.IsTrue(tokens[1].IsWord);
        Assert.IsFalse(tokens[3].IsWord);
    }

    [TestMethod]
    public void Claim_maps_text_offsets_across_segments()
    {
        var c = new Claim { FullText = "abcdefghij klmno" };
        c.Segments.Add(new TextSegment(0, 100, 10));
        c.Segments.Add(new TextSegment(11, 200, 5));
        Assert.AreEqual(103, c.ToDocOffset(3));
        Assert.AreEqual(110, c.ToDocOffset(10));
        Assert.AreEqual(200, c.ToDocOffset(11));
        Assert.AreEqual(201, c.ToDocOffset(12));
    }

    [TestMethod]
    public void Finding_At_computes_document_range_and_text()
    {
        var c = new Claim { Number = 4, FullText = "abcdefghij klmno" };
        c.Segments.Add(new TextSegment(0, 100, 10));
        c.Segments.Add(new TextSegment(11, 200, 5));
        var f = Finding.At(c, 12, 3, FindingType.Antecedent, Severity.Error, "msg");
        Assert.AreEqual(4, f.ClaimNumber);
        Assert.AreEqual(201, f.DocStart);
        Assert.AreEqual(3, f.Length);
        Assert.AreEqual("lmn", f.Text);
        Assert.AreEqual("Rules", f.Source);
    }
}
