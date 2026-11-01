using LLDOCX.Claims;

namespace LLDOCX.Claims.Tests;

[TestClass]
public class ClaimParserTests
{
    static readonly string[] DeviceClaims =
    {
        "What is claimed is:",
        "1. A device comprising:",
        "a housing; and",
        "a lid attached to the housing.",
        "2. The device of claim 1, wherein the lid is hinged."
    };

    [TestMethod]
    public void Parses_typed_numbers_and_joins_continuation_paragraphs()
    {
        var claims = T.Claims(DeviceClaims);
        Assert.AreEqual(2, claims.Count);
        Assert.AreEqual(1, claims[0].Number);
        Assert.AreEqual("A device comprising: a housing; and a lid attached to the housing.", claims[0].FullText);
        Assert.IsTrue(claims[0].IsIndependent);
        CollectionAssert.AreEqual(new[] { 1 }, claims[1].DependsOn);
    }

    [TestMethod]
    public void Maps_claim_text_offsets_back_to_document_offsets()
    {
        var claims = T.Claims(DeviceClaims);
        string doc = T.Joined(DeviceClaims);
        int textIndex = claims[0].FullText.IndexOf("housing");
        Assert.AreEqual(doc.IndexOf("housing"), claims[0].ToDocOffset(textIndex));
        Assert.AreEqual(doc.IndexOf("A device"), claims[0].DocStart);
    }

    [TestMethod]
    public void Recognizes_status_markers()
    {
        var claims = T.Claims("1. (Canceled)", "2. (Currently Amended) A method comprising: heating a fluid.");
        Assert.IsTrue(claims[0].IsCanceled);
        Assert.AreEqual("", claims[0].FullText);
        Assert.AreEqual("Currently Amended", claims[1].Status);
        Assert.IsTrue(claims[1].FullText.StartsWith("A method"));
    }

    [DataTestMethod]
    [DataRow("2. The method of claim 1, wherein x.", "1", false, false)]
    [DataRow("3. The method according to claim 1 or 2, wherein x.", "1,2", true, false)]
    [DataRow("4. The method of any one of claims 1-3, wherein x.", "1,2,3", true, false)]
    [DataRow("4. The method of claims 1 and 2, wherein x.", "1,2", true, true)]
    [DataRow("5. The method of any of the preceding claims, wherein x.", "1,2,3,4", true, false)]
    [DataRow("6. The method as claimed in claims 1, 2, or 3, wherein x.", "1,2,3", true, false)]
    public void Parses_dependency_forms(string line, string deps, bool multiple, bool cumulative)
    {
        var c = T.Claims(line)[0];
        Assert.AreEqual(deps, string.Join(",", c.DependsOn));
        Assert.AreEqual(multiple, c.IsMultipleDependent);
        Assert.AreEqual(cumulative, c.HasCumulativeDependency);
        Assert.IsTrue(c.DependencyRefStart >= 0);
    }

    [DataTestMethod]
    [DataRow("1. A method of operating a device, comprising: x.", ClaimCategory.Method, "method")]
    [DataRow("1. A system for processing data, comprising: x.", ClaimCategory.Apparatus, "system")]
    [DataRow("1. A non-transitory computer-readable medium storing instructions that cause: x.", ClaimCategory.ComputerReadableMedium, "medium")]
    [DataRow("1. A composition comprising: x.", ClaimCategory.Composition, "composition")]
    [DataRow("2. The fluid dispensing apparatus of claim 1, wherein x.", ClaimCategory.Apparatus, "apparatus")]
    public void Detects_category_and_preamble_noun(string line, ClaimCategory category, string noun)
    {
        var c = T.Claims(line)[0];
        Assert.AreEqual(category, c.Category);
        Assert.AreEqual(noun, c.PreambleNoun);
    }

    [TestMethod]
    public void Finds_dependency_reference_outside_the_preamble()
    {
        var claims = T.Claims("1. A method comprising heating.",
            "2. A system comprising a processor configured to perform the method of claim 1.");
        var c = claims[1];
        CollectionAssert.AreEqual(new[] { 1 }, c.DependsOn);
        Assert.IsFalse(c.IsIndependent);
        Assert.AreEqual(c.FullText.IndexOf("claim 1"), c.DependencyRefStart);
        Assert.AreEqual("claim 1".Length, c.DependencyRefLength);
        Assert.AreEqual("A system", c.Preamble.Trim().Substring(0, 8));
        Assert.AreEqual("system", c.PreambleNoun);
        Assert.IsTrue(claims[0].IsIndependent);
    }

    [DataTestMethod]
    [DataRow("5. The method of any preceding claim, wherein x.")]
    [DataRow("5. The method of any previous claim, wherein x.")]
    [DataRow("5. The method of any one of the preceding claims, wherein x.")]
    public void Recognizes_singular_and_one_of_preceding_forms(string line)
    {
        var c = T.Claims(line)[0];
        Assert.AreEqual("1,2,3,4", string.Join(",", c.DependsOn));
        Assert.IsTrue(c.IsMultipleDependent);
    }

    [DataTestMethod]
    [DataRow("1. A built-in sensor comprising a lens.", "sensor")]
    [DataRow("1. A system-on-chip comprising a core.", "system-on-chip")]
    [DataRow("1. A sensor for a lens.", "sensor")]
    public void Hyphenated_words_are_not_split_by_stop_words(string line, string noun)
    {
        Assert.AreEqual(noun, T.Claims(line)[0].PreambleNoun);
    }

    [TestMethod]
    public void Huge_claim_numbers_do_not_throw_or_loop()
    {
        var c = T.Claims("2. The method of claims 1-99999999999.")[0];
        Assert.IsTrue(c.DependsOn.Count <= 1);
        var d = T.Claims("3. The method of claims 1-9999.")[0];
        Assert.AreEqual(0, d.DependsOn.Count);
    }

    [TestMethod]
    public void Ignores_text_before_first_claim_and_blank_paragraphs()
    {
        var claims = T.Claims("CLAIMS", "", "1. A cup comprising a handle.", "", "2. The cup of claim 1, wherein the handle is round.");
        Assert.AreEqual(2, claims.Count);
        Assert.AreEqual("A cup comprising a handle.", claims[0].FullText);
    }

    [TestMethod]
    public void Does_not_treat_decimal_numbers_as_claim_starts()
    {
        var claims = T.Claims("1. A rod comprising a diameter of", "2.5 mm.");
        Assert.AreEqual(1, claims.Count);
        Assert.AreEqual("A rod comprising a diameter of 2.5 mm.", claims[0].FullText);
    }
}
