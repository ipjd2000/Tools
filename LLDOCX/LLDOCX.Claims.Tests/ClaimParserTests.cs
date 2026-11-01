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
