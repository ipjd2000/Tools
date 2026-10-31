# Patent Claims Checker Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Add a "Check Claims" tool to the LLDOCX Word add-in. It finds the patent claims in the active document and reports antecedent basis, claim dependency, and terminology/typo problems in a task pane. An optional "Deep Check" sends the claims to Claude.

**Architecture:** A pure C# engine (`LLDOCX.Claims`, netstandard2.0, no Office references) parses claims and runs the checks on plain text with document offsets. `LLDOCX.DeepCheck` (netstandard2.0) wraps the Anthropic SDK call and its prompt and response protocol. The VSTO add-in (`LLDOCX`, retargeted to .NET Framework 4.7.2) reads paragraphs from Word, calls the engine on a background thread, and shows the findings in a new task pane. `LLDOCX.Claims.Tests` (net8.0, MSTest) tests the engine and the Deep Check protocol with `dotnet test`.

**Tech Stack:** C# / .NET Framework 4.7.2 VSTO (Word interop, WinForms), netstandard2.0 class libraries, MSTest on net8.0, `Anthropic` NuGet SDK 13.3.0, DPAPI (`System.Security.Cryptography.ProtectedData`).

**Spec:** `LLDOCX/docs/superpowers/specs/2026-10-30-patent-claims-checker-design.md`

## Global Constraints

- All paths below are relative to the repository root `Tools/`. Shell commands run from `Tools/LLDOCX/` unless a step says otherwise.
- The add-in target framework becomes **.NET Framework 4.7.2**. The `Anthropic` 13.3.0 SDK ships `netstandard2.0`, and its dependencies need .NET Framework 4.6.2 or later.
- `LLDOCX.Claims` must not reference Office, WinForms, or the Anthropic SDK.
- Default Claude model: `claude-opus-5-5`. Deep Check uses `output_config.effort = high`, structured outputs (`output_config.format`), and server-side refusal fallback (`fallbacks: "default"` with beta `server-side-fallback-2026-07-01`).
- No claims text is sent to Claude without the confirmation dialog. The only exception is a document the user marked "Don't ask again" during the current session.
- The Claude API key is stored only in `%APPDATA%\LL-Docx\claude.key`, encrypted with DPAPI (CurrentUser scope). It is never written to the BinaryFormatter settings file `LLDocx.dll`.
- The tool never modifies the document.
- Existing Scan behavior is unchanged.
- `Severity` enum values are ordered: `Error = 0`, `Warning = 1`, `Info = 2`. A severity filter keeps findings with `Severity <= minimum`.
- Commit after every task. Each commit message ends with:
  `Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>`
- Use the branch `feature/patent-claims-checker`. It already exists and holds the spec commit.

## Build and test environment

- **Engine and tests (Tasks 2–9):** need the .NET 8 SDK. This machine currently has only the `dotnet` runtime. Task 2 installs the SDK.
- **VSTO add-in (Tasks 1, 10, 11, 12):** needs Visual Studio 2022 with the **Office/SharePoint development** workload and the **.NET Framework 4.7.2 targeting pack**. This machine has only VS Build Tools with neither, so those tasks are built, run, and verified in Visual Studio by a human (F5 starts Word with the add-in loaded). Each of those tasks ends with an explicit manual verification checklist.

## File Structure

```
LLDOCX/
  .gitignore                                   (new) ignore .vs/, bin/, obj/, *.user, *.pfx
  LLDOCX.sln                                   (modify) add 3 projects
  LLDOCX.Claims/                               (new) engine, netstandard2.0
    LLDOCX.Claims.csproj
    Models.cs              enums, RawParagraph, TextSegment, Claim, Finding
    TextUtil.cs            Token, Tokenize, Singularize, IsPlural, Levenshtein
    ClaimParser.cs         paragraphs -> claims (numbering, status, preamble, deps, category)
    DependencyChecker.cs   dependency/numbering findings
    NounPhraseExtractor.cs NounPhrase, DeterminerKind, Extract(text)
    AntecedentChecker.cs   antecedent findings (+ singular/plural drift)
    TerminologyChecker.cs  spelling variants, said/the mix, formatting
    ClaimsAnalyzer.cs      ClaimsCheckOptions, AnalysisResult, Analyze()
  LLDOCX.DeepCheck/                            (new) Claude integration, netstandard2.0
    LLDOCX.DeepCheck.csproj
    DeepCheckException.cs
    DeepCheckPrompt.cs     system prompt, JSON schema, user message builder
    DeepCheckResponseParser.cs  JSON -> findings + disputes
    DeepCheckClient.cs     Anthropic SDK call + error mapping
  LLDOCX.Claims.Tests/                         (new) MSTest, net8.0
    LLDOCX.Claims.Tests.csproj
    TestHelpers.cs
    TextUtilTests.cs, ClaimParserTests.cs, DependencyCheckerTests.cs,
    NounPhraseExtractorTests.cs, AntecedentCheckerTests.cs,
    TerminologyCheckerTests.cs, ClaimsAnalyzerTests.cs, DeepCheckProtocolTests.cs
  LLDOCX/                                      (existing VSTO add-in)
    LLDOCX.csproj                              (modify) 4.7.2, LangVersion, refs, new files
    ThisAddIn.cs                               (modify) claims pane, controller, sync context
    RibbonExplorer.xml / RibbonExplorer.cs     (modify) two buttons + callbacks
    Options/SettingOption.cs                   (modify) claims settings fields
    CustomUserControl/SettingUserControl.cs    (modify) preserve claims settings on save
    Options/ApiKeyStore.cs                     (new) DPAPI key storage
    Helpers/ClaimsExtractor.cs                 (new) Word -> RawParagraph list
    Helpers/ClaimsController.cs                (new) orchestrates check + deep check
    CustomUserControl/ClaimsPaneControl.cs     (new) findings tab + settings tab
    CustomUserControl/ClaimsSettingsPanel.cs   (new) claims settings UI
    CustomUserControl/ApiKeyDialog.cs          (new) API key entry dialog
    CustomUserControl/DeepCheckConfirmDialog.cs (new) consent dialog
  LLDOCXSetup/Product.wxs                      (modify) ship new DLLs + LLDOCX.dll.config
```

---

### Task 1: Track the LLDOCX source and retarget to .NET Framework 4.7.2

**Files:**
- Create: `LLDOCX/.gitignore`
- Modify: `LLDOCX/LLDOCX/LLDOCX.csproj` (TargetFrameworkVersion, LangVersion, binding redirects)

**Interfaces:**
- Consumes: nothing
- Produces: the add-in source tracked in git; the add-in builds against .NET Framework 4.7.2 with `LangVersion` latest. Later tasks add references to this csproj.

- [ ] **Step 1: Create `LLDOCX/.gitignore`**

```gitignore
.vs/
bin/
obj/
*.user
*.suo
*.pfx
packages/
TestResults/
```

- [ ] **Step 2: Retarget and enable binding redirects in `LLDOCX/LLDOCX/LLDOCX.csproj`**

Change line 30 from:

```xml
    <TargetFrameworkVersion>v4.6.1</TargetFrameworkVersion>
```

to:

```xml
    <TargetFrameworkVersion>v4.7.2</TargetFrameworkVersion>
    <LangVersion>latest</LangVersion>
    <AutoGenerateBindingRedirects>true</AutoGenerateBindingRedirects>
    <GenerateBindingRedirectsOutputType>true</GenerateBindingRedirectsOutputType>
    <RestoreProjectStyle>PackageReference</RestoreProjectStyle>
```

- [ ] **Step 3: Verify in Visual Studio (manual)**

Open `LLDOCX/LLDOCX.sln` in Visual Studio 2022. Build → Rebuild Solution.
Expected: build succeeds. Press F5: Word opens and the Home tab shows the LL-Docx group with **Scan** and **Settings**. Click **Scan** on any document: it behaves as before.

- [ ] **Step 4: Commit**

Run from `Tools/`:

```bash
git add LLDOCX/.gitignore LLDOCX/LLDOCX.sln LLDOCX/LLDOCX LLDOCX/LLDOCXSetup
git status --short   # confirm LLDOCX_TemporaryKey.pfx and .vs/ are NOT listed
git commit -m "chore(lldocx): track add-in source, retarget to .NET Framework 4.7.2

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 2: Engine scaffold, core models, and text utilities

**Files:**
- Create: `LLDOCX/LLDOCX.Claims/LLDOCX.Claims.csproj`
- Create: `LLDOCX/LLDOCX.Claims/Models.cs`
- Create: `LLDOCX/LLDOCX.Claims/TextUtil.cs`
- Create: `LLDOCX/LLDOCX.Claims.Tests/LLDOCX.Claims.Tests.csproj`
- Create: `LLDOCX/LLDOCX.Claims.Tests/TestHelpers.cs`
- Test: `LLDOCX/LLDOCX.Claims.Tests/TextUtilTests.cs`
- Modify: `LLDOCX/LLDOCX.sln` (via `dotnet sln add`)

**Interfaces:**
- Consumes: nothing
- Produces (namespace `LLDOCX.Claims`):
  - `enum ClaimCategory { Unknown, Method, Apparatus, Composition, ComputerReadableMedium }`
  - `enum FindingType { Antecedent, Dependency, Terminology, Formatting }`
  - `enum Severity { Error = 0, Warning = 1, Info = 2 }`
  - `class RawParagraph(string text, int docStart)` with `Text`, `DocStart`
  - `class TextSegment(int textOffset, int docStart, int length)`
  - `class Claim`: `Number`, `Status`, `IsCanceled`, `FullText`, `Preamble`, `DependsOn` (List<int>), `IsIndependent`, `IsMultipleDependent`, `HasCumulativeDependency`, `DependencyRefStart`, `DependencyRefLength`, `Category`, `PreambleNoun`, `DocStart`, `Segments`, `int ToDocOffset(int textIndex)`
  - `class Finding`: `ClaimNumber`, `Type`, `Severity`, `Message`, `Suggestion`, `DocStart`, `Length`, `Text`, `Source` ("Rules"/"AI"), `DisputedByAI`, `DisputeReason`, `static Finding At(Claim, int start, int length, FindingType, Severity, string message, string suggestion = null)`
  - `struct Token(string value, int index)` with `Value`, `Index`, `End`, `IsWord`, `Lower`
  - `static class TextUtil`: `List<Token> Tokenize(string)`, `string Singularize(string lowerWord)`, `bool IsPlural(string lowerWord)`, `int Levenshtein(string, string)`
  - Test helpers `T.Paras(params string[])`, `T.Claims(params string[])`, `T.Joined(params string[])`, `T.Has(findings, claim, severity, text)`, `T.Dump(findings)`

- [ ] **Step 1: Install the .NET 8 SDK**

Run: `winget install --id Microsoft.DotNet.SDK.8 -e --accept-source-agreements --accept-package-agreements`
Then, in a new shell: `dotnet --list-sdks`
Expected: a line starting with `8.0.`

- [ ] **Step 2: Create the engine project file `LLDOCX/LLDOCX.Claims/LLDOCX.Claims.csproj`**

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFramework>netstandard2.0</TargetFramework>
    <LangVersion>7.3</LangVersion>
    <RootNamespace>LLDOCX.Claims</RootNamespace>
  </PropertyGroup>
</Project>
```

- [ ] **Step 3: Create the test project `LLDOCX/LLDOCX.Claims.Tests/LLDOCX.Claims.Tests.csproj`**

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFramework>net8.0</TargetFramework>
    <ImplicitUsings>enable</ImplicitUsings>
    <Nullable>disable</Nullable>
    <IsPackable>false</IsPackable>
  </PropertyGroup>
  <ItemGroup>
    <PackageReference Include="Microsoft.NET.Test.Sdk" Version="17.11.1" />
    <PackageReference Include="MSTest.TestAdapter" Version="3.6.1" />
    <PackageReference Include="MSTest.TestFramework" Version="3.6.1" />
  </ItemGroup>
  <ItemGroup>
    <ProjectReference Include="..\LLDOCX.Claims\LLDOCX.Claims.csproj" />
  </ItemGroup>
</Project>
```

- [ ] **Step 4: Add both projects to the solution**

Run:

```bash
dotnet sln LLDOCX.sln add LLDOCX.Claims/LLDOCX.Claims.csproj LLDOCX.Claims.Tests/LLDOCX.Claims.Tests.csproj
```

Expected: `Project ... added to the solution.` twice.

- [ ] **Step 5: Write test helpers `LLDOCX/LLDOCX.Claims.Tests/TestHelpers.cs`**

```csharp
global using Microsoft.VisualStudio.TestTools.UnitTesting;
using LLDOCX.Claims;

namespace LLDOCX.Claims.Tests;

internal static class T
{
    /// <summary>One RawParagraph per line; DocStart is the line's offset in the "\n"-joined text.</summary>
    public static List<RawParagraph> Paras(params string[] lines)
    {
        var list = new List<RawParagraph>();
        int pos = 0;
        foreach (var line in lines)
        {
            list.Add(new RawParagraph(line, pos));
            pos += line.Length + 1;
        }
        return list;
    }

    public static string Joined(params string[] lines) => string.Join("\n", lines);

    public static List<Claim> Claims(params string[] lines) => throw new NotImplementedException(); // restored in Task 3

    public static bool Has(IEnumerable<Finding> findings, int claim, Severity severity, string text) =>
        findings.Any(f => f.ClaimNumber == claim && f.Severity == severity &&
                          f.Message.Contains(text, StringComparison.OrdinalIgnoreCase));

    public static string Dump(IEnumerable<Finding> findings) =>
        string.Join("\n", findings.Select(f => $"{f.ClaimNumber} {f.Severity} {f.Type}: {f.Message} | {f.Suggestion}"));
}
```

- [ ] **Step 6: Write the failing tests `LLDOCX/LLDOCX.Claims.Tests/TextUtilTests.cs`**

```csharp
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
```

- [ ] **Step 7: Run tests to verify they fail**

Run: `dotnet test LLDOCX.Claims.Tests`
Expected: build FAILS with errors such as `The name 'TextUtil' does not exist` and `The type or namespace name 'Claim' could not be found`.

- [ ] **Step 8: Write `LLDOCX/LLDOCX.Claims/Models.cs`**

```csharp
using System;
using System.Collections.Generic;

namespace LLDOCX.Claims
{
    public enum ClaimCategory { Unknown, Method, Apparatus, Composition, ComputerReadableMedium }

    public enum FindingType { Antecedent, Dependency, Terminology, Formatting }

    /// <summary>Ordered so that a filter can keep findings with Severity &lt;= minimum.</summary>
    public enum Severity { Error = 0, Warning = 1, Info = 2 }

    /// <summary>One paragraph of document text. DocStart is the document offset of Text[0].</summary>
    public class RawParagraph
    {
        public RawParagraph(string text, int docStart)
        {
            Text = text ?? "";
            DocStart = docStart;
        }

        public string Text { get; }
        public int DocStart { get; }
    }

    /// <summary>Maps a run of Claim.FullText (starting at TextOffset) to the document (starting at DocStart).</summary>
    public class TextSegment
    {
        public TextSegment(int textOffset, int docStart, int length)
        {
            TextOffset = textOffset;
            DocStart = docStart;
            Length = length;
        }

        public int TextOffset { get; }
        public int DocStart { get; }
        public int Length { get; }
    }

    public class Claim
    {
        public int Number { get; set; }
        public string Status { get; set; }
        public bool IsCanceled => Status == "Canceled";
        public string FullText { get; set; } = "";
        public string Preamble { get; set; } = "";
        public List<int> DependsOn { get; } = new List<int>();
        public bool IsIndependent => DependsOn.Count == 0;
        public bool IsMultipleDependent { get; set; }
        public bool HasCumulativeDependency { get; set; }
        public int DependencyRefStart { get; set; } = -1;
        public int DependencyRefLength { get; set; }
        public ClaimCategory Category { get; set; }
        public string PreambleNoun { get; set; } = "";
        /// <summary>Document offset of the claim's content (after its number and status).</summary>
        public int DocStart { get; set; }
        public List<TextSegment> Segments { get; } = new List<TextSegment>();

        public int ToDocOffset(int textIndex)
        {
            if (Segments.Count == 0) return DocStart;
            foreach (var s in Segments)
            {
                if (textIndex < s.TextOffset) return s.DocStart;
                if (textIndex <= s.TextOffset + s.Length) return s.DocStart + (textIndex - s.TextOffset);
            }
            var last = Segments[Segments.Count - 1];
            return last.DocStart + last.Length;
        }
    }

    public class Finding
    {
        public int ClaimNumber { get; set; }
        public FindingType Type { get; set; }
        public Severity Severity { get; set; }
        public string Message { get; set; } = "";
        public string Suggestion { get; set; }
        /// <summary>Document offset of the flagged text, or -1 when it could not be located.</summary>
        public int DocStart { get; set; } = -1;
        public int Length { get; set; }
        /// <summary>The flagged text as it appeared when checked; used for stale detection.</summary>
        public string Text { get; set; } = "";
        public string Source { get; set; } = "Rules";
        public bool DisputedByAI { get; set; }
        public string DisputeReason { get; set; }

        public static Finding At(Claim claim, int start, int length, FindingType type, Severity severity,
            string message, string suggestion = null)
        {
            string text = claim.FullText ?? "";
            start = Math.Max(0, Math.Min(start, text.Length));
            length = Math.Max(0, Math.Min(length, text.Length - start));
            int docStart = claim.ToDocOffset(start);
            return new Finding
            {
                ClaimNumber = claim.Number,
                Type = type,
                Severity = severity,
                Message = message,
                Suggestion = suggestion,
                DocStart = docStart,
                Length = claim.ToDocOffset(start + length) - docStart,
                Text = text.Substring(start, length)
            };
        }
    }
}
```

- [ ] **Step 9: Write `LLDOCX/LLDOCX.Claims/TextUtil.cs`**

```csharp
using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace LLDOCX.Claims
{
    public struct Token
    {
        public Token(string value, int index)
        {
            Value = value;
            Index = index;
        }

        public string Value { get; }
        public int Index { get; }
        public int End => Index + Value.Length;
        public bool IsWord => Value.Length > 0 && char.IsLetter(Value[0]);
        public string Lower => Value.ToLowerInvariant();
    }

    public static class TextUtil
    {
        static readonly Regex TokenRx = new Regex(@"[A-Za-z][A-Za-z0-9'\-]*|\d+(?:\.\d+)?|[^\sA-Za-z0-9]", RegexOptions.Compiled);

        public static List<Token> Tokenize(string text)
        {
            var list = new List<Token>();
            foreach (Match m in TokenRx.Matches(text ?? ""))
                list.Add(new Token(m.Value, m.Index));
            return list;
        }

        /// <summary>Rule-based singular form of a lowercase English word.</summary>
        public static string Singularize(string word)
        {
            if (string.IsNullOrEmpty(word) || word.Length <= 3) return word;
            if (word.EndsWith("ies") && word.Length > 4) return word.Substring(0, word.Length - 3) + "y";
            if (word.EndsWith("sses")) return word.Substring(0, word.Length - 2);
            if (word.EndsWith("xes") || word.EndsWith("ches") || word.EndsWith("shes")) return word.Substring(0, word.Length - 2);
            if (word.EndsWith("ss") || word.EndsWith("us") || word.EndsWith("is")) return word;
            if (word.EndsWith("s")) return word.Substring(0, word.Length - 1);
            return word;
        }

        public static bool IsPlural(string word) => Singularize(word) != word;

        public static int Levenshtein(string a, string b)
        {
            a = a ?? "";
            b = b ?? "";
            var prev = new int[b.Length + 1];
            var cur = new int[b.Length + 1];
            for (int j = 0; j <= b.Length; j++) prev[j] = j;
            for (int i = 1; i <= a.Length; i++)
            {
                cur[0] = i;
                for (int j = 1; j <= b.Length; j++)
                {
                    int cost = a[i - 1] == b[j - 1] ? 0 : 1;
                    cur[j] = Math.Min(Math.Min(cur[j - 1] + 1, prev[j] + 1), prev[j - 1] + cost);
                }
                var tmp = prev; prev = cur; cur = tmp;
            }
            return prev[b.Length];
        }
    }
}
```

- [ ] **Step 10: Run tests to verify they pass**

Run: `dotnet test LLDOCX.Claims.Tests`
Expected: `Passed!` with `Failed: 0`.

- [ ] **Step 11: Commit**

```bash
git add LLDOCX/LLDOCX.sln LLDOCX/LLDOCX.Claims LLDOCX/LLDOCX.Claims.Tests
git commit -m "feat(claims): add engine scaffold, core models, and text utilities

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 3: ClaimParser

**Files:**
- Create: `LLDOCX/LLDOCX.Claims/ClaimParser.cs`
- Modify: `LLDOCX/LLDOCX.Claims.Tests/TestHelpers.cs` (restore `T.Claims`)
- Test: `LLDOCX/LLDOCX.Claims.Tests/ClaimParserTests.cs`

**Interfaces:**
- Consumes: `RawParagraph`, `Claim`, `TextSegment`, `ClaimCategory`, `TextUtil.Singularize` (Task 2)
- Produces: `static class ClaimParser { static List<Claim> Parse(IList<RawParagraph> paragraphs); }`. Each `Claim` comes out with Number, Status, FullText, Segments, DocStart, Preamble, DependsOn, IsMultipleDependent, HasCumulativeDependency, DependencyRefStart/Length, Category, and PreambleNoun (lowercase, singular) populated.

- [ ] **Step 1: Restore `T.Claims` in `TestHelpers.cs`**

Replace the `Claims` line with:

```csharp
    public static List<Claim> Claims(params string[] lines) => ClaimParser.Parse(Paras(lines));
```

- [ ] **Step 2: Write the failing tests `LLDOCX/LLDOCX.Claims.Tests/ClaimParserTests.cs`**

```csharp
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
```

- [ ] **Step 3: Run tests to verify they fail**

Run: `dotnet test LLDOCX.Claims.Tests --filter ClaimParserTests`
Expected: build FAILS with `The name 'ClaimParser' does not exist in the current context`.

- [ ] **Step 4: Write `LLDOCX/LLDOCX.Claims/ClaimParser.cs`**

```csharp
using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace LLDOCX.Claims
{
    public static class ClaimParser
    {
        const RegexOptions Opts = RegexOptions.Compiled | RegexOptions.IgnoreCase;

        static readonly Regex ClaimStart = new Regex(@"^\s*(\d{1,3})\s*[.)](?!\d)\s*", RegexOptions.Compiled);
        static readonly Regex StatusRx = new Regex(@"^\((canceled|cancelled|withdrawn|original|previously presented|currently amended|new)\)\s*", Opts);
        static readonly Regex Transition = new Regex(@"\b(?:comprising|consisting essentially of|consisting of|including|wherein|characterized in that|having)\b|:", Opts);
        static readonly Regex Preceding = new Regex(@"\b(?:any|one)\s+(?:one\s+)?of\s+the\s+(?:preceding|previous|foregoing)\s+claims\b", Opts);
        static readonly Regex DepRef = new Regex(@"\bclaims?\s+(\d+(?:\s*(?:-|–|to|through|,|or|and|, or|, and)\s*\d+)*)", Opts);
        static readonly Regex NumberOrRange = new Regex(@"(\d+)(?:\s*(?:-|–|to|through)\s*(\d+))?", RegexOptions.Compiled);
        static readonly Regex AndRx = new Regex(@"\band\b", Opts);
        static readonly Regex LeadingDeterminer = new Regex(@"^\s*(?:a|an|the|one)\s+", Opts);
        static readonly Regex NounStop = new Regex(@"\b(?:for|of|according|as|in|comprising|consisting|including|having|wherein|that|which|configured|to|storing|with|on)\b|[,:;]", Opts);

        static readonly (Regex Pattern, ClaimCategory Category)[] CategoryKeywords =
        {
            (new Regex(@"computer[- ]readable|machine[- ]readable|processor[- ]readable|storage (?:medium|media)|program product", Opts), ClaimCategory.ComputerReadableMedium),
            (new Regex(@"\b(?:method|process)\b", Opts), ClaimCategory.Method),
            (new Regex(@"\b(?:composition|compound|formulation|mixture)\b", Opts), ClaimCategory.Composition),
            (new Regex(@"\b(?:system|apparatus|device|assembly|machine|circuit|kit|vehicle|tool|article|structure|instrument)\b", Opts), ClaimCategory.Apparatus),
        };

        public static List<Claim> Parse(IList<RawParagraph> paragraphs)
        {
            var claims = new List<Claim>();
            Claim current = null;
            StringBuilder sb = null;
            foreach (var p in paragraphs)
            {
                string text = p.Text.TrimEnd('\r', '\n', '\a', ' ', '\t');
                if (text.Trim().Length == 0) continue;

                var m = ClaimStart.Match(text);
                if (m.Success)
                {
                    if (current != null) Finish(current, sb);
                    current = new Claim { Number = int.Parse(m.Groups[1].Value) };
                    sb = new StringBuilder();
                    int contentStart = m.Length;
                    var sm = StatusRx.Match(text.Substring(contentStart));
                    if (sm.Success)
                    {
                        current.Status = CanonicalStatus(sm.Groups[1].Value);
                        contentStart += sm.Length;
                    }
                    current.DocStart = p.DocStart + contentStart;
                    Append(current, sb, text, contentStart, p.DocStart);
                    claims.Add(current);
                }
                else if (current != null)
                {
                    int lead = text.Length - text.TrimStart().Length;
                    Append(current, sb, text, lead, p.DocStart);
                }
            }
            if (current != null) Finish(current, sb);
            return claims;
        }

        static void Append(Claim claim, StringBuilder sb, string text, int start, int docStart)
        {
            string piece = text.Substring(start);
            if (piece.Length == 0) return;
            if (sb.Length > 0) sb.Append(' ');
            claim.Segments.Add(new TextSegment(sb.Length, docStart + start, piece.Length));
            sb.Append(piece);
        }

        static void Finish(Claim claim, StringBuilder sb)
        {
            claim.FullText = sb.ToString();
            var t = Transition.Match(claim.FullText);
            claim.Preamble = t.Success ? claim.FullText.Substring(0, t.Index) : claim.FullText;
            ParseDependencies(claim);
            claim.Category = DetectCategory(claim.Preamble);
            claim.PreambleNoun = DetectPreambleNoun(claim.Preamble);
        }

        static string CanonicalStatus(string status)
        {
            switch (status.ToLowerInvariant())
            {
                case "canceled":
                case "cancelled": return "Canceled";
                case "withdrawn": return "Withdrawn";
                case "original": return "Original";
                case "previously presented": return "Previously Presented";
                case "currently amended": return "Currently Amended";
                default: return "New";
            }
        }

        static void ParseDependencies(Claim claim)
        {
            string region = claim.Preamble;
            var pm = Preceding.Match(region);
            if (pm.Success)
            {
                for (int i = 1; i < claim.Number; i++) claim.DependsOn.Add(i);
                claim.IsMultipleDependent = true;
                claim.DependencyRefStart = pm.Index;
                claim.DependencyRefLength = pm.Length;
                return;
            }

            var matches = DepRef.Matches(region);
            if (matches.Count == 0) return;
            var first = matches[0];
            var last = matches[matches.Count - 1];
            claim.DependencyRefStart = first.Index;
            claim.DependencyRefLength = last.Index + last.Length - first.Index;

            bool hasRange = false;
            foreach (Match m in matches)
            {
                foreach (Match r in NumberOrRange.Matches(m.Groups[1].Value))
                {
                    int a = int.Parse(r.Groups[1].Value);
                    if (r.Groups[2].Success)
                    {
                        hasRange = true;
                        int b = int.Parse(r.Groups[2].Value);
                        for (int i = Math.Min(a, b); i <= Math.Max(a, b); i++) AddDependency(claim, i);
                    }
                    else AddDependency(claim, a);
                }
            }
            claim.IsMultipleDependent = hasRange || claim.DependsOn.Count > 1;
            string span = region.Substring(claim.DependencyRefStart, claim.DependencyRefLength);
            claim.HasCumulativeDependency = claim.IsMultipleDependent && AndRx.IsMatch(span);
        }

        static void AddDependency(Claim claim, int number)
        {
            if (!claim.DependsOn.Contains(number)) claim.DependsOn.Add(number);
        }

        static ClaimCategory DetectCategory(string preamble)
        {
            int best = int.MaxValue;
            var result = ClaimCategory.Unknown;
            foreach (var (pattern, category) in CategoryKeywords)
            {
                var m = pattern.Match(preamble);
                if (m.Success && m.Index < best)
                {
                    best = m.Index;
                    result = category;
                }
            }
            return result;
        }

        static string DetectPreambleNoun(string preamble)
        {
            string p = LeadingDeterminer.Replace(preamble, "", 1);
            var stop = NounStop.Match(p);
            string phrase = stop.Success ? p.Substring(0, stop.Index) : p;
            var words = phrase.Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries);
            return words.Length == 0 ? "" : TextUtil.Singularize(words[words.Length - 1].ToLowerInvariant());
        }
    }
}
```

- [ ] **Step 5: Run tests to verify they pass**

Run: `dotnet test LLDOCX.Claims.Tests`
Expected: all tests pass, 0 failed.

- [ ] **Step 6: Commit**

```bash
git add LLDOCX/LLDOCX.Claims/ClaimParser.cs LLDOCX/LLDOCX.Claims.Tests
git commit -m "feat(claims): parse claim numbering, status, dependencies, and category

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 4: DependencyChecker

**Files:**
- Create: `LLDOCX/LLDOCX.Claims/DependencyChecker.cs`
- Test: `LLDOCX/LLDOCX.Claims.Tests/DependencyCheckerTests.cs`

**Interfaces:**
- Consumes: `ClaimParser.Parse`, `Claim`, `Finding.At` (Tasks 2–3)
- Produces: `static class DependencyChecker { static List<Finding> Check(IList<Claim> claims); }`. All findings have `Type = FindingType.Dependency`.

- [ ] **Step 1: Write the failing tests `LLDOCX/LLDOCX.Claims.Tests/DependencyCheckerTests.cs`**

```csharp
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
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `dotnet test LLDOCX.Claims.Tests --filter DependencyCheckerTests`
Expected: build FAILS with `The name 'DependencyChecker' does not exist in the current context`.

- [ ] **Step 3: Write `LLDOCX/LLDOCX.Claims/DependencyChecker.cs`**

```csharp
using System;
using System.Collections.Generic;
using System.Linq;

namespace LLDOCX.Claims
{
    public static class DependencyChecker
    {
        public static List<Finding> Check(IList<Claim> claims)
        {
            var findings = new List<Finding>();
            var byNumber = new Dictionary<int, Claim>();
            foreach (var c in claims)
                if (!byNumber.ContainsKey(c.Number)) byNumber[c.Number] = c;

            for (int i = 0; i < claims.Count; i++)
            {
                var c = claims[i];
                if (!ReferenceEquals(byNumber[c.Number], c))
                    findings.Add(AtStart(c, Severity.Error, $"Claim number {c.Number} is used more than once."));
                if (i > 0)
                {
                    int prev = claims[i - 1].Number;
                    if (c.Number != prev && c.Number != prev + 1)
                        findings.Add(AtStart(c, Severity.Warning, $"Claim numbering jumps from {prev} to {c.Number}."));
                }
            }

            foreach (var c in claims)
            {
                if (c.IsCanceled || c.IsIndependent) continue;

                foreach (var d in c.DependsOn)
                {
                    string message = null;
                    if (d == c.Number) message = $"Claim {c.Number} depends on itself.";
                    else if (d > c.Number) message = $"Claim {c.Number} depends on claim {d}, which comes after it. A claim may only refer to a preceding claim.";
                    else if (!byNumber.ContainsKey(d)) message = $"Claim {c.Number} refers to claim {d}, which does not exist.";
                    else if (byNumber[d].IsCanceled) message = $"Claim {c.Number} depends on canceled claim {d}.";
                    if (message != null) findings.Add(AtReference(c, Severity.Error, message));
                }

                if (c.HasCumulativeDependency)
                    findings.Add(AtReference(c, Severity.Error,
                        $"Claim {c.Number} refers to more than one claim cumulatively ('and'). A multiple dependent claim may refer to other claims only in the alternative ('or') (37 CFR 1.75(c))."));

                if (c.IsMultipleDependent)
                {
                    var multiParents = c.DependsOn
                        .Where(d => d < c.Number && byNumber.ContainsKey(d) && byNumber[d].IsMultipleDependent)
                        .ToList();
                    if (multiParents.Count > 0)
                        findings.Add(AtReference(c, Severity.Error,
                            $"Multiple dependent claim {c.Number} depends on multiple dependent claim(s) {string.Join(", ", multiParents)}, which is not permitted (37 CFR 1.75(c))."));
                    else
                        findings.Add(AtReference(c, Severity.Info,
                            $"Claim {c.Number} is a multiple dependent claim (an additional fee applies)."));
                }

                var parent = FirstValidParent(c, byNumber);
                if (parent == null) continue;
                var root = Root(parent, byNumber);
                if (c.Category != ClaimCategory.Unknown && parent.Category != ClaimCategory.Unknown && c.Category != parent.Category)
                {
                    findings.Add(AtReference(c, Severity.Warning,
                        $"Claim {c.Number} claims a different category ({Name(c.Category)}) than claim {parent.Number} ({Name(parent.Category)}), on which it depends."));
                }
                else if (c.Category == root.Category && c.PreambleNoun.Length > 0 && root.PreambleNoun.Length > 0 && c.PreambleNoun != root.PreambleNoun)
                {
                    findings.Add(AtReference(c, Severity.Warning,
                        $"Claim {c.Number} refers to 'the {c.PreambleNoun}' but claim {root.Number} claims 'a {root.PreambleNoun}'."));
                }
            }
            return findings;
        }

        static Claim FirstValidParent(Claim c, Dictionary<int, Claim> byNumber)
        {
            foreach (var d in c.DependsOn)
            {
                Claim p;
                if (d < c.Number && byNumber.TryGetValue(d, out p) && !p.IsCanceled) return p;
            }
            return null;
        }

        static Claim Root(Claim c, Dictionary<int, Claim> byNumber)
        {
            var visited = new HashSet<int>();
            var current = c;
            while (!current.IsIndependent && visited.Add(current.Number))
            {
                var next = FirstValidParent(current, byNumber);
                if (next == null) break;
                current = next;
            }
            return current;
        }

        static string Name(ClaimCategory category)
        {
            switch (category)
            {
                case ClaimCategory.Method: return "method";
                case ClaimCategory.Apparatus: return "apparatus/system";
                case ClaimCategory.Composition: return "composition";
                case ClaimCategory.ComputerReadableMedium: return "computer-readable medium";
                default: return "unknown";
            }
        }

        static Finding AtStart(Claim c, Severity severity, string message) =>
            Finding.At(c, 0, Math.Min(20, c.FullText.Length), FindingType.Dependency, severity, message);

        static Finding AtReference(Claim c, Severity severity, string message) =>
            c.DependencyRefStart >= 0
                ? Finding.At(c, c.DependencyRefStart, c.DependencyRefLength, FindingType.Dependency, severity, message)
                : AtStart(c, severity, message);
    }
}
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `dotnet test LLDOCX.Claims.Tests`
Expected: all tests pass, 0 failed.

- [ ] **Step 5: Commit**

```bash
git add LLDOCX/LLDOCX.Claims/DependencyChecker.cs LLDOCX/LLDOCX.Claims.Tests/DependencyCheckerTests.cs
git commit -m "feat(claims): check claim dependencies and numbering

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 5: NounPhraseExtractor

**Files:**
- Create: `LLDOCX/LLDOCX.Claims/NounPhraseExtractor.cs`
- Test: `LLDOCX/LLDOCX.Claims.Tests/NounPhraseExtractorTests.cs`

**Interfaces:**
- Consumes: `TextUtil.Tokenize`, `TextUtil.Singularize`, `TextUtil.IsPlural`, `Token` (Task 2)
- Produces:
  - `enum DeterminerKind { Introducing, Referring }`
  - `class NounPhrase`: `Kind`, `Determiner` (lowercase, e.g. `"the at least one"`), `Words` (lowercase list), `Start`, `Length`, `WordsStart`, `Text`, `WordsText`, `IsClaimReference`, `Head`, `HeadKey`, `Key`, `ModifierKeys` (List<string>), `IsPluralForm`
  - `static class NounPhraseExtractor { static List<NounPhrase> Extract(string text); }`

**Heuristics:** The extractor is deliberately simple. A phrase is a determiner, then up to 6 words, ending at the first boundary word or punctuation mark. After the first non-ordinal word, a word ending in `-ing`, `-ed`, `-ly`, or a plural `-s` also ends the phrase, unless it is in `NounExceptions`. Errors from these rules are consistent between the introduction and the later reference ("a plurality of temperature sensors" and "the temperature sensors" both reduce to "temperature"), so they rarely cause false antecedent errors.

- [ ] **Step 1: Write the failing tests `LLDOCX/LLDOCX.Claims.Tests/NounPhraseExtractorTests.cs`**

```csharp
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
}
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `dotnet test LLDOCX.Claims.Tests --filter NounPhraseExtractorTests`
Expected: build FAILS with `The name 'NounPhraseExtractor' does not exist in the current context`.

- [ ] **Step 3: Write `LLDOCX/LLDOCX.Claims/NounPhraseExtractor.cs`**

```csharp
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace LLDOCX.Claims
{
    public enum DeterminerKind { Introducing, Referring }

    public class NounPhrase
    {
        public DeterminerKind Kind { get; set; }
        public string Determiner { get; set; } = "";
        public List<string> Words { get; } = new List<string>();
        /// <summary>Index of the determiner in the claim text.</summary>
        public int Start { get; set; }
        /// <summary>Length from the determiner through the last word.</summary>
        public int Length { get; set; }
        public int WordsStart { get; set; }
        public string Text { get; set; } = "";
        public string WordsText { get; set; } = "";
        /// <summary>True for "the method" in "The method of claim 1": a reference to another claim, not to an element.</summary>
        public bool IsClaimReference { get; set; }

        public string Head => Words[Words.Count - 1];
        public string HeadKey => TextUtil.Singularize(Head);
        public string Key => string.Join(" ", Words.Select(TextUtil.Singularize));
        public List<string> ModifierKeys => Words.Take(Words.Count - 1).Select(TextUtil.Singularize).ToList();
        public bool IsPluralForm => TextUtil.IsPlural(Head);
    }

    public static class NounPhraseExtractor
    {
        const int MaxWords = 6;

        static readonly (string[] Words, DeterminerKind Kind)[] Determiners = BuildDeterminers();

        static readonly HashSet<string> Boundary = new HashSet<string>
        {
            "of", "to", "for", "in", "on", "with", "from", "by", "at", "between", "into", "onto", "through", "via",
            "over", "under", "within", "along", "about", "against", "during", "until", "upon", "after", "before",
            "above", "below", "across", "around", "toward", "towards", "beyond", "near",
            "and", "or", "but", "nor",
            "is", "are", "was", "were", "be", "been", "being", "has", "have", "had", "having", "can", "may", "will",
            "must", "should", "shall", "would", "could", "does", "do", "did",
            "comprising", "comprises", "comprise", "including", "includes", "include", "containing", "contains",
            "contain", "consisting", "consists",
            "wherein", "whereby", "which", "that", "who", "whom", "whose", "when", "where", "while", "whether", "if",
            "than", "then", "thereby", "therein", "thereof", "thereto", "therefrom", "thereon", "such", "so", "as",
            "configured", "adapted", "operable", "arranged", "coupled", "connected", "disposed", "positioned",
            "located", "mounted", "attached", "defined", "provided", "formed",
            "a", "an", "the", "said", "another", "each", "every", "any", "all", "both", "either", "neither", "no",
            "not", "only", "also", "further", "respectively", "claim", "claims"
        };

        static readonly HashSet<string> Ordinals = new HashSet<string>
        {
            "first", "second", "third", "fourth", "fifth", "sixth", "seventh", "eighth", "ninth", "tenth"
        };

        /// <summary>Nouns that look like verbs/adverbs/plurals and must not end a phrase.</summary>
        static readonly HashSet<string> NounExceptions = new HashSet<string>
        {
            "housing", "bearing", "coating", "opening", "spring", "ring", "string", "fitting", "casing", "wiring",
            "lining", "setting", "building", "ceiling", "drawing", "mapping", "padding", "thing", "wing", "awning",
            "tubing", "piping", "molding", "moulding", "flooring", "siding", "packaging", "clothing", "railing",
            "fencing", "lighting",
            "bed", "red", "speed", "feed", "seed", "need", "shed", "reed",
            "assembly", "subassembly", "supply", "family", "anomaly", "reply", "poly",
            "means", "series", "species", "lens", "gas", "chassis", "bias", "canvas", "atlas"
        };

        static readonly Regex ClaimRefAfter = new Regex(
            @"\G\s*(?:,\s*)?(?:of|according\s+to|as\s+(?:claimed|recited|defined|set\s+forth|described)\s+in|in)\s+(?:any\s+(?:one\s+)?of\s+)?(?:the\s+(?:preceding|previous|foregoing)\s+)?claims?\b",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        static (string[] Words, DeterminerKind Kind)[] BuildDeterminers()
        {
            var list = new List<(string[] Words, DeterminerKind Kind)>();
            foreach (var d in new[] { "each of said", "each of the", "the at least one", "the one or more", "the plurality of",
                                      "said at least one", "said one or more", "said plurality of", "the", "said" })
                list.Add((d.Split(' '), DeterminerKind.Referring));
            foreach (var d in new[] { "a plurality of", "at least one", "one or more", "plurality of", "another", "a", "an" })
                list.Add((d.Split(' '), DeterminerKind.Introducing));
            return list.OrderByDescending(x => x.Words.Length).ToArray();
        }

        public static List<NounPhrase> Extract(string text)
        {
            var result = new List<NounPhrase>();
            if (string.IsNullOrEmpty(text)) return result;
            var tokens = TextUtil.Tokenize(text);
            int i = 0;
            while (i < tokens.Count)
            {
                int at = i;
                var det = Determiners.FirstOrDefault(d => Matches(tokens, at, d.Words));
                if (det.Words == null)
                {
                    i++;
                    continue;
                }

                var phrase = new NounPhrase
                {
                    Kind = det.Kind,
                    Determiner = string.Join(" ", det.Words),
                    Start = tokens[i].Index
                };
                int j = i + det.Words.Length;
                int nonOrdinal = 0;
                int lastEnd = -1;
                while (j < tokens.Count && phrase.Words.Count < MaxWords)
                {
                    var tok = tokens[j];
                    if (!tok.IsWord) break;
                    string w = tok.Lower;
                    if (Boundary.Contains(w)) break;
                    if (nonOrdinal > 0 && !NounExceptions.Contains(w) &&
                        (w.EndsWith("ing") || w.EndsWith("ed") || w.EndsWith("ly") || TextUtil.IsPlural(w)))
                        break;
                    if (phrase.Words.Count == 0) phrase.WordsStart = tok.Index;
                    phrase.Words.Add(w);
                    if (!Ordinals.Contains(w)) nonOrdinal++;
                    lastEnd = tok.End;
                    j++;
                }

                if (phrase.Words.Count > 0 && nonOrdinal > 0)
                {
                    phrase.Length = lastEnd - phrase.Start;
                    phrase.Text = text.Substring(phrase.Start, phrase.Length);
                    phrase.WordsText = text.Substring(phrase.WordsStart, lastEnd - phrase.WordsStart);
                    phrase.IsClaimReference = ClaimRefAfter.Match(text, lastEnd).Success;
                    result.Add(phrase);
                    i = j;
                }
                else
                {
                    i += det.Words.Length;
                }
            }
            return result;
        }

        static bool Matches(List<Token> tokens, int i, string[] words)
        {
            if (i + words.Length > tokens.Count) return false;
            for (int k = 0; k < words.Length; k++)
            {
                var t = tokens[i + k];
                if (!t.IsWord || t.Lower != words[k]) return false;
            }
            return true;
        }
    }
}
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `dotnet test LLDOCX.Claims.Tests`
Expected: all tests pass, 0 failed.

- [ ] **Step 5: Commit**

```bash
git add LLDOCX/LLDOCX.Claims/NounPhraseExtractor.cs LLDOCX/LLDOCX.Claims.Tests/NounPhraseExtractorTests.cs
git commit -m "feat(claims): extract introducing and referring noun phrases

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 6: AntecedentChecker

**Files:**
- Create: `LLDOCX/LLDOCX.Claims/AntecedentChecker.cs`
- Test: `LLDOCX/LLDOCX.Claims.Tests/AntecedentCheckerTests.cs`

**Interfaces:**
- Consumes: `Claim`, `Finding.At`, `NounPhraseExtractor.Extract`, `NounPhrase`, `DeterminerKind`, `TextUtil.Levenshtein`, `TextUtil.Singularize` (Tasks 2, 5)
- Produces: `static class AntecedentChecker { static List<Finding> Check(IList<Claim> claims, IEnumerable<string> allowlist, bool reportPlurality = true); }`. Findings have `Type = Antecedent`. The singular/plural drift findings have `Type = Terminology` and are emitted only when `reportPlurality` is true.

- [ ] **Step 1: Write the failing tests `LLDOCX/LLDOCX.Claims.Tests/AntecedentCheckerTests.cs`**

```csharp
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
}
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `dotnet test LLDOCX.Claims.Tests --filter AntecedentCheckerTests`
Expected: build FAILS with `The name 'AntecedentChecker' does not exist in the current context`.

- [ ] **Step 3: Write `LLDOCX/LLDOCX.Claims/AntecedentChecker.cs`**

```csharp
using System;
using System.Collections.Generic;
using System.Linq;

namespace LLDOCX.Claims
{
    public static class AntecedentChecker
    {
        /// <summary>Generic nouns that are routinely introduced more than once ("a portion of X ... a portion of Y").</summary>
        static readonly HashSet<string> GenericNouns = new HashSet<string>
        {
            "portion", "part", "end", "side", "surface", "number", "amount", "time", "set", "group", "pair",
            "value", "level", "degree", "piece", "section", "region", "area", "edge", "point", "step"
        };

        static readonly HashSet<string> Units = new HashSet<string>
        {
            "mm", "cm", "m", "um", "nm", "kg", "g", "mg", "ml", "l", "degree", "percent", "second", "minute",
            "hour", "hz", "khz", "mhz", "ghz", "volt", "v", "watt", "w", "celsius", "fahrenheit"
        };

        static readonly HashSet<string> PluralCapable = new HashSet<string>
        {
            "a plurality of", "plurality of", "one or more", "at least one"
        };

        public static List<Finding> Check(IList<Claim> claims, IEnumerable<string> allowlist, bool reportPlurality = true)
        {
            var findings = new List<Finding>();
            var allow = (allowlist ?? Enumerable.Empty<string>()).Select(NormalizeAllowEntry).Where(a => a.Length > 0).ToList();

            var byNumber = new Dictionary<int, Claim>();
            foreach (var c in claims)
                if (!c.IsCanceled && !byNumber.ContainsKey(c.Number)) byNumber[c.Number] = c;
            var phrases = byNumber.Values.ToDictionary(
                c => c.Number,
                c => NounPhraseExtractor.Extract(c.FullText).Where(p => !p.IsClaimReference).ToList());
            var context = new Context(byNumber, phrases);

            foreach (var claim in byNumber.Values.OrderBy(c => c.Number))
            {
                var known = new List<NounPhrase>(context.Inherited(claim));
                foreach (var p in phrases[claim.Number])
                {
                    if (p.Kind == DeterminerKind.Introducing)
                    {
                        if (p.Determiner != "another" && !GenericNouns.Contains(p.HeadKey) && known.Any(k => k.Key == p.Key))
                            findings.Add(Finding.At(claim, p.Start, p.Length, FindingType.Antecedent, Severity.Warning,
                                $"'{p.Text}' introduces a term that already has antecedent basis.",
                                $"If the same element is meant, use 'the {p.WordsText}'."));
                        known.Add(p);
                    }
                    else
                    {
                        CheckReference(claim, p, known, allow, reportPlurality, findings);
                    }
                }
            }
            return findings;
        }

        static void CheckReference(Claim claim, NounPhrase p, List<NounPhrase> known, List<string> allow,
            bool reportPlurality, List<Finding> findings)
        {
            if (IsAllowed(p, allow)) return;

            var exact = known.FirstOrDefault(k => k.Key == p.Key);
            if (exact != null)
            {
                if (reportPlurality) CheckPlurality(claim, p, exact, findings);
                return;
            }

            var mods = p.ModifierKeys;
            var candidates = known
                .Where(k => k.HeadKey == p.HeadKey && mods.All(m => k.ModifierKeys.Contains(m)))
                .GroupBy(k => k.Key).Select(g => g.First()).ToList();
            if (candidates.Count == 1)
            {
                if (reportPlurality) CheckPlurality(claim, p, candidates[0], findings);
                return;
            }
            if (candidates.Count > 1)
            {
                string options = string.Join(" or ", candidates.Select(k => $"'{k.Text}'"));
                findings.Add(Finding.At(claim, p.Start, p.Length, FindingType.Antecedent, Severity.Error,
                    $"'{p.Text}' is ambiguous: it could refer to {options}.",
                    $"Say which one, e.g. 'the {candidates[0].WordsText}'."));
                return;
            }

            var weak = known.FirstOrDefault(k => k.HeadKey == p.HeadKey);
            if (weak != null)
            {
                findings.Add(Finding.At(claim, p.Start, p.Length, FindingType.Antecedent, Severity.Info,
                    $"'{p.Text}' only partly matches '{weak.Text}' (same noun, different modifiers). Check its antecedent basis."));
                return;
            }

            var close = known
                .Where(k => k.Key.Length >= 5 && TextUtil.Levenshtein(k.Key, p.Key) <= 2)
                .OrderBy(k => TextUtil.Levenshtein(k.Key, p.Key))
                .FirstOrDefault();
            string suggestion = close != null
                ? $"Did you mean 'the {close.WordsText}'?"
                : $"Introduce it first (e.g. 'a {p.WordsText}') or correct the term.";
            findings.Add(Finding.At(claim, p.Start, p.Length, FindingType.Antecedent, Severity.Error,
                $"No antecedent basis for '{p.Text}'.", suggestion));
        }

        static void CheckPlurality(Claim claim, NounPhrase p, NounPhrase basis, List<Finding> findings)
        {
            if (PluralCapable.Contains(basis.Determiner) || p.Determiner.StartsWith("each")) return;
            if (p.IsPluralForm == basis.IsPluralForm) return;
            string pForm = p.IsPluralForm ? "plural" : "singular";
            string bForm = basis.IsPluralForm ? "plural" : "singular";
            findings.Add(Finding.At(claim, p.Start, p.Length, FindingType.Terminology, Severity.Warning,
                $"'{p.Text}' is {pForm} but its antecedent '{basis.Text}' is {bForm}.",
                $"Use 'the {basis.WordsText}'."));
        }

        static bool IsAllowed(NounPhrase p, List<string> allow)
        {
            if (Units.Contains(p.HeadKey)) return true;
            string key = p.Key;
            return allow.Any(a => key == a || key.StartsWith(a + " "));
        }

        static string NormalizeAllowEntry(string entry)
        {
            var words = (entry ?? "").ToLowerInvariant()
                .Split(new[] { ' ', '\t' }, StringSplitOptions.RemoveEmptyEntries).ToList();
            if (words.Count > 0 && (words[0] == "the" || words[0] == "said")) words.RemoveAt(0);
            return string.Join(" ", words.Select(TextUtil.Singularize));
        }

        sealed class Context
        {
            readonly Dictionary<int, Claim> _byNumber;
            readonly Dictionary<int, List<NounPhrase>> _phrases;
            readonly Dictionary<int, List<NounPhrase>> _closure = new Dictionary<int, List<NounPhrase>>();
            readonly HashSet<int> _visiting = new HashSet<int>();

            public Context(Dictionary<int, Claim> byNumber, Dictionary<int, List<NounPhrase>> phrases)
            {
                _byNumber = byNumber;
                _phrases = phrases;
            }

            /// <summary>Terms introduced by the claim's ancestors; for multiple dependency, only terms every branch provides.</summary>
            public List<NounPhrase> Inherited(Claim claim)
            {
                var parents = claim.DependsOn.Where(d => d < claim.Number && _byNumber.ContainsKey(d)).Distinct().ToList();
                if (parents.Count == 0) return new List<NounPhrase>();
                var sets = parents.Select(Closure).ToList();
                return sets[0].Where(p => sets.Skip(1).All(s => s.Any(q => q.Key == p.Key))).ToList();
            }

            List<NounPhrase> Closure(int number)
            {
                List<NounPhrase> cached;
                if (_closure.TryGetValue(number, out cached)) return cached;
                if (!_visiting.Add(number)) return new List<NounPhrase>();
                var result = Inherited(_byNumber[number]);
                result.AddRange(_phrases[number].Where(p => p.Kind == DeterminerKind.Introducing));
                _visiting.Remove(number);
                _closure[number] = result;
                return result;
            }
        }
    }
}
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `dotnet test LLDOCX.Claims.Tests`
Expected: all tests pass, 0 failed.

- [ ] **Step 5: Commit**

```bash
git add LLDOCX/LLDOCX.Claims/AntecedentChecker.cs LLDOCX/LLDOCX.Claims.Tests/AntecedentCheckerTests.cs
git commit -m "feat(claims): check antecedent basis across dependency chains

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 7: TerminologyChecker

**Files:**
- Create: `LLDOCX/LLDOCX.Claims/TerminologyChecker.cs`
- Test: `LLDOCX/LLDOCX.Claims.Tests/TerminologyCheckerTests.cs`

**Interfaces:**
- Consumes: `Claim`, `Finding.At`, `NounPhraseExtractor`, `TextUtil.Levenshtein` (Tasks 2, 5)
- Produces: `static class TerminologyChecker { static List<Finding> Check(IList<Claim> claims); }`. Spelling variants and the said/the mix are `Type = Terminology`. Repeated words and punctuation are `Type = Formatting`.

- [ ] **Step 1: Write the failing tests `LLDOCX/LLDOCX.Claims.Tests/TerminologyCheckerTests.cs`**

```csharp
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
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `dotnet test LLDOCX.Claims.Tests --filter TerminologyCheckerTests`
Expected: build FAILS with `The name 'TerminologyChecker' does not exist in the current context`.

- [ ] **Step 3: Write `LLDOCX/LLDOCX.Claims/TerminologyChecker.cs`**

```csharp
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace LLDOCX.Claims
{
    public static class TerminologyChecker
    {
        static readonly Regex WordRx = new Regex(@"[A-Za-z][A-Za-z\-]*", RegexOptions.Compiled);
        static readonly Regex RepeatedRx = new Regex(@"\b([A-Za-z]+)\s+\1\b", RegexOptions.Compiled | RegexOptions.IgnoreCase);
        static readonly Regex InnerPeriodRx = new Regex(
            @"(?<!\b(?:e\.g|i\.e|etc|approx|no|fig|figs|wt|vol|ca|vs|cf|al|inc|ltd|co))\.(?=\s+\S)",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);
        static readonly string[] InflectionSuffixes = { "ing", "ed", "es", "s", "er" };

        public static List<Finding> Check(IList<Claim> claims)
        {
            var active = claims.Where(c => !c.IsCanceled && c.FullText.Length > 0).ToList();
            var findings = new List<Finding>();
            CheckSpellingVariants(active, findings);
            CheckSaidTheMix(active, findings);
            foreach (var c in active) CheckFormatting(c, findings);
            return findings;
        }

        static void CheckSpellingVariants(List<Claim> claims, List<Finding> findings)
        {
            var occurrences = new Dictionary<string, List<KeyValuePair<Claim, Match>>>();
            foreach (var c in claims)
            {
                foreach (Match m in WordRx.Matches(c.FullText))
                {
                    string w = m.Value.ToLowerInvariant();
                    if (w.Length < 5) continue;
                    List<KeyValuePair<Claim, Match>> list;
                    if (!occurrences.TryGetValue(w, out list))
                        occurrences[w] = list = new List<KeyValuePair<Claim, Match>>();
                    list.Add(new KeyValuePair<Claim, Match>(c, m));
                }
            }

            foreach (var rare in occurrences)
            {
                string best = null;
                int bestCount = 0;
                foreach (var common in occurrences)
                {
                    int count = common.Value.Count;
                    if (common.Key == rare.Key || count < 3 * rare.Value.Count || count <= bestCount) continue;
                    if (IsInflection(rare.Key, common.Key)) continue;
                    int maxDistance = Math.Max(rare.Key.Length, common.Key.Length) >= 8 ? 2 : 1;
                    if (Math.Abs(rare.Key.Length - common.Key.Length) > maxDistance) continue;
                    if (TextUtil.Levenshtein(rare.Key, common.Key) > maxDistance) continue;
                    best = common.Key;
                    bestCount = count;
                }
                if (best == null) continue;
                foreach (var occ in rare.Value)
                    findings.Add(Finding.At(occ.Key, occ.Value.Index, occ.Value.Length, FindingType.Terminology, Severity.Warning,
                        $"'{occ.Value.Value}' may be a typo or variant of '{best}' (used {bestCount} times).",
                        $"Use '{best}' consistently."));
            }
        }

        static bool IsInflection(string a, string b) => a.StartsWith(b) || b.StartsWith(a) || Stem(a) == Stem(b);

        static string Stem(string word)
        {
            foreach (var s in InflectionSuffixes)
                if (word.EndsWith(s) && word.Length - s.Length >= 3) return word.Substring(0, word.Length - s.Length);
            return word;
        }

        static void CheckSaidTheMix(List<Claim> claims, List<Finding> findings)
        {
            var refs = new List<KeyValuePair<Claim, NounPhrase>>();
            foreach (var c in claims)
                foreach (var p in NounPhraseExtractor.Extract(c.FullText))
                    if (p.Kind == DeterminerKind.Referring && !p.IsClaimReference)
                        refs.Add(new KeyValuePair<Claim, NounPhrase>(c, p));

            int said = refs.Count(r => r.Value.Determiner.Contains("said"));
            int the = refs.Count - said;
            if (said == 0 || the == 0) return;
            bool saidIsMinority = said <= the;
            var first = refs.First(r => r.Value.Determiner.Contains("said") == saidIsMinority);
            findings.Add(Finding.At(first.Key, first.Value.Start, first.Value.Length, FindingType.Terminology, Severity.Info,
                $"The claims use both 'said' ({said}x) and 'the' ({the}x) to refer back to elements.",
                "Consider using one consistently."));
        }

        static void CheckFormatting(Claim c, List<Finding> findings)
        {
            string text = c.FullText;
            foreach (Match m in RepeatedRx.Matches(text))
                findings.Add(Finding.At(c, m.Index, m.Length, FindingType.Formatting, Severity.Error,
                    $"Repeated word '{m.Groups[1].Value}'.", $"Delete the extra '{m.Groups[1].Value}'."));

            string trimmed = text.TrimEnd();
            if (!trimmed.EndsWith("."))
                findings.Add(Finding.At(c, Math.Max(0, trimmed.Length - 1), 1, FindingType.Formatting, Severity.Warning,
                    $"Claim {c.Number} does not end with a period."));
            else if (trimmed.EndsWith(".."))
                findings.Add(Finding.At(c, trimmed.Length - 2, 2, FindingType.Formatting, Severity.Warning,
                    $"Claim {c.Number} ends with more than one period."));

            foreach (Match m in InnerPeriodRx.Matches(text))
                findings.Add(Finding.At(c, m.Index, 1, FindingType.Formatting, Severity.Warning,
                    $"Claim {c.Number} contains a period before its end. A claim must be a single sentence."));

            if (text.Count(ch => ch == '(') != text.Count(ch => ch == ')') ||
                text.Count(ch => ch == '[') != text.Count(ch => ch == ']'))
                findings.Add(Finding.At(c, 0, Math.Min(20, text.Length), FindingType.Formatting, Severity.Warning,
                    $"Unbalanced parentheses or brackets in claim {c.Number}."));
        }
    }
}
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `dotnet test LLDOCX.Claims.Tests`
Expected: all tests pass, 0 failed.

- [ ] **Step 5: Commit**

```bash
git add LLDOCX/LLDOCX.Claims/TerminologyChecker.cs LLDOCX/LLDOCX.Claims.Tests/TerminologyCheckerTests.cs
git commit -m "feat(claims): check terminology variants and claim formatting

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 8: ClaimsAnalyzer facade and false-positive guard

**Files:**
- Create: `LLDOCX/LLDOCX.Claims/ClaimsAnalyzer.cs`
- Test: `LLDOCX/LLDOCX.Claims.Tests/ClaimsAnalyzerTests.cs`
- Create: `LLDOCX/LLDOCX.Claims.Tests/Fixtures/real1.txt`, `real2.txt`, `real3.txt`

**Interfaces:**
- Consumes: `ClaimParser`, `DependencyChecker`, `AntecedentChecker`, `TerminologyChecker` (Tasks 3–7)
- Produces:
  - `class ClaimsCheckOptions { bool CheckAntecedent = true; bool CheckDependency = true; bool CheckTerminology = true; Severity MinSeverity = Info; List<string> Allowlist; static readonly string[] DefaultAllowlist; }`
  - `class AnalysisResult { List<Claim> Claims; List<Finding> Findings; }`
  - `static class ClaimsAnalyzer { static AnalysisResult Analyze(IList<RawParagraph> paragraphs, ClaimsCheckOptions options); }`. Findings are filtered by `MinSeverity` and sorted by claim number, then by document offset.

- [ ] **Step 1: Write the failing tests `LLDOCX/LLDOCX.Claims.Tests/ClaimsAnalyzerTests.cs`**

```csharp
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
}
```

- [ ] **Step 2: Run tests to verify they fail**

Run: `dotnet test LLDOCX.Claims.Tests --filter ClaimsAnalyzerTests`
Expected: build FAILS with `The name 'ClaimsAnalyzer' does not exist in the current context`.

- [ ] **Step 3: Write `LLDOCX/LLDOCX.Claims/ClaimsAnalyzer.cs`**

```csharp
using System.Collections.Generic;
using System.Linq;

namespace LLDOCX.Claims
{
    public class ClaimsCheckOptions
    {
        public static readonly string[] DefaultAllowlist =
        {
            "the user", "the art", "the same", "the following", "the foregoing", "the group",
            "the presence", "the absence", "the range", "the order"
        };

        public bool CheckAntecedent { get; set; } = true;
        public bool CheckDependency { get; set; } = true;
        public bool CheckTerminology { get; set; } = true;
        public Severity MinSeverity { get; set; } = Severity.Info;
        public List<string> Allowlist { get; set; } = new List<string>(DefaultAllowlist);
    }

    public class AnalysisResult
    {
        public AnalysisResult(List<Claim> claims, List<Finding> findings)
        {
            Claims = claims;
            Findings = findings;
        }

        public List<Claim> Claims { get; }
        public List<Finding> Findings { get; }
    }

    public static class ClaimsAnalyzer
    {
        public static AnalysisResult Analyze(IList<RawParagraph> paragraphs, ClaimsCheckOptions options)
        {
            options = options ?? new ClaimsCheckOptions();
            var claims = ClaimParser.Parse(paragraphs);
            var findings = new List<Finding>();
            if (options.CheckDependency) findings.AddRange(DependencyChecker.Check(claims));
            if (options.CheckAntecedent) findings.AddRange(AntecedentChecker.Check(claims, options.Allowlist, options.CheckTerminology));
            if (options.CheckTerminology) findings.AddRange(TerminologyChecker.Check(claims));
            var visible = findings
                .Where(f => f.Severity <= options.MinSeverity)
                .OrderBy(f => f.ClaimNumber)
                .ThenBy(f => f.DocStart)
                .ToList();
            return new AnalysisResult(claims, visible);
        }
    }
}
```

- [ ] **Step 4: Run tests to verify they pass**

Run: `dotnet test LLDOCX.Claims.Tests`
Expected: all tests pass, 0 failed.

If `Clean_claim_set_produces_no_errors_or_warnings` fails, the dump lists each false positive. Fix the rule that caused it, usually by adding a word to `Boundary` or `NounExceptions` in `NounPhraseExtractor.cs`. Do not weaken the test.

- [ ] **Step 5: Add three real claim sets as additional false-positive guards**

Copy the full claim text of three issued US patents in different fields (one mechanical, one software/method, one chemical composition) from Google Patents. Save them as `LLDOCX/LLDOCX.Claims.Tests/Fixtures/real1.txt`, `real2.txt`, and `real3.txt`, with one claim paragraph per line in the form `1. A ...`. Add to the test csproj:

```xml
  <ItemGroup>
    <None Update="Fixtures\*.txt" CopyToOutputDirectory="PreserveNewest" />
  </ItemGroup>
```

Add this test to `ClaimsAnalyzerTests`:

```csharp
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
```

Run: `dotnet test LLDOCX.Claims.Tests`
Expected: all pass. For each Error the dump shows, decide whether it is:
- **a real drafting problem in the issued patent:** acceptable, but replace that fixture with a different patent, since fixtures must be clean;
- **a heuristic false positive:** fix it in the extractor or checker word lists, add a focused unit test that reproduces it in the relevant `*Tests.cs`, and re-run.

- [ ] **Step 6: Commit**

```bash
git add LLDOCX/LLDOCX.Claims/ClaimsAnalyzer.cs LLDOCX/LLDOCX.Claims.Tests
git commit -m "feat(claims): add analyzer facade with options and false-positive guards

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 9: Deep Check library (prompt, response parser, Claude client)

**Files:**
- Create: `LLDOCX/LLDOCX.DeepCheck/LLDOCX.DeepCheck.csproj`
- Create: `LLDOCX/LLDOCX.DeepCheck/DeepCheckException.cs`
- Create: `LLDOCX/LLDOCX.DeepCheck/DeepCheckPrompt.cs`
- Create: `LLDOCX/LLDOCX.DeepCheck/DeepCheckResponseParser.cs`
- Create: `LLDOCX/LLDOCX.DeepCheck/DeepCheckClient.cs`
- Modify: `LLDOCX/LLDOCX.Claims.Tests/LLDOCX.Claims.Tests.csproj` (add project reference)
- Test: `LLDOCX/LLDOCX.Claims.Tests/DeepCheckProtocolTests.cs`

**Interfaces:**
- Consumes: `Claim`, `Finding`, `Finding.At`, `FindingType`, `Severity`, `ClaimParser` (Tasks 2–3)
- Produces (namespace `LLDOCX.DeepCheck`):
  - `class DeepCheckException : Exception` (user-facing message)
  - `static class DeepCheckPrompt { const string SystemPrompt; const string ResponseSchemaJson; static string BuildUserMessage(IList<Claim> claims, IList<Finding> ruleFindings); }`
  - `static class DeepCheckResponseParser { static List<Finding> Parse(string json, IList<Claim> claims, IList<Finding> ruleFindings); }`. It returns new AI findings (`Source = "AI"`, `DocStart = -1` when the quote isn't found) and sets `DisputedByAI` and `DisputeReason` on the disputed `ruleFindings` in place.
  - `static class DeepCheckClient { const string DefaultModel = "claude-opus-5-5"; static Task<List<Finding>> RunAsync(string apiKey, string model, IList<Claim> claims, IList<Finding> ruleFindings, CancellationToken ct); }`

- [ ] **Step 1: Create `LLDOCX/LLDOCX.DeepCheck/LLDOCX.DeepCheck.csproj` and wire up references**

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFramework>netstandard2.0</TargetFramework>
    <LangVersion>latest</LangVersion>
    <RootNamespace>LLDOCX.DeepCheck</RootNamespace>
  </PropertyGroup>
  <ItemGroup>
    <PackageReference Include="Anthropic" Version="13.3.0" />
  </ItemGroup>
  <ItemGroup>
    <ProjectReference Include="..\LLDOCX.Claims\LLDOCX.Claims.csproj" />
  </ItemGroup>
</Project>
```

Add to the test csproj's project-reference `ItemGroup`:

```xml
    <ProjectReference Include="..\LLDOCX.DeepCheck\LLDOCX.DeepCheck.csproj" />
```

Run: `dotnet sln LLDOCX.sln add LLDOCX.DeepCheck/LLDOCX.DeepCheck.csproj`
Expected: `Project ... added to the solution.`

- [ ] **Step 2: Write the failing tests `LLDOCX/LLDOCX.Claims.Tests/DeepCheckProtocolTests.cs`**

```csharp
using LLDOCX.Claims;
using LLDOCX.DeepCheck;

namespace LLDOCX.Claims.Tests;

[TestClass]
public class DeepCheckProtocolTests
{
    const string Line = "1. A lamp comprising a base, wherein the shade is green.";

    [TestMethod]
    public void User_message_lists_claims_and_indexed_rule_findings()
    {
        var claims = T.Claims(Line, "2. (Canceled)");
        var rule = new List<Finding> { new Finding { ClaimNumber = 1, Type = FindingType.Antecedent, Severity = Severity.Error, Message = "No antecedent basis for 'the shade'." } };
        string msg = DeepCheckPrompt.BuildUserMessage(claims, rule);
        StringAssert.Contains(msg, "1. A lamp comprising a base");
        StringAssert.Contains(msg, "2. (Canceled)");
        StringAssert.Contains(msg, "[0] claim 1, Antecedent, Error: No antecedent basis for 'the shade'.");
    }

    [TestMethod]
    public void Parser_maps_quotes_to_document_offsets_and_applies_disputes()
    {
        var claims = T.Claims(Line);
        var rule = new List<Finding> { new Finding { ClaimNumber = 1, Message = "x" } };
        const string json = """
            {"findings":[{"claim":1,"type":"Terminology","severity":"Warning","quote":"the shade","message":"m","suggestion":"s"}],
             "disputes":[{"index":0,"reason":"r"}]}
            """;
        var added = DeepCheckResponseParser.Parse(json, claims, rule);
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
        var claims = T.Claims(Line);
        const string json = """{"findings":[{"claim":1,"type":"Antecedent","severity":"Error","quote":"not there","message":"m","suggestion":""}],"disputes":[]}""";
        var added = DeepCheckResponseParser.Parse(json, claims, new List<Finding>());
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

    [TestMethod]
    public void Schema_is_valid_json_with_required_fields()
    {
        using var doc = System.Text.Json.JsonDocument.Parse(DeepCheckPrompt.ResponseSchemaJson);
        var required = doc.RootElement.GetProperty("required").EnumerateArray().Select(e => e.GetString()).ToArray();
        CollectionAssert.AreEquivalent(new[] { "findings", "disputes" }, required);
    }
}
```

- [ ] **Step 3: Run tests to verify they fail**

Run: `dotnet test LLDOCX.Claims.Tests --filter DeepCheckProtocolTests`
Expected: build FAILS with `The type or namespace name 'DeepCheckPrompt' could not be found`.

- [ ] **Step 4: Write `LLDOCX/LLDOCX.DeepCheck/DeepCheckException.cs`**

```csharp
using System;

namespace LLDOCX.DeepCheck
{
    /// <summary>A Deep Check failure whose Message is safe to show to the user.</summary>
    public class DeepCheckException : Exception
    {
        public DeepCheckException(string message, Exception inner = null) : base(message, inner) { }
    }
}
```

- [ ] **Step 5: Write `LLDOCX/LLDOCX.DeepCheck/DeepCheckPrompt.cs`**

```csharp
using System.Collections.Generic;
using System.Text;
using LLDOCX.Claims;

namespace LLDOCX.DeepCheck
{
    public static class DeepCheckPrompt
    {
        public const string SystemPrompt =
            "You review the claims of a US patent application for three kinds of problems only:\n" +
            "1. Antecedent basis: every element referred to with 'the' or 'said' must be introduced earlier in the same claim " +
            "or in a claim it depends on (for a multiple dependent claim, in every claim it refers to); references must not be ambiguous.\n" +
            "2. Claim dependency: references to missing, later, or canceled claims; multiple dependent claims that refer to claims " +
            "cumulatively or depend on another multiple dependent claim (37 CFR 1.75(c)); dependent claims whose category or preamble " +
            "does not match the claim they depend on.\n" +
            "3. Terminology: typos and inconsistent names for the same element.\n\n" +
            "An automated rule checker has already reviewed the claims. Its findings are listed with indexes in <rule_findings>. " +
            "Put in `findings` only problems the rule checker missed. Put in `disputes` the index of each rule finding that is a " +
            "false positive, with a one-sentence reason. For every finding, `quote` must be an exact substring of that claim's text, " +
            "as short as possible while unique within the claim. `suggestion` is a concrete fix, or an empty string. " +
            "Do not comment on patentability, claim scope, or style. If there is nothing to report, return empty arrays.";

        public const string ResponseSchemaJson = @"{
  ""type"": ""object"",
  ""properties"": {
    ""findings"": {
      ""type"": ""array"",
      ""items"": {
        ""type"": ""object"",
        ""properties"": {
          ""claim"": { ""type"": ""integer"" },
          ""type"": { ""type"": ""string"", ""enum"": [""Antecedent"", ""Dependency"", ""Terminology"", ""Formatting""] },
          ""severity"": { ""type"": ""string"", ""enum"": [""Error"", ""Warning"", ""Info""] },
          ""quote"": { ""type"": ""string"" },
          ""message"": { ""type"": ""string"" },
          ""suggestion"": { ""type"": ""string"" }
        },
        ""required"": [""claim"", ""type"", ""severity"", ""quote"", ""message"", ""suggestion""],
        ""additionalProperties"": false
      }
    },
    ""disputes"": {
      ""type"": ""array"",
      ""items"": {
        ""type"": ""object"",
        ""properties"": {
          ""index"": { ""type"": ""integer"" },
          ""reason"": { ""type"": ""string"" }
        },
        ""required"": [""index"", ""reason""],
        ""additionalProperties"": false
      }
    }
  },
  ""required"": [""findings"", ""disputes""],
  ""additionalProperties"": false
}";

        public static string BuildUserMessage(IList<Claim> claims, IList<Finding> ruleFindings)
        {
            var sb = new StringBuilder();
            sb.AppendLine("<claims>");
            foreach (var c in claims)
                sb.AppendLine(c.IsCanceled ? $"{c.Number}. (Canceled)" : $"{c.Number}. {c.FullText}");
            sb.AppendLine("</claims>");
            sb.AppendLine("<rule_findings>");
            for (int i = 0; i < ruleFindings.Count; i++)
            {
                var f = ruleFindings[i];
                sb.AppendLine($"[{i}] claim {f.ClaimNumber}, {f.Type}, {f.Severity}: {f.Message}");
            }
            sb.AppendLine("</rule_findings>");
            return sb.ToString();
        }
    }
}
```

- [ ] **Step 6: Write `LLDOCX/LLDOCX.DeepCheck/DeepCheckResponseParser.cs`**

```csharp
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using LLDOCX.Claims;

namespace LLDOCX.DeepCheck
{
    public static class DeepCheckResponseParser
    {
        public static List<Finding> Parse(string json, IList<Claim> claims, IList<Finding> ruleFindings)
        {
            JsonDocument doc;
            try
            {
                doc = JsonDocument.Parse(json);
            }
            catch (JsonException ex)
            {
                throw new DeepCheckException("Claude returned a response that could not be read. The rule-based findings are unchanged.", ex);
            }

            using (doc)
            {
                var added = new List<Finding>();
                var root = doc.RootElement;
                if (root.TryGetProperty("findings", out var findings) && findings.ValueKind == JsonValueKind.Array)
                    foreach (var f in findings.EnumerateArray())
                        added.Add(ToFinding(f, claims));

                if (root.TryGetProperty("disputes", out var disputes) && disputes.ValueKind == JsonValueKind.Array)
                {
                    foreach (var d in disputes.EnumerateArray())
                    {
                        int index = d.TryGetProperty("index", out var i) && i.ValueKind == JsonValueKind.Number ? i.GetInt32() : -1;
                        if (index < 0 || index >= ruleFindings.Count) continue;
                        ruleFindings[index].DisputedByAI = true;
                        ruleFindings[index].DisputeReason = Str(d, "reason");
                    }
                }
                return added;
            }
        }

        static Finding ToFinding(JsonElement f, IList<Claim> claims)
        {
            int number = f.TryGetProperty("claim", out var n) && n.ValueKind == JsonValueKind.Number ? n.GetInt32() : 0;
            var type = ParseEnum(Str(f, "type"), FindingType.Antecedent);
            var severity = ParseEnum(Str(f, "severity"), Severity.Warning);
            string quote = Str(f, "quote");
            string message = Str(f, "message");
            string suggestion = Str(f, "suggestion");
            if (suggestion.Length == 0) suggestion = null;

            var claim = claims.FirstOrDefault(c => c.Number == number);
            int idx = claim == null || quote.Length == 0 ? -1 : claim.FullText.IndexOf(quote, StringComparison.OrdinalIgnoreCase);
            var result = idx >= 0
                ? Finding.At(claim, idx, quote.Length, type, severity, message, suggestion)
                : new Finding { ClaimNumber = number, Type = type, Severity = severity, Message = message, Suggestion = suggestion, DocStart = -1, Text = quote };
            result.Source = "AI";
            return result;
        }

        static string Str(JsonElement e, string name) =>
            e.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.String ? v.GetString() ?? "" : "";

        static T ParseEnum<T>(string s, T fallback) where T : struct =>
            Enum.TryParse(s, true, out T value) ? value : fallback;
    }
}
```

- [ ] **Step 7: Run tests to verify they pass**

Run: `dotnet test LLDOCX.Claims.Tests`
Expected: all tests pass, 0 failed.

- [ ] **Step 8: Write `LLDOCX/LLDOCX.DeepCheck/DeepCheckClient.cs`**

```csharp
using System;
using System.Collections.Generic;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using Anthropic;
using Anthropic.Exceptions;
using Anthropic.Models.Beta.Messages;
using LLDOCX.Claims;

namespace LLDOCX.DeepCheck
{
    public static class DeepCheckClient
    {
        public const string DefaultModel = "claude-opus-5-5";

        public static async Task<List<Finding>> RunAsync(string apiKey, string model, IList<Claim> claims,
            IList<Finding> ruleFindings, CancellationToken ct)
        {
            if (string.IsNullOrWhiteSpace(apiKey))
                throw new DeepCheckException("No Claude API key is set. Add one on the Settings tab of the Claims Check pane.");

            string modelId = string.IsNullOrWhiteSpace(model) ? DefaultModel : model.Trim();
            var schema = JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(DeepCheckPrompt.ResponseSchemaJson);
            var client = new AnthropicClient { ApiKey = apiKey };
            var parameters = new MessageCreateParams
            {
                Model = modelId,
                MaxTokens = 16000,
                Betas = ["server-side-fallback-2026-07-01"],
                Fallbacks = new Default(),
                System = DeepCheckPrompt.SystemPrompt,
                OutputConfig = new BetaOutputConfig
                {
                    Effort = Effort.High,
                    Format = new BetaJsonOutputFormat { Schema = schema },
                },
                Messages = [new BetaUserMessageParam { Content = DeepCheckPrompt.BuildUserMessage(claims, ruleFindings) }],
            };

            BetaMessage response;
            try
            {
                response = await client.Beta.Messages.Create(parameters, cancellationToken: ct).ConfigureAwait(false);
            }
            catch (AnthropicUnauthorizedException ex)
            {
                throw new DeepCheckException("The Claude API key was rejected. Check it on the Settings tab of the Claims Check pane.", ex);
            }
            catch (AnthropicForbiddenException ex)
            {
                throw new DeepCheckException("The Claude API key does not have access to this model. Check the key and model ID in Settings.", ex);
            }
            catch (AnthropicNotFoundException ex)
            {
                throw new DeepCheckException($"Model '{modelId}' was not found. Check the model ID in Settings.", ex);
            }
            catch (AnthropicRateLimitException ex)
            {
                throw new DeepCheckException("Claude is rate-limiting requests. Wait a minute and try again.", ex);
            }
            catch (Anthropic5xxException ex)
            {
                throw new DeepCheckException("The Claude service had an error. Try again shortly.", ex);
            }
            catch (AnthropicIOException ex)
            {
                throw new DeepCheckException("Could not reach the Claude API. Check your network connection.", ex);
            }
            catch (AnthropicApiException ex)
            {
                throw new DeepCheckException("Claude API error: " + ex.Message, ex);
            }

            if (response.StopReason == "refusal")
                throw new DeepCheckException("Claude declined to review these claims. The rule-based findings are unchanged.");
            if (response.StopReason == "max_tokens")
                throw new DeepCheckException("Claude's response was cut off because the claim set is very large. Check a selection of claims instead.");

            var text = new StringBuilder();
            foreach (var block in response.Content)
                if (block.TryPickText(out BetaTextBlock t)) text.Append(t.Text);
            return DeepCheckResponseParser.Parse(text.ToString(), claims, ruleFindings);
        }
    }
}
```

- [ ] **Step 9: Build and fix SDK binding names**

Run: `dotnet build LLDOCX.DeepCheck`
Expected: `Build succeeded.` The SDK type and member names above come from the SDK's documented C# patterns and its XML docs. If the compiler reports an unknown name, look up the real one in `%USERPROFILE%\.nuget\packages\anthropic\13.3.0\lib\netstandard2.0\Anthropic.xml`, for example with `grep -n "BetaFallbacksParam\|T:Anthropic.Models.Beta.Messages.Default\|Exceptions\." Anthropic.xml`. Likely adjustments:
- `Fallbacks = new Default()`: if `Default` has no public constructor, look for a static instance or an implicit conversion on `BetaFallbacksParam`. The intent is the `"default"` mode of the `fallbacks` request field.
- The `Create(parameters, cancellationToken: ct)` parameter name.
- `response.StopReason == "refusal"`: if `StopReason` is an `ApiEnum`, compare its raw string value.
- Exception class names: their namespace is `Anthropic.Exceptions`.

Re-run until `Build succeeded.`, then run `dotnet test LLDOCX.Claims.Tests` again. Expected: all pass.

- [ ] **Step 10: Live smoke test (optional, needs an API key, costs one request)**

Create a scratch console app outside the repo, for example in `%TEMP%\deepcheck-smoke`. Run `dotnet new console`, add a project reference to `LLDOCX.DeepCheck`, and put this in `Program.cs`:

```csharp
using LLDOCX.Claims;
using LLDOCX.DeepCheck;
var result = ClaimsAnalyzer.Analyze(new List<RawParagraph> {
    new("1. A lamp comprising a base, wherein the shade is green.", 0) }, new ClaimsCheckOptions());
var added = await DeepCheckClient.RunAsync(Environment.GetEnvironmentVariable("ANTHROPIC_API_KEY"), null,
    result.Claims, result.Findings, CancellationToken.None);
Console.WriteLine($"AI findings: {added.Count}; disputed: {result.Findings.Count(f => f.DisputedByAI)}");
```

Run: `dotnet run`
Expected: a line like `AI findings: 0; disputed: 0`, with no exception. Delete the scratch app afterwards.

- [ ] **Step 11: Commit**

```bash
git add LLDOCX/LLDOCX.sln LLDOCX/LLDOCX.DeepCheck LLDOCX/LLDOCX.Claims.Tests
git commit -m "feat(deepcheck): add Claude Deep Check prompt, parser, and client

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 10: Add-in settings and API key storage

**Files:**
- Modify: `LLDOCX/LLDOCX/LLDOCX.csproj` (project refs, System.Security, new Compile items)
- Modify: `LLDOCX/LLDOCX/Options/SettingOption.cs`
- Modify: `LLDOCX/LLDOCX/CustomUserControl/SettingUserControl.cs:107` (`SaveAllSettng`)
- Create: `LLDOCX/LLDOCX/Options/ApiKeyStore.cs`

**Interfaces:**
- Consumes: `Severity`, `ClaimsCheckOptions.DefaultAllowlist` (Task 8), `DeepCheckClient.DefaultModel` (Task 9)
- Produces:
  - `SettingOption` properties: `bool ClaimsCheckAntecedent`, `bool ClaimsCheckDependency`, `bool ClaimsCheckTerminology`, `Severity ClaimsMinSeverity`, `List<string> ClaimsAllowlist`, `string DeepCheckModel`; method `void CopyClaimsSettingsFrom(SettingOption other)`; method `ClaimsCheckOptions ToClaimsCheckOptions()`
  - `static class ApiKeyStore { string Load(); void Save(string key); void Clear(); }` (namespace `LLDOCX.Options`)

- [ ] **Step 1: Add references and new files to `LLDOCX/LLDOCX/LLDOCX.csproj`**

In the `ItemGroup` that holds `<Reference Include="System" />`, add:

```xml
    <Reference Include="System.Security" />
```

After the `ItemGroup` with `<Compile Include="CustomEvents\MissingHeadingHandler.cs" />`, add a new `ItemGroup`:

```xml
  <ItemGroup>
    <ProjectReference Include="..\LLDOCX.Claims\LLDOCX.Claims.csproj" />
    <ProjectReference Include="..\LLDOCX.DeepCheck\LLDOCX.DeepCheck.csproj" />
  </ItemGroup>
```

In the `ItemGroup` with the other `Compile` items, add:

```xml
    <Compile Include="Options\ApiKeyStore.cs" />
```

- [ ] **Step 2: Replace `LLDOCX/LLDOCX/Options/SettingOption.cs`**

```csharp
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Serialization;
using System.Text;
using System.Threading.Tasks;
using LLDOCX.Claims;
using LLDOCX.DeepCheck;

namespace LLDOCX.Options
{
    [Serializable]
   public class SettingOption
    {
        public bool CheckPageSize { get; set; } = true;
        public bool CheckPageOrientation { get; set; } = true;
        public bool CheckMarkup { get; set; } = true;
        public bool CheckFont { get; set; } = true;
        public bool CheckHyperlink { get; set; } = true;
        public bool CheckComment { get; set; } = true;
        public bool CheckBookmarker { get; set; } = true;
        public bool RemoveDocumentProperty { get; set; } = true;
        public bool KeepTrackingOn { get; set; } = true;
        public bool HeadingValidation { get; set; } = true;
        public bool CheckedAll { get; set; } = true;
        public string DefaultFontName { get; set; } = "Times New Roman";

        // Claims Check settings (added in version 2; OptionalField keeps older settings files loadable)
        [field: OptionalField(VersionAdded = 2)]
        public bool ClaimsCheckAntecedent { get; set; } = true;
        [field: OptionalField(VersionAdded = 2)]
        public bool ClaimsCheckDependency { get; set; } = true;
        [field: OptionalField(VersionAdded = 2)]
        public bool ClaimsCheckTerminology { get; set; } = true;
        [field: OptionalField(VersionAdded = 2)]
        public Severity ClaimsMinSeverity { get; set; } = Severity.Info;
        [field: OptionalField(VersionAdded = 2)]
        public List<string> ClaimsAllowlist { get; set; } = new List<string>(ClaimsCheckOptions.DefaultAllowlist);
        [field: OptionalField(VersionAdded = 2)]
        public string DeepCheckModel { get; set; } = DeepCheckClient.DefaultModel;

        // Property initializers do not run during BinaryFormatter deserialization; set defaults
        // here so fields missing from an older settings file get them.
        [OnDeserializing]
        private void SetClaimsDefaults(StreamingContext context)
        {
            ClaimsCheckAntecedent = true;
            ClaimsCheckDependency = true;
            ClaimsCheckTerminology = true;
            ClaimsMinSeverity = Severity.Info;
            ClaimsAllowlist = new List<string>(ClaimsCheckOptions.DefaultAllowlist);
            DeepCheckModel = DeepCheckClient.DefaultModel;
        }

        public void CopyClaimsSettingsFrom(SettingOption other)
        {
            ClaimsCheckAntecedent = other.ClaimsCheckAntecedent;
            ClaimsCheckDependency = other.ClaimsCheckDependency;
            ClaimsCheckTerminology = other.ClaimsCheckTerminology;
            ClaimsMinSeverity = other.ClaimsMinSeverity;
            ClaimsAllowlist = new List<string>(other.ClaimsAllowlist ?? new List<string>(ClaimsCheckOptions.DefaultAllowlist));
            DeepCheckModel = other.DeepCheckModel;
        }

        public ClaimsCheckOptions ToClaimsCheckOptions()
        {
            return new ClaimsCheckOptions
            {
                CheckAntecedent = ClaimsCheckAntecedent,
                CheckDependency = ClaimsCheckDependency,
                CheckTerminology = ClaimsCheckTerminology,
                MinSeverity = ClaimsMinSeverity,
                Allowlist = new List<string>(ClaimsAllowlist ?? new List<string>(ClaimsCheckOptions.DefaultAllowlist))
            };
        }
    }
}
```

- [ ] **Step 3: Preserve claims settings when the existing Settings pane saves**

In `LLDOCX/LLDOCX/CustomUserControl/SettingUserControl.cs`, change the start of `SaveAllSettng()` from:

```csharp
        private void SaveAllSettng()
        {
            _option.CheckPageSize = ckbPageSize.Checked;
```

to:

```csharp
        private void SaveAllSettng()
        {
            // The Claims Check pane saves its own settings; don't overwrite them with this pane's stale copy.
            _option.CopyClaimsSettingsFrom(ProcessData.GetData());
            _option.CheckPageSize = ckbPageSize.Checked;
```

- [ ] **Step 4: Create `LLDOCX/LLDOCX/Options/ApiKeyStore.cs`**

```csharp
using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace LLDOCX.Options
{
    /// <summary>Stores the Claude API key encrypted with DPAPI for the current Windows user.</summary>
    static class ApiKeyStore
    {
        static readonly byte[] Entropy = Encoding.UTF8.GetBytes("LL-Docx Claude API key v1");

        static string KeyFile
        {
            get
            {
                string folder = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "LL-Docx");
                if (!Directory.Exists(folder)) Directory.CreateDirectory(folder);
                return Path.Combine(folder, "claude.key");
            }
        }

        public static string Load()
        {
            string file = KeyFile;
            if (!File.Exists(file)) return null;
            try
            {
                byte[] data = ProtectedData.Unprotect(File.ReadAllBytes(file), Entropy, DataProtectionScope.CurrentUser);
                string key = Encoding.UTF8.GetString(data);
                return string.IsNullOrWhiteSpace(key) ? null : key;
            }
            catch (CryptographicException)
            {
                return null;
            }
        }

        public static void Save(string key)
        {
            byte[] data = ProtectedData.Protect(Encoding.UTF8.GetBytes(key.Trim()), Entropy, DataProtectionScope.CurrentUser);
            File.WriteAllBytes(KeyFile, data);
        }

        public static void Clear()
        {
            string file = KeyFile;
            if (File.Exists(file)) File.Delete(file);
        }
    }
}
```

- [ ] **Step 5: Verify in Visual Studio (manual)**

Rebuild the solution. Expected: build succeeds, and `bin\Debug` of the add-in contains `LLDOCX.Claims.dll`, `LLDOCX.DeepCheck.dll`, `Anthropic.dll`, and `LLDOCX.dll.config`. Press F5. In Word, click **Settings**, toggle any checkbox, and click Save: the message "Saved Setting successful" appears with no exception. This confirms that an existing settings file in `%APPDATA%\LL-Docx\LLDocx.dll` still loads with the new fields.

- [ ] **Step 6: Commit**

```bash
git add LLDOCX/LLDOCX/LLDOCX.csproj LLDOCX/LLDOCX/Options LLDOCX/LLDOCX/CustomUserControl/SettingUserControl.cs
git commit -m "feat(lldocx): add claims settings and DPAPI API key storage

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 11: Word integration (extractor, controller, pane, dialogs, ribbon)

**Files:**
- Create: `LLDOCX/LLDOCX/Helpers/ClaimsExtractor.cs`
- Create: `LLDOCX/LLDOCX/Helpers/ClaimsController.cs`
- Create: `LLDOCX/LLDOCX/CustomUserControl/ClaimsPaneControl.cs`
- Create: `LLDOCX/LLDOCX/CustomUserControl/ClaimsSettingsPanel.cs`
- Create: `LLDOCX/LLDOCX/CustomUserControl/ApiKeyDialog.cs`
- Create: `LLDOCX/LLDOCX/CustomUserControl/DeepCheckConfirmDialog.cs`
- Modify: `LLDOCX/LLDOCX/ThisAddIn.cs`
- Modify: `LLDOCX/LLDOCX/RibbonExplorer.xml`
- Modify: `LLDOCX/LLDOCX/RibbonExplorer.cs:41-47` (`GetImages`) and add callbacks
- Modify: `LLDOCX/LLDOCX/LLDOCX.csproj` (Compile items)

**Interfaces:**
- Consumes: `ClaimsAnalyzer.Analyze`, `AnalysisResult`, `Finding`, `Severity`, `RawParagraph` (Task 8); `DeepCheckClient.RunAsync`, `DeepCheckException` (Task 9); `SettingOption.ToClaimsCheckOptions`, `ApiKeyStore`, `ProcessData` (Task 10)
- Produces:
  - `ClaimsExtractor.Extract(Word.Document doc, out bool fromSelection) : List<RawParagraph>` (returns null when no claims section is found)
  - `ClaimsController`: `Task RunCheckAsync()`, `Task RunDeepCheckAsync()`, `AnalysisResult LastResult`, `string LastDocumentName`
  - `ThisAddIn.ClaimsController` and `ThisAddIn.ShowClaimsPane() : ClaimsPaneControl`
  - `ClaimsPaneControl`: `ShowResult(AnalysisResult, bool fromSelection)`, `ShowStatus(string)`, `ShowBusy(string, Action cancel)`, `HideBusy()`, `SelectSettingsTab()`

- [ ] **Step 1: Add the Compile items to `LLDOCX/LLDOCX/LLDOCX.csproj`**

In the `ItemGroup` with the other `Compile` items, add:

```xml
    <Compile Include="Helpers\ClaimsExtractor.cs" />
    <Compile Include="Helpers\ClaimsController.cs" />
    <Compile Include="CustomUserControl\ClaimsPaneControl.cs">
      <SubType>UserControl</SubType>
    </Compile>
    <Compile Include="CustomUserControl\ClaimsSettingsPanel.cs">
      <SubType>UserControl</SubType>
    </Compile>
    <Compile Include="CustomUserControl\ApiKeyDialog.cs">
      <SubType>Form</SubType>
    </Compile>
    <Compile Include="CustomUserControl\DeepCheckConfirmDialog.cs">
      <SubType>Form</SubType>
    </Compile>
```

- [ ] **Step 2: Create `LLDOCX/LLDOCX/Helpers/ClaimsExtractor.cs`**

```csharp
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using LLDOCX.Claims;
using Word = Microsoft.Office.Interop.Word;

namespace LLDOCX.Helpers
{
    /// <summary>Reads the claims section (or the selection) of a Word document as RawParagraphs.</summary>
    static class ClaimsExtractor
    {
        static readonly Regex ClaimsHeading = new Regex(
            @"^\s*(?:claims|what is claimed is|what is claimed|i claim|we claim|the claims are)\s*:?\s*$",
            RegexOptions.IgnoreCase);

        static readonly string[] AbstractHeadings =
        {
            "ABSTRACT", "ABSTRACT OF THE INVENTION", "INVENTION ABSTRACT", "DISCLOSURE ABSTRACT",
            "ABSTRACT OF THE DISCLOSURE", "AMENDMENT TO THE ABSTRACT", "SUBSTITUTE ABSTRACT"
        };

        public static List<RawParagraph> Extract(Word.Document doc, out bool fromSelection)
        {
            fromSelection = false;
            Word.Range scope = null;
            Word.Selection selection = doc.Application.Selection;
            if (selection != null && selection.Range.End > selection.Range.Start)
            {
                scope = selection.Range;
                fromSelection = true;
            }
            else
            {
                scope = FindClaimsSection(doc);
            }
            if (scope == null) return null;

            var list = new List<RawParagraph>();
            foreach (Word.Paragraph p in scope.Paragraphs)
            {
                Word.Range r = p.Range;
                string text = r.Text ?? "";
                string prefix = "";
                var listType = r.ListFormat.ListType;
                if (listType != Word.WdListType.wdListNoNumbering && listType != Word.WdListType.wdListBullet &&
                    listType != Word.WdListType.wdListPictureBullet && !string.IsNullOrEmpty(r.ListFormat.ListString))
                {
                    // Auto-numbered paragraphs: the number is not part of Range.Text. Prepend it and
                    // shift DocStart so offsets into the paragraph text still match the document.
                    prefix = r.ListFormat.ListString + " ";
                }
                list.Add(new RawParagraph(prefix + text, r.Start - prefix.Length));
            }
            return list;
        }

        static Word.Range FindClaimsSection(Word.Document doc)
        {
            var paragraphs = new List<(string Text, int Start, int End)>();
            foreach (Word.Paragraph p in doc.Paragraphs)
            {
                Word.Range r = p.Range;
                paragraphs.Add(((r.Text ?? "").Trim('\r', '\n', '\a', ' ', '\t'), r.Start, r.End));
            }

            var headings = Enumerable.Range(0, paragraphs.Count).Where(i => ClaimsHeading.IsMatch(paragraphs[i].Text)).ToList();
            if (headings.Count == 0) return null;
            var abstracts = Enumerable.Range(0, paragraphs.Count)
                .Where(i => AbstractHeadings.Contains(paragraphs[i].Text.ToUpperInvariant()))
                .ToList();

            // Prefer the last claims heading that is followed by an abstract heading; otherwise the last one.
            var beforeAbstract = headings.Where(h => abstracts.Any(a => a > h)).ToList();
            int heading = beforeAbstract.Count > 0 ? beforeAbstract.Last() : headings.Last();
            if (heading + 1 >= paragraphs.Count) return null;

            int endIndex = abstracts.Where(a => a > heading).DefaultIfEmpty(-1).First();
            int start = paragraphs[heading + 1].Start;
            int end = endIndex >= 0 ? paragraphs[endIndex].Start : paragraphs[paragraphs.Count - 1].End;
            return end > start ? doc.Range(start, end) : null;
        }
    }
}
```

- [ ] **Step 3: Create `LLDOCX/LLDOCX/CustomUserControl/ApiKeyDialog.cs`**

```csharp
using System;
using System.Drawing;
using System.Windows.Forms;

namespace LLDOCX.CustomUserControl
{
    class ApiKeyDialog : Form
    {
        readonly TextBox _key = new TextBox { UseSystemPasswordChar = true, Width = 360 };

        public ApiKeyDialog()
        {
            Text = LLDocxConstant.AppName + " – Claude API key";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MaximizeBox = false;
            MinimizeBox = false;
            AutoSize = true;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;
            Padding = new Padding(12);

            var ok = new Button { Text = "Save", DialogResult = DialogResult.OK, Enabled = false };
            var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel };
            _key.TextChanged += (s, e) => ok.Enabled = _key.Text.Trim().Length > 0;

            var layout = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, AutoSize = true, WrapContents = false };
            layout.Controls.Add(new Label { Text = "Paste your Claude API key. It is stored encrypted for your Windows account.", AutoSize = true, MaximumSize = new Size(360, 0) });
            layout.Controls.Add(_key);
            var buttons = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, AutoSize = true, Width = 360 };
            buttons.Controls.Add(cancel);
            buttons.Controls.Add(ok);
            layout.Controls.Add(buttons);
            Controls.Add(layout);
            AcceptButton = ok;
            CancelButton = cancel;
        }

        public string ApiKey => _key.Text.Trim();
    }
}
```

- [ ] **Step 4: Create `LLDOCX/LLDOCX/CustomUserControl/DeepCheckConfirmDialog.cs`**

```csharp
using System.Drawing;
using System.Windows.Forms;

namespace LLDOCX.CustomUserControl
{
    class DeepCheckConfirmDialog : Form
    {
        readonly CheckBox _dontAsk = new CheckBox { Text = "Don't ask again for this document", AutoSize = true };

        public DeepCheckConfirmDialog(int claimCount, int characterCount)
        {
            Text = LLDocxConstant.AppName + " – Deep Check";
            FormBorderStyle = FormBorderStyle.FixedDialog;
            StartPosition = FormStartPosition.CenterParent;
            MaximizeBox = false;
            MinimizeBox = false;
            AutoSize = true;
            AutoSizeMode = AutoSizeMode.GrowAndShrink;
            Padding = new Padding(12);

            string message =
                $"Deep Check will send the text of {claimCount} claims (about {characterCount:N0} characters) " +
                "to Anthropic's Claude API for review.\n\n" +
                "Only the claims are sent, not the rest of the document. The document is not changed.\n\n" +
                "Continue?";
            var send = new Button { Text = "Send", DialogResult = DialogResult.OK };
            var cancel = new Button { Text = "Cancel", DialogResult = DialogResult.Cancel };

            var layout = new FlowLayoutPanel { FlowDirection = FlowDirection.TopDown, AutoSize = true, WrapContents = false };
            layout.Controls.Add(new Label { Text = message, AutoSize = true, MaximumSize = new Size(380, 0) });
            layout.Controls.Add(_dontAsk);
            var buttons = new FlowLayoutPanel { FlowDirection = FlowDirection.RightToLeft, AutoSize = true, Width = 380 };
            buttons.Controls.Add(cancel);
            buttons.Controls.Add(send);
            layout.Controls.Add(buttons);
            Controls.Add(layout);
            AcceptButton = send;
            CancelButton = cancel;
        }

        public bool DontAskAgain => _dontAsk.Checked;
    }
}
```

- [ ] **Step 5: Create `LLDOCX/LLDOCX/CustomUserControl/ClaimsSettingsPanel.cs`**

```csharp
using System;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using LLDOCX.Claims;
using LLDOCX.Options;

namespace LLDOCX.CustomUserControl
{
    class ClaimsSettingsPanel : UserControl
    {
        const int ControlWidth = 380;
        readonly CheckBox _antecedent = new CheckBox { Text = "Check antecedent basis", AutoSize = true };
        readonly CheckBox _dependency = new CheckBox { Text = "Check claim dependency", AutoSize = true };
        readonly CheckBox _terminology = new CheckBox { Text = "Check terminology and typos", AutoSize = true };
        readonly ComboBox _minSeverity = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 200 };
        readonly TextBox _allowlist = new TextBox { Multiline = true, ScrollBars = ScrollBars.Vertical, Width = ControlWidth, Height = 120 };
        readonly TextBox _model = new TextBox { Width = ControlWidth };
        readonly Label _keyStatus = new Label { AutoSize = true };

        public ClaimsSettingsPanel()
        {
            _minSeverity.Items.AddRange(new object[] { "Errors", "Errors + Warnings", "All" });

            var setKey = new Button { Text = "Set key…", AutoSize = true };
            var clearKey = new Button { Text = "Clear key", AutoSize = true };
            var save = new Button { Text = "Save", Width = 110, Height = 34 };
            setKey.Click += OnSetKey;
            clearKey.Click += (s, e) => { ApiKeyStore.Clear(); UpdateKeyStatus(); };
            save.Click += OnSave;

            var layout = new FlowLayoutPanel
            {
                Dock = DockStyle.Fill,
                FlowDirection = FlowDirection.TopDown,
                WrapContents = false,
                AutoScroll = true,
                Padding = new Padding(8)
            };
            layout.Controls.Add(_antecedent);
            layout.Controls.Add(_dependency);
            layout.Controls.Add(_terminology);
            layout.Controls.Add(new Label { Text = "Show by default:", AutoSize = true, Margin = new Padding(3, 12, 3, 3) });
            layout.Controls.Add(_minSeverity);
            layout.Controls.Add(new Label { Text = "Phrases that need no antecedent (one per line):", AutoSize = true, Margin = new Padding(3, 12, 3, 3) });
            layout.Controls.Add(_allowlist);
            layout.Controls.Add(new Label { Text = "Deep Check model:", AutoSize = true, Margin = new Padding(3, 12, 3, 3) });
            layout.Controls.Add(_model);
            layout.Controls.Add(new Label { Text = "Claude API key:", AutoSize = true, Margin = new Padding(3, 12, 3, 3) });
            layout.Controls.Add(_keyStatus);
            var keyButtons = new FlowLayoutPanel { AutoSize = true };
            keyButtons.Controls.Add(setKey);
            keyButtons.Controls.Add(clearKey);
            layout.Controls.Add(keyButtons);
            layout.Controls.Add(save);
            Controls.Add(layout);

            Load += (s, e) => LoadSettings();
        }

        void LoadSettings()
        {
            var o = ProcessData.GetData();
            _antecedent.Checked = o.ClaimsCheckAntecedent;
            _dependency.Checked = o.ClaimsCheckDependency;
            _terminology.Checked = o.ClaimsCheckTerminology;
            _minSeverity.SelectedIndex = (int)o.ClaimsMinSeverity;
            _allowlist.Text = string.Join(Environment.NewLine, o.ClaimsAllowlist ?? ClaimsCheckOptions.DefaultAllowlist.ToList());
            _model.Text = o.DeepCheckModel;
            UpdateKeyStatus();
        }

        void UpdateKeyStatus()
        {
            _keyStatus.Text = ApiKeyStore.Load() != null ? "A key is saved." : "No key saved. Deep Check is unavailable.";
        }

        void OnSetKey(object sender, EventArgs e)
        {
            using (var dlg = new ApiKeyDialog())
            {
                if (dlg.ShowDialog(this) == DialogResult.OK)
                {
                    ApiKeyStore.Save(dlg.ApiKey);
                    UpdateKeyStatus();
                }
            }
        }

        void OnSave(object sender, EventArgs e)
        {
            var o = ProcessData.GetData();
            o.ClaimsCheckAntecedent = _antecedent.Checked;
            o.ClaimsCheckDependency = _dependency.Checked;
            o.ClaimsCheckTerminology = _terminology.Checked;
            o.ClaimsMinSeverity = (Severity)Math.Max(0, _minSeverity.SelectedIndex);
            o.ClaimsAllowlist = _allowlist.Lines.Select(l => l.Trim()).Where(l => l.Length > 0).ToList();
            o.DeepCheckModel = _model.Text.Trim();
            ProcessData.SaveData(o);
            MessageBox.Show("Claims settings saved.", LLDocxConstant.AppName, MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
    }
}
```

- [ ] **Step 6: Create `LLDOCX/LLDOCX/CustomUserControl/ClaimsPaneControl.cs`**

```csharp
using System;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using LLDOCX.Claims;
using LLDOCX.Options;
using Word = Microsoft.Office.Interop.Word;

namespace LLDOCX.CustomUserControl
{
    public class ClaimsPaneControl : UserControl
    {
        readonly TabControl _tabs = new TabControl { Dock = DockStyle.Fill };
        readonly TabPage _findingsTab = new TabPage("Findings");
        readonly TabPage _settingsTab = new TabPage("Settings");
        readonly Label _summary = new Label { Dock = DockStyle.Top, Height = 40, Padding = new Padding(4) };
        readonly ComboBox _severityFilter = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 130 };
        readonly ComboBox _grouping = new ComboBox { DropDownStyle = ComboBoxStyle.DropDownList, Width = 70 };
        readonly ListView _list = new ListView { Dock = DockStyle.Fill, View = View.Details, FullRowSelect = true, HideSelection = false, MultiSelect = false };
        readonly TextBox _details = new TextBox { Dock = DockStyle.Bottom, Height = 90, Multiline = true, ReadOnly = true, ScrollBars = ScrollBars.Vertical };
        readonly Panel _busyPanel = new Panel { Dock = DockStyle.Top, Height = 34, Visible = false };
        readonly Label _busyLabel = new Label { AutoSize = true, Location = new Point(6, 9) };
        readonly ContextMenuStrip _menu = new ContextMenuStrip();
        readonly ToolStripMenuItem _dismiss = new ToolStripMenuItem("Dismiss AI finding");
        Action _cancelAction;
        AnalysisResult _result;
        bool _fromSelection;

        public ClaimsPaneControl()
        {
            _severityFilter.Items.AddRange(new object[] { "Errors", "Errors + Warnings", "All" });
            _grouping.Items.AddRange(new object[] { "Claim", "Type" });
            _grouping.SelectedIndex = 0;

            var rerun = new Button { Text = "Re-run", AutoSize = true };
            rerun.Click += async (s, e) => await Globals.ThisAddIn.ClaimsController.RunCheckAsync();
            var toolbar = new FlowLayoutPanel { Dock = DockStyle.Top, Height = 34, WrapContents = false };
            toolbar.Controls.Add(new Label { Text = "Show:", AutoSize = true, Margin = new Padding(3, 8, 0, 0) });
            toolbar.Controls.Add(_severityFilter);
            toolbar.Controls.Add(new Label { Text = "Group:", AutoSize = true, Margin = new Padding(6, 8, 0, 0) });
            toolbar.Controls.Add(_grouping);
            toolbar.Controls.Add(rerun);

            var cancel = new Button { Text = "Cancel", AutoSize = true, Location = new Point(250, 4) };
            cancel.Click += (s, e) => _cancelAction?.Invoke();
            _busyPanel.Controls.Add(_busyLabel);
            _busyPanel.Controls.Add(cancel);

            _list.Columns.Add("Severity", 65);
            _list.Columns.Add("Claim", 45);
            _list.Columns.Add("Type", 85);
            _list.Columns.Add("Message", 320);
            _list.Columns.Add("Source", 50);
            _list.SelectedIndexChanged += (s, e) => UpdateDetails();
            _list.MouseClick += (s, e) => { var f = SelectedFinding(); if (f != null) NavigateTo(f); };
            _menu.Items.Add(_dismiss);
            _menu.Opening += (s, e) => { var f = SelectedFinding(); _dismiss.Enabled = f != null && f.Source == "AI"; };
            _dismiss.Click += (s, e) => { var f = SelectedFinding(); if (f != null) { _result.Findings.Remove(f); Populate(); } };
            _list.ContextMenuStrip = _menu;

            _severityFilter.SelectedIndexChanged += (s, e) => Populate();
            _grouping.SelectedIndexChanged += (s, e) => Populate();

            // Dock order: the Fill control is added first; later edge-docked controls take space before it.
            _findingsTab.Controls.Add(_list);
            _findingsTab.Controls.Add(_details);
            _findingsTab.Controls.Add(_busyPanel);
            _findingsTab.Controls.Add(toolbar);
            _findingsTab.Controls.Add(_summary);
            _settingsTab.Controls.Add(new ClaimsSettingsPanel { Dock = DockStyle.Fill });
            _tabs.TabPages.Add(_findingsTab);
            _tabs.TabPages.Add(_settingsTab);
            Controls.Add(_tabs);

            Load += (s, e) => _severityFilter.SelectedIndex = (int)ProcessData.GetData().ClaimsMinSeverity;
        }

        public void ShowStatus(string text) => _summary.Text = text;

        public void ShowBusy(string text, Action cancel)
        {
            _busyLabel.Text = text;
            _cancelAction = cancel;
            _busyPanel.Visible = true;
        }

        public void HideBusy()
        {
            _busyPanel.Visible = false;
            _cancelAction = null;
        }

        public void SelectSettingsTab() => _tabs.SelectedTab = _settingsTab;

        public void ShowResult(AnalysisResult result, bool fromSelection)
        {
            _result = result;
            _fromSelection = fromSelection;
            _tabs.SelectedTab = _findingsTab;
            Populate();
        }

        void Populate()
        {
            _list.BeginUpdate();
            _list.Items.Clear();
            _list.Groups.Clear();
            _details.Text = "";
            if (_result == null)
            {
                _list.EndUpdate();
                return;
            }

            var max = (Severity)Math.Max(0, _severityFilter.SelectedIndex);
            int active = _result.Claims.Count(c => !c.IsCanceled);
            int independent = _result.Claims.Count(c => !c.IsCanceled && c.IsIndependent);
            Func<Severity, int> count = s => _result.Findings.Count(f => f.Severity == s && !f.DisputedByAI);
            _summary.Text = $"{active} claims ({independent} independent){(_fromSelection ? " in selection" : "")}: " +
                            $"{count(Severity.Error)} errors, {count(Severity.Warning)} warnings, {count(Severity.Info)} info";

            bool byClaim = _grouping.SelectedIndex != 1;
            var visible = _result.Findings.Where(f => f.Severity <= max)
                .OrderBy(f => byClaim ? f.ClaimNumber : (int)f.Type)
                .ThenBy(f => f.Severity)
                .ThenBy(f => f.DocStart);
            foreach (var f in visible)
            {
                string groupKey = byClaim ? "Claim " + f.ClaimNumber : f.Type.ToString();
                var group = _list.Groups[groupKey] ?? _list.Groups.Add(groupKey, groupKey);
                string message = f.Message + (string.IsNullOrEmpty(f.Suggestion) ? "" : "  → " + f.Suggestion);
                var item = new ListViewItem(new[] { f.Severity.ToString(), f.ClaimNumber.ToString(), f.Type.ToString(), message, f.Source })
                {
                    Group = group,
                    Tag = f,
                    ForeColor = f.Severity == Severity.Error ? Color.DarkRed
                              : f.Severity == Severity.Warning ? Color.DarkOrange
                              : SystemColors.GrayText
                };
                if (f.DisputedByAI) item.Font = new Font(_list.Font, FontStyle.Strikeout);
                _list.Items.Add(item);
            }
            _list.EndUpdate();
        }

        Finding SelectedFinding() => _list.SelectedItems.Count == 0 ? null : _list.SelectedItems[0].Tag as Finding;

        void UpdateDetails()
        {
            var f = SelectedFinding();
            if (f == null)
            {
                _details.Text = "";
                return;
            }
            string text = $"Claim {f.ClaimNumber} – {f.Type} – {f.Severity} ({f.Source})" + Environment.NewLine + f.Message;
            if (!string.IsNullOrEmpty(f.Suggestion)) text += Environment.NewLine + "Suggestion: " + f.Suggestion;
            if (f.DisputedByAI) text += Environment.NewLine + "Deep Check disputes this: " + f.DisputeReason;
            if (f.DocStart < 0) text += Environment.NewLine + "(This text could not be located in the document.)";
            _details.Text = text;
        }

        void NavigateTo(Finding f)
        {
            if (f.DocStart < 0) return;
            Word.Document doc;
            try { doc = Globals.ThisAddIn.Application.ActiveDocument; }
            catch (COMException) { return; }
            if (!string.Equals(doc.FullName, Globals.ThisAddIn.ClaimsController.LastDocumentName, StringComparison.OrdinalIgnoreCase))
            {
                MessageBox.Show("These findings belong to another document. Switch to that document or click Re-run.",
                    LLDocxConstant.AppName, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }
            Word.Range range = doc.Range(f.DocStart, f.DocStart + f.Length);
            if (Normalize(range.Text) != Normalize(f.Text))
            {
                var item = _list.SelectedItems[0];
                if (!item.SubItems[3].Text.StartsWith("(stale"))
                    item.SubItems[3].Text = "(stale – re-run) " + item.SubItems[3].Text;
                item.ForeColor = SystemColors.GrayText;
                return;
            }
            range.Select();
        }

        static string Normalize(string s) => Regex.Replace(s ?? "", @"\s+", " ").Trim();
    }
}
```

- [ ] **Step 7: Create `LLDOCX/LLDOCX/Helpers/ClaimsController.cs`**

```csharp
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using LLDOCX.Claims;
using LLDOCX.CustomUserControl;
using LLDOCX.DeepCheck;
using LLDOCX.Options;
using Word = Microsoft.Office.Interop.Word;

namespace LLDOCX.Helpers
{
    public class ClaimsController
    {
        readonly HashSet<string> _deepCheckApproved = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        CancellationTokenSource _deepCheckCts;

        public AnalysisResult LastResult { get; private set; }
        public string LastDocumentName { get; private set; }

        public async Task RunCheckAsync()
        {
            Word.Document doc = ActiveDocument();
            if (doc == null) return;
            var pane = Globals.ThisAddIn.ShowClaimsPane();

            bool fromSelection;
            List<RawParagraph> paragraphs = ClaimsExtractor.Extract(doc, out fromSelection);
            if (paragraphs == null)
            {
                MessageBox.Show("No claims section was found (for example, a \"CLAIMS\" or \"What is claimed is:\" heading).\n\n" +
                                "Select the claims and click Check Claims again to check the selected text.",
                    LLDocxConstant.AppName, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return;
            }

            var options = ProcessData.GetData().ToClaimsCheckOptions();
            pane.ShowStatus("Checking claims…");
            try
            {
                var result = await Task.Run(() => ClaimsAnalyzer.Analyze(paragraphs, options));
                LastResult = result;
                LastDocumentName = doc.FullName;
                pane.ShowResult(result, fromSelection);
            }
            catch (Exception ex)
            {
                Trace.TraceError("LL-DOCX claims check failed: " + ex);
                pane.ShowStatus("");
                MessageBox.Show("Claims check failed: " + ex.Message, LLDocxConstant.AppName, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        public async Task RunDeepCheckAsync()
        {
            if (_deepCheckCts != null) return; // already running

            Word.Document doc = ActiveDocument();
            if (doc == null) return;
            if (LastResult == null || !string.Equals(LastDocumentName, doc.FullName, StringComparison.OrdinalIgnoreCase))
            {
                await RunCheckAsync();
                if (LastResult == null || !string.Equals(LastDocumentName, doc.FullName, StringComparison.OrdinalIgnoreCase)) return;
            }

            var pane = Globals.ThisAddIn.ShowClaimsPane();
            string apiKey = ApiKeyStore.Load();
            if (apiKey == null)
            {
                MessageBox.Show("No Claude API key is set. Add one on the Settings tab of the Claims Check pane.",
                    LLDocxConstant.AppName, MessageBoxButtons.OK, MessageBoxIcon.Information);
                pane.SelectSettingsTab();
                return;
            }

            var result = LastResult;
            if (!_deepCheckApproved.Contains(LastDocumentName))
            {
                int chars = result.Claims.Sum(c => c.FullText.Length);
                using (var dlg = new DeepCheckConfirmDialog(result.Claims.Count, chars))
                {
                    if (dlg.ShowDialog() != DialogResult.OK) return;
                    if (dlg.DontAskAgain) _deepCheckApproved.Add(LastDocumentName);
                }
            }

            // Replace any earlier AI results and disputes.
            result.Findings.RemoveAll(f => f.Source == "AI");
            foreach (var f in result.Findings) { f.DisputedByAI = false; f.DisputeReason = null; }
            var ruleFindings = result.Findings.ToList();
            string model = ProcessData.GetData().DeepCheckModel;

            _deepCheckCts = new CancellationTokenSource();
            var token = _deepCheckCts.Token;
            pane.ShowBusy("Deep Check running…", () => _deepCheckCts?.Cancel());
            try
            {
                var added = await Task.Run(() => DeepCheckClient.RunAsync(apiKey, model, result.Claims, ruleFindings, token));
                result.Findings.AddRange(added);
                pane.ShowResult(result, false);
            }
            catch (OperationCanceledException)
            {
                pane.ShowResult(result, false);
                pane.ShowStatus("Deep Check canceled.");
            }
            catch (DeepCheckException ex)
            {
                pane.ShowResult(result, false);
                MessageBox.Show(ex.Message, LLDocxConstant.AppName, MessageBoxButtons.OK, MessageBoxIcon.Warning);
            }
            catch (Exception ex)
            {
                Trace.TraceError("LL-DOCX Deep Check failed: " + ex);
                pane.ShowResult(result, false);
                MessageBox.Show("Deep Check failed: " + ex.Message, LLDocxConstant.AppName, MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                pane.HideBusy();
                _deepCheckCts.Dispose();
                _deepCheckCts = null;
            }
        }

        static Word.Document ActiveDocument()
        {
            try
            {
                return Globals.ThisAddIn.Application.ActiveDocument;
            }
            catch (COMException)
            {
                MessageBox.Show("Open a document first.", LLDocxConstant.AppName, MessageBoxButtons.OK, MessageBoxIcon.Information);
                return null;
            }
        }
    }
}
```

- [ ] **Step 8: Wire the pane and controller into `LLDOCX/LLDOCX/ThisAddIn.cs`**

Add `using System.Threading;` and `using System.Windows.Forms;` to the usings. Replace the class body above `#region VSTO generated code` with:

```csharp
        private SettingUserControl _settingControl;
        private Microsoft.Office.Tools.CustomTaskPane myCustomTaskPane;
        private ClaimsPaneControl _claimsControl;
        private Microsoft.Office.Tools.CustomTaskPane _claimsTaskPane;
        public MissingHeadingHandler MissingHeadingHandler { get; set; }
        public List<DocumentHeadingData> DocumentHeadingData { get; set; }
        public ClaimsController ClaimsController { get; private set; }
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            // Lets async ribbon callbacks resume on the Office UI thread after awaiting background work.
            if (!(SynchronizationContext.Current is WindowsFormsSynchronizationContext))
                SynchronizationContext.SetSynchronizationContext(new WindowsFormsSynchronizationContext());
            MissingHeadingHandler = new MissingHeadingHandler();
            DocumentHeadingData = new List<DocumentHeadingData>();
            ClaimsController = new ClaimsController();
        }
        public void ProcessSideBarPanel(bool isVisible)
        {
            if (_settingControl == null)
            {
                _settingControl = new SettingUserControl();
            }
            if (myCustomTaskPane == null)
            {
                myCustomTaskPane = this.CustomTaskPanes.Add(_settingControl, "LL-DOCX Setting", this.Application.ActiveWindow);
                myCustomTaskPane.Width = 450;
            }
            myCustomTaskPane.Visible = isVisible;
        }
        public ClaimsPaneControl ShowClaimsPane()
        {
            if (_claimsControl == null)
            {
                _claimsControl = new ClaimsPaneControl();
            }
            if (_claimsTaskPane == null)
            {
                _claimsTaskPane = this.CustomTaskPanes.Add(_claimsControl, "LL-DOCX Claims Check", this.Application.ActiveWindow);
                _claimsTaskPane.Width = 450;
            }
            _claimsTaskPane.Visible = true;
            return _claimsControl;
        }
        protected override Office.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new RibbonExplorer();
        }
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }
```

The `if (CustomTaskPanes.Count == 0)` check becomes `if (myCustomTaskPane == null)`. Without that change, the Settings pane would never be created once the Claims pane exists.

- [ ] **Step 9: Add the ribbon buttons to `LLDOCX/LLDOCX/RibbonExplorer.xml`**

After the `btnPanel` toggleButton line, add:

```xml
          <button id="btnCheckClaims" label="Check Claims" size="large" onAction="OnCheckClaims" getImage="GetImages"
                  screentip="Check patent claims"
                  supertip="Checks the claims for antecedent basis, claim dependency, and terminology problems. The document is not changed."/>
          <button id="btnDeepCheck" label="Deep Check" size="large" onAction="OnDeepCheck" getImage="GetImages"
                  screentip="Deep Check with Claude"
                  supertip="Sends the claims to Claude for a second review. Asks before sending anything."/>
```

- [ ] **Step 10: Add the ribbon callbacks in `LLDOCX/LLDOCX/RibbonExplorer.cs`**

Replace `GetImages` with:

```csharp
        public Bitmap GetImages(Office.IRibbonControl control)
        {
            if (control.Id == "btnScan" || control.Id == "btnCheckClaims" || control.Id == "btnDeepCheck")
                return Properties.Resources.LL_logo;
            if (isPressed)
                return Properties.Resources.toggle_on;
            return Properties.Resources.toggle_off;
        }
```

Add after `ShowSetting`:

```csharp
        public async void OnCheckClaims(Office.IRibbonControl control)
        {
            try { await Globals.ThisAddIn.ClaimsController.RunCheckAsync(); }
            catch (Exception ex) { System.Windows.Forms.MessageBox.Show("Claims check failed: " + ex.Message, LLDocxConstant.AppName); }
        }
        public async void OnDeepCheck(Office.IRibbonControl control)
        {
            try { await Globals.ThisAddIn.ClaimsController.RunDeepCheckAsync(); }
            catch (Exception ex) { System.Windows.Forms.MessageBox.Show("Deep Check failed: " + ex.Message, LLDocxConstant.AppName); }
        }
```

- [ ] **Step 11: Build and verify in Word (manual)**

Rebuild in Visual Studio. Expected: build succeeds. Press F5, then check each item:

1. The Home tab shows **Check Claims** and **Deep Check** in the LL-Docx group.
2. Create a document with a heading `What is claimed is:`, then typed claims:
   `1. A lamp comprising a base, wherein the shade is green.` / `2. The lamp of claim 3, wherein the base is round.` / `3. The lamp of claim 1, wherein the bulb is bright.`
   Then a heading `ABSTRACT` and one paragraph. Click **Check Claims**. Expected: the "LL-DOCX Claims Check" pane opens. Claim 1 shows "No antecedent basis for 'the shade'". Claim 2 shows "comes after it". Claim 3 shows "the bulb". The abstract is not checked.
3. Click a finding: the flagged words are selected in the document.
4. Edit the flagged words, then click the same finding: the row shows "(stale – re-run)" and nothing is selected. Click **Re-run**: the findings refresh.
5. Convert the claims to a Word numbered list (Home → Numbering) with the typed numbers removed. Click **Check Claims**. Expected: the same findings. Clicking still selects the correct words.
6. Delete the "What is claimed is:" heading and click **Check Claims** with no selection. Expected: the "No claims section was found" message. Select the claims and click again: the findings appear, and the summary says "in selection".
7. Change "Show:" to "Errors" and "Group:" to "Type": the list filters and regroups.
8. On the Settings tab, untick "Check antecedent basis", click Save, and click Re-run: the antecedent findings are gone. Re-tick the box and save.
9. Click **Settings** in the ribbon (the existing pane) and click its Save. Then reopen the Claims Settings tab: your claims settings are unchanged.
10. Click **Deep Check** with no key: the message points to Settings, and the Settings tab is selected. Click **Set key…**, enter an invalid key `sk-ant-invalid`, and click **Deep Check** → **Send**. Expected: "The Claude API key was rejected…", and the rule findings stay in the pane.
11. Set a valid key and click **Deep Check**. The confirmation dialog shows the claim count. Click Send: "Deep Check running…" appears with Cancel. When it finishes, any AI findings show Source "AI", and any disputed rule findings are struck through. Right-click an AI finding → **Dismiss AI finding**: it disappears.
12. Disconnect the network and click **Deep Check**: "Could not reach the Claude API…" appears, and the findings remain.
13. Click **Scan**: it behaves exactly as before.

If step 11 fails with `FileLoadException` or `FileNotFoundException` for an assembly such as `System.Text.Json` or `System.Memory`: confirm that `bin\Debug\LLDOCX.dll.config` exists and contains `bindingRedirect` entries, and that `AutoGenerateBindingRedirects` and `GenerateBindingRedirectsOutputType` from Task 1 are in the csproj. Rebuild.

- [ ] **Step 12: Commit**

```bash
git add LLDOCX/LLDOCX
git commit -m "feat(lldocx): add Check Claims and Deep Check to the Word ribbon

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

### Task 12: Installer

**Files:**
- Modify: `LLDOCX/LLDOCXSetup/Product.wxs:61-90` (ComponentGroup `ProductComponents`)

**Interfaces:**
- Consumes: Release build output of the add-in (Tasks 1–11)
- Produces: an MSI that installs every assembly the add-in needs, plus `LLDOCX.dll.config`

- [ ] **Step 1: List the assemblies the add-in now ships**

In Visual Studio, switch to **Release** and rebuild. Then run:

```bash
ls LLDOCX/bin/Release/*.dll LLDOCX/bin/Release/*.config
```

Expected: `LLDOCX.dll`, `LLDOCX.dll.config`, `LLDOCX.Claims.dll`, `LLDOCX.DeepCheck.dll`, `Anthropic.dll`, and their dependencies (for example `System.Text.Json.dll`, `System.Memory.dll`, `System.Buffers.dll`, `Microsoft.Bcl.AsyncInterfaces.dll`, `System.Runtime.CompilerServices.Unsafe.dll`, `System.Threading.Tasks.Extensions.dll`, `System.Text.Encodings.Web.dll`, `Microsoft.Extensions.AI.Abstractions.dll`, `System.Collections.Immutable.dll`, `System.Diagnostics.DiagnosticSource.dll`, `System.Net.ServerSentEvents.dll`, `System.Numerics.Vectors.dll`, `System.IO.Pipelines.dll`). The exact list comes from this command's output.

- [ ] **Step 2: Add a component for each new file**

In `Product.wxs`, uncomment the existing `LLDOCX_dll.config_Component` block (lines 76–79). Then, inside `<ComponentGroup Id="ProductComponents" ...>`, add one block per DLL from Step 1 that isn't already listed. Use the file name with `.` and `-` replaced by `_` in the Ids, for example:

```xml
			<Component Id="LLDOCX_Claims_dll_Component">
				<File Id="LLDOCX_Claims_dll" KeyPath="yes"
							  Name="LLDOCX.Claims.dll" Source="$(var.AddinFiles)" />
			</Component>
			<Component Id="LLDOCX_DeepCheck_dll_Component">
				<File Id="LLDOCX_DeepCheck_dll" KeyPath="yes"
							  Name="LLDOCX.DeepCheck.dll" Source="$(var.AddinFiles)" />
			</Component>
			<Component Id="Anthropic_dll_Component">
				<File Id="Anthropic_dll" KeyPath="yes"
							  Name="Anthropic.dll" Source="$(var.AddinFiles)" />
			</Component>
```

Repeat the same three-line pattern for every remaining DLL in the Step 1 list.

- [ ] **Step 3: Build and test the installer (manual)**

Build the `LLDOCXSetup` project in Release. Install the MSI on a test machine, or after uninstalling the debug add-in. Open Word and run steps 2 and 11 from the Task 11 checklist.
Expected: Check Claims and Deep Check both work from the installed add-in.

- [ ] **Step 4: Commit**

```bash
git add LLDOCX/LLDOCXSetup/Product.wxs
git commit -m "build(lldocx): ship claims checker assemblies in the installer

Co-Authored-By: Claude Opus 5.5 <noreply@anthropic.com>"
```

---

## Spec updates made with this plan

These changes were made to the spec while writing this plan, and the spec file has been updated to match:

- The engine's input type is `RawParagraph`, one per document paragraph. `ClaimsAnalyzer.Analyze` returns an `AnalysisResult` that holds both the parsed claims and the findings, so the pane can show claim counts.
- The Claude client lives in a separate `LLDOCX.DeepCheck` project, which keeps the engine free of the SDK. The tests run on net8.0, and the add-in moves to .NET Framework 4.7.2. This is now decided, not conditional.
- Introducing determiners: `each`, `any`, `one`, and bare plurals were dropped to cut false positives. Phrases like "The method of claim 1" are skipped automatically, so the allowlist doesn't need "the method of claim" entries, and "the group consisting of" becomes "the group".
- Typo rule: edit distance 1 for words of 5–7 letters and 2 for 8+ letters; inflections are excluded instead of being flagged.
- Singular/plural drift is detected during the antecedent pass and controlled by the terminology toggle.
- The Claims settings live on a Settings tab inside the Claims Check pane rather than in the fixed-layout existing Settings pane.
