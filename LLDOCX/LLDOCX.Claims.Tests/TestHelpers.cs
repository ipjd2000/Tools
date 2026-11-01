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
