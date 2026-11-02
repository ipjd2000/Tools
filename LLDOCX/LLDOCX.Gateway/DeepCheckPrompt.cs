using System.Text;
using LLDOCX.DeepCheck;

namespace LLDOCX.Gateway;

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

    public const string ResponseSchemaJson = """
        {
          "type": "object",
          "properties": {
            "findings": {
              "type": "array",
              "items": {
                "type": "object",
                "properties": {
                  "claim": { "type": "integer" },
                  "type": { "type": "string", "enum": ["Antecedent", "Dependency", "Terminology", "Formatting"] },
                  "severity": { "type": "string", "enum": ["Error", "Warning", "Info"] },
                  "quote": { "type": "string" },
                  "message": { "type": "string" },
                  "suggestion": { "type": "string" }
                },
                "required": ["claim", "type", "severity", "quote", "message", "suggestion"],
                "additionalProperties": false
              }
            },
            "disputes": {
              "type": "array",
              "items": {
                "type": "object",
                "properties": {
                  "index": { "type": "integer" },
                  "reason": { "type": "string" }
                },
                "required": ["index", "reason"],
                "additionalProperties": false
              }
            }
          },
          "required": ["findings", "disputes"],
          "additionalProperties": false
        }
        """;

    public static string BuildUserMessage(DeepCheckRequest request)
    {
        var sb = new StringBuilder();
        sb.AppendLine("<claims>");
        foreach (var c in request.Claims)
            sb.AppendLine(c.Canceled ? $"{c.Number}. (Canceled)" : $"{c.Number}. {c.Text}");
        sb.AppendLine("</claims>");
        sb.AppendLine("<rule_findings>");
        for (int i = 0; i < request.RuleFindings.Count; i++)
        {
            var f = request.RuleFindings[i];
            sb.AppendLine($"[{i}] claim {f.Claim}, {f.Type}, {f.Severity}: {f.Message}");
        }
        sb.AppendLine("</rule_findings>");
        return sb.ToString();
    }
}
