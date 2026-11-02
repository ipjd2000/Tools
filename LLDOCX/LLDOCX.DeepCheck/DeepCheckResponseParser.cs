using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json;
using LLDOCX.Claims;

namespace LLDOCX.DeepCheck
{
    public static class DeepCheckResponseParser
    {
        const string Unreadable = "Deep Check returned a response that could not be read. The rule-based findings are unchanged.";

        public static List<Finding> Parse(string json, IList<Claim> claims, IList<Finding> ruleFindings)
        {
            JsonDocument doc;
            try
            {
                doc = JsonDocument.Parse(json);
            }
            catch (JsonException ex)
            {
                throw new DeepCheckException(Unreadable, ex);
            }

            try
            {
            using (doc)
            {
                var root = doc.RootElement;
                if (root.ValueKind != JsonValueKind.Object) throw new DeepCheckException(Unreadable);

                var added = new List<Finding>();
                if (root.TryGetProperty("findings", out var findings) && findings.ValueKind == JsonValueKind.Array)
                    foreach (var f in findings.EnumerateArray())
                    {
                        if (f.ValueKind != JsonValueKind.Object) continue;
                        added.Add(ToFinding(f, claims));
                    }

                if (root.TryGetProperty("disputes", out var disputes) && disputes.ValueKind == JsonValueKind.Array)
                {
                    foreach (var d in disputes.EnumerateArray())
                    {
                        if (d.ValueKind != JsonValueKind.Object) continue;
                        int index = -1;
                        if (d.TryGetProperty("index", out var i) && i.ValueKind == JsonValueKind.Number && !i.TryGetInt32(out index)) throw new DeepCheckException(Unreadable);
                        if (index < 0 || index >= ruleFindings.Count) continue;
                        ruleFindings[index].DisputedByAI = true;
                        ruleFindings[index].DisputeReason = Str(d, "reason");
                    }
                }
                return added;
            }
            }
            catch (InvalidOperationException ex)
            {
                throw new DeepCheckException(Unreadable, ex);
            }
            catch (FormatException ex)
            {
                throw new DeepCheckException(Unreadable, ex);
            }
        }

        /// <summary>True when json has the response schema's shape. The gateway checks this before returning Claude's output.</summary>
        public static bool IsWellFormed(string json)
        {
            try
            {
                using (var doc = JsonDocument.Parse(json))
                {
                    var root = doc.RootElement;
                    if (root.ValueKind != JsonValueKind.Object) return false;
                    if (!root.TryGetProperty("findings", out var findings) || findings.ValueKind != JsonValueKind.Array) return false;
                    if (!root.TryGetProperty("disputes", out var disputes) || disputes.ValueKind != JsonValueKind.Array) return false;
                    foreach (var f in findings.EnumerateArray())
                    {
                        if (f.ValueKind != JsonValueKind.Object || !IsInt(f, "claim")) return false;
                        foreach (var name in new[] { "type", "severity", "quote", "message", "suggestion" })
                            if (!Is(f, name, JsonValueKind.String)) return false;
                    }
                    foreach (var d in disputes.EnumerateArray())
                        if (d.ValueKind != JsonValueKind.Object || !IsInt(d, "index") || !Is(d, "reason", JsonValueKind.String))
                            return false;
                    return true;
                }
            }
            catch (JsonException)
            {
                return false;
            }
        }

        static bool Is(JsonElement e, string name, JsonValueKind kind) => e.TryGetProperty(name, out var v) && v.ValueKind == kind;

        static bool IsInt(JsonElement e, string name) => e.TryGetProperty(name, out var v) && v.ValueKind == JsonValueKind.Number && v.TryGetInt32(out _);

        static Finding ToFinding(JsonElement f, IList<Claim> claims)
        {
            int number = 0;
            if (f.TryGetProperty("claim", out var n) && n.ValueKind == JsonValueKind.Number)
            {
                if (!n.TryGetInt32(out number)) throw new DeepCheckException(Unreadable);
            }
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
