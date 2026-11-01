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
