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
