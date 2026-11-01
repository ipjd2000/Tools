using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

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
                        CheckReference(claim, p, known, allow, reportPlurality, context, findings);
                    }
                }
            }
            return findings;
        }

        static void CheckReference(Claim claim, NounPhrase p, List<NounPhrase> known, List<string> allow,
            bool reportPlurality, Context context, List<Finding> findings)
        {
            if (IsAllowed(p, allow)) return;
            // The subject of a dependent claim ("The sodium salt of the compound of claim 1") is the claim's own term.
            if (p.Start == 0 && claim.DependsOn.Count > 0 && AnchoredToClaimReference(claim.FullText, p)) return;

            var exact = known.FirstOrDefault(k => k.Key == p.Key);
            if (exact != null)
            {
                if (reportPlurality) CheckPlurality(claim, p, exact, findings);
                return;
            }

            // The extractor stops before a plural-looking word ("said conduit | programs"); accept it if the longer term is known.
            if (ExtendsToKnown(claim.FullText, p, known)) return;

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

            var weak = known.FirstOrDefault(k => k.HeadKey == p.HeadKey)
                       ?? known.FirstOrDefault(k => k.ModifierKeys.Contains(p.HeadKey) && IsCoordinatedModifier(claim.FullText, p, k));
            if (weak != null)
            {
                findings.Add(Finding.At(claim, p.Start, p.Length, FindingType.Antecedent, Severity.Info,
                    $"'{p.Text}' only partly matches '{weak.Text}' (same noun, different modifiers). Check its antecedent basis."));
                return;
            }

            if (context.AppearsBareBefore(claim, p))
            {
                findings.Add(Finding.At(claim, p.Start, p.Length, FindingType.Antecedent, Severity.Info,
                    $"'{p.Text}' refers to '{p.Head}', which appears earlier without an article. Check its antecedent basis."));
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

        static readonly Regex ClaimAnchor = new Regex(
            @"\G\s*(?:of|according\s+to|as\s+(?:claimed|recited|defined|set\s+forth|described)\s+in|in)\s+(?:(?:the|a|an)\s+)?(?:[\w-]+\s+){0,5}?claims?\b",
            RegexOptions.Compiled | RegexOptions.IgnoreCase);

        /// <summary>True for "The X of claim N" / "The X of the Y of claim N": the opening subject is tied to a claim reference.</summary>
        static bool AnchoredToClaimReference(string text, NounPhrase p) =>
            ClaimAnchor.Match(text, p.Start + p.Length).Success;

        /// <summary>True for elliptical coordination: "said front and rear sections" where "front" modifies the known "front section".</summary>
        static bool IsCoordinatedModifier(string text, NounPhrase p, NounPhrase known)
        {
            var tokens = TextUtil.Tokenize(text);
            int end = p.Start + p.Length;
            int i = 0;
            while (i < tokens.Count && tokens[i].Index < end) i++;
            int guard = 0;
            while (i < tokens.Count && guard++ < 4 && !(tokens[i].IsWord && (tokens[i].Lower == "and" || tokens[i].Lower == "or"))) i++;
            if (i >= tokens.Count || !tokens[i].IsWord || (tokens[i].Lower != "and" && tokens[i].Lower != "or")) return false;
            for (int k = i + 1, n = 0; k < tokens.Count && n < 3 && tokens[k].IsWord; k++, n++)
                if (TextUtil.Singularize(tokens[k].Lower) == known.HeadKey) return true;
            return false;
        }

        static bool ExtendsToKnown(string text, NounPhrase p, List<NounPhrase> known)
        {
            var tokens = TextUtil.Tokenize(text);
            int end = p.Start + p.Length;
            int i = 0;
            while (i < tokens.Count && tokens[i].Index < end) i++;
            if (i >= tokens.Count || !tokens[i].IsWord) return false;
            string extended = p.Key + " " + TextUtil.Singularize(tokens[i].Lower);
            return known.Any(k => k.Key == extended);
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
            if (p.Words.Count == 1 && (p.Head == "latter" || p.Head == "former")) return true;
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

            /// <summary>True if the phrase's head noun occurs earlier in the claim or its ancestors without a determiner (a bare plural or mass-noun introduction).</summary>
            public bool AppearsBareBefore(Claim claim, NounPhrase p)
            {
                if (HasBare(claim.FullText, p.HeadKey, p.Start)) return true;
                var seen = new HashSet<int>();
                var stack = new Stack<int>(claim.DependsOn);
                while (stack.Count > 0)
                {
                    int n = stack.Pop();
                    Claim parent;
                    if (!seen.Add(n) || !_byNumber.TryGetValue(n, out parent)) continue;
                    if (HasBare(parent.FullText, p.HeadKey, parent.FullText.Length)) return true;
                    foreach (var d in parent.DependsOn) stack.Push(d);
                }
                return false;
            }

            static readonly HashSet<string> Articles = new HashSet<string>
            {
                "the", "said", "a", "an", "another", "each", "every", "any", "one", "more", "plurality"
            };

            static bool HasBare(string text, string headKey, int limit)
            {
                var tokens = TextUtil.Tokenize(text);
                for (int i = 0; i < tokens.Count && tokens[i].Index < limit; i++)
                {
                    if (!tokens[i].IsWord || TextUtil.Singularize(tokens[i].Lower) != headKey) continue;
                    string prev = i > 0 && tokens[i - 1].IsWord ? tokens[i - 1].Lower : "";
                    if (!Articles.Contains(prev)) return true;
                }
                return false;
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
