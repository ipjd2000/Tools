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

        /// <summary>Adjectives (positional or participial) that look like -ing/-ed words and may precede a plural noun.</summary>
        static readonly HashSet<string> ModifierWords = new HashSet<string>
        {
            "upper", "lower", "left", "right", "outer", "inner", "front", "rear", "top", "bottom", "corresponding",
            "underlying", "overlying", "respective", "remaining", "existing", "resulting"
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
                    bool afterModifier = phrase.Words.Count > 0 && ModifierWords.Contains(phrase.Words[phrase.Words.Count - 1]);
                    if (nonOrdinal > 0 && !NounExceptions.Contains(w) && !ModifierWords.Contains(w) &&
                        (w.EndsWith("ing") || w.EndsWith("ed") || w.EndsWith("ly") || (TextUtil.IsPlural(w) && !afterModifier)))
                        break;
                    if (phrase.Words.Count == 0) phrase.WordsStart = tok.Index;
                    phrase.Words.Add(w);
                    bool participialLead = phrase.Words.Count == 1 && nonOrdinal == 0 && IsParticiple(w) &&
                                           j + 1 < tokens.Count && tokens[j + 1].IsWord && !Boundary.Contains(tokens[j + 1].Lower);
                    if (!Ordinals.Contains(w) && !participialLead) nonOrdinal++;
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

        static bool IsParticiple(string w) =>
            (w.EndsWith("ed") || w.EndsWith("ing")) && !NounExceptions.Contains(w);

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
