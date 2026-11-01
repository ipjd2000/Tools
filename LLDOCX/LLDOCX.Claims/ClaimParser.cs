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
