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
