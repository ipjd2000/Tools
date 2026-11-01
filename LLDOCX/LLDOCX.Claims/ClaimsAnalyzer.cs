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
