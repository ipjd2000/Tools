using System.Collections.Generic;
using System.Linq;
using LLDOCX.Claims;

namespace LLDOCX.DeepCheck
{
    public static class DeepCheckRequestBuilder
    {
        public static DeepCheckRequest Build(IList<Claim> claims, IList<Finding> ruleFindings)
        {
            var request = new DeepCheckRequest();
            foreach (var c in claims)
                request.Claims.Add(new ClaimDto { Number = c.Number, Text = c.IsCanceled ? "" : c.FullText, Canceled = c.IsCanceled });
            foreach (var f in ruleFindings)
                request.RuleFindings.Add(new RuleFindingDto
                {
                    Claim = f.ClaimNumber,
                    Type = f.Type.ToString(),
                    Severity = f.Severity.ToString(),
                    Message = f.Message
                });
            return request;
        }

        public static int TextLength(DeepCheckRequest request) =>
            request.Claims == null ? 0 : request.Claims.Sum(c => c?.Text?.Length ?? 0);
    }
}
