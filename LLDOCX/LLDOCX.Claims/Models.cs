using System;
using System.Collections.Generic;

namespace LLDOCX.Claims
{
    public enum ClaimCategory { Unknown, Method, Apparatus, Composition, ComputerReadableMedium }

    public enum FindingType { Antecedent, Dependency, Terminology, Formatting }

    /// <summary>Ordered so that a filter can keep findings with Severity &lt;= minimum.</summary>
    public enum Severity { Error = 0, Warning = 1, Info = 2 }

    /// <summary>One paragraph of document text. DocStart is the document offset of Text[0].</summary>
    public class RawParagraph
    {
        public RawParagraph(string text, int docStart)
        {
            Text = text ?? "";
            DocStart = docStart;
        }

        public string Text { get; }
        public int DocStart { get; }
    }

    /// <summary>Maps a run of Claim.FullText (starting at TextOffset) to the document (starting at DocStart).</summary>
    public class TextSegment
    {
        public TextSegment(int textOffset, int docStart, int length)
        {
            TextOffset = textOffset;
            DocStart = docStart;
            Length = length;
        }

        public int TextOffset { get; }
        public int DocStart { get; }
        public int Length { get; }
    }

    public class Claim
    {
        public int Number { get; set; }
        public string Status { get; set; }
        public bool IsCanceled => Status == "Canceled";
        public string FullText { get; set; } = "";
        public string Preamble { get; set; } = "";
        public List<int> DependsOn { get; } = new List<int>();
        public bool IsIndependent => DependsOn.Count == 0;
        public bool IsMultipleDependent { get; set; }
        public bool HasCumulativeDependency { get; set; }
        internal bool UsesPrecedingForm { get; set; }
        internal HashSet<int> SingleRefNumbers { get; } = new HashSet<int>();
        public int DependencyRefStart { get; set; } = -1;
        public int DependencyRefLength { get; set; }
        public ClaimCategory Category { get; set; }
        public string PreambleNoun { get; set; } = "";
        /// <summary>Document offset of the claim's content (after its number and status).</summary>
        public int DocStart { get; set; }
        public List<TextSegment> Segments { get; } = new List<TextSegment>();

        public int ToDocOffset(int textIndex)
        {
            if (Segments.Count == 0) return DocStart;
            foreach (var s in Segments)
            {
                if (textIndex < s.TextOffset) return s.DocStart;
                if (textIndex <= s.TextOffset + s.Length) return s.DocStart + (textIndex - s.TextOffset);
            }
            var last = Segments[Segments.Count - 1];
            return last.DocStart + last.Length;
        }
    }

    public class Finding
    {
        public int ClaimNumber { get; set; }
        public FindingType Type { get; set; }
        public Severity Severity { get; set; }
        public string Message { get; set; } = "";
        public string Suggestion { get; set; }
        /// <summary>Document offset of the flagged text, or -1 when it could not be located.</summary>
        public int DocStart { get; set; } = -1;
        public int Length { get; set; }
        /// <summary>The flagged text as it appeared when checked; used for stale detection.</summary>
        public string Text { get; set; } = "";
        public string Source { get; set; } = "Rules";
        public bool DisputedByAI { get; set; }
        public string DisputeReason { get; set; }

        public static Finding At(Claim claim, int start, int length, FindingType type, Severity severity,
            string message, string suggestion = null)
        {
            string text = claim.FullText ?? "";
            start = Math.Max(0, Math.Min(start, text.Length));
            length = Math.Max(0, Math.Min(length, text.Length - start));
            int docStart = claim.ToDocOffset(start);
            return new Finding
            {
                ClaimNumber = claim.Number,
                Type = type,
                Severity = severity,
                Message = message,
                Suggestion = suggestion,
                DocStart = docStart,
                Length = claim.ToDocOffset(start + length) - docStart,
                Text = text.Substring(start, length)
            };
        }
    }
}
