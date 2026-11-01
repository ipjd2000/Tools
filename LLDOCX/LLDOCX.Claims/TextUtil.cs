using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace LLDOCX.Claims
{
    public struct Token
    {
        public Token(string value, int index)
        {
            Value = value;
            Index = index;
        }

        public string Value { get; }
        public int Index { get; }
        public int End => Index + Value.Length;
        public bool IsWord => Value.Length > 0 && char.IsLetter(Value[0]);
        public string Lower => Value.ToLowerInvariant();
    }

    public static class TextUtil
    {
        static readonly Regex TokenRx = new Regex(@"[A-Za-z][A-Za-z0-9'\-]*|\d+(?:\.\d+)?|[^\sA-Za-z0-9]", RegexOptions.Compiled);

        public static List<Token> Tokenize(string text)
        {
            var list = new List<Token>();
            foreach (Match m in TokenRx.Matches(text ?? ""))
                list.Add(new Token(m.Value, m.Index));
            return list;
        }

        /// <summary>Rule-based singular form of a lowercase English word.</summary>
        public static string Singularize(string word)
        {
            if (string.IsNullOrEmpty(word) || word.Length <= 3) return word;
            if (word.EndsWith("ies") && word.Length > 4) return word.Substring(0, word.Length - 3) + "y";
            if (word.EndsWith("sses")) return word.Substring(0, word.Length - 2);
            if (word.EndsWith("xes") || word.EndsWith("ches") || word.EndsWith("shes")) return word.Substring(0, word.Length - 2);
            if (word.EndsWith("ss") || word.EndsWith("us") || word.EndsWith("is")) return word;
            if (word.EndsWith("s")) return word.Substring(0, word.Length - 1);
            return word;
        }

        public static bool IsPlural(string word) => Singularize(word) != word;

        public static int Levenshtein(string a, string b)
        {
            a = a ?? "";
            b = b ?? "";
            var prev = new int[b.Length + 1];
            var cur = new int[b.Length + 1];
            for (int j = 0; j <= b.Length; j++) prev[j] = j;
            for (int i = 1; i <= a.Length; i++)
            {
                cur[0] = i;
                for (int j = 1; j <= b.Length; j++)
                {
                    int cost = a[i - 1] == b[j - 1] ? 0 : 1;
                    cur[j] = Math.Min(Math.Min(cur[j - 1] + 1, prev[j] + 1), prev[j - 1] + cost);
                }
                var tmp = prev; prev = cur; cur = tmp;
            }
            return prev[b.Length];
        }
    }
}
