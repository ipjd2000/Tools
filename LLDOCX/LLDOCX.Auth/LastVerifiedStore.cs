using System;
using System.Globalization;
using System.IO;
using System.Security.Cryptography;
using System.Text;

namespace LLDOCX.Auth
{
    public interface ILastVerifiedStore
    {
        /// <summary>Returns null when there is no record or it can't be read.</summary>
        VerifiedRecord Load();

        void Save(VerifiedRecord record);

        void Clear();
    }

    /// <summary>Stores the last-verified record encrypted with DPAPI for the current Windows user.</summary>
    public sealed class DpapiLastVerifiedStore : ILastVerifiedStore
    {
        static readonly byte[] Entropy = Encoding.UTF8.GetBytes("LL-Docx gate record v1");
        readonly string _filePath;

        public DpapiLastVerifiedStore(string filePath)
        {
            _filePath = filePath;
        }

        public VerifiedRecord Load()
        {
            try
            {
                if (!File.Exists(_filePath)) return null;
                byte[] plain = ProtectedData.Unprotect(File.ReadAllBytes(_filePath), Entropy, DataProtectionScope.CurrentUser);
                string[] lines = Encoding.UTF8.GetString(plain).Split('\n');
                if (lines.Length != 4 || lines[0] != "v1") return null;
                long ticks = long.Parse(lines[1], CultureInfo.InvariantCulture);
                return new VerifiedRecord(new DateTime(ticks, DateTimeKind.Utc), lines[2], lines[3]);
            }
            catch (Exception ex) when (ex is CryptographicException || ex is FormatException || ex is IOException ||
                                       ex is UnauthorizedAccessException || ex is OverflowException || ex is ArgumentException)
            {
                return null;
            }
        }

        public void Save(VerifiedRecord record)
        {
            string text = string.Join("\n", "v1",
                record.VerifiedUtc.ToUniversalTime().Ticks.ToString(CultureInfo.InvariantCulture),
                Clean(record.AccountId), Clean(record.Username));
            byte[] cipher = ProtectedData.Protect(Encoding.UTF8.GetBytes(text), Entropy, DataProtectionScope.CurrentUser);
            Directory.CreateDirectory(Path.GetDirectoryName(_filePath));
            File.WriteAllBytes(_filePath, cipher);
        }

        public void Clear()
        {
            if (File.Exists(_filePath)) File.Delete(_filePath);
        }

        static string Clean(string s) => (s ?? "").Replace("\n", " ").Replace("\r", " ");
    }
}
