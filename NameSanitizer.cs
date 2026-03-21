using System;
using System.Collections.Generic;
using System.Text;

namespace XlsToAccessImporter
{
    internal static class NameSanitizer
    {
        public static string SanitizeColumnName(string rawName, int index, ISet<string> usedNames)
        {
            string baseName = SanitizeIdentifier(rawName, "F" + index);
            return EnsureUnique(baseName, usedNames);
        }

        public static string SanitizeTableName(string rawName)
        {
            return SanitizeIdentifier(rawName, "ImportedSheet");
        }

        public static string EscapeIdentifier(string name)
        {
            return "[" + (name ?? string.Empty).Replace("]", "]]" ) + "]";
        }

        private static string SanitizeIdentifier(string rawName, string fallback)
        {
            string candidate = (rawName ?? string.Empty).Trim();
            if (candidate.Length == 0)
            {
                candidate = fallback;
            }

            StringBuilder builder = new StringBuilder(candidate.Length);
            foreach (char ch in candidate)
            {
                if (char.IsLetterOrDigit(ch) || ch == '_' || ch == ' ')
                {
                    builder.Append(ch);
                }
                else
                {
                    builder.Append('_');
                }
            }

            string cleaned = builder.ToString().Trim();
            while (cleaned.Contains("  "))
            {
                cleaned = cleaned.Replace("  ", " ");
            }

            if (cleaned.Length == 0)
            {
                cleaned = fallback;
            }

            if (char.IsDigit(cleaned[0]))
            {
                cleaned = "T_" + cleaned;
            }

            if (cleaned.Length > 50)
            {
                cleaned = cleaned.Substring(0, 50).Trim();
            }

            return cleaned;
        }

        private static string EnsureUnique(string baseName, ISet<string> usedNames)
        {
            string candidate = baseName;
            int suffix = 2;
            while (usedNames.Contains(candidate))
            {
                string suffixText = "_" + suffix;
                int maxBaseLength = 50 - suffixText.Length;
                string prefix = baseName.Length > maxBaseLength ? baseName.Substring(0, maxBaseLength) : baseName;
                candidate = prefix + suffixText;
                suffix++;
            }

            usedNames.Add(candidate);
            return candidate;
        }
    }
}
