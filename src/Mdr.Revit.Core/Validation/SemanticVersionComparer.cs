using System;

namespace Mdr.Revit.Core.Validation
{
    public static class SemanticVersionComparer
    {
        public static int Compare(string leftVersion, string rightVersion)
        {
            Version left = Parse(leftVersion);
            Version right = Parse(rightVersion);
            return left.CompareTo(right);
        }

        public static bool IsGreater(string candidateVersion, string currentVersion)
        {
            return Compare(candidateVersion, currentVersion) > 0;
        }

        private static Version Parse(string value)
        {
            string raw = (value ?? string.Empty).Trim();
            if (raw.StartsWith("v", StringComparison.OrdinalIgnoreCase))
            {
                raw = raw.Substring(1);
            }

            if (string.IsNullOrWhiteSpace(raw))
            {
                return new Version(0, 0, 0, 0);
            }

            int dash = raw.IndexOf('-');
            if (dash >= 0)
            {
                raw = raw.Substring(0, dash);
            }

            string[] parts = raw.Split('.');
            int major = ParsePart(parts, 0);
            int minor = ParsePart(parts, 1);
            int build = ParsePart(parts, 2);
            int revision = ParsePart(parts, 3);
            return new Version(major, minor, build, revision);
        }

        private static int ParsePart(string[] parts, int index)
        {
            if (parts == null || index < 0 || index >= parts.Length)
            {
                return 0;
            }

            if (int.TryParse(parts[index], out int parsed) && parsed >= 0)
            {
                return parsed;
            }

            return 0;
        }
    }
}
