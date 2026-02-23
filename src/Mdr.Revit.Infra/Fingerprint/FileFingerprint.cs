using System;
using System.IO;
using System.Security.Cryptography;

namespace Mdr.Revit.Infra.Fingerprint
{
    public static class FileFingerprint
    {
        public static string ComputeSha256(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                throw new ArgumentException("File path is required.", nameof(filePath));
            }

            using (SHA256 sha = SHA256.Create())
            using (FileStream stream = File.OpenRead(filePath))
            {
                byte[] hash = sha.ComputeHash(stream);
                return ToHex(hash);
            }
        }

        private static string ToHex(byte[] hash)
        {
            char[] chars = new char[hash.Length * 2];

            for (int i = 0; i < hash.Length; i++)
            {
                byte value = hash[i];
                chars[i * 2] = GetHexDigit(value / 16);
                chars[(i * 2) + 1] = GetHexDigit(value % 16);
            }

            return new string(chars);
        }

        private static char GetHexDigit(int value)
        {
            if (value < 10)
            {
                return (char)('0' + value);
            }

            return (char)('a' + (value - 10));
        }
    }
}
