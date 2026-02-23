using System;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json;

namespace Mdr.Revit.Client.Auth
{
    public sealed class GoogleTokenStore
    {
        private readonly string _filePath;

        public GoogleTokenStore(string filePath)
        {
            if (string.IsNullOrWhiteSpace(filePath))
            {
                throw new ArgumentException("Token file path is required.", nameof(filePath));
            }

            _filePath = filePath;
        }

        public GoogleOAuthToken Load()
        {
            if (!File.Exists(_filePath))
            {
                return new GoogleOAuthToken();
            }

            byte[] cipher = File.ReadAllBytes(_filePath);
            if (cipher.Length == 0)
            {
                return new GoogleOAuthToken();
            }

            byte[] plain;
            try
            {
                plain = ProtectedData.Unprotect(cipher, null, DataProtectionScope.CurrentUser);
            }
            catch
            {
                return new GoogleOAuthToken();
            }

            string json = Encoding.UTF8.GetString(plain);
            if (string.IsNullOrWhiteSpace(json))
            {
                return new GoogleOAuthToken();
            }

            GoogleOAuthToken? token = JsonSerializer.Deserialize<GoogleOAuthToken>(json);
            return token ?? new GoogleOAuthToken();
        }

        public void Save(GoogleOAuthToken token)
        {
            if (token == null)
            {
                throw new ArgumentNullException(nameof(token));
            }

            string? directory = Path.GetDirectoryName(_filePath);
            if (!string.IsNullOrWhiteSpace(directory))
            {
                Directory.CreateDirectory(directory);
            }

            string json = JsonSerializer.Serialize(token);
            byte[] plain = Encoding.UTF8.GetBytes(json);
            byte[] cipher = ProtectedData.Protect(plain, null, DataProtectionScope.CurrentUser);
            File.WriteAllBytes(_filePath, cipher);
        }
    }
}
