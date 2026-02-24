using System;
using System.IO;
using System.Reflection;
using System.Text.Json;

namespace Mdr.Revit.Addin.Security
{
    internal readonly struct EmbeddedGoogleOAuthCredentials
    {
        public EmbeddedGoogleOAuthCredentials(string clientId, string clientSecret)
        {
            ClientId = clientId ?? string.Empty;
            ClientSecret = clientSecret ?? string.Empty;
        }

        public string ClientId { get; }

        public string ClientSecret { get; }

        public bool IsConfigured =>
            !string.IsNullOrWhiteSpace(ClientId) &&
            !string.IsNullOrWhiteSpace(ClientSecret);

        public static EmbeddedGoogleOAuthCredentials Empty => new EmbeddedGoogleOAuthCredentials(string.Empty, string.Empty);
    }

    internal static class EmbeddedGoogleOAuthCredentialsProvider
    {
        private const string ResourceSuffix = ".Resources.Google.credentials.json";
        private const string PlaceholderClientId = "YOUR_GOOGLE_CLIENT_ID.apps.googleusercontent.com";
        private const string PlaceholderClientSecret = "YOUR_GOOGLE_CLIENT_SECRET";

        public static EmbeddedGoogleOAuthCredentials Load()
        {
            try
            {
                Assembly assembly = typeof(EmbeddedGoogleOAuthCredentialsProvider).Assembly;
                string? resourceName = FindResourceName(assembly);
                if (string.IsNullOrWhiteSpace(resourceName))
                {
                    return EmbeddedGoogleOAuthCredentials.Empty;
                }

                using Stream? stream = assembly.GetManifestResourceStream(resourceName);
                if (stream == null)
                {
                    return EmbeddedGoogleOAuthCredentials.Empty;
                }

                using JsonDocument document = JsonDocument.Parse(stream);
                if (TryParse(document.RootElement, out EmbeddedGoogleOAuthCredentials credentials))
                {
                    return credentials;
                }
            }
            catch
            {
                return EmbeddedGoogleOAuthCredentials.Empty;
            }

            return EmbeddedGoogleOAuthCredentials.Empty;
        }

        internal static bool TryParse(string json, out EmbeddedGoogleOAuthCredentials credentials)
        {
            credentials = EmbeddedGoogleOAuthCredentials.Empty;
            if (string.IsNullOrWhiteSpace(json))
            {
                return false;
            }

            try
            {
                using JsonDocument document = JsonDocument.Parse(json);
                return TryParse(document.RootElement, out credentials);
            }
            catch
            {
                return false;
            }
        }

        private static string? FindResourceName(Assembly assembly)
        {
            string[] names = assembly.GetManifestResourceNames();
            for (int i = 0; i < names.Length; i++)
            {
                if (names[i].EndsWith(ResourceSuffix, StringComparison.OrdinalIgnoreCase))
                {
                    return names[i];
                }
            }

            return null;
        }

        private static bool TryParse(JsonElement root, out EmbeddedGoogleOAuthCredentials credentials)
        {
            credentials = EmbeddedGoogleOAuthCredentials.Empty;

            if (root.ValueKind != JsonValueKind.Object)
            {
                return false;
            }

            if (root.TryGetProperty("installed", out JsonElement installed) &&
                TryCreate(installed, out credentials))
            {
                return true;
            }

            if (root.TryGetProperty("web", out JsonElement web) &&
                TryCreate(web, out credentials))
            {
                return true;
            }

            return TryCreate(root, out credentials);
        }

        private static bool TryCreate(JsonElement section, out EmbeddedGoogleOAuthCredentials credentials)
        {
            credentials = EmbeddedGoogleOAuthCredentials.Empty;
            if (section.ValueKind != JsonValueKind.Object)
            {
                return false;
            }

            string clientId = ReadString(section, "client_id");
            string clientSecret = ReadString(section, "client_secret");
            if (!IsValidCredential(clientId, clientSecret))
            {
                return false;
            }

            credentials = new EmbeddedGoogleOAuthCredentials(clientId.Trim(), clientSecret.Trim());
            return true;
        }

        private static string ReadString(JsonElement element, string propertyName)
        {
            if (element.TryGetProperty(propertyName, out JsonElement property) &&
                property.ValueKind == JsonValueKind.String)
            {
                return property.GetString() ?? string.Empty;
            }

            return string.Empty;
        }

        private static bool IsValidCredential(string clientId, string clientSecret)
        {
            if (string.IsNullOrWhiteSpace(clientId) || string.IsNullOrWhiteSpace(clientSecret))
            {
                return false;
            }

            return
                !string.Equals(clientId, PlaceholderClientId, StringComparison.OrdinalIgnoreCase) &&
                !string.Equals(clientSecret, PlaceholderClientSecret, StringComparison.OrdinalIgnoreCase);
        }
    }
}
