using System;

namespace Mdr.Revit.Client.Auth
{
    public sealed class GoogleOAuthToken
    {
        public string AccessToken { get; set; } = string.Empty;

        public string RefreshToken { get; set; } = string.Empty;

        public DateTimeOffset ExpiresAtUtc { get; set; } = DateTimeOffset.MinValue;

        public bool HasUsableAccessToken()
        {
            return !string.IsNullOrWhiteSpace(AccessToken) && ExpiresAtUtc > DateTimeOffset.UtcNow.AddMinutes(1);
        }
    }
}
