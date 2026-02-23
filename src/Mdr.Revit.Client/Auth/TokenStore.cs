using System;

namespace Mdr.Revit.Client.Auth
{
    public sealed class TokenStore
    {
        private readonly object _sync = new object();
        private string _accessToken = string.Empty;
        private DateTimeOffset _expiresAtUtc = DateTimeOffset.MinValue;

        public void Set(string token, DateTimeOffset expiresAtUtc)
        {
            if (string.IsNullOrWhiteSpace(token))
            {
                throw new ArgumentException("Token cannot be empty.", nameof(token));
            }

            lock (_sync)
            {
                _accessToken = token;
                _expiresAtUtc = expiresAtUtc;
            }
        }

        public string Get()
        {
            lock (_sync)
            {
                return _accessToken;
            }
        }

        public bool HasValidToken()
        {
            lock (_sync)
            {
                return !string.IsNullOrWhiteSpace(_accessToken) && _expiresAtUtc > DateTimeOffset.UtcNow;
            }
        }

        public void Clear()
        {
            lock (_sync)
            {
                _accessToken = string.Empty;
                _expiresAtUtc = DateTimeOffset.MinValue;
            }
        }
    }
}
