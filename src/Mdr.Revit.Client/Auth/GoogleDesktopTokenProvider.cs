using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace Mdr.Revit.Client.Auth
{
    public sealed class GoogleDesktopTokenProvider : IGoogleTokenProvider, IDisposable
    {
        private readonly string _clientId;
        private readonly string _clientSecret;
        private readonly GoogleTokenStore _tokenStore;
        private readonly HttpClient _httpClient;
        private bool _disposed;

        public GoogleDesktopTokenProvider(
            string clientId,
            string clientSecret,
            GoogleTokenStore tokenStore)
            : this(clientId, clientSecret, tokenStore, new HttpClient())
        {
        }

        public GoogleDesktopTokenProvider(
            string clientId,
            string clientSecret,
            GoogleTokenStore tokenStore,
            HttpClient httpClient)
        {
            _clientId = (clientId ?? string.Empty).Trim();
            _clientSecret = (clientSecret ?? string.Empty).Trim();
            _tokenStore = tokenStore ?? throw new ArgumentNullException(nameof(tokenStore));
            _httpClient = httpClient ?? throw new ArgumentNullException(nameof(httpClient));
        }

        public async Task<string> GetAccessTokenAsync(CancellationToken cancellationToken)
        {
            EnsureNotDisposed();

            GoogleOAuthToken token = _tokenStore.Load();
            if (token.HasUsableAccessToken())
            {
                return token.AccessToken;
            }

            if (string.IsNullOrWhiteSpace(token.RefreshToken))
            {
                throw new InvalidOperationException(
                    "Google refresh token is missing. Complete Google OAuth desktop authorization first.");
            }

            if (string.IsNullOrWhiteSpace(_clientId) || string.IsNullOrWhiteSpace(_clientSecret))
            {
                throw new InvalidOperationException("Google OAuth client credentials are required.");
            }

            using (HttpRequestMessage request = new HttpRequestMessage(
                HttpMethod.Post,
                "https://oauth2.googleapis.com/token"))
            {
                request.Content = new FormUrlEncodedContent(new[]
                {
                    new KeyValuePair<string, string>("client_id", _clientId),
                    new KeyValuePair<string, string>("client_secret", _clientSecret),
                    new KeyValuePair<string, string>("refresh_token", token.RefreshToken),
                    new KeyValuePair<string, string>("grant_type", "refresh_token"),
                });

                using (HttpResponseMessage response = await _httpClient.SendAsync(request, cancellationToken).ConfigureAwait(false))
                {
                    string body = response.Content == null
                        ? string.Empty
                        : await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                    if (!response.IsSuccessStatusCode)
                    {
                        throw new HttpRequestException(
                            "Google token refresh failed with status " + (int)response.StatusCode + ": " + body);
                    }

                    GoogleRefreshTokenResponse? parsed = JsonSerializer.Deserialize<GoogleRefreshTokenResponse>(body);
                    if (parsed == null || string.IsNullOrWhiteSpace(parsed.AccessToken))
                    {
                        throw new InvalidOperationException("Google token refresh response does not contain access_token.");
                    }

                    token.AccessToken = parsed.AccessToken;
                    int expiresIn = parsed.ExpiresIn <= 0 ? 3600 : parsed.ExpiresIn;
                    token.ExpiresAtUtc = DateTimeOffset.UtcNow.AddSeconds(expiresIn);
                    _tokenStore.Save(token);
                    return token.AccessToken;
                }
            }
        }

        public void Dispose()
        {
            if (_disposed)
            {
                return;
            }

            _httpClient.Dispose();
            _disposed = true;
        }

        private void EnsureNotDisposed()
        {
            if (_disposed)
            {
                throw new ObjectDisposedException(nameof(GoogleDesktopTokenProvider));
            }
        }

        private sealed class GoogleRefreshTokenResponse
        {
            public string AccessToken { get; set; } = string.Empty;

            public int ExpiresIn { get; set; }
        }
    }
}
