using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;

namespace Mdr.Revit.Client.Auth
{
    public sealed class GoogleOAuthDesktopFlow
    {
        private const string RedirectHost = "http://127.0.0.1";
        private readonly HttpClient _httpClient;

        public GoogleOAuthDesktopFlow()
            : this(new HttpClient())
        {
        }

        public GoogleOAuthDesktopFlow(HttpClient httpClient)
        {
            _httpClient = httpClient ?? throw new ArgumentNullException(nameof(httpClient));
        }

        public async Task<GoogleOAuthToken> AuthorizeAsync(
            string clientId,
            string clientSecret,
            CancellationToken cancellationToken)
        {
            string normalizedClientId = (clientId ?? string.Empty).Trim();
            string normalizedClientSecret = (clientSecret ?? string.Empty).Trim();

            if (string.IsNullOrWhiteSpace(normalizedClientId) || string.IsNullOrWhiteSpace(normalizedClientSecret))
            {
                throw new InvalidOperationException("Google OAuth client_id and client_secret are required.");
            }

            int port = AllocateLocalPort();
            string redirectUri = $"{RedirectHost}:{port}/callback/";
            string scope = Uri.EscapeDataString("https://www.googleapis.com/auth/spreadsheets");
            string state = Guid.NewGuid().ToString("N");
            string authorizeUrl =
                "https://accounts.google.com/o/oauth2/v2/auth" +
                "?response_type=code" +
                "&access_type=offline" +
                "&prompt=consent" +
                "&scope=" + scope +
                "&redirect_uri=" + Uri.EscapeDataString(redirectUri) +
                "&client_id=" + Uri.EscapeDataString(normalizedClientId) +
                "&state=" + Uri.EscapeDataString(state);

            using (HttpListener listener = new HttpListener())
            {
                listener.Prefixes.Add($"{RedirectHost}:{port}/callback/");
                listener.Start();

                Process.Start(new ProcessStartInfo
                {
                    FileName = authorizeUrl,
                    UseShellExecute = true,
                });

                HttpListenerContext context = await listener.GetContextAsync().ConfigureAwait(false);
                string code = context.Request.QueryString["code"] ?? string.Empty;
                string callbackState = context.Request.QueryString["state"] ?? string.Empty;
                if (!string.Equals(state, callbackState, StringComparison.Ordinal))
                {
                    throw new InvalidOperationException("Google OAuth state validation failed.");
                }

                byte[] responseBytes = Encoding.UTF8.GetBytes("Authorization completed. You can close this window.");
                context.Response.StatusCode = 200;
                context.Response.ContentType = "text/plain";
                context.Response.OutputStream.Write(responseBytes, 0, responseBytes.Length);
                context.Response.OutputStream.Flush();
                context.Response.Close();

                if (string.IsNullOrWhiteSpace(code))
                {
                    throw new InvalidOperationException("Google OAuth did not return an authorization code.");
                }

                using (HttpRequestMessage tokenRequest = new HttpRequestMessage(HttpMethod.Post, "https://oauth2.googleapis.com/token"))
                {
                    tokenRequest.Content = new FormUrlEncodedContent(new[]
                    {
                        new KeyValuePair<string, string>("code", code),
                        new KeyValuePair<string, string>("client_id", normalizedClientId),
                        new KeyValuePair<string, string>("client_secret", normalizedClientSecret),
                        new KeyValuePair<string, string>("redirect_uri", redirectUri),
                        new KeyValuePair<string, string>("grant_type", "authorization_code"),
                    });

                    using (HttpResponseMessage response = await _httpClient.SendAsync(tokenRequest, cancellationToken).ConfigureAwait(false))
                    {
                        string body = response.Content == null
                            ? string.Empty
                            : await response.Content.ReadAsStringAsync().ConfigureAwait(false);
                        if (!response.IsSuccessStatusCode)
                        {
                            throw new HttpRequestException(
                                "Google OAuth token exchange failed with status " + (int)response.StatusCode + ": " + body);
                        }

                        GoogleTokenResponse? parsed = JsonSerializer.Deserialize<GoogleTokenResponse>(body);
                        if (parsed == null || string.IsNullOrWhiteSpace(parsed.AccessToken))
                        {
                            throw new InvalidOperationException("Google OAuth token response is missing access_token.");
                        }

                        return new GoogleOAuthToken
                        {
                            AccessToken = parsed.AccessToken,
                            RefreshToken = parsed.RefreshToken ?? string.Empty,
                            ExpiresAtUtc = DateTimeOffset.UtcNow.AddSeconds(parsed.ExpiresIn <= 0 ? 3600 : parsed.ExpiresIn),
                        };
                    }
                }
            }
        }

        private static int AllocateLocalPort()
        {
            Random random = new Random();
            return random.Next(41000, 48000);
        }

        private sealed class GoogleTokenResponse
        {
            public string AccessToken { get; set; } = string.Empty;

            public string RefreshToken { get; set; } = string.Empty;

            public int ExpiresIn { get; set; }
        }
    }
}
