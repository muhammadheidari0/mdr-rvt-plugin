using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Mdr.Revit.Client.Auth;
using Xunit;

namespace Mdr.Revit.Client.Tests
{
    public sealed class GoogleDesktopTokenProviderTests
    {
        [Fact]
        public async Task GetAccessTokenAsync_ParsesSnakeCaseRefreshResponse()
        {
            string tokenPath = Path.Combine(Path.GetTempPath(), "mdr-google-token-" + Guid.NewGuid().ToString("N") + ".dat");
            try
            {
                GoogleTokenStore store = new GoogleTokenStore(tokenPath);
                store.Save(new GoogleOAuthToken
                {
                    RefreshToken = "refresh-token-value",
                });

                StubHttpMessageHandler handler = new StubHttpMessageHandler(async request =>
                {
                    string content = request.Content == null
                        ? string.Empty
                        : await request.Content.ReadAsStringAsync().ConfigureAwait(false);
                    Assert.Contains("client_id=my-client-id", content);
                    Assert.Contains("client_secret=my-client-secret", content);
                    Assert.Contains("grant_type=refresh_token", content);
                    Assert.Contains("refresh_token=refresh-token-value", content);

                    return new HttpResponseMessage(HttpStatusCode.OK)
                    {
                        Content = new StringContent(
                            "{\"access_token\":\"new-access-token\",\"expires_in\":3599}",
                            Encoding.UTF8,
                            "application/json"),
                    };
                });

                using HttpClient httpClient = new HttpClient(handler);
                using GoogleDesktopTokenProvider provider = new GoogleDesktopTokenProvider(
                    "my-client-id",
                    "my-client-secret",
                    store,
                    httpClient);

                string token = await provider.GetAccessTokenAsync(CancellationToken.None);

                Assert.Equal("new-access-token", token);
                GoogleOAuthToken persisted = store.Load();
                Assert.Equal("new-access-token", persisted.AccessToken);
                Assert.Equal("refresh-token-value", persisted.RefreshToken);
                Assert.True(persisted.ExpiresAtUtc > DateTimeOffset.UtcNow);
            }
            finally
            {
                if (File.Exists(tokenPath))
                {
                    File.Delete(tokenPath);
                }
            }
        }

        private sealed class StubHttpMessageHandler : HttpMessageHandler
        {
            private readonly Func<HttpRequestMessage, Task<HttpResponseMessage>> _handler;

            public StubHttpMessageHandler(Func<HttpRequestMessage, Task<HttpResponseMessage>> handler)
            {
                _handler = handler ?? throw new ArgumentNullException(nameof(handler));
            }

            protected override Task<HttpResponseMessage> SendAsync(HttpRequestMessage request, CancellationToken cancellationToken)
            {
                _ = cancellationToken;
                return _handler(request);
            }
        }
    }
}
