using System;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Mdr.Revit.Client.Http;
using Mdr.Revit.Client.Retry;
using Mdr.Revit.Core.Models;
using Xunit;

namespace Mdr.Revit.Client.Tests
{
    public sealed class GitHubReleaseClientTests
    {
        [Fact]
        public async Task GetLatestAsync_StableChannel_IgnoresPrerelease()
        {
            StubHttpMessageHandler handler = new StubHttpMessageHandler(_ =>
            {
                string payload = "[" +
                    "{\"tag_name\":\"v2.0.0-beta\",\"draft\":false,\"prerelease\":true,\"assets\":[]}," +
                    "{\"tag_name\":\"v1.9.0\",\"draft\":false,\"prerelease\":false,\"assets\":[{\"name\":\"plugin.msi\",\"browser_download_url\":\"https://example/plugin.msi\",\"size\":100}]}" +
                    "]";
                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new StringContent(payload, Encoding.UTF8, "application/json"),
                });
            });

            HttpClient http = new HttpClient(handler);
            GitHubReleaseClient client = new GitHubReleaseClient(http, new RetryPolicy(maxAttempts: 1));

            UpdateManifest manifest = await client.GetLatestAsync(
                new UpdateCheckRequest
                {
                    GithubRepo = "owner/repo",
                    Channel = "stable",
                },
                CancellationToken.None);

            Assert.Equal("1.9.0", manifest.Version);
            Assert.Single(manifest.Assets);
            Assert.Equal("plugin.msi", manifest.Assets[0].Name);
        }

        [Fact]
        public async Task GetLatestAsync_PreviewChannel_UsesPrerelease()
        {
            StubHttpMessageHandler handler = new StubHttpMessageHandler(_ =>
            {
                string payload = "[" +
                    "{\"tag_name\":\"v2.0.0-beta\",\"draft\":false,\"prerelease\":true,\"assets\":[{\"name\":\"plugin-preview.msi\",\"browser_download_url\":\"https://example/plugin-preview.msi\",\"size\":100}]}" +
                    "]";
                return Task.FromResult(new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new StringContent(payload, Encoding.UTF8, "application/json"),
                });
            });

            HttpClient http = new HttpClient(handler);
            GitHubReleaseClient client = new GitHubReleaseClient(http, new RetryPolicy(maxAttempts: 1));

            UpdateManifest manifest = await client.GetLatestAsync(
                new UpdateCheckRequest
                {
                    GithubRepo = "owner/repo",
                    Channel = "preview",
                },
                CancellationToken.None);

            Assert.Equal("2.0.0-beta", manifest.Version);
            Assert.Single(manifest.Assets);
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
