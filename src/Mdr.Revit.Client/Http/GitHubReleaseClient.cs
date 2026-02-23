using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using Mdr.Revit.Client.Retry;
using Mdr.Revit.Core.Contracts;
using Mdr.Revit.Core.Models;

namespace Mdr.Revit.Client.Http
{
    public sealed class GitHubReleaseClient : IUpdateFeedClient, IDisposable
    {
        private readonly HttpClient _httpClient;
        private readonly RetryPolicy _retryPolicy;
        private bool _disposed;

        public GitHubReleaseClient()
            : this(new HttpClient(), new RetryPolicy())
        {
        }

        public GitHubReleaseClient(HttpClient httpClient, RetryPolicy retryPolicy)
        {
            _httpClient = httpClient ?? throw new ArgumentNullException(nameof(httpClient));
            _retryPolicy = retryPolicy ?? throw new ArgumentNullException(nameof(retryPolicy));
        }

        public Task<UpdateManifest> GetLatestAsync(UpdateCheckRequest request, CancellationToken cancellationToken)
        {
            if (request == null)
            {
                throw new ArgumentNullException(nameof(request));
            }

            string repo = (request.GithubRepo ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(repo) || !repo.Contains("/"))
            {
                throw new InvalidOperationException("GithubRepo must be in 'owner/repo' format.");
            }

            return _retryPolicy.ExecuteAsync(async token =>
            {
                EnsureNotDisposed();

                string url = "https://api.github.com/repos/" + repo + "/releases?per_page=20";
                using (HttpRequestMessage httpRequest = new HttpRequestMessage(HttpMethod.Get, url))
                {
                    httpRequest.Headers.UserAgent.ParseAdd("Mdr-Revit-Plugin/1.0");
                    using (HttpResponseMessage response = await _httpClient.SendAsync(httpRequest, token).ConfigureAwait(false))
                    {
                        string body = response.Content == null
                            ? string.Empty
                            : await response.Content.ReadAsStringAsync().ConfigureAwait(false);

                        if (!response.IsSuccessStatusCode)
                        {
                            throw new HttpRequestException(
                                "GitHub release query failed with status " + (int)response.StatusCode + ": " + body);
                        }

                        return ParseLatest(body, request.Channel);
                    }
                }
            }, cancellationToken);
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

        private static UpdateManifest ParseLatest(string body, string channel)
        {
            if (string.IsNullOrWhiteSpace(body))
            {
                return new UpdateManifest();
            }

            bool includePrerelease = !string.Equals(channel, "stable", StringComparison.OrdinalIgnoreCase);
            using (JsonDocument document = JsonDocument.Parse(body))
            {
                if (document.RootElement.ValueKind != JsonValueKind.Array)
                {
                    return new UpdateManifest();
                }

                foreach (JsonElement release in document.RootElement.EnumerateArray())
                {
                    bool draft = GetBool(release, "draft");
                    bool prerelease = GetBool(release, "prerelease");
                    if (draft)
                    {
                        continue;
                    }

                    if (prerelease && !includePrerelease)
                    {
                        continue;
                    }

                    UpdateManifest manifest = new UpdateManifest
                    {
                        Version = NormalizeVersion(GetString(release, "tag_name")),
                        PublishedAtUtc = GetDateTimeOffset(release, "published_at"),
                        ReleaseNotes = GetString(release, "body"),
                    };

                    Dictionary<string, string> shaByAsset = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                    if (release.TryGetProperty("assets", out JsonElement assets) && assets.ValueKind == JsonValueKind.Array)
                    {
                        foreach (JsonElement asset in assets.EnumerateArray())
                        {
                            string name = GetString(asset, "name");
                            string download = GetString(asset, "browser_download_url");
                            if (string.IsNullOrWhiteSpace(name) || string.IsNullOrWhiteSpace(download))
                            {
                                continue;
                            }

                            if (name.EndsWith(".sha256", StringComparison.OrdinalIgnoreCase))
                            {
                                string key = name.Substring(0, name.Length - ".sha256".Length);
                                shaByAsset[key] = download;
                                continue;
                            }

                            manifest.Assets.Add(new UpdateAsset
                            {
                                Name = name,
                                DownloadUrl = download,
                                SizeBytes = GetLong(asset, "size"),
                            });
                        }
                    }

                    foreach (UpdateAsset asset in manifest.Assets)
                    {
                        if (shaByAsset.TryGetValue(asset.Name, out string shaUrl))
                        {
                            asset.Sha256 = shaUrl;
                        }
                    }

                    return manifest;
                }
            }

            return new UpdateManifest();
        }

        private static string NormalizeVersion(string tagName)
        {
            string version = (tagName ?? string.Empty).Trim();
            if (version.StartsWith("v", StringComparison.OrdinalIgnoreCase))
            {
                version = version.Substring(1);
            }

            return version;
        }

        private static string GetString(JsonElement element, string propertyName)
        {
            if (!element.TryGetProperty(propertyName, out JsonElement value))
            {
                return string.Empty;
            }

            return value.GetString() ?? string.Empty;
        }

        private static bool GetBool(JsonElement element, string propertyName)
        {
            if (!element.TryGetProperty(propertyName, out JsonElement value))
            {
                return false;
            }

            return value.ValueKind == JsonValueKind.True;
        }

        private static long GetLong(JsonElement element, string propertyName)
        {
            if (!element.TryGetProperty(propertyName, out JsonElement value))
            {
                return 0;
            }

            return value.TryGetInt64(out long parsed) ? parsed : 0;
        }

        private static DateTimeOffset? GetDateTimeOffset(JsonElement element, string propertyName)
        {
            string raw = GetString(element, propertyName);
            if (string.IsNullOrWhiteSpace(raw))
            {
                return null;
            }

            return DateTimeOffset.TryParse(raw, out DateTimeOffset parsed)
                ? parsed
                : null;
        }

        private void EnsureNotDisposed()
        {
            if (_disposed)
            {
                throw new ObjectDisposedException(nameof(GitHubReleaseClient));
            }
        }
    }
}
