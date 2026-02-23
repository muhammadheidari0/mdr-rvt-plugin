using System;
using System.Collections.Generic;
using System.IO;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Threading;
using System.Threading.Tasks;
using Mdr.Revit.Client.Auth;
using Mdr.Revit.Client.Retry;
using Mdr.Revit.Client.Serialization;
using Mdr.Revit.Core.Contracts;
using Mdr.Revit.Core.Models;

namespace Mdr.Revit.Client.Http
{
    public sealed class ApiClient : IApiClient, IDisposable
    {
        private readonly HttpClient _httpClient;
        private readonly TokenStore _tokenStore;
        private readonly RetryPolicy _retryPolicy;
        private bool _disposed;

        public ApiClient(Uri baseAddress)
            : this(new HttpClient { BaseAddress = baseAddress }, new TokenStore(), new RetryPolicy())
        {
        }

        public ApiClient(HttpClient httpClient, TokenStore tokenStore, RetryPolicy retryPolicy)
        {
            _httpClient = httpClient ?? throw new ArgumentNullException(nameof(httpClient));
            _tokenStore = tokenStore ?? throw new ArgumentNullException(nameof(tokenStore));
            _retryPolicy = retryPolicy ?? throw new ArgumentNullException(nameof(retryPolicy));
        }

        public Task<string> LoginAsync(string username, string password, CancellationToken cancellationToken)
        {
            return _retryPolicy.ExecuteAsync(async token =>
            {
                EnsureNotDisposed();

                using (HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, "/api/v1/auth/login"))
                {
                    request.Content = new FormUrlEncodedContent(new[]
                    {
                        new KeyValuePair<string, string>("username", username ?? string.Empty),
                        new KeyValuePair<string, string>("password", password ?? string.Empty),
                    });

                    using (HttpResponseMessage response = await _httpClient.SendAsync(request, token).ConfigureAwait(false))
                    {
                        LoginResponseDto loginResponse = await ParseResponse<LoginResponseDto>(
                            "/api/v1/auth/login",
                            response).ConfigureAwait(false);

                        string accessToken = !string.IsNullOrWhiteSpace(loginResponse.AccessToken)
                            ? loginResponse.AccessToken
                            : loginResponse.Token;

                        if (string.IsNullOrWhiteSpace(accessToken))
                        {
                            throw new InvalidOperationException("Login response did not contain a token.");
                        }

                        int expiresIn = loginResponse.ExpiresIn <= 0 ? 3600 : loginResponse.ExpiresIn;
                        _tokenStore.Set(accessToken, DateTimeOffset.UtcNow.AddSeconds(expiresIn));
                        return accessToken;
                    }
                }
            }, cancellationToken);
        }

        public Task<PublishBatchResponse> PublishBatchAsync(PublishBatchRequest request, CancellationToken cancellationToken)
        {
            if (request == null)
            {
                throw new ArgumentNullException(nameof(request));
            }

            if (request.HasAnyLocalFile())
            {
                return PostPublishMultipartAsync(request, cancellationToken);
            }

            object payload = BuildPublishBody(request, request.BuildFilesManifest());
            return PostJsonAsync<object, PublishBatchResponse>(
                "/api/v1/bim/edms/publish-batch",
                payload,
                requiresToken: true,
                cancellationToken: cancellationToken);
        }

        public Task<ScheduleIngestResponse> IngestScheduleAsync(ScheduleIngestRequest request, CancellationToken cancellationToken)
        {
            return PostJsonAsync<ScheduleIngestRequest, ScheduleIngestResponse>(
                "/api/v1/bim/schedules/ingest",
                request,
                requiresToken: true,
                cancellationToken: cancellationToken);
        }

        public Task<SiteLogManifestResponse> GetSiteLogManifestAsync(
            SiteLogManifestRequest request,
            CancellationToken cancellationToken)
        {
            string path = BuildManifestPath(request);
            return GetJsonAsync<SiteLogManifestResponse>(path, requiresToken: true, cancellationToken: cancellationToken);
        }

        public Task<SiteLogPullResponse> PullSiteLogRowsAsync(
            SiteLogPullRequest request,
            CancellationToken cancellationToken)
        {
            return PostJsonAsync<SiteLogPullRequest, SiteLogPullResponse>(
                "/api/v1/bim/site-logs/revit/pull",
                request,
                requiresToken: true,
                cancellationToken: cancellationToken);
        }

        public Task<SiteLogAckResponse> AckSiteLogSyncAsync(
            SiteLogAckRequest request,
            CancellationToken cancellationToken)
        {
            return PostJsonAsync<SiteLogAckRequest, SiteLogAckResponse>(
                "/api/v1/bim/site-logs/revit/ack",
                request,
                requiresToken: true,
                cancellationToken: cancellationToken);
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

        private Task<PublishBatchResponse> PostPublishMultipartAsync(
            PublishBatchRequest request,
            CancellationToken cancellationToken)
        {
            return _retryPolicy.ExecuteAsync(async token =>
            {
                EnsureNotDisposed();

                using (HttpRequestMessage httpRequest = CreatePublishMultipartRequest(request))
                {
                    AttachToken(httpRequest);

                    using (HttpResponseMessage response = await _httpClient.SendAsync(httpRequest, token).ConfigureAwait(false))
                    {
                        return await ParseResponse<PublishBatchResponse>(
                            "/api/v1/bim/edms/publish-batch",
                            response).ConfigureAwait(false);
                    }
                }
            }, cancellationToken);
        }

        private HttpRequestMessage CreatePublishMultipartRequest(PublishBatchRequest request)
        {
            MultipartFormDataContent content = new MultipartFormDataContent();
            List<PublishFileManifestItem> manifest = new List<PublishFileManifestItem>();

            AddFormField(content, "run_client_id", request.RunClientId);
            AddFormField(content, "project_code", request.ProjectCode);
            AddOptionalFormField(content, "revit_version", request.RevitVersion);
            AddOptionalFormField(content, "model_guid", request.ModelGuid);
            AddOptionalFormField(content, "model_title", request.ModelTitle);
            AddOptionalFormField(content, "plugin_version", request.PluginVersion);

            List<object> itemRows = new List<object>();
            foreach (PublishSheetItem item in request.Items)
            {
                string pdfUploadName = TryAttachUpload(
                    content,
                    item.PdfFilePath,
                    item.ItemIndex,
                    "pdf",
                    "application/pdf");

                string nativeUploadName = TryAttachUpload(
                    content,
                    item.NativeFilePath,
                    item.ItemIndex,
                    "native",
                    "application/octet-stream");

                itemRows.Add(BuildPublishItemBody(item));

                if (!string.IsNullOrWhiteSpace(pdfUploadName) ||
                    !string.IsNullOrWhiteSpace(nativeUploadName) ||
                    !string.IsNullOrWhiteSpace(item.FileSha256))
                {
                    manifest.Add(new PublishFileManifestItem
                    {
                        ItemIndex = item.ItemIndex,
                        SheetUniqueId = item.SheetUniqueId,
                        PdfFileName = pdfUploadName,
                        NativeFileName = nativeUploadName,
                        FileSha256 = item.FileSha256,
                    });
                }
            }

            request.FilesManifest.Clear();
            foreach (PublishFileManifestItem row in manifest)
            {
                request.FilesManifest.Add(row);
            }

            AddFormField(content, "items_json", JsonSerializer.Serialize(itemRows, JsonOptions.Default));
            AddFormField(content, "files_manifest", JsonSerializer.Serialize(manifest, JsonOptions.Default));

            HttpRequestMessage requestMessage = new HttpRequestMessage(HttpMethod.Post, "/api/v1/bim/edms/publish-batch");
            requestMessage.Content = content;
            return requestMessage;
        }

        private Task<TResponse> GetJsonAsync<TResponse>(
            string relativePath,
            bool requiresToken,
            CancellationToken cancellationToken)
            where TResponse : class, new()
        {
            return _retryPolicy.ExecuteAsync(async token =>
            {
                EnsureNotDisposed();

                using (HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Get, relativePath))
                {
                    if (requiresToken)
                    {
                        AttachToken(request);
                    }

                    using (HttpResponseMessage response = await _httpClient.SendAsync(request, token).ConfigureAwait(false))
                    {
                        return await ParseResponse<TResponse>(relativePath, response).ConfigureAwait(false);
                    }
                }
            }, cancellationToken);
        }

        private Task<TResponse> PostJsonAsync<TRequest, TResponse>(
            string relativePath,
            TRequest payload,
            bool requiresToken,
            CancellationToken cancellationToken)
            where TResponse : class, new()
        {
            return _retryPolicy.ExecuteAsync(async token =>
            {
                EnsureNotDisposed();

                string json = JsonSerializer.Serialize(payload, JsonOptions.Default);

                using (HttpRequestMessage request = new HttpRequestMessage(HttpMethod.Post, relativePath))
                {
                    if (requiresToken)
                    {
                        AttachToken(request);
                    }

                    request.Content = new StringContent(json, Encoding.UTF8, "application/json");

                    using (HttpResponseMessage response = await _httpClient.SendAsync(request, token).ConfigureAwait(false))
                    {
                        return await ParseResponse<TResponse>(relativePath, response).ConfigureAwait(false);
                    }
                }
            }, cancellationToken);
        }

        private static object BuildPublishBody(
            PublishBatchRequest request,
            IReadOnlyList<PublishFileManifestItem> filesManifest)
        {
            List<object> itemRows = new List<object>();
            foreach (PublishSheetItem item in request.Items)
            {
                itemRows.Add(BuildPublishItemBody(item));
            }

            return new
            {
                run_client_id = request.RunClientId,
                project_code = request.ProjectCode,
                revit_version = request.RevitVersion,
                model_guid = request.ModelGuid,
                model_title = request.ModelTitle,
                plugin_version = request.PluginVersion,
                items = itemRows,
                files_manifest = filesManifest,
            };
        }

        private static object BuildPublishItemBody(PublishSheetItem item)
        {
            return new
            {
                item_index = item.ItemIndex,
                sheet_unique_id = item.SheetUniqueId ?? string.Empty,
                sheet_number = item.SheetNumber ?? string.Empty,
                sheet_name = item.SheetName ?? string.Empty,
                doc_number = item.DocNumber ?? string.Empty,
                requested_revision = item.RequestedRevision ?? string.Empty,
                status_code = item.StatusCode ?? string.Empty,
                include_native = item.IncludeNative,
                metadata = item.Metadata ?? new DocumentMetadata(),
                file_sha256 = item.FileSha256 ?? string.Empty,
            };
        }

        private static string TryAttachUpload(
            MultipartFormDataContent content,
            string sourcePath,
            int itemIndex,
            string kind,
            string mimeType)
        {
            if (string.IsNullOrWhiteSpace(sourcePath))
            {
                return string.Empty;
            }

            string normalizedPath = sourcePath.Trim();
            if (!File.Exists(normalizedPath))
            {
                return string.Empty;
            }

            string originalName = Path.GetFileName(normalizedPath);
            if (string.IsNullOrWhiteSpace(originalName))
            {
                return string.Empty;
            }

            string uploadName = BuildUploadFileName(itemIndex, kind, originalName);

            FileStream stream = new FileStream(normalizedPath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
            StreamContent streamContent = new StreamContent(stream);
            streamContent.Headers.ContentType = new MediaTypeHeaderValue(mimeType);
            content.Add(streamContent, "files", uploadName);

            return uploadName;
        }

        private static string BuildUploadFileName(int itemIndex, string kind, string originalName)
        {
            string safeKind = SanitizeToken(kind);
            string safeName = SanitizeToken(originalName);
            return string.Format("i{0}_{1}_{2}", itemIndex, safeKind, safeName);
        }

        private static string SanitizeToken(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return "file";
            }

            char[] input = value.Trim().ToCharArray();
            char[] output = new char[input.Length];
            int length = 0;

            for (int i = 0; i < input.Length; i++)
            {
                char c = input[i];
                if (char.IsLetterOrDigit(c) || c == '.' || c == '_' || c == '-')
                {
                    output[length++] = c;
                }
                else
                {
                    output[length++] = '_';
                }
            }

            string result = new string(output, 0, length);
            return string.IsNullOrWhiteSpace(result) ? "file" : result;
        }

        private static void AddFormField(MultipartFormDataContent content, string key, string value)
        {
            content.Add(new StringContent(value ?? string.Empty, Encoding.UTF8), key);
        }

        private static void AddOptionalFormField(MultipartFormDataContent content, string key, string value)
        {
            if (!string.IsNullOrWhiteSpace(value))
            {
                AddFormField(content, key, value);
            }
        }

        private static async Task<TResponse> ParseResponse<TResponse>(
            string relativePath,
            HttpResponseMessage response)
            where TResponse : class, new()
        {
            string body = response.Content == null
                ? string.Empty
                : await response.Content.ReadAsStringAsync().ConfigureAwait(false);

            if (!response.IsSuccessStatusCode)
            {
                throw new HttpRequestException($"Request to {relativePath} failed with {(int)response.StatusCode}: {body}");
            }

            if (string.IsNullOrWhiteSpace(body))
            {
                return new TResponse();
            }

            TResponse? parsed = JsonSerializer.Deserialize<TResponse>(body, JsonOptions.Default);
            return parsed ?? new TResponse();
        }

        private void AttachToken(HttpRequestMessage request)
        {
            if (!_tokenStore.HasValidToken())
            {
                throw new InvalidOperationException("No valid access token is available. Call LoginAsync first.");
            }

            request.Headers.Authorization = new AuthenticationHeaderValue("Bearer", _tokenStore.Get());
        }

        private static string BuildManifestPath(SiteLogManifestRequest request)
        {
            List<string> query = new List<string>
            {
                "project_code=" + Uri.EscapeDataString(request.ProjectCode ?? string.Empty),
                "limit=" + request.Limit,
                "client_model_guid=" + Uri.EscapeDataString(request.ClientModelGuid ?? string.Empty),
            };

            if (!string.IsNullOrWhiteSpace(request.DisciplineCode))
            {
                query.Add("discipline_code=" + Uri.EscapeDataString(request.DisciplineCode));
            }

            if (request.UpdatedAfterUtc.HasValue)
            {
                query.Add("updated_after=" + Uri.EscapeDataString(request.UpdatedAfterUtc.Value.UtcDateTime.ToString("o")));
            }

            return "/api/v1/bim/site-logs/revit/manifest?" + string.Join("&", query);
        }

        private void EnsureNotDisposed()
        {
            if (_disposed)
            {
                throw new ObjectDisposedException(nameof(ApiClient));
            }
        }

        private sealed class LoginResponseDto
        {
            public string AccessToken { get; set; } = string.Empty;

            public string Token { get; set; } = string.Empty;

            public int ExpiresIn { get; set; }
        }
    }
}
