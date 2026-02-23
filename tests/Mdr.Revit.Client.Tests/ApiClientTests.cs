using System;
using System.IO;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Mdr.Revit.Client.Auth;
using Mdr.Revit.Client.Http;
using Mdr.Revit.Client.Retry;
using Mdr.Revit.Core.Models;
using Xunit;

namespace Mdr.Revit.Client.Tests
{
    public sealed class ApiClientTests
    {
        [Fact]
        public async Task LoginAsync_SendsFormPayload_AndStoresToken()
        {
            string? contentType = null;
            string body = string.Empty;

            StubHttpMessageHandler handler = new StubHttpMessageHandler(async request =>
            {
                contentType = request.Content?.Headers?.ContentType?.MediaType;
                body = request.Content == null
                    ? string.Empty
                    : await request.Content.ReadAsStringAsync().ConfigureAwait(false);

                Assert.Equal(HttpMethod.Post, request.Method);
                Assert.Equal("/api/v1/auth/login", request.RequestUri?.AbsolutePath);

                return JsonResponse("{\"access_token\":\"jwt-token\",\"token_type\":\"bearer\"}");
            });

            HttpClient httpClient = new HttpClient(handler)
            {
                BaseAddress = new Uri("http://localhost:8000"),
            };

            TokenStore tokenStore = new TokenStore();
            ApiClient client = new ApiClient(httpClient, tokenStore, new RetryPolicy(maxAttempts: 1));

            string token = await client.LoginAsync("user@example.com", "secret", CancellationToken.None);

            Assert.Equal("jwt-token", token);
            Assert.Equal("application/x-www-form-urlencoded", contentType);
            Assert.Contains("username=user%40example.com", body);
            Assert.Contains("password=secret", body);
            Assert.True(tokenStore.HasValidToken());
            Assert.Equal("jwt-token", tokenStore.Get());
        }

        [Fact]
        public async Task PublishBatchAsync_WhenLocalPdfExists_UsesMultipartFormData()
        {
            string capturedContentType = string.Empty;
            string capturedBody = string.Empty;
            string? capturedAuth = null;

            StubHttpMessageHandler handler = new StubHttpMessageHandler(async request =>
            {
                capturedContentType = request.Content?.Headers?.ContentType?.MediaType ?? string.Empty;
                capturedBody = request.Content == null
                    ? string.Empty
                    : await request.Content.ReadAsStringAsync().ConfigureAwait(false);
                capturedAuth = request.Headers.Authorization?.ToString();

                return JsonResponse(
                    "{\"run_id\":\"run-1\",\"summary\":{\"requested_count\":1,\"success_count\":1,\"failed_count\":0,\"duplicate_count\":0,\"status\":\"completed\"},\"items\":[{\"item_index\":0,\"state\":\"completed\"}]}");
            });

            HttpClient httpClient = new HttpClient(handler)
            {
                BaseAddress = new Uri("http://localhost:8000"),
            };

            TokenStore tokenStore = new TokenStore();
            tokenStore.Set("jwt-token", DateTimeOffset.UtcNow.AddMinutes(10));
            ApiClient client = new ApiClient(httpClient, tokenStore, new RetryPolicy(maxAttempts: 1));

            string tempPdf = Path.Combine(Path.GetTempPath(), "mdr_publish_test.pdf");
            File.WriteAllBytes(tempPdf, Encoding.UTF8.GetBytes("pdf-binary-content"));

            try
            {
                PublishBatchRequest request = new PublishBatchRequest
                {
                    ProjectCode = "PRJ-001",
                    ModelGuid = "model-1",
                    ModelTitle = "Model",
                    PluginVersion = "0.1.0",
                };
                request.Items.Add(new PublishSheetItem
                {
                    ItemIndex = 0,
                    SheetUniqueId = "sheet-1",
                    SheetNumber = "A-101",
                    SheetName = "Plan",
                    RequestedRevision = "A",
                    StatusCode = "IFA",
                    PdfFilePath = tempPdf,
                });

                PublishBatchResponse response = await client.PublishBatchAsync(request, CancellationToken.None);

                Assert.Equal("multipart/form-data", capturedContentType);
                Assert.True(
                    capturedBody.Contains("name=\"files\"") || capturedBody.Contains("name=files"),
                    "Multipart body must contain files[] part.");
                Assert.True(
                    capturedBody.Contains("name=\"items_json\"") || capturedBody.Contains("name=items_json"),
                    "Multipart body must contain items_json part.");
                Assert.True(
                    capturedBody.Contains("name=\"files_manifest\"") || capturedBody.Contains("name=files_manifest"),
                    "Multipart body must contain files_manifest part.");
                Assert.Contains("sheet_unique_id", capturedBody);
                Assert.Equal("Bearer jwt-token", capturedAuth);
                Assert.Equal("run-1", response.RunId);
                Assert.Equal("completed", response.Summary.Status);
            }
            finally
            {
                if (File.Exists(tempPdf))
                {
                    File.Delete(tempPdf);
                }
            }
        }

        [Fact]
        public async Task PublishBatchAsync_WhenNoLocalFile_UsesJsonPayload()
        {
            string capturedContentType = string.Empty;
            string capturedBody = string.Empty;

            StubHttpMessageHandler handler = new StubHttpMessageHandler(async request =>
            {
                capturedContentType = request.Content?.Headers?.ContentType?.MediaType ?? string.Empty;
                capturedBody = request.Content == null
                    ? string.Empty
                    : await request.Content.ReadAsStringAsync().ConfigureAwait(false);

                return JsonResponse(
                    "{\"run_id\":\"run-2\",\"summary\":{\"requested_count\":1,\"success_count\":1,\"failed_count\":0,\"duplicate_count\":0,\"status\":\"completed\"},\"items\":[{\"item_index\":0,\"state\":\"completed\"}]}");
            });

            HttpClient httpClient = new HttpClient(handler)
            {
                BaseAddress = new Uri("http://localhost:8000"),
            };

            TokenStore tokenStore = new TokenStore();
            tokenStore.Set("jwt-token", DateTimeOffset.UtcNow.AddMinutes(10));
            ApiClient client = new ApiClient(httpClient, tokenStore, new RetryPolicy(maxAttempts: 1));

            PublishBatchRequest request = new PublishBatchRequest
            {
                ProjectCode = "PRJ-001",
            };
            request.Items.Add(new PublishSheetItem
            {
                ItemIndex = 0,
                SheetUniqueId = "sheet-1",
                RequestedRevision = "A",
                FileSha256 = "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa",
            });

            PublishBatchResponse response = await client.PublishBatchAsync(request, CancellationToken.None);

            Assert.Equal("application/json", capturedContentType);
            Assert.Contains("file_sha256", capturedBody);
            Assert.Contains("sheet_unique_id", capturedBody);
            Assert.Equal("run-2", response.RunId);
        }

        private static HttpResponseMessage JsonResponse(string body)
        {
            return new HttpResponseMessage(HttpStatusCode.OK)
            {
                Content = new StringContent(body, Encoding.UTF8, "application/json"),
            };
        }

        private sealed class StubHttpMessageHandler : HttpMessageHandler
        {
            private readonly Func<HttpRequestMessage, Task<HttpResponseMessage>> _handler;

            public StubHttpMessageHandler(Func<HttpRequestMessage, Task<HttpResponseMessage>> handler)
            {
                _handler = handler ?? throw new ArgumentNullException(nameof(handler));
            }

            protected override Task<HttpResponseMessage> SendAsync(
                HttpRequestMessage request,
                CancellationToken cancellationToken)
            {
                return _handler(request);
            }
        }
    }
}
