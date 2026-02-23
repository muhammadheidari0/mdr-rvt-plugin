using System;
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
    public sealed class GoogleSheetsClientTests
    {
        [Fact]
        public async Task ReadRowsAsync_ParsesRowsAndAnchor()
        {
            StubGoogleTokenProvider tokenProvider = new StubGoogleTokenProvider();
            StubHttpMessageHandler handler = new StubHttpMessageHandler(_ =>
            {
                HttpResponseMessage response = new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new StringContent(
                        "{\"values\":[[\"MDR_UNIQUE_ID\",\"Length\"],[\"uid-1\",\"12.5\"]]}",
                        Encoding.UTF8,
                        "application/json"),
                };
                return Task.FromResult(response);
            });

            HttpClient http = new HttpClient(handler) { BaseAddress = new Uri("https://sheets.googleapis.com") };
            GoogleSheetsClient client = new GoogleSheetsClient(http, tokenProvider, new RetryPolicy(maxAttempts: 1));
            GoogleSheetReadResult result = await client.ReadRowsAsync(
                new GoogleSheetSyncProfile
                {
                    SpreadsheetId = "sheet-id",
                    WorksheetName = "Sheet1",
                },
                CancellationToken.None);

            Assert.Single(result.Rows);
            Assert.Equal("uid-1", result.Rows[0].AnchorUniqueId);
            Assert.Equal("12.5", result.Rows[0].Cells["Length"]);
        }

        [Fact]
        public async Task WriteRowsAsync_SendsClearAndUpdateRequests()
        {
            int requestCount = 0;
            StubGoogleTokenProvider tokenProvider = new StubGoogleTokenProvider();
            StubHttpMessageHandler handler = new StubHttpMessageHandler(async request =>
            {
                requestCount++;
                string path = request.RequestUri?.AbsolutePath ?? string.Empty;
                if (path.Contains(":clear"))
                {
                    return new HttpResponseMessage(HttpStatusCode.OK)
                    {
                        Content = new StringContent("{}", Encoding.UTF8, "application/json"),
                    };
                }

                string payload = request.Content == null
                    ? string.Empty
                    : await request.Content.ReadAsStringAsync().ConfigureAwait(false);
                Assert.Contains("MDR_UNIQUE_ID", payload);
                Assert.Contains("uid-1", payload);

                return new HttpResponseMessage(HttpStatusCode.OK)
                {
                    Content = new StringContent(
                        "{\"updatedRows\":2,\"updatedRange\":\"Sheet1!A1:B2\"}",
                        Encoding.UTF8,
                        "application/json"),
                };
            });

            HttpClient http = new HttpClient(handler) { BaseAddress = new Uri("https://sheets.googleapis.com") };
            GoogleSheetsClient client = new GoogleSheetsClient(http, tokenProvider, new RetryPolicy(maxAttempts: 1));

            ScheduleSyncRow row = new ScheduleSyncRow
            {
                AnchorUniqueId = "uid-1",
            };
            row.Cells["Length"] = "12.5";

            GoogleSheetWriteResult result = await client.WriteRowsAsync(
                new GoogleSheetSyncProfile
                {
                    SpreadsheetId = "sheet-id",
                    WorksheetName = "Sheet1",
                },
                new[] { row },
                CancellationToken.None);

            Assert.Equal(2, requestCount);
            Assert.Equal(1, result.UpdatedRows);
        }

        private sealed class StubGoogleTokenProvider : IGoogleTokenProvider
        {
            public Task<string> GetAccessTokenAsync(CancellationToken cancellationToken)
            {
                _ = cancellationToken;
                return Task.FromResult("google-token");
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
