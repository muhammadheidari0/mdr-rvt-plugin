using System;
using System.Collections.Generic;
using System.Net.Http;
using System.Threading;
using System.Threading.Tasks;
using Mdr.Revit.Addin.Commands;
using Mdr.Revit.Core.Contracts;
using Mdr.Revit.Core.Models;
using Mdr.Revit.Infra.Logging;
using Mdr.Revit.RevitAdapter.Extractors;
using Xunit;

namespace Mdr.Revit.Addin.Tests
{
    public sealed class PublishSheetsCommandTests
    {
        [Fact]
        public async Task ExecuteAsync_WhenLoginFails_DoesNotRunExport()
        {
            bool pdfCalled = false;
            bool nativeCalled = false;

            FakeApiClient apiClient = new FakeApiClient
            {
                LoginException = new HttpRequestException("Request to /api/v1/auth/login failed with 401"),
            };

            PublishSheetsCommand command = new PublishSheetsCommand(
                _ => apiClient,
                new FakeExtractor(),
                new PdfExporter((items, output) =>
                {
                    _ = items;
                    _ = output;
                    pdfCalled = true;
                    return Array.Empty<ExportArtifact>();
                }),
                new NativeExporter((items, output) =>
                {
                    _ = items;
                    _ = output;
                    nativeCalled = true;
                    return Array.Empty<ExportArtifact>();
                }),
                new PluginLogger(System.IO.Path.GetTempPath()));

            PublishSheetsCommandRequest request = new PublishSheetsCommandRequest
            {
                BaseUrl = "http://127.0.0.1:8000",
                Username = "user",
                Password = "bad",
                ProjectCode = "PRJ-001",
            };
            request.Items.Add(new PublishSheetItem
            {
                ItemIndex = 0,
                SheetUniqueId = "sheet-1",
                SheetNumber = "A-101",
                RequestedRevision = "A",
            });

            InvalidOperationException ex = await Assert.ThrowsAsync<InvalidOperationException>(() =>
                command.ExecuteAsync(request, CancellationToken.None));

            Assert.Contains("auth_failed", ex.Message);
            Assert.False(pdfCalled);
            Assert.False(nativeCalled);
        }

        [Fact]
        public async Task ExecuteAsync_WhenOneExportFails_PartialContinueIsApplied()
        {
            FakeApiClient apiClient = new FakeApiClient();

            PublishSheetsCommand command = new PublishSheetsCommand(
                _ => apiClient,
                new FakeExtractor(),
                new PdfExporter((items, output) =>
                {
                    _ = output;
                    List<ExportArtifact> artifacts = new List<ExportArtifact>();
                    for (int i = 0; i < items.Count; i++)
                    {
                        PublishSheetItem item = items[i];
                        if (item.ItemIndex == 1)
                        {
                            artifacts.Add(new ExportArtifact
                            {
                                ItemIndex = item.ItemIndex,
                                SheetUniqueId = item.SheetUniqueId,
                                Kind = ExportArtifactKinds.Pdf,
                                ErrorCode = "export_pdf_failed",
                                ErrorMessage = "simulated export failure",
                            });
                        }
                        else
                        {
                            string filePath = System.IO.Path.Combine(
                                System.IO.Path.GetTempPath(),
                                "publish-test-" + item.ItemIndex + ".pdf");
                            System.IO.File.WriteAllText(filePath, "pdf");
                            artifacts.Add(new ExportArtifact
                            {
                                ItemIndex = item.ItemIndex,
                                SheetUniqueId = item.SheetUniqueId,
                                Kind = ExportArtifactKinds.Pdf,
                                FilePath = filePath,
                            });
                        }
                    }

                    return artifacts;
                }),
                new NativeExporter((items, output) =>
                {
                    _ = items;
                    _ = output;
                    return Array.Empty<ExportArtifact>();
                }),
                new PluginLogger(System.IO.Path.GetTempPath()));

            PublishSheetsCommandRequest request = new PublishSheetsCommandRequest
            {
                BaseUrl = "http://127.0.0.1:8000",
                Username = "user",
                Password = "pass",
                ProjectCode = "PRJ-001",
                IncludeNative = false,
                RetryFailedItems = false,
            };
            request.Items.Add(new PublishSheetItem
            {
                ItemIndex = 0,
                SheetUniqueId = "sheet-1",
                SheetNumber = "A-101",
                RequestedRevision = "A",
            });
            request.Items.Add(new PublishSheetItem
            {
                ItemIndex = 1,
                SheetUniqueId = "sheet-2",
                SheetNumber = "A-102",
                RequestedRevision = "A",
            });

            PublishSheetsCommandResult result = await command.ExecuteAsync(request, CancellationToken.None);

            Assert.NotNull(apiClient.LastPublishRequest);
            Assert.Single(apiClient.LastPublishRequest!.Items);
            Assert.Equal(0, apiClient.LastPublishRequest.Items[0].ItemIndex);

            Assert.Equal(2, result.FinalResponse.Summary.RequestedCount);
            Assert.Equal(1, result.FinalResponse.Summary.SuccessCount);
            Assert.Equal(1, result.FinalResponse.Summary.FailedCount);
            Assert.Equal("completed_with_errors", result.FinalResponse.Summary.Status);
            Assert.Contains(result.FinalResponse.Items, x => x.ItemIndex == 1 && x.ErrorCode == "export_pdf_failed");
        }

        private sealed class FakeExtractor : IRevitExtractor
        {
            public IReadOnlyList<PublishSheetItem> GetSelectedSheets()
            {
                return Array.Empty<PublishSheetItem>();
            }

            public IReadOnlyList<ScheduleRow> GetScheduleRows(string profileCode)
            {
                _ = profileCode;
                return Array.Empty<ScheduleRow>();
            }
        }

        private sealed class FakeApiClient : IApiClient
        {
            public Exception? LoginException { get; set; }

            public PublishBatchRequest? LastPublishRequest { get; private set; }

            public Task<string> LoginAsync(string username, string password, CancellationToken cancellationToken)
            {
                _ = username;
                _ = password;
                _ = cancellationToken;
                if (LoginException != null)
                {
                    throw LoginException;
                }

                return Task.FromResult("token");
            }

            public Task<PublishBatchResponse> PublishBatchAsync(PublishBatchRequest request, CancellationToken cancellationToken)
            {
                _ = cancellationToken;
                LastPublishRequest = request;
                PublishBatchResponse response = new PublishBatchResponse
                {
                    RunId = "run-server",
                    Summary = new PublishBatchSummary
                    {
                        RequestedCount = request.Items.Count,
                        SuccessCount = request.Items.Count,
                        FailedCount = 0,
                        DuplicateCount = 0,
                        Status = "completed",
                    },
                };

                for (int i = 0; i < request.Items.Count; i++)
                {
                    response.Items.Add(new PublishItemResult
                    {
                        ItemIndex = request.Items[i].ItemIndex,
                        State = "completed",
                    });
                }

                return Task.FromResult(response);
            }

            public Task<ScheduleIngestResponse> IngestScheduleAsync(ScheduleIngestRequest request, CancellationToken cancellationToken)
            {
                _ = request;
                _ = cancellationToken;
                return Task.FromResult(new ScheduleIngestResponse());
            }

            public Task<SiteLogManifestResponse> GetSiteLogManifestAsync(SiteLogManifestRequest request, CancellationToken cancellationToken)
            {
                _ = request;
                _ = cancellationToken;
                return Task.FromResult(new SiteLogManifestResponse());
            }

            public Task<SiteLogPullResponse> PullSiteLogRowsAsync(SiteLogPullRequest request, CancellationToken cancellationToken)
            {
                _ = request;
                _ = cancellationToken;
                return Task.FromResult(new SiteLogPullResponse());
            }

            public Task<SiteLogAckResponse> AckSiteLogSyncAsync(SiteLogAckRequest request, CancellationToken cancellationToken)
            {
                _ = request;
                _ = cancellationToken;
                return Task.FromResult(new SiteLogAckResponse());
            }
        }
    }
}
