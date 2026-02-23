using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Mdr.Revit.Core.Contracts;
using Mdr.Revit.Core.Models;
using Mdr.Revit.Core.UseCases;
using Xunit;

namespace Mdr.Revit.Core.Tests
{
    public sealed class PublishSheetsUseCaseTests
    {
        [Fact]
        public async Task ExecuteAsync_WhenRequestItemsEmpty_UsesExtractorItems()
        {
            FakeApiClient api = new FakeApiClient();
            FakeExtractor extractor = new FakeExtractor();
            extractor.SelectedSheets.Add(new PublishSheetItem
            {
                SheetUniqueId = "sheet-1",
                SheetNumber = "A-101",
                SheetName = "Plan",
                RequestedRevision = "A",
                FileSha256 = "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa",
            });

            PublishSheetsUseCase useCase = new PublishSheetsUseCase(api, extractor);
            PublishBatchRequest request = new PublishBatchRequest
            {
                ProjectCode = "PRJ-001",
            };

            PublishBatchResponse response = await useCase.ExecuteAsync(request, CancellationToken.None);

            Assert.Single(request.Items);
            Assert.Equal(1, api.PublishCallCount);
            Assert.Equal("run-publish", response.RunId);
        }

        [Fact]
        public async Task ExecuteAsync_WhenPdfPathProvided_ComputesSha256()
        {
            string tempPdf = Path.Combine(Path.GetTempPath(), "mdr_usecase_pdf_" + System.Guid.NewGuid().ToString("N") + ".pdf");
            File.WriteAllBytes(tempPdf, Encoding.ASCII.GetBytes("%PDF-1.4 test"));

            try
            {
                FakeApiClient api = new FakeApiClient();
                FakeExtractor extractor = new FakeExtractor();
                PublishSheetsUseCase useCase = new PublishSheetsUseCase(api, extractor);

                PublishBatchRequest request = new PublishBatchRequest
                {
                    ProjectCode = "PRJ-001",
                };
                request.Items.Add(new PublishSheetItem
                {
                    ItemIndex = 0,
                    SheetUniqueId = "sheet-1",
                    RequestedRevision = "A",
                    PdfFilePath = tempPdf,
                });

                await useCase.ExecuteAsync(request, CancellationToken.None);

                Assert.NotNull(api.LastRequest);
                Assert.False(string.IsNullOrWhiteSpace(api.LastRequest!.Items[0].FileSha256));
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
        public async Task RetryFailedAsync_WhenPreviousRunHasFailures_OnlyRetriesFailedItems()
        {
            FakeApiClient api = new FakeApiClient();
            FakeExtractor extractor = new FakeExtractor();
            PublishSheetsUseCase useCase = new PublishSheetsUseCase(api, extractor);

            PublishBatchRequest source = new PublishBatchRequest
            {
                ProjectCode = "PRJ-001",
                ModelGuid = "model-1",
                ModelTitle = "Model",
                PluginVersion = "0.1.0",
            };
            source.Items.Add(new PublishSheetItem
            {
                ItemIndex = 0,
                SheetUniqueId = "sheet-0",
                RequestedRevision = "A",
                FileSha256 = "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa",
            });
            source.Items.Add(new PublishSheetItem
            {
                ItemIndex = 1,
                SheetUniqueId = "sheet-1",
                RequestedRevision = "A",
                FileSha256 = "bbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbbb",
            });

            PublishBatchResponse previous = new PublishBatchResponse
            {
                RunId = "run-1",
            };
            previous.Items.Add(new PublishItemResult { ItemIndex = 0, State = "failed" });
            previous.Items.Add(new PublishItemResult { ItemIndex = 1, State = "completed" });

            PublishBatchResponse retry = await useCase.RetryFailedAsync(source, previous, CancellationToken.None);

            Assert.Equal(1, api.PublishCallCount);
            Assert.NotNull(api.LastRequest);
            Assert.Single(api.LastRequest!.Items);
            Assert.Equal(0, api.LastRequest.Items[0].ItemIndex);
            Assert.Equal("sheet-0", api.LastRequest.Items[0].SheetUniqueId);
            Assert.Equal("run-publish", retry.RunId);
        }

        private sealed class FakeExtractor : IRevitExtractor
        {
            public List<PublishSheetItem> SelectedSheets { get; } = new List<PublishSheetItem>();

            public IReadOnlyList<PublishSheetItem> GetSelectedSheets()
            {
                return SelectedSheets;
            }

            public IReadOnlyList<ScheduleRow> GetScheduleRows(string profileCode)
            {
                return new List<ScheduleRow>();
            }
        }

        private sealed class FakeApiClient : IApiClient
        {
            public int PublishCallCount { get; private set; }

            public PublishBatchRequest? LastRequest { get; private set; }

            public Task<string> LoginAsync(string username, string password, CancellationToken cancellationToken)
            {
                return Task.FromResult("token");
            }

            public Task<PublishBatchResponse> PublishBatchAsync(PublishBatchRequest request, CancellationToken cancellationToken)
            {
                PublishCallCount++;
                LastRequest = request;
                return Task.FromResult(new PublishBatchResponse
                {
                    RunId = "run-publish",
                });
            }

            public Task<ScheduleIngestResponse> IngestScheduleAsync(ScheduleIngestRequest request, CancellationToken cancellationToken)
            {
                return Task.FromResult(new ScheduleIngestResponse());
            }

            public Task<SiteLogManifestResponse> GetSiteLogManifestAsync(SiteLogManifestRequest request, CancellationToken cancellationToken)
            {
                return Task.FromResult(new SiteLogManifestResponse());
            }

            public Task<SiteLogPullResponse> PullSiteLogRowsAsync(SiteLogPullRequest request, CancellationToken cancellationToken)
            {
                return Task.FromResult(new SiteLogPullResponse());
            }

            public Task<SiteLogAckResponse> AckSiteLogSyncAsync(SiteLogAckRequest request, CancellationToken cancellationToken)
            {
                return Task.FromResult(new SiteLogAckResponse());
            }
        }
    }
}
