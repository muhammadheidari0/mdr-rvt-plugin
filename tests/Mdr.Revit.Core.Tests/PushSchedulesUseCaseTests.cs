using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Mdr.Revit.Core.Contracts;
using Mdr.Revit.Core.Models;
using Mdr.Revit.Core.UseCases;
using Xunit;

namespace Mdr.Revit.Core.Tests
{
    public sealed class PushSchedulesUseCaseTests
    {
        [Fact]
        public async Task ExecuteAsync_WhenRequestRowsEmpty_UsesExtractorRows()
        {
            FakeApiClient api = new FakeApiClient();
            FakeExtractor extractor = new FakeExtractor();
            extractor.ScheduleRows.Add(new ScheduleRow
            {
                RowNo = 1,
                ElementKey = "EL-001",
            });

            PushSchedulesUseCase useCase = new PushSchedulesUseCase(api, extractor);
            ScheduleIngestRequest request = new ScheduleIngestRequest
            {
                ProjectCode = "PRJ-001",
                ProfileCode = ScheduleProfiles.Mto,
            };

            ScheduleIngestResponse response = await useCase.ExecuteAsync(request, CancellationToken.None);

            Assert.Single(request.Rows);
            Assert.Equal(1, api.IngestCallCount);
            Assert.Equal("run-schedule", response.RunId);
        }

        private sealed class FakeExtractor : IRevitExtractor
        {
            public List<ScheduleRow> ScheduleRows { get; } = new List<ScheduleRow>();

            public IReadOnlyList<PublishSheetItem> GetSelectedSheets()
            {
                return new List<PublishSheetItem>();
            }

            public IReadOnlyList<ScheduleRow> GetScheduleRows(string profileCode)
            {
                return ScheduleRows;
            }
        }

        private sealed class FakeApiClient : IApiClient
        {
            public int IngestCallCount { get; private set; }

            public Task<string> LoginAsync(string username, string password, CancellationToken cancellationToken)
            {
                return Task.FromResult("token");
            }

            public Task<PublishBatchResponse> PublishBatchAsync(PublishBatchRequest request, CancellationToken cancellationToken)
            {
                return Task.FromResult(new PublishBatchResponse());
            }

            public Task<ScheduleIngestResponse> IngestScheduleAsync(ScheduleIngestRequest request, CancellationToken cancellationToken)
            {
                IngestCallCount++;
                return Task.FromResult(new ScheduleIngestResponse
                {
                    RunId = "run-schedule",
                });
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
