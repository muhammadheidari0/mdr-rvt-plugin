using System.Threading;
using System.Threading.Tasks;
using Mdr.Revit.Core.Contracts;
using Mdr.Revit.Core.Models;
using Mdr.Revit.Core.UseCases;
using Xunit;

namespace Mdr.Revit.Core.Tests
{
    public sealed class SyncSiteLogsUseCaseTests
    {
        [Fact]
        public async Task ExecuteAsync_WhenManifestHasRows_AcksApplyResult()
        {
            FakeApiClient api = new FakeApiClient();
            api.Manifest.Changes.Add(new SiteLogManifestChange
            {
                LogId = 101,
            });
            api.PullResponse.ManpowerRows.Add(new SiteLogRow
            {
                SyncKey = "site_log:101:MANPOWER:1",
                LogId = 101,
            });

            FakeWriter writer = new FakeWriter();
            SyncSiteLogsUseCase useCase = new SyncSiteLogsUseCase(api, writer);

            SiteLogApplyResult result = await useCase.ExecuteAsync(
                new SiteLogManifestRequest
                {
                    ProjectCode = "PRJ-001",
                    ClientModelGuid = "model-1",
                },
                pluginVersion: "0.1.0",
                cancellationToken: CancellationToken.None);

            Assert.Equal(1, result.AppliedCount);
            Assert.Equal("run-sync", result.RunId);
            Assert.Equal(1, api.AckCount);
        }

        private sealed class FakeWriter : IRevitWriter
        {
            public SiteLogApplyResult ApplySiteLogRows(SiteLogPullResponse pullResponse)
            {
                return new SiteLogApplyResult
                {
                    AppliedCount = pullResponse.ManpowerRows.Count,
                    FailedCount = 0,
                };
            }
        }

        private sealed class FakeApiClient : IApiClient
        {
            public SiteLogManifestResponse Manifest { get; } = new SiteLogManifestResponse { RunId = "run-sync" };

            public SiteLogPullResponse PullResponse { get; } = new SiteLogPullResponse();

            public int AckCount { get; private set; }

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
                return Task.FromResult(new ScheduleIngestResponse());
            }

            public Task<SiteLogManifestResponse> GetSiteLogManifestAsync(SiteLogManifestRequest request, CancellationToken cancellationToken)
            {
                return Task.FromResult(Manifest);
            }

            public Task<SiteLogPullResponse> PullSiteLogRowsAsync(SiteLogPullRequest request, CancellationToken cancellationToken)
            {
                PullResponse.RunId = "run-sync";
                return Task.FromResult(PullResponse);
            }

            public Task<SiteLogAckResponse> AckSiteLogSyncAsync(SiteLogAckRequest request, CancellationToken cancellationToken)
            {
                AckCount++;
                return Task.FromResult(new SiteLogAckResponse { Ok = true });
            }
        }
    }
}
