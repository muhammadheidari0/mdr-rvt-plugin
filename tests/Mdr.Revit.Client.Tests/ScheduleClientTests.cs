using System.Threading;
using System.Threading.Tasks;
using Mdr.Revit.Client.Http;
using Mdr.Revit.Core.Contracts;
using Mdr.Revit.Core.Models;
using Xunit;

namespace Mdr.Revit.Client.Tests
{
    public sealed class ScheduleClientTests
    {
        [Fact]
        public async Task IngestAsync_DelegatesToApiClient()
        {
            FakeApiClient api = new FakeApiClient();
            ScheduleClient client = new ScheduleClient(api);

            ScheduleIngestResponse response = await client.IngestAsync(
                new ScheduleIngestRequest { ProjectCode = "PRJ-001", ProfileCode = ScheduleProfiles.Mto },
                CancellationToken.None);

            Assert.Equal(1, api.IngestCalls);
            Assert.Equal("run-schedule", response.RunId);
        }

        private sealed class FakeApiClient : IApiClient
        {
            public int IngestCalls { get; private set; }

            public Task<string> LoginAsync(string username, string password, CancellationToken cancellationToken)
                => Task.FromResult("token");

            public Task<PublishBatchResponse> PublishBatchAsync(PublishBatchRequest request, CancellationToken cancellationToken)
                => Task.FromResult(new PublishBatchResponse());

            public Task<ScheduleIngestResponse> IngestScheduleAsync(ScheduleIngestRequest request, CancellationToken cancellationToken)
            {
                IngestCalls++;
                return Task.FromResult(new ScheduleIngestResponse { RunId = "run-schedule" });
            }

            public Task<SiteLogManifestResponse> GetSiteLogManifestAsync(SiteLogManifestRequest request, CancellationToken cancellationToken)
                => Task.FromResult(new SiteLogManifestResponse());

            public Task<SiteLogPullResponse> PullSiteLogRowsAsync(SiteLogPullRequest request, CancellationToken cancellationToken)
                => Task.FromResult(new SiteLogPullResponse());

            public Task<SiteLogAckResponse> AckSiteLogSyncAsync(SiteLogAckRequest request, CancellationToken cancellationToken)
                => Task.FromResult(new SiteLogAckResponse());
        }
    }
}
