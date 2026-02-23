using System.Threading;
using System.Threading.Tasks;
using Mdr.Revit.Client.Http;
using Mdr.Revit.Core.Contracts;
using Mdr.Revit.Core.Models;
using Xunit;

namespace Mdr.Revit.Client.Tests
{
    public sealed class SiteLogsSyncClientTests
    {
        [Fact]
        public async Task AckAsync_DelegatesToApiClient()
        {
            FakeApiClient api = new FakeApiClient();
            SiteLogsSyncClient client = new SiteLogsSyncClient(api);

            SiteLogAckResponse response = await client.AckAsync(
                new SiteLogAckRequest { RunId = "run-1", AppliedCount = 2 },
                CancellationToken.None);

            Assert.True(response.Ok);
            Assert.Equal(1, api.AckCalls);
        }

        private sealed class FakeApiClient : IApiClient
        {
            public int AckCalls { get; private set; }

            public Task<string> LoginAsync(string username, string password, CancellationToken cancellationToken)
                => Task.FromResult("token");

            public Task<PublishBatchResponse> PublishBatchAsync(PublishBatchRequest request, CancellationToken cancellationToken)
                => Task.FromResult(new PublishBatchResponse());

            public Task<ScheduleIngestResponse> IngestScheduleAsync(ScheduleIngestRequest request, CancellationToken cancellationToken)
                => Task.FromResult(new ScheduleIngestResponse());

            public Task<SiteLogManifestResponse> GetSiteLogManifestAsync(SiteLogManifestRequest request, CancellationToken cancellationToken)
                => Task.FromResult(new SiteLogManifestResponse());

            public Task<SiteLogPullResponse> PullSiteLogRowsAsync(SiteLogPullRequest request, CancellationToken cancellationToken)
                => Task.FromResult(new SiteLogPullResponse());

            public Task<SiteLogAckResponse> AckSiteLogSyncAsync(SiteLogAckRequest request, CancellationToken cancellationToken)
            {
                AckCalls++;
                return Task.FromResult(new SiteLogAckResponse { Ok = true });
            }
        }
    }
}
