using System.Threading;
using System.Threading.Tasks;
using Mdr.Revit.Client.Http;
using Mdr.Revit.Core.Contracts;
using Mdr.Revit.Core.Models;
using Xunit;

namespace Mdr.Revit.Client.Tests
{
    public sealed class EdmsClientTests
    {
        [Fact]
        public async Task PublishBatchAsync_DelegatesToApiClient()
        {
            FakeApiClient api = new FakeApiClient();
            EdmsClient client = new EdmsClient(api);

            PublishBatchResponse response = await client.PublishBatchAsync(
                new PublishBatchRequest { ProjectCode = "PRJ-001" },
                CancellationToken.None);

            Assert.Equal(1, api.PublishCalls);
            Assert.Equal("run-edms", response.RunId);
        }

        private sealed class FakeApiClient : IApiClient
        {
            public int PublishCalls { get; private set; }

            public Task<string> LoginAsync(string username, string password, CancellationToken cancellationToken)
                => Task.FromResult("token");

            public Task<PublishBatchResponse> PublishBatchAsync(PublishBatchRequest request, CancellationToken cancellationToken)
            {
                PublishCalls++;
                return Task.FromResult(new PublishBatchResponse { RunId = "run-edms" });
            }

            public Task<ScheduleIngestResponse> IngestScheduleAsync(ScheduleIngestRequest request, CancellationToken cancellationToken)
                => Task.FromResult(new ScheduleIngestResponse());

            public Task<SiteLogManifestResponse> GetSiteLogManifestAsync(SiteLogManifestRequest request, CancellationToken cancellationToken)
                => Task.FromResult(new SiteLogManifestResponse());

            public Task<SiteLogPullResponse> PullSiteLogRowsAsync(SiteLogPullRequest request, CancellationToken cancellationToken)
                => Task.FromResult(new SiteLogPullResponse());

            public Task<SiteLogAckResponse> AckSiteLogSyncAsync(SiteLogAckRequest request, CancellationToken cancellationToken)
                => Task.FromResult(new SiteLogAckResponse());
        }
    }
}
