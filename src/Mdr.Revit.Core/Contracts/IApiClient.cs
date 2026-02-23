using System.Threading;
using System.Threading.Tasks;
using Mdr.Revit.Core.Models;

namespace Mdr.Revit.Core.Contracts
{
    public interface IApiClient
    {
        Task<string> LoginAsync(string username, string password, CancellationToken cancellationToken);

        Task<PublishBatchResponse> PublishBatchAsync(PublishBatchRequest request, CancellationToken cancellationToken);

        Task<ScheduleIngestResponse> IngestScheduleAsync(ScheduleIngestRequest request, CancellationToken cancellationToken);

        Task<SiteLogManifestResponse> GetSiteLogManifestAsync(
            SiteLogManifestRequest request,
            CancellationToken cancellationToken);

        Task<SiteLogPullResponse> PullSiteLogRowsAsync(
            SiteLogPullRequest request,
            CancellationToken cancellationToken);

        Task<SiteLogAckResponse> AckSiteLogSyncAsync(
            SiteLogAckRequest request,
            CancellationToken cancellationToken);
    }
}
