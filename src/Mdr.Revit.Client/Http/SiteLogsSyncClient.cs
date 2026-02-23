using System;
using System.Threading;
using System.Threading.Tasks;
using Mdr.Revit.Core.Contracts;
using Mdr.Revit.Core.Models;

namespace Mdr.Revit.Client.Http
{
    public sealed class SiteLogsSyncClient
    {
        private readonly IApiClient _apiClient;

        public SiteLogsSyncClient(IApiClient apiClient)
        {
            _apiClient = apiClient ?? throw new ArgumentNullException(nameof(apiClient));
        }

        public Task<SiteLogManifestResponse> GetManifestAsync(
            SiteLogManifestRequest request,
            CancellationToken cancellationToken)
        {
            return _apiClient.GetSiteLogManifestAsync(request, cancellationToken);
        }

        public Task<SiteLogPullResponse> PullAsync(
            SiteLogPullRequest request,
            CancellationToken cancellationToken)
        {
            return _apiClient.PullSiteLogRowsAsync(request, cancellationToken);
        }

        public Task<SiteLogAckResponse> AckAsync(
            SiteLogAckRequest request,
            CancellationToken cancellationToken)
        {
            return _apiClient.AckSiteLogSyncAsync(request, cancellationToken);
        }
    }
}
