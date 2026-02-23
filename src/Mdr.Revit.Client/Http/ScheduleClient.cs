using System;
using System.Threading;
using System.Threading.Tasks;
using Mdr.Revit.Core.Contracts;
using Mdr.Revit.Core.Models;

namespace Mdr.Revit.Client.Http
{
    public sealed class ScheduleClient
    {
        private readonly IApiClient _apiClient;

        public ScheduleClient(IApiClient apiClient)
        {
            _apiClient = apiClient ?? throw new ArgumentNullException(nameof(apiClient));
        }

        public Task<ScheduleIngestResponse> IngestAsync(
            ScheduleIngestRequest request,
            CancellationToken cancellationToken)
        {
            return _apiClient.IngestScheduleAsync(request, cancellationToken);
        }
    }
}
