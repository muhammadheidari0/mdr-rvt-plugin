using System;
using System.Threading;
using System.Threading.Tasks;
using Mdr.Revit.Core.Contracts;
using Mdr.Revit.Core.Models;

namespace Mdr.Revit.Client.Http
{
    public sealed class EdmsClient
    {
        private readonly IApiClient _apiClient;

        public EdmsClient(IApiClient apiClient)
        {
            _apiClient = apiClient ?? throw new ArgumentNullException(nameof(apiClient));
        }

        public Task<PublishBatchResponse> PublishBatchAsync(
            PublishBatchRequest request,
            CancellationToken cancellationToken)
        {
            return _apiClient.PublishBatchAsync(request, cancellationToken);
        }
    }
}
