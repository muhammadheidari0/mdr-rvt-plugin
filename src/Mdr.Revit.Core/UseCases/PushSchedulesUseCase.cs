using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Mdr.Revit.Core.Contracts;
using Mdr.Revit.Core.Models;
using Mdr.Revit.Core.Validation;

namespace Mdr.Revit.Core.UseCases
{
    public sealed class PushSchedulesUseCase
    {
        private readonly IApiClient _apiClient;
        private readonly IRevitExtractor _revitExtractor;

        public PushSchedulesUseCase(IApiClient apiClient, IRevitExtractor revitExtractor)
        {
            _apiClient = apiClient ?? throw new ArgumentNullException(nameof(apiClient));
            _revitExtractor = revitExtractor ?? throw new ArgumentNullException(nameof(revitExtractor));
        }

        public async Task<ScheduleIngestResponse> ExecuteAsync(
            ScheduleIngestRequest request,
            CancellationToken cancellationToken)
        {
            if (request == null)
            {
                throw new ArgumentNullException(nameof(request));
            }

            if (request.Rows.Count == 0)
            {
                IReadOnlyList<ScheduleRow> extractedRows = _revitExtractor.GetScheduleRows(request.ProfileCode);
                foreach (ScheduleRow row in extractedRows)
                {
                    request.Rows.Add(row);
                }
            }

            BusinessRules.EnsureScheduleRequestIsValid(request);

            ScheduleIngestResponse response = await _apiClient
                .IngestScheduleAsync(request, cancellationToken)
                .ConfigureAwait(false);

            return response;
        }
    }
}
