using System;
using System.Threading;
using System.Threading.Tasks;
using Mdr.Revit.Core.Contracts;
using Mdr.Revit.Core.Models;

namespace Mdr.Revit.Core.UseCases
{
    public sealed class SyncScheduleFromGoogleUseCase
    {
        private readonly IGoogleSheetsClient _googleSheetsClient;
        private readonly IRevitScheduleSyncAdapter _revitScheduleSyncAdapter;

        public SyncScheduleFromGoogleUseCase(
            IGoogleSheetsClient googleSheetsClient,
            IRevitScheduleSyncAdapter revitScheduleSyncAdapter)
        {
            _googleSheetsClient = googleSheetsClient ?? throw new ArgumentNullException(nameof(googleSheetsClient));
            _revitScheduleSyncAdapter = revitScheduleSyncAdapter ?? throw new ArgumentNullException(nameof(revitScheduleSyncAdapter));
        }

        public async Task<GoogleScheduleSyncResult> ExecuteAsync(
            GoogleScheduleSyncRequest request,
            CancellationToken cancellationToken)
        {
            if (request == null)
            {
                throw new ArgumentNullException(nameof(request));
            }

            if (request.Profile == null)
            {
                throw new InvalidOperationException("Google sync profile is required.");
            }

            GoogleSheetReadResult sheetRows = await _googleSheetsClient
                .ReadRowsAsync(request.Profile, cancellationToken);

            ScheduleSyncDiffResult diff = _revitScheduleSyncAdapter.BuildDiff(sheetRows.Rows, request.Profile);
            GoogleScheduleSyncResult result = new GoogleScheduleSyncResult
            {
                Direction = GoogleSyncDirections.Import,
                DiffResult = diff,
            };

            if (!request.ApplyChanges)
            {
                return result;
            }

            result.ApplyResult = _revitScheduleSyncAdapter.ApplyDiff(diff, request.Profile);
            return result;
        }
    }
}
