using System;
using System.Threading;
using System.Threading.Tasks;
using Mdr.Revit.Core.Contracts;
using Mdr.Revit.Core.Models;

namespace Mdr.Revit.Core.UseCases
{
    public sealed class SyncScheduleToGoogleUseCase
    {
        private readonly IGoogleSheetsClient _googleSheetsClient;
        private readonly IRevitScheduleSyncAdapter _revitScheduleSyncAdapter;

        public SyncScheduleToGoogleUseCase(
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

            ValidateProfile(request.Profile);
            string scheduleName = string.IsNullOrWhiteSpace(request.ScheduleName)
                ? string.Empty
                : request.ScheduleName.Trim();

            var rows = _revitScheduleSyncAdapter.ExtractRows(scheduleName, request.Profile);
            GoogleSheetWriteResult write = await _googleSheetsClient
                .WriteRowsAsync(request.Profile, rows, cancellationToken)
                .ConfigureAwait(false);

            return new GoogleScheduleSyncResult
            {
                Direction = GoogleSyncDirections.Export,
                ExportedRows = rows.Count,
                WriteResult = write,
            };
        }

        private static void ValidateProfile(GoogleSheetSyncProfile profile)
        {
            if (profile == null)
            {
                throw new InvalidOperationException("Google sync profile is required.");
            }

            if (string.IsNullOrWhiteSpace(profile.SpreadsheetId))
            {
                throw new InvalidOperationException("SpreadsheetId is required.");
            }

            if (string.IsNullOrWhiteSpace(profile.WorksheetName))
            {
                throw new InvalidOperationException("WorksheetName is required.");
            }
        }
    }
}
