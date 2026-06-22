using System;
using System.Threading;
using System.Threading.Tasks;
using Mdr.Revit.Core.Contracts;
using Mdr.Revit.Core.Models;

namespace Mdr.Revit.Core.UseCases
{
    public sealed class SyncScheduleFromExcelUseCase
    {
        private readonly IExcelWorkbookClient _excelWorkbookClient;
        private readonly IRevitScheduleSyncAdapter _revitScheduleSyncAdapter;

        public SyncScheduleFromExcelUseCase(
            IExcelWorkbookClient excelWorkbookClient,
            IRevitScheduleSyncAdapter revitScheduleSyncAdapter)
        {
            _excelWorkbookClient = excelWorkbookClient ?? throw new ArgumentNullException(nameof(excelWorkbookClient));
            _revitScheduleSyncAdapter = revitScheduleSyncAdapter ?? throw new ArgumentNullException(nameof(revitScheduleSyncAdapter));
        }

        public async Task<ExcelScheduleSyncResult> ExecuteAsync(
            ExcelScheduleSyncRequest request,
            CancellationToken cancellationToken)
        {
            if (request == null)
            {
                throw new ArgumentNullException(nameof(request));
            }

            if (request.Profile == null)
            {
                throw new InvalidOperationException("Excel sync profile is required.");
            }

            ExcelWorkbookReadResult workbookRows = await _excelWorkbookClient
                .ReadRowsAsync(request.Profile, cancellationToken)
                .ConfigureAwait(false);

            GoogleSheetSyncProfile scheduleProfile =
                SyncScheduleToExcelUseCase.ToScheduleProfile(request.Profile, instanceOnlyWrites: true);
            ScheduleSyncDiffResult diff = _revitScheduleSyncAdapter.BuildDiff(workbookRows.Rows, scheduleProfile);
            ExcelScheduleSyncResult result = new ExcelScheduleSyncResult
            {
                Direction = GoogleSyncDirections.Import,
                DiffResult = diff,
            };

            if (!request.ApplyChanges)
            {
                return result;
            }

            result.ApplyResult = _revitScheduleSyncAdapter.ApplyDiff(diff, scheduleProfile);
            return result;
        }
    }
}
