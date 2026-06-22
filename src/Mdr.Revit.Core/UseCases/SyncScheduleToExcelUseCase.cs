using System;
using System.Collections.Generic;
using System.Threading;
using System.Threading.Tasks;
using Mdr.Revit.Core.Contracts;
using Mdr.Revit.Core.Models;

namespace Mdr.Revit.Core.UseCases
{
    public sealed class SyncScheduleToExcelUseCase
    {
        private readonly IExcelWorkbookClient _excelWorkbookClient;
        private readonly IRevitScheduleSyncAdapter _revitScheduleSyncAdapter;

        public SyncScheduleToExcelUseCase(
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

            ValidateProfile(request.Profile);
            string scheduleName = string.IsNullOrWhiteSpace(request.ScheduleName)
                ? string.Empty
                : request.ScheduleName.Trim();

            GoogleSheetSyncProfile scheduleProfile = ToScheduleProfile(request.Profile, instanceOnlyWrites: false);
            IReadOnlyList<ScheduleSyncRow> extractedRows = _revitScheduleSyncAdapter.ExtractRows(scheduleName, scheduleProfile);
            List<ScheduleSyncRow> rowsToWrite = new List<ScheduleSyncRow>(extractedRows.Count);
            Dictionary<string, int> skippedByReason = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

            for (int i = 0; i < extractedRows.Count; i++)
            {
                ScheduleSyncRow row = extractedRows[i];
                if (row.Errors.Count == 0 && !string.IsNullOrWhiteSpace(row.AnchorUniqueId))
                {
                    rowsToWrite.Add(row);
                    continue;
                }

                string reason = ResolveSkipReason(row);
                if (!skippedByReason.TryGetValue(reason, out int count))
                {
                    skippedByReason[reason] = 1;
                }
                else
                {
                    skippedByReason[reason] = count + 1;
                }
            }

            ExcelWorkbookWriteResult write = new ExcelWorkbookWriteResult();
            if (rowsToWrite.Count > 0)
            {
                write = await _excelWorkbookClient
                    .WriteRowsAsync(request.Profile, rowsToWrite, cancellationToken)
                    .ConfigureAwait(false);
            }

            ExcelScheduleSyncResult result = new ExcelScheduleSyncResult
            {
                Direction = GoogleSyncDirections.Export,
                ExportedRows = rowsToWrite.Count,
                SkippedRows = extractedRows.Count - rowsToWrite.Count,
                WriteResult = write,
            };

            foreach (KeyValuePair<string, int> item in skippedByReason)
            {
                result.SkippedByReason[item.Key] = item.Value;
            }

            AddWarnings(result);
            return result;
        }

        internal static GoogleSheetSyncProfile ToScheduleProfile(
            ExcelWorkbookProfile profile,
            bool instanceOnlyWrites)
        {
            GoogleSheetSyncProfile scheduleProfile = new GoogleSheetSyncProfile
            {
                SpreadsheetId = profile.FilePath ?? string.Empty,
                WorksheetName = profile.WorksheetName ?? string.Empty,
                AnchorColumn = string.IsNullOrWhiteSpace(profile.AnchorColumn)
                    ? "MDR_UNIQUE_ID"
                    : profile.AnchorColumn.Trim(),
                InstanceOnlyWrites = instanceOnlyWrites,
            };

            for (int i = 0; i < profile.ColumnMappings.Count; i++)
            {
                GoogleSheetColumnMapping mapping = profile.ColumnMappings[i];
                scheduleProfile.ColumnMappings.Add(new GoogleSheetColumnMapping
                {
                    SheetColumn = mapping.SheetColumn ?? string.Empty,
                    RevitParameter = mapping.RevitParameter ?? string.Empty,
                    IsEditable = mapping.IsEditable,
                });
            }

            for (int i = 0; i < profile.ProtectedColumns.Count; i++)
            {
                scheduleProfile.ProtectedColumns.Add(profile.ProtectedColumns[i] ?? string.Empty);
            }

            return scheduleProfile;
        }

        private static void ValidateProfile(ExcelWorkbookProfile profile)
        {
            if (profile == null)
            {
                throw new InvalidOperationException("Excel sync profile is required.");
            }

            if (string.IsNullOrWhiteSpace(profile.FilePath))
            {
                throw new InvalidOperationException("Excel file path is required.");
            }

            if (string.IsNullOrWhiteSpace(profile.WorksheetName))
            {
                throw new InvalidOperationException("WorksheetName is required.");
            }
        }

        private static string ResolveSkipReason(ScheduleSyncRow row)
        {
            if (row == null || row.Errors.Count == 0)
            {
                return "anchor_missing";
            }

            string code = row.Errors[0].Code ?? string.Empty;
            return string.IsNullOrWhiteSpace(code) ? "anchor_missing" : code.Trim();
        }

        private static void AddWarnings(ExcelScheduleSyncResult result)
        {
            if (result.SkippedByReason.ContainsKey("schedule_not_itemized"))
            {
                result.Warnings.Add("schedule_not_itemized");
            }

            if (result.SkippedByReason.ContainsKey("aggregate_row_skipped"))
            {
                result.Warnings.Add("aggregate_row_skipped");
            }
        }
    }
}
