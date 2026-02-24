using System;
using System.Collections.Generic;
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

            IReadOnlyList<ScheduleSyncRow> extractedRows = _revitScheduleSyncAdapter.ExtractRows(scheduleName, request.Profile);
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

            GoogleSheetWriteResult write = new GoogleSheetWriteResult();
            if (rowsToWrite.Count > 0)
            {
                write = await _googleSheetsClient
                    .WriteRowsAsync(request.Profile, rowsToWrite, cancellationToken);
            }

            GoogleScheduleSyncResult result = new GoogleScheduleSyncResult
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

            if (HasRenamedDuplicateHeaders(rowsToWrite))
            {
                AddWarning(result, "header_duplicate_renamed");
            }

            if (result.SkippedByReason.ContainsKey("schedule_not_itemized"))
            {
                AddWarning(result, "schedule_not_itemized");
            }

            if (result.SkippedByReason.ContainsKey("aggregate_row_skipped"))
            {
                AddWarning(result, "aggregate_row_skipped");
            }

            return result;
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

        private static string ResolveSkipReason(ScheduleSyncRow row)
        {
            if (row == null)
            {
                return "anchor_missing";
            }

            if (row.Errors.Count == 0)
            {
                return "anchor_missing";
            }

            string code = row.Errors[0].Code ?? string.Empty;
            if (string.IsNullOrWhiteSpace(code))
            {
                return "anchor_missing";
            }

            return code.Trim();
        }

        private static bool HasRenamedDuplicateHeaders(IReadOnlyList<ScheduleSyncRow> rows)
        {
            if (rows == null || rows.Count == 0)
            {
                return false;
            }

            HashSet<string> keys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < rows.Count; i++)
            {
                foreach (string key in rows[i].Cells.Keys)
                {
                    keys.Add(key);
                }
            }

            foreach (string key in keys)
            {
                int underscore = key.LastIndexOf('_');
                if (underscore <= 0 || underscore >= key.Length - 1)
                {
                    continue;
                }

                if (!int.TryParse(key.Substring(underscore + 1), out int suffix) || suffix < 2)
                {
                    continue;
                }

                string baseName = key.Substring(0, underscore);
                if (keys.Contains(baseName))
                {
                    return true;
                }
            }

            return false;
        }

        private static void AddWarning(GoogleScheduleSyncResult result, string code)
        {
            string normalized = (code ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return;
            }

            for (int i = 0; i < result.Warnings.Count; i++)
            {
                if (string.Equals(result.Warnings[i], normalized, StringComparison.OrdinalIgnoreCase))
                {
                    return;
                }
            }

            result.Warnings.Add(normalized);
        }
    }
}
