using System;
using System.Collections.Generic;

namespace Mdr.Revit.Core.Models
{
    public static class GoogleSyncDirections
    {
        public const string Export = "export";

        public const string Import = "import";
    }

    public static class ScheduleSyncStates
    {
        public const string Unchanged = "unchanged";

        public const string Modified = "modified";

        public const string Error = "error";
    }

    public sealed class GoogleSheetSyncProfile
    {
        public string SpreadsheetId { get; set; } = string.Empty;

        public string WorksheetName { get; set; } = "Sheet1";

        public string AnchorColumn { get; set; } = "MDR_UNIQUE_ID";

        public List<GoogleSheetColumnMapping> ColumnMappings { get; } = new List<GoogleSheetColumnMapping>();

        public List<string> ProtectedColumns { get; } = new List<string>();
    }

    public sealed class GoogleSheetColumnMapping
    {
        public string SheetColumn { get; set; } = string.Empty;

        public string RevitParameter { get; set; } = string.Empty;

        public bool IsEditable { get; set; } = true;
    }

    public sealed class ScheduleSyncRow
    {
        public string AnchorUniqueId { get; set; } = string.Empty;

        public string ElementId { get; set; } = string.Empty;

        public string ChangeState { get; set; } = ScheduleSyncStates.Unchanged;

        public Dictionary<string, string> Cells { get; } = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

        public List<ScheduleSyncError> Errors { get; } = new List<ScheduleSyncError>();

        public ScheduleSyncRow Clone()
        {
            ScheduleSyncRow clone = new ScheduleSyncRow
            {
                AnchorUniqueId = AnchorUniqueId,
                ElementId = ElementId,
                ChangeState = ChangeState,
            };
            foreach (KeyValuePair<string, string> pair in Cells)
            {
                clone.Cells[pair.Key] = pair.Value ?? string.Empty;
            }

            foreach (ScheduleSyncError error in Errors)
            {
                clone.Errors.Add(new ScheduleSyncError
                {
                    Code = error.Code,
                    Message = error.Message,
                });
            }

            return clone;
        }
    }

    public sealed class ScheduleSyncError
    {
        public string Code { get; set; } = string.Empty;

        public string Message { get; set; } = string.Empty;
    }

    public sealed class ScheduleSyncDiffResult
    {
        public List<ScheduleSyncRow> Rows { get; } = new List<ScheduleSyncRow>();

        public int ChangedRowsCount { get; set; }

        public int ErrorRowsCount { get; set; }
    }

    public sealed class ScheduleSyncApplyResult
    {
        public int AppliedCount { get; set; }

        public int FailedCount { get; set; }

        public List<ScheduleSyncError> Errors { get; } = new List<ScheduleSyncError>();
    }

    public sealed class GoogleSheetReadResult
    {
        public IReadOnlyList<string> Headers { get; set; } = Array.Empty<string>();

        public List<ScheduleSyncRow> Rows { get; } = new List<ScheduleSyncRow>();
    }

    public sealed class GoogleSheetWriteResult
    {
        public int UpdatedRows { get; set; }

        public string UpdatedRange { get; set; } = string.Empty;
    }

    public sealed class GoogleScheduleSyncRequest
    {
        public string Direction { get; set; } = GoogleSyncDirections.Export;

        public string ScheduleName { get; set; } = string.Empty;

        public GoogleSheetSyncProfile Profile { get; set; } = new GoogleSheetSyncProfile();

        public bool ApplyChanges { get; set; } = true;
    }

    public sealed class GoogleScheduleSyncResult
    {
        public string Direction { get; set; } = GoogleSyncDirections.Export;

        public int ExportedRows { get; set; }

        public int SkippedRows { get; set; }

        public Dictionary<string, int> SkippedByReason { get; } =
            new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        public List<string> Warnings { get; } = new List<string>();

        public GoogleSheetWriteResult WriteResult { get; set; } = new GoogleSheetWriteResult();

        public ScheduleSyncDiffResult DiffResult { get; set; } = new ScheduleSyncDiffResult();

        public ScheduleSyncApplyResult ApplyResult { get; set; } = new ScheduleSyncApplyResult();
    }
}
