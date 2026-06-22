using System;
using System.Collections.Generic;

namespace Mdr.Revit.Core.Models
{
    public sealed class ExcelWorkbookProfile
    {
        public string FilePath { get; set; } = string.Empty;

        public string WorksheetName { get; set; } = string.Empty;

        public string AnchorColumn { get; set; } = "MDR_UNIQUE_ID";

        public List<GoogleSheetColumnMapping> ColumnMappings { get; } = new List<GoogleSheetColumnMapping>();

        public List<string> ProtectedColumns { get; } = new List<string>();
    }

    public sealed class ExcelWorkbookReadResult
    {
        public IReadOnlyList<string> Headers { get; set; } = Array.Empty<string>();

        public List<ScheduleSyncRow> Rows { get; } = new List<ScheduleSyncRow>();
    }

    public sealed class ExcelWorkbookWriteResult
    {
        public int UpdatedRows { get; set; }

        public string FilePath { get; set; } = string.Empty;

        public string WorksheetName { get; set; } = string.Empty;
    }

    public sealed class ExcelScheduleSyncRequest
    {
        public string Direction { get; set; } = GoogleSyncDirections.Export;

        public string ScheduleName { get; set; } = string.Empty;

        public ExcelWorkbookProfile Profile { get; set; } = new ExcelWorkbookProfile();

        public bool ApplyChanges { get; set; } = true;
    }

    public sealed class ExcelScheduleSyncResult
    {
        public string Direction { get; set; } = GoogleSyncDirections.Export;

        public int ExportedRows { get; set; }

        public int SkippedRows { get; set; }

        public Dictionary<string, int> SkippedByReason { get; } =
            new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);

        public List<string> Warnings { get; } = new List<string>();

        public ExcelWorkbookWriteResult WriteResult { get; set; } = new ExcelWorkbookWriteResult();

        public ScheduleSyncDiffResult DiffResult { get; set; } = new ScheduleSyncDiffResult();

        public ScheduleSyncApplyResult ApplyResult { get; set; } = new ScheduleSyncApplyResult();
    }
}
