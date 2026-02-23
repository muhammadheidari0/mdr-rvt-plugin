using System;
using System.Collections.Generic;

namespace Mdr.Revit.Core.Models
{
    public static class ScheduleProfiles
    {
        public const string Mto = "MTO";

        public const string Equipment = "EQUIPMENT";
    }

    public sealed class ScheduleIngestRequest
    {
        public string ProjectCode { get; set; } = string.Empty;

        public string ProfileCode { get; set; } = string.Empty;

        public string ModelGuid { get; set; } = string.Empty;

        public string ViewName { get; set; } = string.Empty;

        public string SchemaVersion { get; set; } = "v1";

        public DateTimeOffset ExtractedAtUtc { get; set; } = DateTimeOffset.UtcNow;

        public List<ScheduleRow> Rows { get; } = new List<ScheduleRow>();
    }

    public sealed class ScheduleRow
    {
        public int RowNo { get; set; }

        public string ElementKey { get; set; } = string.Empty;

        public Dictionary<string, string> Values { get; } = new Dictionary<string, string>();
    }

    public sealed class ScheduleIngestResponse
    {
        public string RunId { get; set; } = string.Empty;

        public ValidationSummary ValidationSummary { get; set; } = new ValidationSummary();

        public List<RowError> RowErrors { get; } = new List<RowError>();
    }

    public sealed class ValidationSummary
    {
        public int TotalRows { get; set; }

        public int ValidRows { get; set; }

        public int InvalidRows { get; set; }
    }

    public sealed class RowError
    {
        public int RowNo { get; set; }

        public string ErrorCode { get; set; } = string.Empty;

        public string Message { get; set; } = string.Empty;
    }
}
