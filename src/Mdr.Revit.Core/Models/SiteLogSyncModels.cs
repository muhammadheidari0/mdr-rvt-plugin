using System;
using System.Collections.Generic;

namespace Mdr.Revit.Core.Models
{
    public static class SyncOperations
    {
        public const string Upsert = "upsert";

        public const string Delete = "delete";
    }

    public static class SyncSections
    {
        public const string Manpower = "MANPOWER";

        public const string Equipment = "EQUIPMENT";

        public const string Activity = "ACTIVITY";
    }

    public sealed class SiteLogManifestRequest
    {
        public string ProjectCode { get; set; } = string.Empty;

        public string DisciplineCode { get; set; } = string.Empty;

        public DateTimeOffset? UpdatedAfterUtc { get; set; }

        public int Limit { get; set; } = 500;

        public string ClientModelGuid { get; set; } = string.Empty;
    }

    public sealed class SiteLogManifestResponse
    {
        public string RunId { get; set; } = string.Empty;

        public string NextCursor { get; set; } = string.Empty;

        public List<SiteLogManifestChange> Changes { get; } = new List<SiteLogManifestChange>();
    }

    public sealed class SiteLogManifestChange
    {
        public long LogId { get; set; }

        public string LogNo { get; set; } = string.Empty;

        public DateTimeOffset VerifiedAtUtc { get; set; }

        public string LogHash { get; set; } = string.Empty;

        public string Operation { get; set; } = SyncOperations.Upsert;
    }

    public sealed class SiteLogPullRequest
    {
        public string ProjectCode { get; set; } = string.Empty;

        public string DisciplineCode { get; set; } = string.Empty;

        public string ClientModelGuid { get; set; } = string.Empty;

        public string PluginVersion { get; set; } = string.Empty;

        public List<long> LogIds { get; } = new List<long>();
    }

    public sealed class SiteLogPullResponse
    {
        public string RunId { get; set; } = string.Empty;

        public List<SiteLogRow> ManpowerRows { get; } = new List<SiteLogRow>();

        public List<SiteLogRow> EquipmentRows { get; } = new List<SiteLogRow>();

        public List<SiteLogRow> ActivityRows { get; } = new List<SiteLogRow>();
    }

    public sealed class SiteLogRow
    {
        public string SyncKey { get; set; } = string.Empty;

        public long LogId { get; set; }

        public string LogNo { get; set; } = string.Empty;

        public DateTimeOffset LogDateUtc { get; set; }

        public string SectionCode { get; set; } = string.Empty;

        public string Operation { get; set; } = SyncOperations.Upsert;

        public string RowHash { get; set; } = string.Empty;

        public Dictionary<string, string> Fields { get; } = new Dictionary<string, string>();
    }

    public sealed class SiteLogAckRequest
    {
        public string RunId { get; set; } = string.Empty;

        public int AppliedCount { get; set; }

        public int FailedCount { get; set; }

        public List<SiteLogApplyError> Errors { get; } = new List<SiteLogApplyError>();
    }

    public sealed class SiteLogAckResponse
    {
        public bool Ok { get; set; }

        public DateTimeOffset ReceivedAtUtc { get; set; }
    }

    public sealed class SiteLogApplyResult
    {
        public string RunId { get; set; } = string.Empty;

        public int AppliedCount { get; set; }

        public int FailedCount { get; set; }

        public List<SiteLogApplyError> Errors { get; } = new List<SiteLogApplyError>();

        public static SiteLogApplyResult Empty(string runId)
        {
            return new SiteLogApplyResult
            {
                RunId = runId,
                AppliedCount = 0,
                FailedCount = 0,
            };
        }
    }

    public sealed class SiteLogApplyError
    {
        public string SyncKey { get; set; } = string.Empty;

        public string Message { get; set; } = string.Empty;
    }
}
