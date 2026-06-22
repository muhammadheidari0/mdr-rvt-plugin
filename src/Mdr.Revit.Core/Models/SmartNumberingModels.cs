using System.Collections.Generic;

namespace Mdr.Revit.Core.Models
{
    public sealed class SmartNumberingRule
    {
        public string RuleId { get; set; } = string.Empty;

        public string Mode { get; set; } = SmartNumberingModes.Formula;

        public string Formula { get; set; } = string.Empty;

        public string SelectionFilter { get; set; } = string.Empty;

        public string CategoryBuiltInName { get; set; } = string.Empty;

        public string SelectedBlock { get; set; } = string.Empty;

        public string SelectedLevel { get; set; } = string.Empty;

        public int SequenceWidth { get; set; } = 5;

        public int StartAt { get; set; } = 1;

        public List<string> Targets { get; } = new List<string>();
    }

    public static class SmartNumberingModes
    {
        public const string Formula = "formula";
        public const string Arca = "arca";
    }

    public sealed class SmartNumberingResult
    {
        public int PlannedCount { get; set; }

        public int AppliedCount { get; set; }

        public int SkippedCount { get; set; }

        public int FailedCount { get; set; }

        public bool IsAtomicRollback { get; set; } = true;

        public bool WasRolledBack { get; set; }

        public string FatalErrorCode { get; set; } = string.Empty;

        public string FatalErrorMessage { get; set; } = string.Empty;

        public List<SmartNumberingPreviewItem> Preview { get; } = new List<SmartNumberingPreviewItem>();

        public List<SmartNumberingError> Errors { get; } = new List<SmartNumberingError>();
    }

    public sealed class SmartNumberingPreviewItem
    {
        public string ElementKey { get; set; } = string.Empty;

        public string Target { get; set; } = string.Empty;

        public string CurrentValue { get; set; } = string.Empty;

        public string ProposedValue { get; set; } = string.Empty;

        // Backward-compatibility alias for older tests/callers.
        public string Value
        {
            get
            {
                return ProposedValue;
            }

            set
            {
                ProposedValue = value ?? string.Empty;
            }
        }

        public string Status { get; set; } = SmartNumberingPreviewStates.Planned;

        public string ErrorCode { get; set; } = string.Empty;

        public string ErrorMessage { get; set; } = string.Empty;
    }

    public sealed class SmartNumberingError
    {
        public string ElementKey { get; set; } = string.Empty;

        public string Target { get; set; } = string.Empty;

        public string ErrorCode { get; set; } = string.Empty;

        public string Message { get; set; } = string.Empty;
    }

    public static class SmartNumberingPreviewStates
    {
        public const string Planned = "planned";
        public const string Applied = "applied";
        public const string Error = "error";
        public const string Skipped = "skipped";
    }

    public sealed class SmartNumberingFormula
    {
        public List<SmartNumberingToken> Tokens { get; } = new List<SmartNumberingToken>();

        public int SequenceWidth { get; set; } = 5;
    }

    public sealed class SmartNumberingToken
    {
        public string Kind { get; set; } = string.Empty;

        public string Value { get; set; } = string.Empty;

        public int SequenceWidth { get; set; }
    }

    public sealed class SmartNumberingMetadata
    {
        public List<string> Blocks { get; } = new List<string>();

        public List<string> Levels { get; } = new List<string>();
    }

    public sealed class ArcaSerialNumberingElementSnapshot
    {
        public string ElementKey { get; set; } = string.Empty;

        public int ElementId { get; set; }

        public string Block { get; set; } = string.Empty;

        public string LevelCode { get; set; } = string.Empty;

        public string TypeMark { get; set; } = string.Empty;

        public string Series { get; set; } = string.Empty;

        public string SerialNo { get; set; } = string.Empty;

        public int ScopeSort { get; set; } = 9999;
    }

    public sealed class ArcaSerialNumberingPlanRequest
    {
        public string SelectedBlock { get; set; } = string.Empty;

        public string SelectedLevelCode { get; set; } = string.Empty;

        public List<ArcaSerialNumberingElementSnapshot> Elements { get; } = new List<ArcaSerialNumberingElementSnapshot>();
    }

    public sealed class ArcaSerialNumberingPlan
    {
        public List<SmartNumberingPreviewItem> Preview { get; } = new List<SmartNumberingPreviewItem>();

        public List<ArcaSerialNumberingWrite> Writes { get; } = new List<ArcaSerialNumberingWrite>();

        public int SkippedCount { get; set; }
    }

    public sealed class ArcaSerialNumberingWrite
    {
        public string ElementKey { get; set; } = string.Empty;

        public string Target { get; set; } = string.Empty;

        public string Value { get; set; } = string.Empty;

        public SmartNumberingPreviewItem Preview { get; set; } = new SmartNumberingPreviewItem();
    }
}
