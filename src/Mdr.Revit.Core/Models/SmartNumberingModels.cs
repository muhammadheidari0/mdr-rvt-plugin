using System.Collections.Generic;

namespace Mdr.Revit.Core.Models
{
    public sealed class SmartNumberingRule
    {
        public string RuleId { get; set; } = string.Empty;

        public string Formula { get; set; } = string.Empty;

        public string SelectionFilter { get; set; } = string.Empty;

        public int SequenceWidth { get; set; } = 5;

        public int StartAt { get; set; } = 1;

        public List<string> Targets { get; } = new List<string>();
    }

    public sealed class SmartNumberingResult
    {
        public int AppliedCount { get; set; }

        public int SkippedCount { get; set; }

        public List<SmartNumberingPreviewItem> Preview { get; } = new List<SmartNumberingPreviewItem>();

        public List<SmartNumberingError> Errors { get; } = new List<SmartNumberingError>();
    }

    public sealed class SmartNumberingPreviewItem
    {
        public string ElementKey { get; set; } = string.Empty;

        public string Value { get; set; } = string.Empty;
    }

    public sealed class SmartNumberingError
    {
        public string ElementKey { get; set; } = string.Empty;

        public string ErrorCode { get; set; } = string.Empty;

        public string Message { get; set; } = string.Empty;
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
}
