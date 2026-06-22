using System;
using System.Collections.Generic;
using System.Linq;
using Mdr.Revit.Core.Models;

namespace Mdr.Revit.Addin.UI
{
    public sealed class SmartNumberingWindowViewModel
    {
        private readonly List<SmartNumberingRuleOption> _ruleOptions = new List<SmartNumberingRuleOption>();
        private readonly List<SmartNumberingPreviewRow> _previewRows = new List<SmartNumberingPreviewRow>();
        private readonly List<string> _blockOptions = new List<string>();
        private readonly List<string> _levelOptions = new List<string>();

        public IReadOnlyList<SmartNumberingRuleOption> RuleOptions => _ruleOptions;

        public IReadOnlyList<SmartNumberingPreviewRow> PreviewRows => _previewRows;

        public IReadOnlyList<string> BlockOptions => _blockOptions;

        public IReadOnlyList<string> LevelOptions => _levelOptions;

        public string SelectedRuleId { get; set; } = string.Empty;

        public string Mode { get; set; } = SmartNumberingModes.Formula;

        public string Formula { get; set; } = "{Mark}-{Sequence:5}";

        public string TargetsText { get; set; } = "Serial No";

        public string CategoryBuiltInName { get; set; } = "OST_Walls";

        public string SelectedBlock { get; set; } = string.Empty;

        public string SelectedLevel { get; set; } = string.Empty;

        public int StartAt { get; set; } = 1;

        public int SequenceWidth { get; set; } = 5;

        public string StatusText { get; private set; } = "No preview yet.";

        public bool HasPreviewErrors { get; private set; }

        public bool CanApply => _previewRows.Count > 0 && !HasPreviewErrors;

        public bool IsArcaMode => string.Equals(Mode, SmartNumberingModes.Arca, StringComparison.OrdinalIgnoreCase);

        public void SetRules(IReadOnlyList<SmartNumberingRule> rules, string defaultRuleId)
        {
            _ruleOptions.Clear();
            if (rules != null)
            {
                for (int i = 0; i < rules.Count; i++)
                {
                    SmartNumberingRule? source = rules[i];
                    if (source == null)
                    {
                        continue;
                    }

                    SmartNumberingRule copy = CloneRule(source);
                    string id = string.IsNullOrWhiteSpace(copy.RuleId)
                        ? ("rule_" + (_ruleOptions.Count + 1))
                        : copy.RuleId.Trim();
                    copy.RuleId = id;

                    _ruleOptions.Add(new SmartNumberingRuleOption
                    {
                        RuleId = id,
                        DisplayName = id,
                        Rule = copy,
                    });
                }
            }

            if (_ruleOptions.Count == 0)
            {
                SelectedRuleId = string.Empty;
                return;
            }

            SmartNumberingRuleOption? selected = _ruleOptions.FirstOrDefault(x =>
                string.Equals(x.RuleId, defaultRuleId, StringComparison.OrdinalIgnoreCase));
            if (selected == null && !string.IsNullOrWhiteSpace(SelectedRuleId))
            {
                selected = _ruleOptions.FirstOrDefault(x =>
                    string.Equals(x.RuleId, SelectedRuleId, StringComparison.OrdinalIgnoreCase));
            }

            selected ??= _ruleOptions[0];
            SelectedRuleId = selected.RuleId;
            ApplyRuleTemplate(selected.Rule);
        }

        public void ApplySelectedRuleTemplate()
        {
            SmartNumberingRuleOption? selected = _ruleOptions.FirstOrDefault(x =>
                string.Equals(x.RuleId, SelectedRuleId, StringComparison.OrdinalIgnoreCase));
            if (selected == null)
            {
                return;
            }

            ApplyRuleTemplate(selected.Rule);
        }

        public SmartNumberingRule BuildRule()
        {
            SmartNumberingRuleOption? selected = _ruleOptions.FirstOrDefault(x =>
                string.Equals(x.RuleId, SelectedRuleId, StringComparison.OrdinalIgnoreCase));
            SmartNumberingRule? template = selected?.Rule;

            SmartNumberingRule rule = new SmartNumberingRule
            {
                RuleId = string.IsNullOrWhiteSpace(SelectedRuleId)
                    ? (template?.RuleId ?? "runtime")
                    : SelectedRuleId.Trim(),
                Mode = string.IsNullOrWhiteSpace(Mode) ? SmartNumberingModes.Formula : Mode.Trim(),
                Formula = (Formula ?? string.Empty).Trim(),
                SelectionFilter = template?.SelectionFilter ?? string.Empty,
                CategoryBuiltInName = string.IsNullOrWhiteSpace(CategoryBuiltInName)
                    ? (template?.CategoryBuiltInName ?? string.Empty)
                    : CategoryBuiltInName.Trim(),
                SelectedBlock = (SelectedBlock ?? string.Empty).Trim(),
                SelectedLevel = (SelectedLevel ?? string.Empty).Trim(),
                SequenceWidth = SequenceWidth <= 0 ? 5 : SequenceWidth,
                StartAt = StartAt <= 0 ? 1 : StartAt,
            };

            IReadOnlyList<string> targets = ParseTargets(TargetsText);
            for (int i = 0; i < targets.Count; i++)
            {
                rule.Targets.Add(targets[i]);
            }

            if (rule.Targets.Count == 0)
            {
                rule.Targets.Add("Serial No");
            }

            return rule;
        }

        public void SetMetadata(SmartNumberingMetadata metadata)
        {
            _blockOptions.Clear();
            _levelOptions.Clear();

            if (metadata != null)
            {
                AddDistinct(_blockOptions, metadata.Blocks);
                AddDistinct(_levelOptions, metadata.Levels);
            }

            if (string.IsNullOrWhiteSpace(SelectedBlock) && _blockOptions.Count > 0)
            {
                SelectedBlock = _blockOptions[0];
            }

            if (string.IsNullOrWhiteSpace(SelectedLevel) && _levelOptions.Count > 0)
            {
                SelectedLevel = _levelOptions[0];
            }
        }

        public void UpdatePreview(SmartNumberingResult result)
        {
            _previewRows.Clear();

            if (result?.Preview != null)
            {
                for (int i = 0; i < result.Preview.Count; i++)
                {
                    SmartNumberingPreviewItem item = result.Preview[i];
                    if (item == null)
                    {
                        continue;
                    }

                    string errorText = item.ErrorMessage ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(errorText) && !string.IsNullOrWhiteSpace(item.ErrorCode))
                    {
                        errorText = item.ErrorCode;
                    }

                    _previewRows.Add(new SmartNumberingPreviewRow
                    {
                        Element = item.ElementKey ?? string.Empty,
                        Target = item.Target ?? string.Empty,
                        Current = item.CurrentValue ?? string.Empty,
                        Proposed = item.ProposedValue ?? string.Empty,
                        Status = item.Status ?? string.Empty,
                        Error = errorText,
                    });
                }
            }

            bool hasErrorRows = _previewRows.Any(x =>
                string.Equals(x.Status, SmartNumberingPreviewStates.Error, StringComparison.OrdinalIgnoreCase));
            bool hasErrors = hasErrorRows ||
                (result?.Errors?.Count > 0) ||
                !string.IsNullOrWhiteSpace(result?.FatalErrorCode);
            HasPreviewErrors = hasErrors;

            if (result == null)
            {
                StatusText = "No preview result.";
                return;
            }

            StatusText = "Planned: " + result.PlannedCount +
                " | Applied: " + result.AppliedCount +
                " | Skipped: " + result.SkippedCount +
                " | Failed: " + result.FailedCount;
        }

        private void ApplyRuleTemplate(SmartNumberingRule rule)
        {
            Mode = string.IsNullOrWhiteSpace(rule.Mode) ? SmartNumberingModes.Formula : rule.Mode.Trim();
            Formula = (rule.Formula ?? string.Empty).Trim();
            CategoryBuiltInName = string.IsNullOrWhiteSpace(rule.CategoryBuiltInName)
                ? "OST_Walls"
                : rule.CategoryBuiltInName.Trim();
            SelectedBlock = (rule.SelectedBlock ?? string.Empty).Trim();
            SelectedLevel = (rule.SelectedLevel ?? string.Empty).Trim();
            StartAt = rule.StartAt <= 0 ? 1 : rule.StartAt;
            SequenceWidth = rule.SequenceWidth <= 0 ? 5 : rule.SequenceWidth;
            TargetsText = rule.Targets.Count == 0
                ? "Serial No"
                : string.Join(", ", rule.Targets);
        }

        private static SmartNumberingRule CloneRule(SmartNumberingRule source)
        {
            SmartNumberingRule copy = new SmartNumberingRule
            {
                RuleId = source.RuleId ?? string.Empty,
                Mode = string.IsNullOrWhiteSpace(source.Mode) ? SmartNumberingModes.Formula : source.Mode,
                Formula = source.Formula ?? string.Empty,
                SelectionFilter = source.SelectionFilter ?? string.Empty,
                CategoryBuiltInName = source.CategoryBuiltInName ?? string.Empty,
                SelectedBlock = source.SelectedBlock ?? string.Empty,
                SelectedLevel = source.SelectedLevel ?? string.Empty,
                SequenceWidth = source.SequenceWidth,
                StartAt = source.StartAt,
            };

            for (int i = 0; i < source.Targets.Count; i++)
            {
                string target = (source.Targets[i] ?? string.Empty).Trim();
                if (string.IsNullOrWhiteSpace(target))
                {
                    continue;
                }

                copy.Targets.Add(target);
            }

            return copy;
        }

        private static void AddDistinct(List<string> target, IReadOnlyList<string> values)
        {
            if (values == null)
            {
                return;
            }

            for (int i = 0; i < values.Count; i++)
            {
                string value = (values[i] ?? string.Empty).Trim();
                if (string.IsNullOrWhiteSpace(value))
                {
                    continue;
                }

                if (target.Any(x => string.Equals(x, value, StringComparison.OrdinalIgnoreCase)))
                {
                    continue;
                }

                target.Add(value);
            }
        }

        private static IReadOnlyList<string> ParseTargets(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw))
            {
                return Array.Empty<string>();
            }

            string[] parts = raw
                .Split(new[] { ',', ';', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            List<string> targets = new List<string>(parts.Length);
            for (int i = 0; i < parts.Length; i++)
            {
                string value = (parts[i] ?? string.Empty).Trim();
                if (string.IsNullOrWhiteSpace(value))
                {
                    continue;
                }

                if (targets.Any(x => string.Equals(x, value, StringComparison.OrdinalIgnoreCase)))
                {
                    continue;
                }

                targets.Add(value);
            }

            return targets;
        }
    }

    public sealed class SmartNumberingRuleOption
    {
        public string RuleId { get; set; } = string.Empty;

        public string DisplayName { get; set; } = string.Empty;

        public SmartNumberingRule Rule { get; set; } = new SmartNumberingRule();
    }

    public sealed class SmartNumberingPreviewRow
    {
        public string Element { get; set; } = string.Empty;

        public string Target { get; set; } = string.Empty;

        public string Current { get; set; } = string.Empty;

        public string Proposed { get; set; } = string.Empty;

        public string Status { get; set; } = string.Empty;

        public string Error { get; set; } = string.Empty;
    }
}
