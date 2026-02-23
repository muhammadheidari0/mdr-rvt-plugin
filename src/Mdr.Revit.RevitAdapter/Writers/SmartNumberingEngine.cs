using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Mdr.Revit.Core.Contracts;
using Mdr.Revit.Core.Models;
using Mdr.Revit.Core.Validation;
using Mdr.Revit.RevitAdapter.Helpers;

namespace Mdr.Revit.RevitAdapter.Writers
{
    public sealed class SmartNumberingEngine : ISmartNumberingEngine
    {
        private readonly UIDocument? _uiDocument;
        private readonly ParameterAccessor _parameterAccessor;
        private readonly LocalSequenceScanner _sequenceScanner;
        private readonly SmartNumberingFormulaParser _formulaParser;

        public SmartNumberingEngine()
            : this(null, new ParameterAccessor(), new LocalSequenceScanner(), new SmartNumberingFormulaParser())
        {
        }

        public SmartNumberingEngine(UIDocument uiDocument)
            : this(uiDocument, new ParameterAccessor(), new LocalSequenceScanner(), new SmartNumberingFormulaParser())
        {
        }

        internal SmartNumberingEngine(
            UIDocument? uiDocument,
            ParameterAccessor parameterAccessor,
            LocalSequenceScanner sequenceScanner,
            SmartNumberingFormulaParser formulaParser)
        {
            _uiDocument = uiDocument;
            _parameterAccessor = parameterAccessor ?? throw new ArgumentNullException(nameof(parameterAccessor));
            _sequenceScanner = sequenceScanner ?? throw new ArgumentNullException(nameof(sequenceScanner));
            _formulaParser = formulaParser ?? throw new ArgumentNullException(nameof(formulaParser));
        }

        public SmartNumberingResult Apply(SmartNumberingRule rule, bool previewOnly)
        {
            SmartNumberingResult result = new SmartNumberingResult
            {
                IsAtomicRollback = true,
            };
            if (rule == null)
            {
                AddError(result, string.Empty, string.Empty, "validation_error", "Smart numbering rule is required.");
                return result;
            }

            if (_uiDocument?.Document == null)
            {
                AddError(result, string.Empty, string.Empty, "revit_context_missing", "Revit document context is unavailable.");
                return result;
            }

            if (string.IsNullOrWhiteSpace(rule.Formula))
            {
                AddError(result, string.Empty, string.Empty, "formula_invalid", "Smart numbering formula is required.");
                return result;
            }

            ICollection<ElementId> selectedIds = _uiDocument.Selection.GetElementIds();
            if (selectedIds == null || selectedIds.Count == 0)
            {
                AddError(result, string.Empty, string.Empty, "selection_required", "Select one or more elements before running smart numbering.");
                return result;
            }

            SmartNumberingFormula formula;
            try
            {
                formula = _formulaParser.Parse(rule.Formula, rule.SequenceWidth);
            }
            catch (Exception ex)
            {
                AddError(result, string.Empty, string.Empty, "formula_invalid", ex.Message);
                return result;
            }

            IReadOnlyList<string> targets = ResolveTargets(rule);
            List<Element> selectedElements = selectedIds
                .Select(id => _uiDocument.Document.GetElement(id))
                .Where(x => x != null)
                .Cast<Element>()
                .OrderBy(x => x.Id.Value)
                .ToList();
            if (selectedElements.Count == 0)
            {
                return result;
            }

            HashSet<ElementId> categories = new HashSet<ElementId>(selectedElements
                .Where(x => x.Category != null)
                .Select(x => x.Category.Id));
            Dictionary<string, int> nextByPrefix = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            int width = formula.SequenceWidth <= 0 ? 5 : formula.SequenceWidth;
            int startAt = rule.StartAt <= 0 ? 1 : rule.StartAt;
            List<PlannedAssignment> assignments = new List<PlannedAssignment>(selectedElements.Count);

            for (int i = 0; i < selectedElements.Count; i++)
            {
                Element element = selectedElements[i];
                string elementKey = GetElementKey(element);
                List<string> missingPlaceholders = new List<string>();
                string prefix = BuildPrefix(formula, element, missingPlaceholders);
                if (missingPlaceholders.Count > 0)
                {
                    string message = "Missing placeholder values: " + string.Join(", ", missingPlaceholders.Distinct(StringComparer.OrdinalIgnoreCase));
                    result.FailedCount++;
                    result.SkippedCount++;
                    AddError(result, elementKey, string.Join(", ", targets), "placeholder_missing", message);
                    result.Preview.Add(new SmartNumberingPreviewItem
                    {
                        ElementKey = elementKey,
                        Target = string.Join(", ", targets),
                        CurrentValue = _parameterAccessor.ReadValue(element, targets[0]),
                        ProposedValue = string.Empty,
                        Status = SmartNumberingPreviewStates.Error,
                        ErrorCode = "placeholder_missing",
                        ErrorMessage = message,
                    });
                    continue;
                }

                if (string.IsNullOrWhiteSpace(prefix))
                {
                    result.FailedCount++;
                    result.SkippedCount++;
                    AddError(result, elementKey, string.Join(", ", targets), "prefix_empty", "Generated prefix is empty.");
                    result.Preview.Add(new SmartNumberingPreviewItem
                    {
                        ElementKey = elementKey,
                        Target = string.Join(", ", targets),
                        CurrentValue = _parameterAccessor.ReadValue(element, targets[0]),
                        ProposedValue = string.Empty,
                        Status = SmartNumberingPreviewStates.Error,
                        ErrorCode = "prefix_empty",
                        ErrorMessage = "Generated prefix is empty.",
                    });
                    continue;
                }

                if (!nextByPrefix.TryGetValue(prefix, out int next))
                {
                    int maxExisting = _sequenceScanner.FindMaxSequence(
                        _uiDocument.Document,
                        prefix,
                        targets,
                        categories);
                    next = Math.Max(maxExisting + 1, startAt);
                    nextByPrefix[prefix] = next;
                }

                string value = prefix + next.ToString("D" + width, CultureInfo.InvariantCulture);
                nextByPrefix[prefix] = next + 1;
                SmartNumberingPreviewItem previewItem = new SmartNumberingPreviewItem
                {
                    ElementKey = elementKey,
                    Target = string.Join(", ", targets),
                    CurrentValue = _parameterAccessor.ReadValue(element, targets[0]),
                    ProposedValue = value,
                    Status = SmartNumberingPreviewStates.Planned,
                };
                assignments.Add(new PlannedAssignment(element, value, previewItem));
                result.Preview.Add(previewItem);
            }

            result.PlannedCount = assignments.Count;
            if (previewOnly)
            {
                return result;
            }

            if (result.Errors.Count > 0)
            {
                AbortForValidation(result, assignments, "apply_aborted_validation", "Apply aborted because preview contains errors.");
                return result;
            }

            if (!ValidateAssignments(assignments, targets, result))
            {
                AbortForValidation(result, assignments, "apply_aborted_validation", "Apply aborted because one or more targets are invalid.");
                return result;
            }

            if (assignments.Count == 0)
            {
                return result;
            }

            bool writeFailed = false;
            string writeFailureCode = "apply_failed";
            string writeFailureMessage = "Unexpected error while writing smart numbering values.";
            try
            {
                using (TransactionGroup group = new TransactionGroup(_uiDocument.Document, "MDR Smart Numbering"))
                {
                    group.Start();
                    using (Transaction transaction = new Transaction(_uiDocument.Document, "MDR Smart Numbering Apply"))
                    {
                        transaction.Start();
                        for (int i = 0; i < assignments.Count; i++)
                        {
                            PlannedAssignment assignment = assignments[i];
                            for (int targetIndex = 0; targetIndex < targets.Count; targetIndex++)
                            {
                                if (!_parameterAccessor.TryWriteValue(
                                    assignment.Element,
                                    targets[targetIndex],
                                    assignment.Value,
                                    out string errorCode,
                                    out string errorMessage))
                                {
                                    writeFailed = true;
                                    writeFailureCode = string.IsNullOrWhiteSpace(errorCode)
                                        ? "apply_failed"
                                        : errorCode;
                                    writeFailureMessage = string.IsNullOrWhiteSpace(errorMessage)
                                        ? "Failed to write parameter " + targets[targetIndex]
                                        : errorMessage;
                                    assignment.Preview.Status = SmartNumberingPreviewStates.Error;
                                    assignment.Preview.ErrorCode = writeFailureCode;
                                    assignment.Preview.ErrorMessage = writeFailureMessage;
                                    AddError(
                                        result,
                                        assignment.Preview.ElementKey,
                                        targets[targetIndex],
                                        writeFailureCode,
                                        writeFailureMessage);
                                    break;
                                }
                            }

                            if (writeFailed)
                            {
                                break;
                            }

                            assignment.Preview.Status = SmartNumberingPreviewStates.Applied;
                            assignment.Preview.ErrorCode = string.Empty;
                            assignment.Preview.ErrorMessage = string.Empty;
                        }

                        if (writeFailed)
                        {
                            transaction.RollBack();
                            group.RollBack();
                        }
                        else
                        {
                            transaction.Commit();
                            group.Assimilate();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                writeFailed = true;
                writeFailureCode = "apply_failed";
                writeFailureMessage = ex.Message;
                AddError(result, string.Empty, string.Join(", ", targets), writeFailureCode, writeFailureMessage);
            }

            if (writeFailed)
            {
                result.FailedCount++;
                AbortForRollback(result, assignments, "transaction_rolled_back", writeFailureMessage);
                return result;
            }

            result.AppliedCount = assignments.Count;
            return result;
        }

        private static IReadOnlyList<string> ResolveTargets(SmartNumberingRule rule)
        {
            if (rule.Targets.Count > 0)
            {
                return rule.Targets;
            }

            return new[] { "Serial No" };
        }

        private string BuildPrefix(
            SmartNumberingFormula formula,
            Element element,
            List<string> missingPlaceholders)
        {
            StringBuilder builder = new StringBuilder();
            for (int i = 0; i < formula.Tokens.Count; i++)
            {
                SmartNumberingToken token = formula.Tokens[i];
                if (token.Kind.Equals("sequence", StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                if (token.Kind.Equals("literal", StringComparison.OrdinalIgnoreCase))
                {
                    builder.Append(token.Value);
                    continue;
                }

                if (!TryResolvePlaceholderValue(element, token.Value, out string value))
                {
                    missingPlaceholders.Add(token.Value ?? string.Empty);
                    continue;
                }

                builder.Append(value ?? string.Empty);
            }

            return builder.ToString();
        }

        private bool TryResolvePlaceholderValue(Element element, string placeholder, out string value)
        {
            string key = (placeholder ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(key))
            {
                value = string.Empty;
                return false;
            }

            if (key.Equals("CategoryCode", StringComparison.OrdinalIgnoreCase))
            {
                value = ToShortCode(element.Category?.Name ?? string.Empty, 2);
                return !string.IsNullOrWhiteSpace(value);
            }

            if (key.Equals("SubcategoryCode", StringComparison.OrdinalIgnoreCase))
            {
                value = ToShortCode(_parameterAccessor.ReadValue(element, "Type Mark"), 2);
                return !string.IsNullOrWhiteSpace(value);
            }

            if (key.Equals("ElementId", StringComparison.OrdinalIgnoreCase))
            {
                value = element.Id.Value.ToString(CultureInfo.InvariantCulture);
                return true;
            }

            string parameterValue = _parameterAccessor.ReadValue(element, key);
            if (!string.IsNullOrWhiteSpace(parameterValue))
            {
                value = Sanitize(parameterValue);
                return !string.IsNullOrWhiteSpace(value);
            }

            if (key.Equals("Level", StringComparison.OrdinalIgnoreCase))
            {
                Level? level = element.Document.GetElement(element.LevelId) as Level;
                if (level != null)
                {
                    value = Sanitize(level.Name ?? string.Empty);
                    return !string.IsNullOrWhiteSpace(value);
                }
            }

            value = string.Empty;
            return false;
        }

        private static string ToShortCode(string value, int maxLength)
        {
            string text = Sanitize(value);
            if (string.IsNullOrWhiteSpace(text))
            {
                return string.Empty;
            }

            if (text.Length <= maxLength)
            {
                return text;
            }

            return text.Substring(0, maxLength);
        }

        private static string Sanitize(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            string trimmed = value.Trim().ToUpperInvariant();
            StringBuilder builder = new StringBuilder(trimmed.Length);
            for (int i = 0; i < trimmed.Length; i++)
            {
                char c = trimmed[i];
                if (char.IsLetterOrDigit(c))
                {
                    builder.Append(c);
                }
            }

            return builder.ToString();
        }

        private bool ValidateAssignments(
            IReadOnlyList<PlannedAssignment> assignments,
            IReadOnlyList<string> targets,
            SmartNumberingResult result)
        {
            bool hasValidationErrors = false;
            for (int i = 0; i < assignments.Count; i++)
            {
                PlannedAssignment assignment = assignments[i];
                for (int targetIndex = 0; targetIndex < targets.Count; targetIndex++)
                {
                    string target = targets[targetIndex];
                    if (_parameterAccessor.CanWriteValue(
                        assignment.Element,
                        target,
                        assignment.Value,
                        out string errorCode,
                        out string errorMessage))
                    {
                        continue;
                    }

                    hasValidationErrors = true;
                    assignment.Preview.Status = SmartNumberingPreviewStates.Error;
                    assignment.Preview.ErrorCode = string.IsNullOrWhiteSpace(errorCode)
                        ? "parameter_read_only"
                        : errorCode;
                    assignment.Preview.ErrorMessage = string.IsNullOrWhiteSpace(errorMessage)
                        ? "Target parameter is not writable: " + target
                        : errorMessage;
                    AddError(
                        result,
                        assignment.Preview.ElementKey,
                        target,
                        assignment.Preview.ErrorCode,
                        assignment.Preview.ErrorMessage);
                    result.FailedCount++;
                    break;
                }
            }

            return !hasValidationErrors;
        }

        private static void AbortForValidation(
            SmartNumberingResult result,
            IReadOnlyList<PlannedAssignment> assignments,
            string errorCode,
            string message)
        {
            result.WasRolledBack = true;
            result.FatalErrorCode = errorCode;
            result.FatalErrorMessage = message;
            AddError(result, string.Empty, string.Empty, errorCode, message);
            result.SkippedCount += assignments.Count;
            MarkRemainingAsSkipped(assignments, message);
        }

        private static void AbortForRollback(
            SmartNumberingResult result,
            IReadOnlyList<PlannedAssignment> assignments,
            string errorCode,
            string message)
        {
            result.WasRolledBack = true;
            result.FatalErrorCode = errorCode;
            result.FatalErrorMessage = message;
            AddError(result, string.Empty, string.Empty, errorCode, message);
            result.SkippedCount += assignments.Count;
            MarkRemainingAsSkipped(assignments, "Changes were rolled back.");
            result.AppliedCount = 0;
        }

        private static void MarkRemainingAsSkipped(
            IReadOnlyList<PlannedAssignment> assignments,
            string message)
        {
            for (int i = 0; i < assignments.Count; i++)
            {
                SmartNumberingPreviewItem preview = assignments[i].Preview;
                if (preview.Status.Equals(SmartNumberingPreviewStates.Error, StringComparison.OrdinalIgnoreCase))
                {
                    continue;
                }

                preview.Status = SmartNumberingPreviewStates.Skipped;
                if (string.IsNullOrWhiteSpace(preview.ErrorMessage))
                {
                    preview.ErrorMessage = message;
                }
            }
        }

        private static void AddError(
            SmartNumberingResult result,
            string elementKey,
            string target,
            string errorCode,
            string message)
        {
            result.Errors.Add(new SmartNumberingError
            {
                ElementKey = elementKey ?? string.Empty,
                Target = target ?? string.Empty,
                ErrorCode = string.IsNullOrWhiteSpace(errorCode) ? "internal_error" : errorCode,
                Message = message ?? string.Empty,
            });
        }

        private static string GetElementKey(Element element)
        {
            return !string.IsNullOrWhiteSpace(element.UniqueId)
                ? element.UniqueId
                : element.Id.Value.ToString(CultureInfo.InvariantCulture);
        }

        private sealed class PlannedAssignment
        {
            public PlannedAssignment(Element element, string value, SmartNumberingPreviewItem preview)
            {
                Element = element ?? throw new ArgumentNullException(nameof(element));
                Value = value ?? string.Empty;
                Preview = preview ?? throw new ArgumentNullException(nameof(preview));
            }

            public Element Element { get; }

            public string Value { get; }

            public SmartNumberingPreviewItem Preview { get; }
        }
    }
}
