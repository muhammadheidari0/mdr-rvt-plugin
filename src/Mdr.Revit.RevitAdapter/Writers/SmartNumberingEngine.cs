using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
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
        private readonly ArcaSerialNumberingPlanner _arcaPlanner;
        private readonly ArcaLevelCodeResolver _levelCodeResolver;

        public SmartNumberingEngine()
            : this(
                null,
                new ParameterAccessor(),
                new LocalSequenceScanner(),
                new SmartNumberingFormulaParser(),
                new ArcaSerialNumberingPlanner(),
                new ArcaLevelCodeResolver())
        {
        }

        public SmartNumberingEngine(UIDocument uiDocument)
            : this(
                uiDocument,
                new ParameterAccessor(),
                new LocalSequenceScanner(),
                new SmartNumberingFormulaParser(),
                new ArcaSerialNumberingPlanner(),
                new ArcaLevelCodeResolver())
        {
        }

        internal SmartNumberingEngine(
            UIDocument? uiDocument,
            ParameterAccessor parameterAccessor,
            LocalSequenceScanner sequenceScanner,
            SmartNumberingFormulaParser formulaParser,
            ArcaSerialNumberingPlanner arcaPlanner,
            ArcaLevelCodeResolver levelCodeResolver)
        {
            _uiDocument = uiDocument;
            _parameterAccessor = parameterAccessor ?? throw new ArgumentNullException(nameof(parameterAccessor));
            _sequenceScanner = sequenceScanner ?? throw new ArgumentNullException(nameof(sequenceScanner));
            _formulaParser = formulaParser ?? throw new ArgumentNullException(nameof(formulaParser));
            _arcaPlanner = arcaPlanner ?? throw new ArgumentNullException(nameof(arcaPlanner));
            _levelCodeResolver = levelCodeResolver ?? throw new ArgumentNullException(nameof(levelCodeResolver));
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

            if (IsArcaRule(rule))
            {
                return ApplyArca(rule, previewOnly, result);
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

        private SmartNumberingResult ApplyArca(
            SmartNumberingRule rule,
            bool previewOnly,
            SmartNumberingResult result)
        {
            Document document = _uiDocument!.Document;
            if (!TryResolveBuiltInCategory(rule.CategoryBuiltInName, out BuiltInCategory category))
            {
                AddError(result, string.Empty, string.Empty, "category_invalid", "Invalid ARCA category: " + rule.CategoryBuiltInName);
                return result;
            }

            string selectedBlock = (rule.SelectedBlock ?? string.Empty).Trim();
            string selectedLevel = ArcaSerialNumberingPlanner.NormalizeLevelCode(rule.SelectedLevel);
            if (string.IsNullOrWhiteSpace(selectedBlock))
            {
                AddError(result, string.Empty, string.Empty, "block_required", "Select a Block before running ARCA numbering.");
                return result;
            }

            if (string.IsNullOrWhiteSpace(selectedLevel))
            {
                AddError(result, string.Empty, string.Empty, "level_required", "Select a Level before running ARCA numbering.");
                return result;
            }

            List<Element> categoryElements = new FilteredElementCollector(document)
                .OfCategory(category)
                .WhereElementIsNotElementType()
                .Cast<Element>()
                .OrderBy(x => x.Id.Value)
                .ToList();
            Dictionary<string, Element> elementsByKey = categoryElements
                .GroupBy(GetElementKey, StringComparer.OrdinalIgnoreCase)
                .ToDictionary(x => x.Key, x => x.First(), StringComparer.OrdinalIgnoreCase);
            IReadOnlyList<ScopeBoxInfo> scopeBoxes = CollectScopeBoxes(document);

            ArcaSerialNumberingPlanRequest planRequest = new ArcaSerialNumberingPlanRequest
            {
                SelectedBlock = selectedBlock,
                SelectedLevelCode = selectedLevel,
            };
            for (int i = 0; i < categoryElements.Count; i++)
            {
                Element element = categoryElements[i];
                planRequest.Elements.Add(new ArcaSerialNumberingElementSnapshot
                {
                    ElementKey = GetElementKey(element),
                    ElementId = (int)element.Id.Value,
                    Block = _parameterAccessor.ReadValue(element, "Block"),
                    LevelCode = _levelCodeResolver.GetEffectiveLevelCode(element),
                    TypeMark = ReadTypeMark(element),
                    Series = _parameterAccessor.ReadValue(element, "Series"),
                    SerialNo = _parameterAccessor.ReadValue(element, "Serial No"),
                    ScopeSort = ResolveScopeSort(element, scopeBoxes),
                });
            }

            ArcaSerialNumberingPlan plan = _arcaPlanner.Plan(planRequest);
            for (int i = 0; i < plan.Preview.Count; i++)
            {
                result.Preview.Add(plan.Preview[i]);
            }

            result.PlannedCount = plan.Writes.Count;
            result.SkippedCount = plan.SkippedCount;
            if (previewOnly)
            {
                return result;
            }

            List<ParameterAssignment> assignments = new List<ParameterAssignment>(plan.Writes.Count);
            for (int i = 0; i < plan.Writes.Count; i++)
            {
                ArcaSerialNumberingWrite write = plan.Writes[i];
                if (!elementsByKey.TryGetValue(write.ElementKey, out Element? element))
                {
                    write.Preview.Status = SmartNumberingPreviewStates.Error;
                    write.Preview.ErrorCode = "element_not_found";
                    write.Preview.ErrorMessage = "Element was not found in Revit document.";
                    AddError(result, write.ElementKey, write.Target, "element_not_found", write.Preview.ErrorMessage);
                    result.FailedCount++;
                    continue;
                }

                assignments.Add(new ParameterAssignment(element, write.Target, write.Value, write.Preview));
            }

            if (result.Errors.Count > 0)
            {
                AbortParameterAssignmentsForValidation(result, assignments, "apply_aborted_validation", "Apply aborted because preview contains errors.");
                return result;
            }

            if (!ValidateParameterAssignments(assignments, result))
            {
                AbortParameterAssignmentsForValidation(result, assignments, "apply_aborted_validation", "Apply aborted because one or more ARCA targets are invalid.");
                return result;
            }

            if (assignments.Count == 0)
            {
                return result;
            }

            bool writeFailed = false;
            string writeFailureMessage = "Unexpected error while writing ARCA numbering values.";
            try
            {
                using (TransactionGroup group = new TransactionGroup(document, "MDR ARCA Smart Numbering"))
                {
                    group.Start();
                    using (Transaction transaction = new Transaction(document, "MDR ARCA Smart Numbering Apply"))
                    {
                        transaction.Start();
                        for (int i = 0; i < assignments.Count; i++)
                        {
                            ParameterAssignment assignment = assignments[i];
                            if (!_parameterAccessor.TryWriteValue(
                                assignment.Element,
                                assignment.Target,
                                assignment.Value,
                                out string errorCode,
                                out string errorMessage))
                            {
                                writeFailed = true;
                                writeFailureMessage = string.IsNullOrWhiteSpace(errorMessage)
                                    ? "Failed to write parameter " + assignment.Target
                                    : errorMessage;
                                assignment.Preview.Status = SmartNumberingPreviewStates.Error;
                                assignment.Preview.ErrorCode = string.IsNullOrWhiteSpace(errorCode) ? "apply_failed" : errorCode;
                                assignment.Preview.ErrorMessage = writeFailureMessage;
                                AddError(result, assignment.Preview.ElementKey, assignment.Target, assignment.Preview.ErrorCode, writeFailureMessage);
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
                writeFailureMessage = ex.Message;
                AddError(result, string.Empty, "Series, Serial No", "apply_failed", writeFailureMessage);
            }

            if (writeFailed)
            {
                result.FailedCount++;
                AbortParameterAssignmentsForRollback(result, assignments, "transaction_rolled_back", writeFailureMessage);
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

        private static bool IsArcaRule(SmartNumberingRule rule)
        {
            return string.Equals(rule.Mode, SmartNumberingModes.Arca, StringComparison.OrdinalIgnoreCase);
        }

        private static bool TryResolveBuiltInCategory(string categoryName, out BuiltInCategory category)
        {
            string value = string.IsNullOrWhiteSpace(categoryName)
                ? "OST_Walls"
                : categoryName.Trim();
            return Enum.TryParse(value, out category);
        }

        private string ReadTypeMark(Element element)
        {
            try
            {
                ElementId typeId = element.GetTypeId();
                if (typeId != ElementId.InvalidElementId)
                {
                    Element? typeElement = element.Document.GetElement(typeId);
                    if (typeElement != null)
                    {
                        string typeValue = _parameterAccessor.ReadValue(typeElement, "Type Mark");
                        if (!string.IsNullOrWhiteSpace(typeValue))
                        {
                            return typeValue.Trim();
                        }
                    }
                }
            }
            catch
            {
            }

            return _parameterAccessor.ReadValue(element, "Type Mark");
        }

        private static IReadOnlyList<ScopeBoxInfo> CollectScopeBoxes(Document document)
        {
            List<ScopeBoxInfo> scopeBoxes = new List<ScopeBoxInfo>();
            try
            {
                FilteredElementCollector collector = new FilteredElementCollector(document)
                    .OfCategory(BuiltInCategory.OST_VolumeOfInterest)
                    .WhereElementIsNotElementType();
                foreach (Element scopeBox in collector)
                {
                    Match match = Regex.Match(scopeBox.Name ?? string.Empty, "SC(\\d+)", RegexOptions.IgnoreCase);
                    if (!match.Success || !int.TryParse(match.Groups[1].Value, out int scopeNumber))
                    {
                        continue;
                    }

                    BoundingBoxXYZ? boundingBox = scopeBox.get_BoundingBox(null);
                    if (boundingBox == null)
                    {
                        continue;
                    }

                    scopeBoxes.Add(new ScopeBoxInfo(scopeNumber, boundingBox));
                }
            }
            catch
            {
                return scopeBoxes;
            }

            return scopeBoxes.OrderBy(x => x.ScopeNumber).ToList();
        }

        private static int ResolveScopeSort(Element element, IReadOnlyList<ScopeBoxInfo> scopeBoxes)
        {
            BoundingBoxXYZ? elementBox = null;
            try
            {
                elementBox = element.get_BoundingBox(null);
            }
            catch
            {
                return 9999;
            }

            if (elementBox == null)
            {
                return 9999;
            }

            for (int i = 0; i < scopeBoxes.Count; i++)
            {
                ScopeBoxInfo scopeBox = scopeBoxes[i];
                if (Overlaps(elementBox, scopeBox.BoundingBox))
                {
                    return scopeBox.ScopeNumber;
                }
            }

            return 9999;
        }

        private static bool Overlaps(BoundingBoxXYZ first, BoundingBoxXYZ second)
        {
            return first.Min.X <= second.Max.X &&
                first.Max.X >= second.Min.X &&
                first.Min.Y <= second.Max.Y &&
                first.Max.Y >= second.Min.Y &&
                first.Min.Z <= second.Max.Z &&
                first.Max.Z >= second.Min.Z;
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

        private bool ValidateParameterAssignments(
            IReadOnlyList<ParameterAssignment> assignments,
            SmartNumberingResult result)
        {
            bool hasValidationErrors = false;
            for (int i = 0; i < assignments.Count; i++)
            {
                ParameterAssignment assignment = assignments[i];
                if (_parameterAccessor.CanWriteValue(
                    assignment.Element,
                    assignment.Target,
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
                    ? "Target parameter is not writable: " + assignment.Target
                    : errorMessage;
                AddError(
                    result,
                    assignment.Preview.ElementKey,
                    assignment.Target,
                    assignment.Preview.ErrorCode,
                    assignment.Preview.ErrorMessage);
                result.FailedCount++;
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

        private static void AbortParameterAssignmentsForValidation(
            SmartNumberingResult result,
            IReadOnlyList<ParameterAssignment> assignments,
            string errorCode,
            string message)
        {
            result.WasRolledBack = true;
            result.FatalErrorCode = errorCode;
            result.FatalErrorMessage = message;
            AddError(result, string.Empty, string.Empty, errorCode, message);
            result.SkippedCount += assignments.Count;
            MarkRemainingParameterAssignmentsAsSkipped(assignments, message);
        }

        private static void AbortParameterAssignmentsForRollback(
            SmartNumberingResult result,
            IReadOnlyList<ParameterAssignment> assignments,
            string errorCode,
            string message)
        {
            result.WasRolledBack = true;
            result.FatalErrorCode = errorCode;
            result.FatalErrorMessage = message;
            AddError(result, string.Empty, string.Empty, errorCode, message);
            result.SkippedCount += assignments.Count;
            MarkRemainingParameterAssignmentsAsSkipped(assignments, "Changes were rolled back.");
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

        private static void MarkRemainingParameterAssignmentsAsSkipped(
            IReadOnlyList<ParameterAssignment> assignments,
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

        private sealed class ParameterAssignment
        {
            public ParameterAssignment(Element element, string target, string value, SmartNumberingPreviewItem preview)
            {
                Element = element ?? throw new ArgumentNullException(nameof(element));
                Target = target ?? string.Empty;
                Value = value ?? string.Empty;
                Preview = preview ?? throw new ArgumentNullException(nameof(preview));
            }

            public Element Element { get; }

            public string Target { get; }

            public string Value { get; }

            public SmartNumberingPreviewItem Preview { get; }
        }

        private sealed class ScopeBoxInfo
        {
            public ScopeBoxInfo(int scopeNumber, BoundingBoxXYZ boundingBox)
            {
                ScopeNumber = scopeNumber;
                BoundingBox = boundingBox ?? throw new ArgumentNullException(nameof(boundingBox));
            }

            public int ScopeNumber { get; }

            public BoundingBoxXYZ BoundingBox { get; }
        }
    }
}
