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
            SmartNumberingResult result = new SmartNumberingResult();
            if (rule == null)
            {
                result.Errors.Add(new SmartNumberingError
                {
                    ErrorCode = "validation_error",
                    Message = "Smart numbering rule is required.",
                });
                return result;
            }

            if (_uiDocument?.Document == null)
            {
                result.Errors.Add(new SmartNumberingError
                {
                    ErrorCode = "revit_context_missing",
                    Message = "Revit document context is unavailable.",
                });
                return result;
            }

            ICollection<ElementId> selectedIds = _uiDocument.Selection.GetElementIds();
            if (selectedIds == null || selectedIds.Count == 0)
            {
                result.Errors.Add(new SmartNumberingError
                {
                    ErrorCode = "selection_required",
                    Message = "Select one or more elements before running smart numbering.",
                });
                return result;
            }

            SmartNumberingFormula formula;
            try
            {
                formula = _formulaParser.Parse(rule.Formula, rule.SequenceWidth);
            }
            catch (Exception ex)
            {
                result.Errors.Add(new SmartNumberingError
                {
                    ErrorCode = "formula_invalid",
                    Message = ex.Message,
                });
                return result;
            }

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
            Dictionary<string, List<(Element Element, string Value)>> assignments =
                new Dictionary<string, List<(Element Element, string Value)>>(StringComparer.OrdinalIgnoreCase);

            int width = formula.SequenceWidth <= 0 ? 5 : formula.SequenceWidth;
            int startAt = rule.StartAt <= 0 ? 1 : rule.StartAt;
            for (int i = 0; i < selectedElements.Count; i++)
            {
                Element element = selectedElements[i];
                string prefix = BuildPrefix(formula, element);
                if (string.IsNullOrWhiteSpace(prefix))
                {
                    result.SkippedCount++;
                    result.Errors.Add(new SmartNumberingError
                    {
                        ElementKey = element.UniqueId ?? element.Id.Value.ToString(CultureInfo.InvariantCulture),
                        ErrorCode = "prefix_empty",
                        Message = "Generated prefix is empty.",
                    });
                    continue;
                }

                if (!nextByPrefix.TryGetValue(prefix, out int next))
                {
                    int maxExisting = _sequenceScanner.FindMaxSequence(
                        _uiDocument.Document,
                        prefix,
                        ResolveTargets(rule),
                        categories);
                    next = Math.Max(maxExisting + 1, startAt);
                    nextByPrefix[prefix] = next;
                }

                string value = prefix + next.ToString("D" + width, CultureInfo.InvariantCulture);
                nextByPrefix[prefix] = next + 1;

                if (!assignments.TryGetValue(prefix, out List<(Element Element, string Value)> items))
                {
                    items = new List<(Element Element, string Value)>();
                    assignments[prefix] = items;
                }

                items.Add((element, value));
                result.Preview.Add(new SmartNumberingPreviewItem
                {
                    ElementKey = element.UniqueId ?? string.Empty,
                    Value = value,
                });
            }

            if (previewOnly)
            {
                return result;
            }

            using (Transaction transaction = new Transaction(_uiDocument.Document, "MDR Smart Numbering"))
            {
                transaction.Start();
                IReadOnlyList<string> targets = ResolveTargets(rule);

                foreach (KeyValuePair<string, List<(Element Element, string Value)>> pair in assignments)
                {
                    _ = pair;
                    foreach ((Element Element, string Value) item in pair.Value)
                    {
                        bool failed = false;
                        for (int targetIndex = 0; targetIndex < targets.Count; targetIndex++)
                        {
                            if (!_parameterAccessor.TryWriteValue(
                                item.Element,
                                targets[targetIndex],
                                item.Value,
                                out string errorCode,
                                out string errorMessage))
                            {
                                failed = true;
                                result.Errors.Add(new SmartNumberingError
                                {
                                    ElementKey = item.Element.UniqueId ?? string.Empty,
                                    ErrorCode = string.IsNullOrWhiteSpace(errorCode) ? "apply_failed" : errorCode,
                                    Message = string.IsNullOrWhiteSpace(errorMessage)
                                        ? "Failed to write parameter " + targets[targetIndex]
                                        : errorMessage,
                                });
                                break;
                            }
                        }

                        if (failed)
                        {
                            result.SkippedCount++;
                        }
                        else
                        {
                            result.AppliedCount++;
                        }
                    }
                }

                transaction.Commit();
            }

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

        private string BuildPrefix(SmartNumberingFormula formula, Element element)
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

                string value = ResolvePlaceholderValue(element, token.Value);
                builder.Append(value);
            }

            return builder.ToString();
        }

        private string ResolvePlaceholderValue(Element element, string placeholder)
        {
            string key = (placeholder ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(key))
            {
                return string.Empty;
            }

            if (key.Equals("CategoryCode", StringComparison.OrdinalIgnoreCase))
            {
                return ToShortCode(element.Category?.Name ?? string.Empty, 2);
            }

            if (key.Equals("SubcategoryCode", StringComparison.OrdinalIgnoreCase))
            {
                return ToShortCode(_parameterAccessor.ReadValue(element, "Type Mark"), 2);
            }

            if (key.Equals("ElementId", StringComparison.OrdinalIgnoreCase))
            {
                return element.Id.Value.ToString(CultureInfo.InvariantCulture);
            }

            string value = _parameterAccessor.ReadValue(element, key);
            if (!string.IsNullOrWhiteSpace(value))
            {
                return Sanitize(value);
            }

            if (key.Equals("Level", StringComparison.OrdinalIgnoreCase))
            {
                Level? level = element.Document.GetElement(element.LevelId) as Level;
                if (level != null)
                {
                    return Sanitize(level.Name ?? string.Empty);
                }
            }

            return string.Empty;
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
    }
}
