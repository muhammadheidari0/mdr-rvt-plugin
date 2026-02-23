using System;
using System.Collections.Generic;
using System.Globalization;
using Autodesk.Revit.DB;

namespace Mdr.Revit.RevitAdapter.Helpers
{
    public sealed class LocalSequenceScanner
    {
        public int ExtractSequence(string fullValue, string prefix)
        {
            return ParseSequence(fullValue, prefix);
        }

        public int FindMaxSequence(
            Document document,
            string prefix,
            IReadOnlyList<string> parameterNames,
            ISet<ElementId>? categoryIds)
        {
            if (document == null)
            {
                throw new ArgumentNullException(nameof(document));
            }

            if (string.IsNullOrWhiteSpace(prefix))
            {
                return 0;
            }

            IReadOnlyList<string> targets = parameterNames ?? Array.Empty<string>();
            if (targets.Count == 0)
            {
                targets = new[] { "Serial No", "Type Mark" };
            }

            int max = 0;
            FilteredElementCollector collector = new FilteredElementCollector(document).WhereElementIsNotElementType();
            IEnumerable<Element> elements = collector;
            foreach (Element element in elements)
            {
                if (element == null)
                {
                    continue;
                }

                if (categoryIds != null && categoryIds.Count > 0)
                {
                    ElementId? categoryId = element.Category?.Id;
                    if (categoryId == null || !categoryIds.Contains(categoryId))
                    {
                        continue;
                    }
                }

                for (int i = 0; i < targets.Count; i++)
                {
                    string text = ReadParameterText(element, targets[i]);
                    if (string.IsNullOrWhiteSpace(text))
                    {
                        continue;
                    }

                    int parsed = ParseSequence(text, prefix);
                    if (parsed > max)
                    {
                        max = parsed;
                    }
                }
            }

            return max;
        }

        private static int ParseSequence(string fullValue, string prefix)
        {
            string input = (fullValue ?? string.Empty).Trim();
            if (!input.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
            {
                return 0;
            }

            string suffix = input.Substring(prefix.Length);
            if (string.IsNullOrWhiteSpace(suffix))
            {
                return 0;
            }

            string trimmed = suffix.Trim();
            for (int i = 0; i < trimmed.Length; i++)
            {
                if (!char.IsDigit(trimmed[i]))
                {
                    return 0;
                }
            }

            if (string.IsNullOrWhiteSpace(trimmed))
            {
                return 0;
            }

            return int.TryParse(trimmed, NumberStyles.Integer, CultureInfo.InvariantCulture, out int parsed)
                ? parsed
                : 0;
        }

        private static string ReadParameterText(Element element, string parameterName)
        {
            if (string.IsNullOrWhiteSpace(parameterName))
            {
                return string.Empty;
            }

            Parameter? parameter = element.LookupParameter(parameterName);
            if (parameter == null)
            {
                Element typeElement = element.Document.GetElement(element.GetTypeId());
                parameter = typeElement?.LookupParameter(parameterName);
            }

            if (parameter == null)
            {
                return string.Empty;
            }

            string value = parameter.AsString();
            if (!string.IsNullOrWhiteSpace(value))
            {
                return value.Trim();
            }

            value = parameter.AsValueString();
            return value?.Trim() ?? string.Empty;
        }
    }
}
