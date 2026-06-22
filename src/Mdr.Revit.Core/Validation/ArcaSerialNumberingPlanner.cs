using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;
using Mdr.Revit.Core.Models;

namespace Mdr.Revit.Core.Validation
{
    public sealed class ArcaSerialNumberingPlanner
    {
        private const string SeriesTarget = "Series";
        private const string SerialTarget = "Serial No";

        public ArcaSerialNumberingPlan Plan(ArcaSerialNumberingPlanRequest request)
        {
            if (request == null)
            {
                throw new ArgumentNullException(nameof(request));
            }

            ArcaSerialNumberingPlan plan = new ArcaSerialNumberingPlan();
            string selectedBlock = SafeTrim(request.SelectedBlock);
            string selectedLevel = NormalizeLevelCode(request.SelectedLevelCode);
            List<WorkingElement> elements = request.Elements
                .Where(x => x != null)
                .Select(x => new WorkingElement(x))
                .Where(x =>
                    string.Equals(SafeTrim(x.Source.Block), selectedBlock, StringComparison.OrdinalIgnoreCase) &&
                    string.Equals(NormalizeLevelCode(x.Source.LevelCode), selectedLevel, StringComparison.OrdinalIgnoreCase))
                .OrderBy(x => x.Source.ElementId)
                .ToList();

            for (int i = 0; i < elements.Count; i++)
            {
                WorkingElement element = elements[i];
                if (string.IsNullOrWhiteSpace(element.SerialNo))
                {
                    continue;
                }

                string extracted = ExtractLastThreeDigits(element.SerialNo);
                if (string.IsNullOrWhiteSpace(extracted))
                {
                    AddSkipped(plan, element, SeriesTarget, element.Series, element.Series, "Serial No exists but does not end with 3 digits.");
                    continue;
                }

                element.EffectiveSeries = extracted;
                AddWriteIfChanged(plan, element, SeriesTarget, element.Series, extracted);
            }

            Dictionary<string, int> maxSeriesByTypeMark = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < elements.Count; i++)
            {
                WorkingElement element = elements[i];
                if (string.IsNullOrWhiteSpace(element.TypeMark))
                {
                    continue;
                }

                if (!int.TryParse(element.EffectiveSeries, out int seriesNumber))
                {
                    continue;
                }

                if (!maxSeriesByTypeMark.TryGetValue(element.TypeMark, out int current) || seriesNumber > current)
                {
                    maxSeriesByTypeMark[element.TypeMark] = seriesNumber;
                }
            }

            List<WorkingElement> toNumber = elements
                .Where(x =>
                    string.IsNullOrWhiteSpace(x.SerialNo) &&
                    string.IsNullOrWhiteSpace(x.EffectiveSeries) &&
                    !string.IsNullOrWhiteSpace(x.TypeMark))
                .OrderBy(x => x.Source.ScopeSort)
                .ThenBy(x => x.TypeMark, StringComparer.OrdinalIgnoreCase)
                .ThenBy(x => x.Source.ElementId)
                .ToList();
            for (int i = 0; i < toNumber.Count; i++)
            {
                WorkingElement element = toNumber[i];
                int next = maxSeriesByTypeMark.TryGetValue(element.TypeMark, out int current)
                    ? current + 1
                    : 1;
                maxSeriesByTypeMark[element.TypeMark] = next;
                string series = next.ToString("D3");
                element.EffectiveSeries = series;
                AddWriteIfChanged(plan, element, SeriesTarget, element.Series, series);
            }

            for (int i = 0; i < elements.Count; i++)
            {
                WorkingElement element = elements[i];
                if (string.IsNullOrWhiteSpace(element.TypeMark))
                {
                    AddSkipped(plan, element, SerialTarget, element.SerialNo, string.Empty, "Skipped Serial No because Type Mark is empty.");
                    continue;
                }

                if (string.IsNullOrWhiteSpace(element.EffectiveSeries))
                {
                    AddSkipped(plan, element, SerialTarget, element.SerialNo, string.Empty, "Skipped Serial No because Series is empty.");
                    continue;
                }

                if (!int.TryParse(element.EffectiveSeries, out int numericSeries))
                {
                    AddSkipped(plan, element, SerialTarget, element.SerialNo, element.EffectiveSeries, "Skipped Serial No because Series is not numeric.");
                    continue;
                }

                string series = numericSeries.ToString("D3");
                string expectedSerial = selectedBlock + selectedLevel + "-" + element.TypeMark + series;
                AddWriteIfChanged(plan, element, SerialTarget, element.SerialNo, expectedSerial);
            }

            return plan;
        }

        public string ExtractLastThreeDigits(string value)
        {
            string text = SafeTrim(value);
            if (string.IsNullOrWhiteSpace(text))
            {
                return string.Empty;
            }

            Match match = Regex.Match(text, "(\\d{3})$");
            return match.Success ? match.Groups[1].Value : string.Empty;
        }

        public static string NormalizeLevelCode(string value)
        {
            string text = SafeTrim(value);
            int dash = text.IndexOf("-", StringComparison.Ordinal);
            return dash >= 0 ? text.Substring(0, dash).Trim() : text;
        }

        private static void AddWriteIfChanged(
            ArcaSerialNumberingPlan plan,
            WorkingElement element,
            string target,
            string current,
            string proposed)
        {
            if (string.Equals(SafeTrim(current), SafeTrim(proposed), StringComparison.Ordinal))
            {
                AddSkipped(plan, element, target, current, proposed, target + " is already correct.");
                return;
            }

            SmartNumberingPreviewItem preview = new SmartNumberingPreviewItem
            {
                ElementKey = element.Source.ElementKey,
                Target = target,
                CurrentValue = current ?? string.Empty,
                ProposedValue = proposed ?? string.Empty,
                Status = SmartNumberingPreviewStates.Planned,
            };
            plan.Preview.Add(preview);
            plan.Writes.Add(new ArcaSerialNumberingWrite
            {
                ElementKey = element.Source.ElementKey,
                Target = target,
                Value = proposed ?? string.Empty,
                Preview = preview,
            });
        }

        private static void AddSkipped(
            ArcaSerialNumberingPlan plan,
            WorkingElement element,
            string target,
            string current,
            string proposed,
            string message)
        {
            plan.SkippedCount++;
            plan.Preview.Add(new SmartNumberingPreviewItem
            {
                ElementKey = element.Source.ElementKey,
                Target = target,
                CurrentValue = current ?? string.Empty,
                ProposedValue = proposed ?? string.Empty,
                Status = SmartNumberingPreviewStates.Skipped,
                ErrorMessage = message ?? string.Empty,
            });
        }

        private static string SafeTrim(string value)
        {
            return (value ?? string.Empty).Trim();
        }

        private sealed class WorkingElement
        {
            public WorkingElement(ArcaSerialNumberingElementSnapshot source)
            {
                Source = source ?? throw new ArgumentNullException(nameof(source));
                TypeMark = SafeTrim(source.TypeMark);
                Series = SafeTrim(source.Series);
                EffectiveSeries = Series;
                SerialNo = SafeTrim(source.SerialNo);
            }

            public ArcaSerialNumberingElementSnapshot Source { get; }

            public string TypeMark { get; }

            public string Series { get; }

            public string EffectiveSeries { get; set; }

            public string SerialNo { get; }
        }
    }
}
