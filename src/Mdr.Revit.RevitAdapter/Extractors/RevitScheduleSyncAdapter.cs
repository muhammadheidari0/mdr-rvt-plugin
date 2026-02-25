using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Mdr.Revit.Core.Contracts;
using Mdr.Revit.Core.Models;
using Mdr.Revit.RevitAdapter.Helpers;

namespace Mdr.Revit.RevitAdapter.Extractors
{
    public sealed class RevitScheduleSyncAdapter : IRevitScheduleSyncAdapter
    {
        private readonly UIDocument? _uiDocument;
        private readonly ParameterAccessor _parameterAccessor;

        public RevitScheduleSyncAdapter()
            : this(null, new ParameterAccessor())
        {
        }

        public RevitScheduleSyncAdapter(UIDocument? uiDocument)
            : this(uiDocument, new ParameterAccessor())
        {
        }

        internal RevitScheduleSyncAdapter(
            UIDocument? uiDocument,
            ParameterAccessor parameterAccessor)
        {
            _uiDocument = uiDocument;
            _parameterAccessor = parameterAccessor ?? throw new ArgumentNullException(nameof(parameterAccessor));
        }

        public IReadOnlyList<string> GetAvailableScheduleNames()
        {
            if (_uiDocument?.Document == null)
            {
                return Array.Empty<string>();
            }

            return ResolveSchedules(_uiDocument.Document)
                .Select(x => x.Name ?? string.Empty)
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(x => x, StringComparer.OrdinalIgnoreCase)
                .ToList();
        }

        public IReadOnlyList<GoogleSheetColumnMapping> GetScheduleColumnMappings(string scheduleName)
        {
            if (_uiDocument?.Document == null)
            {
                return Array.Empty<GoogleSheetColumnMapping>();
            }

            ViewSchedule? schedule = ResolveSchedule(_uiDocument.Document, scheduleName);
            if (schedule == null)
            {
                return Array.Empty<GoogleSheetColumnMapping>();
            }

            List<ScheduleColumnDefinition> columns = ResolveColumnsFromDefinition(schedule);
            List<GoogleSheetColumnMapping> mappings = new List<GoogleSheetColumnMapping>(columns.Count);
            for (int i = 0; i < columns.Count; i++)
            {
                ScheduleColumnDefinition column = columns[i];
                if (string.IsNullOrWhiteSpace(column.Header))
                {
                    continue;
                }

                mappings.Add(new GoogleSheetColumnMapping
                {
                    SheetColumn = column.Header,
                    RevitParameter = string.IsNullOrWhiteSpace(column.ParameterName)
                        ? column.Header
                        : column.ParameterName,
                    IsEditable = true,
                });
            }

            return mappings;
        }

        public IReadOnlyList<ScheduleSyncRow> ExtractRows(
            string scheduleName,
            GoogleSheetSyncProfile profile)
        {
            if (_uiDocument?.Document == null)
            {
                return Array.Empty<ScheduleSyncRow>();
            }

            ViewSchedule? schedule = ResolveSchedule(_uiDocument.Document, scheduleName);
            if (schedule == null)
            {
                return Array.Empty<ScheduleSyncRow>();
            }

            if (!schedule.Definition.IsItemized)
            {
                return new[]
                {
                    BuildErrorRow(
                        "schedule_not_itemized",
                        "Schedule is not itemized. Enable 'Itemize every instance' and retry."),
                };
            }

            List<ScheduleColumnDefinition> columns = ResolveColumnsFromDefinition(schedule);
            string anchorColumn = string.IsNullOrWhiteSpace(profile?.AnchorColumn)
                ? "MDR_UNIQUE_ID"
                : profile.AnchorColumn.Trim();

            List<Element> elements = new FilteredElementCollector(_uiDocument.Document, schedule.Id)
                .WhereElementIsNotElementType()
                .ToElements()
                .Where(x => x != null)
                .OrderBy(x => x.Id.Value)
                .ToList();

            List<ScheduleSyncRow> rows = new List<ScheduleSyncRow>(elements.Count);
            for (int i = 0; i < elements.Count; i++)
            {
                Element element = elements[i];
                ScheduleSyncRow data = new ScheduleSyncRow();
                for (int c = 0; c < columns.Count; c++)
                {
                    ScheduleColumnDefinition column = columns[c];
                    string value = _parameterAccessor.ReadValue(element, column.ParameterName);
                    if (string.IsNullOrWhiteSpace(value) &&
                        !string.Equals(column.ParameterName, column.Header, StringComparison.OrdinalIgnoreCase))
                    {
                        value = _parameterAccessor.ReadValue(element, column.Header);
                    }

                    data.Cells[column.Header] = (value ?? string.Empty).Trim();
                }

                string elementIdText = element.Id.Value.ToString(CultureInfo.InvariantCulture);
                data.ElementId = elementIdText;
                data.AnchorUniqueId = (element.UniqueId ?? string.Empty).Trim();
                data.Cells[anchorColumn] = data.AnchorUniqueId;
                data.Cells["MDR_ELEMENT_ID"] = elementIdText;

                if (string.IsNullOrWhiteSpace(data.AnchorUniqueId))
                {
                    data.ChangeState = ScheduleSyncStates.Error;
                    data.Errors.Add(new ScheduleSyncError
                    {
                        Code = "anchor_missing",
                        Message = "Unable to resolve element unique id for schedule row.",
                    });
                }

                rows.Add(data);
            }

            int aggregateRows = CountAggregateRows(schedule, rows.Count);
            for (int i = 0; i < aggregateRows; i++)
            {
                rows.Add(BuildErrorRow(
                    "aggregate_row_skipped",
                    "A schedule row could not be mapped to a single element and was skipped."));
            }

            return rows;
        }

        public ScheduleSyncDiffResult BuildDiff(
            IReadOnlyList<ScheduleSyncRow> incomingRows,
            GoogleSheetSyncProfile profile)
        {
            ScheduleSyncDiffResult diff = new ScheduleSyncDiffResult();
            if (_uiDocument?.Document == null || incomingRows == null || incomingRows.Count == 0)
            {
                return diff;
            }

            for (int i = 0; i < incomingRows.Count; i++)
            {
                ScheduleSyncRow row = incomingRows[i].Clone();
                if (row.Errors.Count > 0 || string.IsNullOrWhiteSpace(row.AnchorUniqueId))
                {
                    row.ChangeState = ScheduleSyncStates.Error;
                    row.Errors.Add(new ScheduleSyncError
                    {
                        Code = "anchor_missing",
                        Message = "Anchor is required.",
                    });
                    diff.ErrorRowsCount++;
                    diff.Rows.Add(row);
                    continue;
                }

                Element? element = _uiDocument.Document.GetElement(row.AnchorUniqueId);
                if (element == null)
                {
                    row.ChangeState = ScheduleSyncStates.Error;
                    row.Errors.Add(new ScheduleSyncError
                    {
                        Code = "element_not_found",
                        Message = "Element not found in Revit model.",
                    });
                    diff.ErrorRowsCount++;
                    diff.Rows.Add(row);
                    continue;
                }

                bool changed = false;
                for (int m = 0; m < profile.ColumnMappings.Count; m++)
                {
                    GoogleSheetColumnMapping mapping = profile.ColumnMappings[m];
                    if (!mapping.IsEditable)
                    {
                        continue;
                    }

                    if (!row.Cells.TryGetValue(mapping.SheetColumn, out string incomingValue))
                    {
                        continue;
                    }

                    string current = _parameterAccessor.ReadValue(element, mapping.RevitParameter);
                    if (string.Equals(current, incomingValue ?? string.Empty, StringComparison.Ordinal))
                    {
                        continue;
                    }

                    if (!_parameterAccessor.CanWriteValue(
                        element,
                        mapping.RevitParameter,
                        incomingValue ?? string.Empty,
                        out string code,
                        out string message))
                    {
                        row.ChangeState = ScheduleSyncStates.Error;
                        row.Errors.Add(new ScheduleSyncError
                        {
                            Code = string.IsNullOrWhiteSpace(code) ? "type_mismatch" : code,
                            Message = string.IsNullOrWhiteSpace(message) ? "Value cannot be applied." : message,
                        });
                        continue;
                    }

                    changed = true;
                }

                if (row.Errors.Count > 0)
                {
                    row.ChangeState = ScheduleSyncStates.Error;
                    diff.ErrorRowsCount++;
                }
                else if (changed)
                {
                    row.ChangeState = ScheduleSyncStates.Modified;
                    diff.ChangedRowsCount++;
                }
                else
                {
                    row.ChangeState = ScheduleSyncStates.Unchanged;
                }

                diff.Rows.Add(row);
            }

            return diff;
        }

        public ScheduleSyncApplyResult ApplyDiff(
            ScheduleSyncDiffResult diff,
            GoogleSheetSyncProfile profile)
        {
            ScheduleSyncApplyResult result = new ScheduleSyncApplyResult();
            if (_uiDocument?.Document == null || diff == null || diff.Rows.Count == 0)
            {
                return result;
            }

            bool hasChanges = diff.Rows.Any(x => string.Equals(x.ChangeState, ScheduleSyncStates.Modified, StringComparison.OrdinalIgnoreCase));
            if (!hasChanges)
            {
                return result;
            }

            using (Transaction transaction = new Transaction(_uiDocument.Document, "MDR Google Sheets Sync Apply"))
            {
                transaction.Start();
                for (int i = 0; i < diff.Rows.Count; i++)
                {
                    ScheduleSyncRow row = diff.Rows[i];
                    if (!string.Equals(row.ChangeState, ScheduleSyncStates.Modified, StringComparison.OrdinalIgnoreCase))
                    {
                        continue;
                    }

                    Element? element = _uiDocument.Document.GetElement(row.AnchorUniqueId);
                    if (element == null)
                    {
                        result.FailedCount++;
                        result.Errors.Add(new ScheduleSyncError
                        {
                            Code = "element_not_found",
                            Message = "Element not found for anchor " + row.AnchorUniqueId,
                        });
                        continue;
                    }

                    bool failed = false;
                    for (int m = 0; m < profile.ColumnMappings.Count; m++)
                    {
                        GoogleSheetColumnMapping mapping = profile.ColumnMappings[m];
                        if (!mapping.IsEditable)
                        {
                            continue;
                        }

                        if (!row.Cells.TryGetValue(mapping.SheetColumn, out string value))
                        {
                            continue;
                        }

                        string current = _parameterAccessor.ReadValue(element, mapping.RevitParameter);
                        if (string.Equals(current, value ?? string.Empty, StringComparison.Ordinal))
                        {
                            continue;
                        }

                        if (!_parameterAccessor.TryWriteValue(
                            element,
                            mapping.RevitParameter,
                            value ?? string.Empty,
                            out string code,
                            out string message))
                        {
                            failed = true;
                            result.Errors.Add(new ScheduleSyncError
                            {
                                Code = string.IsNullOrWhiteSpace(code) ? "apply_failed" : code,
                                Message = string.IsNullOrWhiteSpace(message) ? "Failed to apply parameter update." : message,
                            });
                            break;
                        }
                    }

                    if (failed)
                    {
                        result.FailedCount++;
                    }
                    else
                    {
                        result.AppliedCount++;
                    }
                }

                transaction.Commit();
            }

            return result;
        }

        private static ViewSchedule? ResolveSchedule(Document document, string scheduleName)
        {
            List<ViewSchedule> schedules = ResolveSchedules(document);

            if (schedules.Count == 0)
            {
                return null;
            }

            string normalizedName = (scheduleName ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(normalizedName))
            {
                return schedules[0];
            }

            ViewSchedule? exact = schedules.FirstOrDefault(x =>
                string.Equals(x.Name ?? string.Empty, normalizedName, StringComparison.OrdinalIgnoreCase));
            if (exact != null)
            {
                return exact;
            }

            return schedules.FirstOrDefault(x =>
                (x.Name ?? string.Empty).IndexOf(normalizedName, StringComparison.OrdinalIgnoreCase) >= 0);
        }

        private static List<ScheduleColumnDefinition> ResolveColumnsFromDefinition(ViewSchedule schedule)
        {
            List<string> headers = new List<string>();
            List<string> parameterNames = new List<string>();
            IList<ScheduleFieldId> fieldOrder = schedule.Definition.GetFieldOrder();
            for (int i = 0; i < fieldOrder.Count; i++)
            {
                ScheduleField field = schedule.Definition.GetField(fieldOrder[i]);
                if (field == null || field.IsHidden)
                {
                    continue;
                }

                string heading = NormalizeHeaderLabel(field.ColumnHeading);
                string parameterName = ResolveFieldParameterName(field, heading);
                headers.Add(heading);
                parameterNames.Add(parameterName);
            }

            IReadOnlyList<string> uniqueHeaders = BuildUniqueHeaders(headers);
            List<ScheduleColumnDefinition> columns = new List<ScheduleColumnDefinition>(uniqueHeaders.Count);
            for (int i = 0; i < uniqueHeaders.Count; i++)
            {
                columns.Add(new ScheduleColumnDefinition
                {
                    Header = uniqueHeaders[i],
                    ParameterName = parameterNames[i],
                });
            }

            return columns;
        }

        internal static IReadOnlyList<string> BuildUniqueHeaders(IReadOnlyList<string> rawHeaders)
        {
            List<string> unique = new List<string>();
            if (rawHeaders == null || rawHeaders.Count == 0)
            {
                return unique;
            }

            Dictionary<string, int> counters = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < rawHeaders.Count; i++)
            {
                string normalized = NormalizeHeaderLabel(rawHeaders[i]);
                if (string.IsNullOrWhiteSpace(normalized))
                {
                    normalized = "COL_" + (i + 1).ToString(CultureInfo.InvariantCulture);
                }

                if (!counters.TryGetValue(normalized, out int count))
                {
                    counters[normalized] = 1;
                    unique.Add(normalized);
                    continue;
                }

                count++;
                counters[normalized] = count;
                unique.Add(normalized + "_" + count.ToString(CultureInfo.InvariantCulture));
            }

            return unique;
        }

        private static string ResolveFieldParameterName(ScheduleField field, string fallback)
        {
            string value = string.Empty;
            try
            {
                value = NormalizeHeaderLabel(field.GetName());
            }
            catch
            {
                value = string.Empty;
            }

            if (string.IsNullOrWhiteSpace(value))
            {
                return fallback;
            }

            return value;
        }

        private static string NormalizeHeaderLabel(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            string[] tokens = value
                .Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            return string.Join(" ", tokens).Trim();
        }

        private static int CountAggregateRows(ViewSchedule schedule, int elementRows)
        {
            try
            {
                TableData table = schedule.GetTableData();
                TableSectionData body = table.GetSectionData(SectionType.Body);
                if (body == null || body.NumberOfRows <= 0)
                {
                    return 0;
                }

                int bodyRows = body.LastRowNumber - body.FirstRowNumber + 1;
                if (bodyRows <= elementRows)
                {
                    return 0;
                }

                return bodyRows - elementRows;
            }
            catch
            {
                return 0;
            }
        }

        private static ScheduleSyncRow BuildErrorRow(string code, string message)
        {
            ScheduleSyncRow row = new ScheduleSyncRow
            {
                ChangeState = ScheduleSyncStates.Error,
            };
            row.Errors.Add(new ScheduleSyncError
            {
                Code = code ?? string.Empty,
                Message = message ?? string.Empty,
            });
            return row;
        }

        private static List<ViewSchedule> ResolveSchedules(Document document)
        {
            return new FilteredElementCollector(document)
                .OfClass(typeof(ViewSchedule))
                .Cast<ViewSchedule>()
                .Where(x => !x.IsTemplate)
                .ToList();
        }

        private sealed class ScheduleColumnDefinition
        {
            public string Header { get; set; } = string.Empty;

            public string ParameterName { get; set; } = string.Empty;
        }
    }
}
