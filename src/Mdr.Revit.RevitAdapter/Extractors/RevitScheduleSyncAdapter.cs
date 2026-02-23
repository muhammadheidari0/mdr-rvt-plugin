using System;
using System.Collections.Generic;
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

            TableData tableData = schedule.GetTableData();
            TableSectionData body = tableData.GetSectionData(SectionType.Body);
            if (body == null || body.NumberOfRows <= 0 || body.NumberOfColumns <= 0)
            {
                return Array.Empty<ScheduleSyncRow>();
            }

            TableSectionData headerSection = tableData.GetSectionData(SectionType.Header);
            int headerRow = headerSection != null && headerSection.NumberOfRows > 0
                ? headerSection.LastRowNumber
                : body.FirstRowNumber;

            List<string> headers = new List<string>();
            for (int col = body.FirstColumnNumber; col <= body.LastColumnNumber; col++)
            {
                string header = schedule.GetCellText(SectionType.Header, headerRow, col);
                if (string.IsNullOrWhiteSpace(header))
                {
                    header = "COL_" + (col - body.FirstColumnNumber + 1);
                }

                headers.Add(header.Trim());
            }

            List<ScheduleSyncRow> rows = new List<ScheduleSyncRow>();
            for (int row = body.FirstRowNumber; row <= body.LastRowNumber; row++)
            {
                ScheduleSyncRow data = new ScheduleSyncRow();
                for (int col = body.FirstColumnNumber; col <= body.LastColumnNumber; col++)
                {
                    string key = headers[col - body.FirstColumnNumber];
                    string value = schedule.GetCellText(SectionType.Body, row, col) ?? string.Empty;
                    data.Cells[key] = value.Trim();
                }

                string elementIdText = ResolveElementId(data.Cells);
                data.ElementId = elementIdText;
                data.AnchorUniqueId = ResolveUniqueId(_uiDocument.Document, elementIdText, data.Cells, profile.AnchorColumn);
                data.Cells[profile.AnchorColumn] = data.AnchorUniqueId;
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
            List<ViewSchedule> schedules = new FilteredElementCollector(document)
                .OfClass(typeof(ViewSchedule))
                .Cast<ViewSchedule>()
                .Where(x => !x.IsTemplate)
                .ToList();

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

        private static string ResolveElementId(IReadOnlyDictionary<string, string> cells)
        {
            string[] candidateKeys =
            {
                "Element Id",
                "ElementID",
                "ID",
                "ElementId",
                "MDR_ELEMENT_ID",
            };

            for (int i = 0; i < candidateKeys.Length; i++)
            {
                if (cells.TryGetValue(candidateKeys[i], out string value) && !string.IsNullOrWhiteSpace(value))
                {
                    return value.Trim();
                }
            }

            return string.Empty;
        }

        private static string ResolveUniqueId(
            Document document,
            string elementIdText,
            IReadOnlyDictionary<string, string> cells,
            string anchorColumn)
        {
            if (int.TryParse(elementIdText, out int elementId))
            {
                Element? element = document.GetElement(new ElementId(elementId));
                if (element != null && !string.IsNullOrWhiteSpace(element.UniqueId))
                {
                    return element.UniqueId;
                }
            }

            if (!string.IsNullOrWhiteSpace(anchorColumn) &&
                cells.TryGetValue(anchorColumn, out string anchorValue) &&
                !string.IsNullOrWhiteSpace(anchorValue))
            {
                return anchorValue.Trim();
            }

            return string.Empty;
        }
    }
}
