using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using ClosedXML.Excel;
using Mdr.Revit.Core.Contracts;
using Mdr.Revit.Core.Models;

namespace Mdr.Revit.Client.Excel
{
    public sealed class ExcelWorkbookClient : IExcelWorkbookClient
    {
        public Task<ExcelWorkbookReadResult> ReadRowsAsync(
            ExcelWorkbookProfile profile,
            CancellationToken cancellationToken)
        {
            if (profile == null)
            {
                throw new ArgumentNullException(nameof(profile));
            }

            cancellationToken.ThrowIfCancellationRequested();
            string filePath = NormalizeRequiredPath(profile.FilePath);
            if (!File.Exists(filePath))
            {
                throw new FileNotFoundException("Excel workbook was not found.", filePath);
            }

            using (XLWorkbook workbook = new XLWorkbook(filePath))
            {
                IXLWorksheet worksheet = ResolveWorksheet(workbook, profile.WorksheetName);
                ExcelWorkbookReadResult result = ParseWorksheet(worksheet, profile.AnchorColumn);
                return Task.FromResult(result);
            }
        }

        public Task<ExcelWorkbookWriteResult> WriteRowsAsync(
            ExcelWorkbookProfile profile,
            IReadOnlyList<ScheduleSyncRow> rows,
            CancellationToken cancellationToken)
        {
            if (profile == null)
            {
                throw new ArgumentNullException(nameof(profile));
            }

            cancellationToken.ThrowIfCancellationRequested();
            string filePath = NormalizeRequiredPath(profile.FilePath);
            string worksheetName = NormalizeWorksheetName(profile.WorksheetName);
            rows ??= Array.Empty<ScheduleSyncRow>();

            string? directory = Path.GetDirectoryName(filePath);
            if (!string.IsNullOrWhiteSpace(directory))
            {
                Directory.CreateDirectory(directory);
            }

            IReadOnlyList<string> headers = ResolveHeaders(profile, rows);
            using (XLWorkbook workbook = new XLWorkbook())
            {
                IXLWorksheet worksheet = workbook.Worksheets.Add(worksheetName);
                WriteHeaders(worksheet, headers);
                WriteRows(worksheet, headers, rows);
                worksheet.Columns().AdjustToContents();
                workbook.SaveAs(filePath);
            }

            return Task.FromResult(new ExcelWorkbookWriteResult
            {
                FilePath = filePath,
                WorksheetName = worksheetName,
                UpdatedRows = rows.Count,
            });
        }

        private static ExcelWorkbookReadResult ParseWorksheet(
            IXLWorksheet worksheet,
            string anchorColumn)
        {
            ExcelWorkbookReadResult result = new ExcelWorkbookReadResult();
            IXLRange? usedRange = worksheet.RangeUsed();
            if (usedRange == null)
            {
                return result;
            }

            int firstRow = usedRange.RangeAddress.FirstAddress.RowNumber;
            int lastRow = usedRange.RangeAddress.LastAddress.RowNumber;
            int firstColumn = usedRange.RangeAddress.FirstAddress.ColumnNumber;
            int lastColumn = usedRange.RangeAddress.LastAddress.ColumnNumber;
            if (lastRow < firstRow || lastColumn < firstColumn)
            {
                return result;
            }

            List<string> rawHeaders = new List<string>();
            for (int column = firstColumn; column <= lastColumn; column++)
            {
                rawHeaders.Add(NormalizeHeader(worksheet.Cell(firstRow, column).GetFormattedString()));
            }

            IReadOnlyList<string> headers = BuildUniqueHeaders(rawHeaders);
            result.Headers = headers.ToArray();
            int anchorIndex = FindHeaderIndex(headers, anchorColumn);
            int elementIdIndex = FindHeaderIndex(headers, "MDR_ELEMENT_ID");
            HashSet<string> seenAnchors = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            for (int row = firstRow + 1; row <= lastRow; row++)
            {
                ScheduleSyncRow parsed = new ScheduleSyncRow();
                bool hasAnyValue = false;
                for (int offset = 0; offset < headers.Count; offset++)
                {
                    string header = headers[offset] ?? string.Empty;
                    if (string.IsNullOrWhiteSpace(header))
                    {
                        continue;
                    }

                    string value = worksheet.Cell(row, firstColumn + offset).GetFormattedString() ?? string.Empty;
                    if (!string.IsNullOrWhiteSpace(value))
                    {
                        hasAnyValue = true;
                    }

                    parsed.Cells[header] = value;
                }

                if (!hasAnyValue)
                {
                    continue;
                }

                if (elementIdIndex >= 0 && elementIdIndex < headers.Count &&
                    parsed.Cells.TryGetValue(headers[elementIdIndex], out string elementId))
                {
                    parsed.ElementId = elementId ?? string.Empty;
                }

                if (anchorIndex < 0 || anchorIndex >= headers.Count ||
                    !parsed.Cells.TryGetValue(headers[anchorIndex], out string anchorValue) ||
                    string.IsNullOrWhiteSpace(anchorValue))
                {
                    parsed.ChangeState = ScheduleSyncStates.Error;
                    parsed.Errors.Add(new ScheduleSyncError
                    {
                        Code = "anchor_missing",
                        Message = "Anchor column value is missing.",
                    });
                }
                else
                {
                    parsed.AnchorUniqueId = anchorValue.Trim();
                    if (!seenAnchors.Add(parsed.AnchorUniqueId))
                    {
                        parsed.ChangeState = ScheduleSyncStates.Error;
                        parsed.Errors.Add(new ScheduleSyncError
                        {
                            Code = "anchor_duplicate",
                            Message = "Anchor value is duplicated in Excel workbook.",
                        });
                    }
                }

                result.Rows.Add(parsed);
            }

            return result;
        }

        private static void WriteHeaders(IXLWorksheet worksheet, IReadOnlyList<string> headers)
        {
            for (int column = 0; column < headers.Count; column++)
            {
                IXLCell cell = worksheet.Cell(1, column + 1);
                cell.Value = headers[column] ?? string.Empty;
                cell.Style.Font.Bold = true;
            }

            worksheet.SheetView.FreezeRows(1);
        }

        private static void WriteRows(
            IXLWorksheet worksheet,
            IReadOnlyList<string> headers,
            IReadOnlyList<ScheduleSyncRow> rows)
        {
            for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++)
            {
                ScheduleSyncRow row = rows[rowIndex];
                for (int headerIndex = 0; headerIndex < headers.Count; headerIndex++)
                {
                    string header = headers[headerIndex];
                    string value;
                    if (header.Equals("MDR_UNIQUE_ID", StringComparison.OrdinalIgnoreCase))
                    {
                        value = row.AnchorUniqueId ?? string.Empty;
                    }
                    else if (header.Equals("MDR_ELEMENT_ID", StringComparison.OrdinalIgnoreCase))
                    {
                        value = row.ElementId ?? string.Empty;
                    }
                    else if (!row.Cells.TryGetValue(header, out value))
                    {
                        value = string.Empty;
                    }

                    worksheet.Cell(rowIndex + 2, headerIndex + 1).Value = value ?? string.Empty;
                }
            }
        }

        private static IReadOnlyList<string> ResolveHeaders(
            ExcelWorkbookProfile profile,
            IReadOnlyList<ScheduleSyncRow> rows)
        {
            List<string> headers = new List<string>();
            AddHeader(headers, profile.AnchorColumn);
            AddHeader(headers, "MDR_ELEMENT_ID");

            for (int i = 0; i < profile.ColumnMappings.Count; i++)
            {
                AddHeader(headers, profile.ColumnMappings[i].SheetColumn);
            }

            for (int i = 0; i < rows.Count; i++)
            {
                foreach (KeyValuePair<string, string> pair in rows[i].Cells)
                {
                    AddHeader(headers, pair.Key);
                }
            }

            return headers;
        }

        private static IXLWorksheet ResolveWorksheet(XLWorkbook workbook, string worksheetName)
        {
            if (!string.IsNullOrWhiteSpace(worksheetName) &&
                workbook.Worksheets.TryGetWorksheet(worksheetName.Trim(), out IXLWorksheet? requested))
            {
                return requested;
            }

            IXLWorksheet? first = workbook.Worksheets.FirstOrDefault();
            if (first == null)
            {
                throw new InvalidOperationException("Excel workbook does not contain any worksheets.");
            }

            return first;
        }

        private static IReadOnlyList<string> BuildUniqueHeaders(IReadOnlyList<string> rawHeaders)
        {
            List<string> unique = new List<string>();
            Dictionary<string, int> counters = new Dictionary<string, int>(StringComparer.OrdinalIgnoreCase);
            for (int i = 0; i < rawHeaders.Count; i++)
            {
                string header = NormalizeHeader(rawHeaders[i]);
                if (string.IsNullOrWhiteSpace(header))
                {
                    header = "COL_" + (i + 1).ToString(System.Globalization.CultureInfo.InvariantCulture);
                }

                if (!counters.TryGetValue(header, out int count))
                {
                    counters[header] = 1;
                    unique.Add(header);
                    continue;
                }

                count++;
                counters[header] = count;
                unique.Add(header + "_" + count.ToString(System.Globalization.CultureInfo.InvariantCulture));
            }

            return unique;
        }

        private static int FindHeaderIndex(IReadOnlyList<string> headers, string expected)
        {
            for (int i = 0; i < headers.Count; i++)
            {
                if (string.Equals(headers[i], expected, StringComparison.OrdinalIgnoreCase))
                {
                    return i;
                }
            }

            return -1;
        }

        private static string NormalizeRequiredPath(string filePath)
        {
            string value = (filePath ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(value))
            {
                throw new InvalidOperationException("Excel file path is required.");
            }

            return Path.GetFullPath(Environment.ExpandEnvironmentVariables(value.Replace("/", "\\")));
        }

        private static string NormalizeWorksheetName(string worksheetName)
        {
            string value = (worksheetName ?? string.Empty).Trim();
            if (string.IsNullOrWhiteSpace(value))
            {
                value = "Sheet1";
            }

            char[] invalid = { '\\', '/', '?', '*', '[', ']', ':' };
            for (int i = 0; i < invalid.Length; i++)
            {
                value = value.Replace(invalid[i], '_');
            }

            return value.Length <= 31 ? value : value.Substring(0, 31);
        }

        private static string NormalizeHeader(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return string.Empty;
            }

            string[] tokens = value
                .Split(new[] { ' ', '\t', '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries);
            return string.Join(" ", tokens).Trim();
        }

        private static void AddHeader(List<string> headers, string value)
        {
            string normalized = NormalizeHeader(value);
            if (string.IsNullOrWhiteSpace(normalized))
            {
                return;
            }

            if (headers.Any(x => x.Equals(normalized, StringComparison.OrdinalIgnoreCase)))
            {
                return;
            }

            headers.Add(normalized);
        }
    }
}
