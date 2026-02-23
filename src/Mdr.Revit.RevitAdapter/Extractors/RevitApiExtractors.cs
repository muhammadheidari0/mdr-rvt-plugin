using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Mdr.Revit.Core.Contracts;
using Mdr.Revit.Core.Models;
using Mdr.Revit.RevitAdapter.Writers;

namespace Mdr.Revit.RevitAdapter.Extractors
{
    public static class RevitApiExtractors
    {
        public static IRevitExtractor CreateExtractor(UIDocument uiDocument)
        {
            if (uiDocument == null)
            {
                throw new ArgumentNullException(nameof(uiDocument));
            }

            Document document = uiDocument.Document ?? throw new ArgumentNullException(nameof(uiDocument.Document));
            return new RevitExtractorAdapter(
                new SheetExtractor(() => ExtractSelectedSheets(uiDocument)),
                new ScheduleExtractor(profileCode => ExtractScheduleRows(document, profileCode)));
        }

        public static PdfExporter CreatePdfExporter(UIDocument uiDocument)
        {
            if (uiDocument == null)
            {
                throw new ArgumentNullException(nameof(uiDocument));
            }

            Document document = uiDocument.Document ?? throw new ArgumentNullException(nameof(uiDocument.Document));
            return new PdfExporter((sheetIds, outputDirectory) =>
                ExportSheetsToPdf(document, sheetIds, outputDirectory));
        }

        public static NativeExporter CreateNativeExporter(UIDocument uiDocument)
        {
            if (uiDocument == null)
            {
                throw new ArgumentNullException(nameof(uiDocument));
            }

            Document document = uiDocument.Document ?? throw new ArgumentNullException(nameof(uiDocument.Document));
            return new NativeExporter((sheetIds, outputDirectory) =>
                ExportNativeFiles(document, sheetIds, outputDirectory));
        }

        public static IRevitScheduleSyncAdapter CreateScheduleSyncAdapter(UIDocument uiDocument)
        {
            if (uiDocument == null)
            {
                throw new ArgumentNullException(nameof(uiDocument));
            }

            return new RevitScheduleSyncAdapter(uiDocument);
        }

        public static ISmartNumberingEngine CreateSmartNumberingEngine(UIDocument uiDocument)
        {
            if (uiDocument == null)
            {
                throw new ArgumentNullException(nameof(uiDocument));
            }

            return new SmartNumberingEngine(uiDocument);
        }

        private static IReadOnlyList<PublishSheetItem> ExtractSelectedSheets(UIDocument uiDocument)
        {
            Document document = uiDocument.Document;
            ICollection<ElementId> selectedIds = uiDocument.Selection.GetElementIds();
            if (selectedIds == null || selectedIds.Count == 0)
            {
                return Array.Empty<PublishSheetItem>();
            }

            List<ViewSheet> sheets = selectedIds
                .Select(id => document.GetElement(id) as ViewSheet)
                .Where(sheet => sheet != null && !sheet.IsPlaceholder)
                .Cast<ViewSheet>()
                .OrderBy(sheet => sheet.SheetNumber ?? string.Empty, StringComparer.OrdinalIgnoreCase)
                .ToList();

            List<PublishSheetItem> rows = new List<PublishSheetItem>(sheets.Count);
            for (int i = 0; i < sheets.Count; i++)
            {
                ViewSheet sheet = sheets[i];
                rows.Add(new PublishSheetItem
                {
                    ItemIndex = i,
                    SheetUniqueId = sheet.UniqueId ?? string.Empty,
                    SheetNumber = sheet.SheetNumber ?? string.Empty,
                    SheetName = sheet.Name ?? string.Empty,
                    RequestedRevision = ResolveRequestedRevision(sheet),
                    StatusCode = string.Empty,
                    Metadata = new DocumentMetadata
                    {
                        Subject = sheet.Name ?? string.Empty,
                    },
                });
            }

            return rows;
        }

        private static string ResolveRequestedRevision(ViewSheet sheet)
        {
            string[] parameterNames =
            {
                "Current Revision",
                "Revision",
                "REVISION",
            };

            for (int i = 0; i < parameterNames.Length; i++)
            {
                Parameter? parameter = sheet.LookupParameter(parameterNames[i]);
                if (parameter == null)
                {
                    continue;
                }

                string value = parameter.AsString();
                if (string.IsNullOrWhiteSpace(value))
                {
                    value = parameter.AsValueString();
                }

                if (!string.IsNullOrWhiteSpace(value))
                {
                    return value.Trim();
                }
            }

            return "A";
        }

        private static IReadOnlyList<ScheduleRow> ExtractScheduleRows(Document document, string profileCode)
        {
            ViewSchedule? schedule = ResolveSchedule(document, profileCode);
            if (schedule == null)
            {
                return Array.Empty<ScheduleRow>();
            }

            return ExtractScheduleRows(schedule);
        }

        private static ViewSchedule? ResolveSchedule(Document document, string profileCode)
        {
            string profile = profileCode.Trim().ToUpperInvariant();
            string[] preferredTokens = profile == ScheduleProfiles.Equipment
                ? new[] { "EQUIPMENT", "EQP" }
                : new[] { "MTO", "MATERIAL", "QUANTITY" };

            List<ViewSchedule> schedules = new FilteredElementCollector(document)
                .OfClass(typeof(ViewSchedule))
                .Cast<ViewSchedule>()
                .Where(x => !x.IsTemplate)
                .ToList();

            for (int i = 0; i < preferredTokens.Length; i++)
            {
                string token = preferredTokens[i];
                ViewSchedule? match = schedules.FirstOrDefault(x =>
                    (x.Name ?? string.Empty).IndexOf(token, StringComparison.OrdinalIgnoreCase) >= 0);
                if (match != null)
                {
                    return match;
                }
            }

            return schedules.FirstOrDefault();
        }

        private static IReadOnlyList<ScheduleRow> ExtractScheduleRows(ViewSchedule schedule)
        {
            List<ScheduleRow> rows = new List<ScheduleRow>();

            TableData tableData = schedule.GetTableData();
            TableSectionData body = tableData.GetSectionData(SectionType.Body);
            if (body == null || body.NumberOfRows <= 0 || body.NumberOfColumns <= 0)
            {
                return rows;
            }

            TableSectionData header = tableData.GetSectionData(SectionType.Header);
            int headerRow = header != null && header.NumberOfRows > 0
                ? header.LastRowNumber
                : body.FirstRowNumber;

            List<string> headers = new List<string>(body.NumberOfColumns);
            for (int col = body.FirstColumnNumber; col <= body.LastColumnNumber; col++)
            {
                string title = schedule.GetCellText(SectionType.Header, headerRow, col);
                if (string.IsNullOrWhiteSpace(title))
                {
                    title = "COL_" + (col - body.FirstColumnNumber + 1);
                }

                headers.Add(title.Trim());
            }

            int rowNo = 1;
            for (int row = body.FirstRowNumber; row <= body.LastRowNumber; row++)
            {
                ScheduleRow dataRow = new ScheduleRow
                {
                    RowNo = rowNo++,
                };

                for (int col = body.FirstColumnNumber; col <= body.LastColumnNumber; col++)
                {
                    string key = headers[col - body.FirstColumnNumber];
                    string value = schedule.GetCellText(SectionType.Body, row, col) ?? string.Empty;
                    dataRow.Values[key] = value.Trim();
                }

                dataRow.ElementKey = ResolveElementKey(dataRow.Values);
                if (string.IsNullOrWhiteSpace(dataRow.ElementKey))
                {
                    dataRow.ElementKey = "ROW-" + dataRow.RowNo.ToString();
                }

                rows.Add(dataRow);
            }

            return rows;
        }

        private static string ResolveElementKey(Dictionary<string, string> values)
        {
            string[] keys =
            {
                "ElementKey",
                "Element Key",
                "Element Id",
                "ElementID",
                "ID",
            };

            for (int i = 0; i < keys.Length; i++)
            {
                if (!values.TryGetValue(keys[i], out string value))
                {
                    continue;
                }

                if (!string.IsNullOrWhiteSpace(value))
                {
                    return value.Trim();
                }
            }

            foreach (KeyValuePair<string, string> pair in values)
            {
                if (pair.Key.IndexOf("element", StringComparison.OrdinalIgnoreCase) >= 0 &&
                    !string.IsNullOrWhiteSpace(pair.Value))
                {
                    return pair.Value.Trim();
                }
            }

            return string.Empty;
        }

        private static IReadOnlyList<string> ExportSheetsToPdf(
            Document document,
            IReadOnlyList<string> sheetIds,
            string outputDirectory)
        {
            List<ViewSheet> sheets = ResolveSheets(document, sheetIds);
            if (sheets.Count == 0)
            {
                return Array.Empty<string>();
            }

            HashSet<string> before = new HashSet<string>(
                Directory.GetFiles(outputDirectory, "*.pdf", SearchOption.TopDirectoryOnly),
                StringComparer.OrdinalIgnoreCase);

            List<ElementId> ids = sheets.Select(x => x.Id).ToList();
            PDFExportOptions options = new PDFExportOptions
            {
                Combine = false,
            };

            bool ok = document.Export(outputDirectory, ids, options);
            if (!ok)
            {
                throw new InvalidOperationException("Revit PDF export failed.");
            }

            List<string> produced = Directory.GetFiles(outputDirectory, "*.pdf", SearchOption.TopDirectoryOnly)
                .Where(x => !before.Contains(x))
                .OrderBy(x => x, StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (produced.Count == 0)
            {
                return Directory.GetFiles(outputDirectory, "*.pdf", SearchOption.TopDirectoryOnly)
                    .OrderByDescending(File.GetLastWriteTimeUtc)
                    .Take(sheetIds.Count)
                    .OrderBy(x => x, StringComparer.OrdinalIgnoreCase)
                    .ToList();
            }

            return produced;
        }

        private static IReadOnlyList<string> ExportNativeFiles(
            Document document,
            IReadOnlyList<string> sheetIds,
            string outputDirectory)
        {
            List<ViewSheet> sheets = ResolveSheets(document, sheetIds);
            if (sheets.Count == 0)
            {
                return Array.Empty<string>();
            }

            List<string> results = new List<string>(sheets.Count);
            for (int i = 0; i < sheets.Count; i++)
            {
                ViewSheet sheet = sheets[i];
                string exportName = "i" + i + "_native_" + SanitizeToken(sheet.SheetNumber ?? sheet.UniqueId);

                bool ok = document.Export(
                    outputDirectory,
                    exportName,
                    new List<ElementId> { sheet.Id },
                    new DWGExportOptions());
                if (!ok)
                {
                    continue;
                }

                string candidate = Path.Combine(outputDirectory, exportName + ".dwg");
                if (File.Exists(candidate))
                {
                    results.Add(candidate);
                    continue;
                }

                string[] allDwgs = Directory.GetFiles(outputDirectory, "*.dwg", SearchOption.TopDirectoryOnly);
                string? latest = allDwgs
                    .OrderByDescending(File.GetLastWriteTimeUtc)
                    .FirstOrDefault();
                if (!string.IsNullOrWhiteSpace(latest))
                {
                    results.Add(latest);
                }
            }

            return results;
        }

        private static List<ViewSheet> ResolveSheets(Document document, IReadOnlyList<string> sheetIds)
        {
            Dictionary<string, ViewSheet> byUniqueId = new Dictionary<string, ViewSheet>(StringComparer.OrdinalIgnoreCase);
            Dictionary<string, ViewSheet> bySheetNumber = new Dictionary<string, ViewSheet>(StringComparer.OrdinalIgnoreCase);

            List<ViewSheet> allSheets = new FilteredElementCollector(document)
                .OfClass(typeof(ViewSheet))
                .Cast<ViewSheet>()
                .Where(x => !x.IsPlaceholder)
                .ToList();

            for (int i = 0; i < allSheets.Count; i++)
            {
                ViewSheet sheet = allSheets[i];
                if (!string.IsNullOrWhiteSpace(sheet.UniqueId))
                {
                    byUniqueId[sheet.UniqueId] = sheet;
                }

                if (!string.IsNullOrWhiteSpace(sheet.SheetNumber))
                {
                    bySheetNumber[sheet.SheetNumber] = sheet;
                }
            }

            List<ViewSheet> resolved = new List<ViewSheet>(sheetIds.Count);
            for (int i = 0; i < sheetIds.Count; i++)
            {
                string token = (sheetIds[i] ?? string.Empty).Trim();
                if (string.IsNullOrWhiteSpace(token))
                {
                    continue;
                }

                if (byUniqueId.TryGetValue(token, out ViewSheet uniqueSheet))
                {
                    resolved.Add(uniqueSheet);
                    continue;
                }

                if (bySheetNumber.TryGetValue(token, out ViewSheet numberSheet))
                {
                    resolved.Add(numberSheet);
                }
            }

            return resolved;
        }

        private static string SanitizeToken(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
            {
                return "sheet";
            }

            char[] input = value.Trim().ToCharArray();
            char[] output = new char[input.Length];
            int length = 0;

            for (int i = 0; i < input.Length; i++)
            {
                char c = input[i];
                if (char.IsLetterOrDigit(c) || c == '.' || c == '_' || c == '-')
                {
                    output[length++] = c;
                }
                else
                {
                    output[length++] = '_';
                }
            }

            string result = new string(output, 0, length);
            return string.IsNullOrWhiteSpace(result) ? "sheet" : result;
        }
    }
}
