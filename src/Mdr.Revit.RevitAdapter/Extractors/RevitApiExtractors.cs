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
            return new PdfExporter((items, outputDirectory) =>
                ExportSheetsToPdf(document, items, outputDirectory));
        }

        public static NativeExporter CreateNativeExporter(UIDocument uiDocument)
        {
            if (uiDocument == null)
            {
                throw new ArgumentNullException(nameof(uiDocument));
            }

            Document document = uiDocument.Document ?? throw new ArgumentNullException(nameof(uiDocument.Document));
            return new NativeExporter((items, outputDirectory) =>
                ExportNativeFiles(document, items, outputDirectory));
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

        private static IReadOnlyList<ExportArtifact> ExportSheetsToPdf(
            Document document,
            IReadOnlyList<PublishSheetItem> items,
            string outputDirectory)
        {
            if (items == null || items.Count == 0)
            {
                return Array.Empty<ExportArtifact>();
            }

            Dictionary<int, ViewSheet> sheets = ResolveSheetsByItemIndex(document, items);
            List<ExportArtifact> artifacts = new List<ExportArtifact>(items.Count);
            PDFExportOptions options = new PDFExportOptions
            {
                Combine = false,
            };

            for (int i = 0; i < items.Count; i++)
            {
                PublishSheetItem item = items[i] ?? new PublishSheetItem { ItemIndex = i };
                int itemIndex = item.ItemIndex < 0 ? i : item.ItemIndex;
                string sheetUniqueId = item.SheetUniqueId ?? string.Empty;
                if (!sheets.TryGetValue(itemIndex, out ViewSheet? sheet))
                {
                    artifacts.Add(CreateExportFailure(itemIndex, sheetUniqueId, ExportArtifactKinds.Pdf, "sheet_not_found", "Sheet not found in Revit model."));
                    continue;
                }

                try
                {
                    HashSet<string> before = SnapshotFiles(outputDirectory, "*.pdf");
                    bool ok = document.Export(
                        outputDirectory,
                        new List<ElementId> { sheet.Id },
                        options);
                    if (!ok)
                    {
                        artifacts.Add(CreateExportFailure(itemIndex, sheetUniqueId, ExportArtifactKinds.Pdf, "export_pdf_failed", "Revit PDF export returned false."));
                        continue;
                    }

                    string? producedFile = ResolveProducedFile(outputDirectory, "*.pdf", before);
                    if (string.IsNullOrWhiteSpace(producedFile) || !File.Exists(producedFile))
                    {
                        artifacts.Add(CreateExportFailure(itemIndex, sheetUniqueId, ExportArtifactKinds.Pdf, "export_pdf_failed", "Revit did not produce a PDF file."));
                        continue;
                    }

                    string targetName = "i" + itemIndex + "_pdf_" + SanitizeToken(sheet.SheetNumber ?? sheet.UniqueId) + ".pdf";
                    string targetPath = Path.Combine(outputDirectory, targetName);
                    if (!string.Equals(producedFile, targetPath, StringComparison.OrdinalIgnoreCase))
                    {
                        if (File.Exists(targetPath))
                        {
                            File.Delete(targetPath);
                        }

                        File.Move(producedFile, targetPath);
                    }

                    artifacts.Add(new ExportArtifact
                    {
                        ItemIndex = itemIndex,
                        SheetUniqueId = sheetUniqueId,
                        Kind = ExportArtifactKinds.Pdf,
                        FilePath = targetPath,
                        FileSha256 = ComputeSha256(targetPath),
                    });
                }
                catch (Exception ex)
                {
                    artifacts.Add(CreateExportFailure(itemIndex, sheetUniqueId, ExportArtifactKinds.Pdf, "export_pdf_failed", ex.Message));
                }
            }

            return artifacts;
        }

        private static IReadOnlyList<ExportArtifact> ExportNativeFiles(
            Document document,
            IReadOnlyList<PublishSheetItem> items,
            string outputDirectory)
        {
            if (items == null || items.Count == 0)
            {
                return Array.Empty<ExportArtifact>();
            }

            Dictionary<int, ViewSheet> sheets = ResolveSheetsByItemIndex(document, items);
            List<ExportArtifact> artifacts = new List<ExportArtifact>(items.Count);

            for (int i = 0; i < items.Count; i++)
            {
                PublishSheetItem item = items[i] ?? new PublishSheetItem { ItemIndex = i };
                int itemIndex = item.ItemIndex < 0 ? i : item.ItemIndex;
                string sheetUniqueId = item.SheetUniqueId ?? string.Empty;
                if (!sheets.TryGetValue(itemIndex, out ViewSheet? sheet))
                {
                    artifacts.Add(CreateExportFailure(itemIndex, sheetUniqueId, ExportArtifactKinds.Native, "sheet_not_found", "Sheet not found in Revit model."));
                    continue;
                }

                try
                {
                    string exportName = "i" + itemIndex + "_native_" + SanitizeToken(sheet.SheetNumber ?? sheet.UniqueId);
                    HashSet<string> before = SnapshotFiles(outputDirectory, "*.dwg");

                    bool ok = document.Export(
                        outputDirectory,
                        exportName,
                        new List<ElementId> { sheet.Id },
                        new DWGExportOptions());
                    if (!ok)
                    {
                        artifacts.Add(CreateExportFailure(itemIndex, sheetUniqueId, ExportArtifactKinds.Native, "export_native_failed", "Revit DWG export returned false."));
                        continue;
                    }

                    string candidate = Path.Combine(outputDirectory, exportName + ".dwg");
                    if (!File.Exists(candidate))
                    {
                        string? produced = ResolveProducedFile(outputDirectory, "*.dwg", before);
                        if (!string.IsNullOrWhiteSpace(produced))
                        {
                            if (File.Exists(candidate))
                            {
                                File.Delete(candidate);
                            }

                            File.Move(produced, candidate);
                        }
                    }

                    if (!File.Exists(candidate))
                    {
                        artifacts.Add(CreateExportFailure(itemIndex, sheetUniqueId, ExportArtifactKinds.Native, "export_native_failed", "Revit did not produce a DWG file."));
                        continue;
                    }

                    artifacts.Add(new ExportArtifact
                    {
                        ItemIndex = itemIndex,
                        SheetUniqueId = sheetUniqueId,
                        Kind = ExportArtifactKinds.Native,
                        FilePath = candidate,
                    });
                }
                catch (Exception ex)
                {
                    artifacts.Add(CreateExportFailure(itemIndex, sheetUniqueId, ExportArtifactKinds.Native, "export_native_failed", ex.Message));
                }
            }

            return artifacts;
        }

        private static Dictionary<int, ViewSheet> ResolveSheetsByItemIndex(
            Document document,
            IReadOnlyList<PublishSheetItem> items)
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

            Dictionary<int, ViewSheet> resolved = new Dictionary<int, ViewSheet>();
            for (int i = 0; i < items.Count; i++)
            {
                PublishSheetItem item = items[i] ?? new PublishSheetItem { ItemIndex = i };
                int itemIndex = item.ItemIndex < 0 ? i : item.ItemIndex;

                if (!string.IsNullOrWhiteSpace(item.SheetUniqueId) &&
                    byUniqueId.TryGetValue(item.SheetUniqueId.Trim(), out ViewSheet? uniqueSheet))
                {
                    resolved[itemIndex] = uniqueSheet;
                    continue;
                }

                if (!string.IsNullOrWhiteSpace(item.SheetNumber) &&
                    bySheetNumber.TryGetValue(item.SheetNumber.Trim(), out ViewSheet? numberSheet))
                {
                    resolved[itemIndex] = numberSheet;
                }
            }

            return resolved;
        }

        private static HashSet<string> SnapshotFiles(string directory, string pattern)
        {
            return new HashSet<string>(
                Directory.GetFiles(directory, pattern, SearchOption.TopDirectoryOnly),
                StringComparer.OrdinalIgnoreCase);
        }

        private static string? ResolveProducedFile(
            string directory,
            string pattern,
            HashSet<string> beforeFiles)
        {
            string[] files = Directory.GetFiles(directory, pattern, SearchOption.TopDirectoryOnly);
            string? newestAdded = files
                .Where(x => !beforeFiles.Contains(x))
                .OrderByDescending(File.GetLastWriteTimeUtc)
                .FirstOrDefault();
            if (!string.IsNullOrWhiteSpace(newestAdded))
            {
                return newestAdded;
            }

            return files
                .OrderByDescending(File.GetLastWriteTimeUtc)
                .FirstOrDefault();
        }

        private static ExportArtifact CreateExportFailure(
            int itemIndex,
            string sheetUniqueId,
            string kind,
            string errorCode,
            string errorMessage)
        {
            return new ExportArtifact
            {
                ItemIndex = itemIndex,
                SheetUniqueId = sheetUniqueId ?? string.Empty,
                Kind = kind ?? string.Empty,
                ErrorCode = errorCode ?? string.Empty,
                ErrorMessage = errorMessage ?? string.Empty,
            };
        }

        private static string ComputeSha256(string filePath)
        {
            using (System.Security.Cryptography.SHA256 sha = System.Security.Cryptography.SHA256.Create())
            using (FileStream stream = File.OpenRead(filePath))
            {
                byte[] hash = sha.ComputeHash(stream);
                char[] chars = new char[hash.Length * 2];
                int offset = 0;
                for (int i = 0; i < hash.Length; i++)
                {
                    string hex = hash[i].ToString("x2");
                    chars[offset++] = hex[0];
                    chars[offset++] = hex[1];
                }

                return new string(chars);
            }
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
