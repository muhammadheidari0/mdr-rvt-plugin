using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Mdr.Revit.Core.Models;

namespace Mdr.Revit.RevitAdapter.Extractors
{
    public sealed class NativeExporter
    {
        private readonly Func<IReadOnlyList<PublishSheetItem>, string, IReadOnlyList<ExportArtifact>>? _revitExporter;

        public NativeExporter()
        {
        }

        public NativeExporter(Func<IReadOnlyList<PublishSheetItem>, string, IReadOnlyList<ExportArtifact>> revitExporter)
        {
            _revitExporter = revitExporter ?? throw new ArgumentNullException(nameof(revitExporter));
        }

        public IReadOnlyList<ExportArtifact> ExportNativeFiles(IReadOnlyList<PublishSheetItem> items, string outputDirectory)
        {
            if (items == null)
            {
                throw new ArgumentNullException(nameof(items));
            }

            if (string.IsNullOrWhiteSpace(outputDirectory))
            {
                throw new ArgumentException("Output directory is required.", nameof(outputDirectory));
            }

            Directory.CreateDirectory(outputDirectory);

            if (_revitExporter == null)
            {
                return ExportPlaceholder(items, outputDirectory);
            }

            return _revitExporter(items, outputDirectory);
        }

        private static IReadOnlyList<ExportArtifact> ExportPlaceholder(IReadOnlyList<PublishSheetItem> items, string outputDirectory)
        {
            List<ExportArtifact> results = new List<ExportArtifact>(items.Count);
            for (int i = 0; i < items.Count; i++)
            {
                PublishSheetItem item = items[i] ?? new PublishSheetItem { ItemIndex = i };
                int itemIndex = item.ItemIndex < 0 ? i : item.ItemIndex;
                string sheetId = string.IsNullOrWhiteSpace(item.SheetUniqueId) ? ("sheet_" + itemIndex) : item.SheetUniqueId.Trim();
                string fileName = "i" + itemIndex + "_native_" + SanitizeToken(sheetId) + ".dwg";
                string filePath = Path.Combine(outputDirectory, fileName);

                try
                {
                    string text =
                        "MDR_NATIVE_PLACEHOLDER\n" +
                        "SHEET_ID=" + sheetId + "\n" +
                        "GENERATED_UTC=" + DateTimeOffset.UtcNow.ToString("o") + "\n";
                    File.WriteAllBytes(filePath, Encoding.UTF8.GetBytes(text));

                    results.Add(new ExportArtifact
                    {
                        ItemIndex = itemIndex,
                        SheetUniqueId = sheetId,
                        Kind = ExportArtifactKinds.Native,
                        FilePath = filePath,
                    });
                }
                catch (Exception ex)
                {
                    results.Add(new ExportArtifact
                    {
                        ItemIndex = itemIndex,
                        SheetUniqueId = sheetId,
                        Kind = ExportArtifactKinds.Native,
                        ErrorCode = "export_native_failed",
                        ErrorMessage = ex.Message,
                    });
                }
            }

            return results;
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
