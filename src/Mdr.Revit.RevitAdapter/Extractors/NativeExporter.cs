using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Mdr.Revit.RevitAdapter.Extractors
{
    public sealed class NativeExporter
    {
        private readonly Func<IReadOnlyList<string>, string, IReadOnlyList<string>>? _revitExporter;

        public NativeExporter()
        {
        }

        public NativeExporter(Func<IReadOnlyList<string>, string, IReadOnlyList<string>> revitExporter)
        {
            _revitExporter = revitExporter ?? throw new ArgumentNullException(nameof(revitExporter));
        }

        public IReadOnlyList<string> ExportNativeFiles(IReadOnlyList<string> sheetIds, string outputDirectory)
        {
            if (sheetIds == null)
            {
                throw new ArgumentNullException(nameof(sheetIds));
            }

            if (string.IsNullOrWhiteSpace(outputDirectory))
            {
                throw new ArgumentException("Output directory is required.", nameof(outputDirectory));
            }

            Directory.CreateDirectory(outputDirectory);

            if (_revitExporter == null)
            {
                return ExportPlaceholder(sheetIds, outputDirectory);
            }

            return _revitExporter(sheetIds, outputDirectory);
        }

        private static IReadOnlyList<string> ExportPlaceholder(IReadOnlyList<string> sheetIds, string outputDirectory)
        {
            List<string> results = new List<string>(sheetIds.Count);
            for (int i = 0; i < sheetIds.Count; i++)
            {
                string sheetId = string.IsNullOrWhiteSpace(sheetIds[i]) ? ("sheet_" + i) : sheetIds[i].Trim();
                string fileName = "i" + i + "_native_" + SanitizeToken(sheetId) + ".dwg";
                string filePath = Path.Combine(outputDirectory, fileName);

                string text =
                    "MDR_NATIVE_PLACEHOLDER\n" +
                    "SHEET_ID=" + sheetId + "\n" +
                    "GENERATED_UTC=" + DateTimeOffset.UtcNow.ToString("o") + "\n";
                File.WriteAllBytes(filePath, Encoding.UTF8.GetBytes(text));

                results.Add(filePath);
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
