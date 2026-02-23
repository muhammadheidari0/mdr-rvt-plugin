using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace Mdr.Revit.RevitAdapter.Extractors
{
    public sealed class PdfExporter
    {
        private readonly Func<IReadOnlyList<string>, string, IReadOnlyList<string>>? _revitExporter;

        public PdfExporter()
        {
        }

        public PdfExporter(Func<IReadOnlyList<string>, string, IReadOnlyList<string>> revitExporter)
        {
            _revitExporter = revitExporter ?? throw new ArgumentNullException(nameof(revitExporter));
        }

        public IReadOnlyList<string> ExportSheetsToPdf(IReadOnlyList<string> sheetIds, string outputDirectory)
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
                string fileName = "i" + i + "_pdf_" + SanitizeToken(sheetId) + ".pdf";
                string filePath = Path.Combine(outputDirectory, fileName);

                WriteMinimalPdf(filePath, sheetId);
                results.Add(filePath);
            }

            return results;
        }

        private static void WriteMinimalPdf(string path, string sheetId)
        {
            string safeLabel = string.IsNullOrWhiteSpace(sheetId) ? "unknown" : sheetId;
            string body =
                "%PDF-1.4\n" +
                "1 0 obj\n" +
                "<< /Type /Catalog /Pages 2 0 R >>\n" +
                "endobj\n" +
                "2 0 obj\n" +
                "<< /Type /Pages /Count 1 /Kids [3 0 R] >>\n" +
                "endobj\n" +
                "3 0 obj\n" +
                "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] >>\n" +
                "endobj\n" +
                "% " + safeLabel + "\n" +
                "trailer\n" +
                "<< /Root 1 0 R >>\n" +
                "%%EOF\n";

            File.WriteAllBytes(path, Encoding.ASCII.GetBytes(body));
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
