using System;
using System.Collections.Generic;
using System.IO;
using System.Security.Cryptography;
using System.Text;
using Mdr.Revit.Core.Models;

namespace Mdr.Revit.RevitAdapter.Extractors
{
    public sealed class PdfExporter
    {
        private readonly Func<IReadOnlyList<PublishSheetItem>, string, IReadOnlyList<ExportArtifact>>? _revitExporter;

        public PdfExporter()
        {
        }

        public PdfExporter(Func<IReadOnlyList<PublishSheetItem>, string, IReadOnlyList<ExportArtifact>> revitExporter)
        {
            _revitExporter = revitExporter ?? throw new ArgumentNullException(nameof(revitExporter));
        }

        public IReadOnlyList<ExportArtifact> ExportSheetsToPdf(IReadOnlyList<PublishSheetItem> items, string outputDirectory)
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
                string fileName = "i" + itemIndex + "_pdf_" + SanitizeToken(sheetId) + ".pdf";
                string filePath = Path.Combine(outputDirectory, fileName);

                try
                {
                    WriteMinimalPdf(filePath, sheetId);
                    results.Add(new ExportArtifact
                    {
                        ItemIndex = itemIndex,
                        SheetUniqueId = sheetId,
                        Kind = ExportArtifactKinds.Pdf,
                        FilePath = filePath,
                        FileSha256 = ComputeSha256(filePath),
                    });
                }
                catch (Exception ex)
                {
                    results.Add(new ExportArtifact
                    {
                        ItemIndex = itemIndex,
                        SheetUniqueId = sheetId,
                        Kind = ExportArtifactKinds.Pdf,
                        ErrorCode = "export_pdf_failed",
                        ErrorMessage = ex.Message,
                    });
                }
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

        private static string ComputeSha256(string filePath)
        {
            using (SHA256 sha = SHA256.Create())
            using (FileStream stream = File.OpenRead(filePath))
            {
                byte[] hash = sha.ComputeHash(stream);
                StringBuilder builder = new StringBuilder(hash.Length * 2);
                for (int i = 0; i < hash.Length; i++)
                {
                    builder.Append(hash[i].ToString("x2"));
                }

                return builder.ToString();
            }
        }
    }
}
