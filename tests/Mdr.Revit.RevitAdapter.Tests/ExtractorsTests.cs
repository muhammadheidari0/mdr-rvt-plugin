using System;
using System.IO;
using System.Text;
using Mdr.Revit.Core.Models;
using Mdr.Revit.RevitAdapter.Extractors;
using Xunit;

namespace Mdr.Revit.RevitAdapter.Tests
{
    public sealed class ExtractorsTests
    {
        [Fact]
        public void RevitExtractorAdapter_DelegatesToSheetAndScheduleExtractors()
        {
            SheetExtractor sheetExtractor = new SheetExtractor(() =>
            {
                return new[]
                {
                    new PublishSheetItem
                    {
                        SheetUniqueId = "sheet-1",
                        RequestedRevision = "A",
                        FileSha256 = "aaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaaa",
                    },
                };
            });

            ScheduleExtractor scheduleExtractor = new ScheduleExtractor(profile =>
            {
                return new[]
                {
                    new ScheduleRow
                    {
                        RowNo = 1,
                        ElementKey = profile + "-EL-1",
                    },
                };
            });

            RevitExtractorAdapter adapter = new RevitExtractorAdapter(sheetExtractor, scheduleExtractor);

            var sheets = adapter.GetSelectedSheets();
            var rows = adapter.GetScheduleRows(ScheduleProfiles.Mto);

            Assert.Single(sheets);
            Assert.Equal("sheet-1", sheets[0].SheetUniqueId);
            Assert.Single(rows);
            Assert.Equal("MTO-EL-1", rows[0].ElementKey);
        }

        [Fact]
        public void PdfExporter_ExportsOnePdfPerSheet()
        {
            string outputDirectory = Path.Combine(Path.GetTempPath(), "mdr_pdf_export_" + Guid.NewGuid().ToString("N"));
            try
            {
                PdfExporter exporter = new PdfExporter();
                var files = exporter.ExportSheetsToPdf(
                    new[]
                    {
                        new PublishSheetItem { ItemIndex = 0, SheetUniqueId = "A-101" },
                        new PublishSheetItem { ItemIndex = 1, SheetUniqueId = "A-102" },
                    },
                    outputDirectory);

                Assert.Equal(2, files.Count);
                Assert.All(files, file => Assert.True(file.IsSuccess()));
                Assert.All(files, file => Assert.True(File.Exists(file.FilePath)));

                byte[] bytes = File.ReadAllBytes(files[0].FilePath);
                string prefix = Encoding.ASCII.GetString(bytes, 0, Math.Min(8, bytes.Length));
                Assert.Contains("%PDF-1.4", prefix);
                Assert.False(string.IsNullOrWhiteSpace(files[0].FileSha256));
            }
            finally
            {
                if (Directory.Exists(outputDirectory))
                {
                    Directory.Delete(outputDirectory, recursive: true);
                }
            }
        }

        [Fact]
        public void NativeExporter_ExportsOneNativeFilePerSheet()
        {
            string outputDirectory = Path.Combine(Path.GetTempPath(), "mdr_native_export_" + Guid.NewGuid().ToString("N"));
            try
            {
                NativeExporter exporter = new NativeExporter();
                var files = exporter.ExportNativeFiles(
                    new[]
                    {
                        new PublishSheetItem { ItemIndex = 0, SheetUniqueId = "A-201" },
                    },
                    outputDirectory);

                Assert.Single(files);
                Assert.True(files[0].IsSuccess());
                Assert.True(File.Exists(files[0].FilePath));

                string text = File.ReadAllText(files[0].FilePath);
                Assert.Contains("MDR_NATIVE_PLACEHOLDER", text);
                Assert.Contains("SHEET_ID=A-201", text);
            }
            finally
            {
                if (Directory.Exists(outputDirectory))
                {
                    Directory.Delete(outputDirectory, recursive: true);
                }
            }
        }
    }
}
