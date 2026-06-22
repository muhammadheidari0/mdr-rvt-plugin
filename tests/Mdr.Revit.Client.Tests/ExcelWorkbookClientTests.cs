using System;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Mdr.Revit.Client.Excel;
using Mdr.Revit.Core.Models;
using Xunit;

namespace Mdr.Revit.Client.Tests
{
    public sealed class ExcelWorkbookClientTests
    {
        [Fact]
        public async Task WriteThenReadRows_PreservesSystemColumnsAndCells()
        {
            string path = Path.Combine(Path.GetTempPath(), "mdr-excel-" + Guid.NewGuid().ToString("N") + ".xlsx");
            try
            {
                ExcelWorkbookClient client = new ExcelWorkbookClient();
                ExcelWorkbookProfile profile = NewProfile(path);
                ScheduleSyncRow row = new ScheduleSyncRow
                {
                    AnchorUniqueId = "uid-1",
                    ElementId = "42",
                };
                row.Cells["Comments"] = "Ready";

                ExcelWorkbookWriteResult write = await client.WriteRowsAsync(
                    profile,
                    new[] { row },
                    CancellationToken.None);
                ExcelWorkbookReadResult read = await client.ReadRowsAsync(profile, CancellationToken.None);

                Assert.Equal(1, write.UpdatedRows);
                Assert.True(File.Exists(path));
                Assert.Contains("MDR_UNIQUE_ID", read.Headers);
                Assert.Contains("MDR_ELEMENT_ID", read.Headers);
                Assert.Contains("Comments", read.Headers);
                ScheduleSyncRow parsed = Assert.Single(read.Rows);
                Assert.Equal("uid-1", parsed.AnchorUniqueId);
                Assert.Equal("42", parsed.ElementId);
                Assert.Equal("Ready", parsed.Cells["Comments"]);
            }
            finally
            {
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
            }
        }

        [Fact]
        public async Task WriteRows_TrimsAndDeduplicatesHeaders()
        {
            string path = Path.Combine(Path.GetTempPath(), "mdr-excel-" + Guid.NewGuid().ToString("N") + ".xlsx");
            try
            {
                ExcelWorkbookClient client = new ExcelWorkbookClient();
                ExcelWorkbookProfile profile = NewProfile(path);
                profile.ColumnMappings.Add(new GoogleSheetColumnMapping
                {
                    SheetColumn = "  Room   Name  ",
                    RevitParameter = "Name",
                    IsEditable = true,
                });
                ScheduleSyncRow row = new ScheduleSyncRow
                {
                    AnchorUniqueId = "uid-1",
                    ElementId = "1",
                };
                row.Cells["Room Name"] = "A101";

                await client.WriteRowsAsync(profile, new[] { row }, CancellationToken.None);
                ExcelWorkbookReadResult read = await client.ReadRowsAsync(profile, CancellationToken.None);

                Assert.Contains("Room Name", read.Headers);
                Assert.Equal(read.Headers.Count, read.Headers.Distinct(StringComparer.OrdinalIgnoreCase).Count());
            }
            finally
            {
                if (File.Exists(path))
                {
                    File.Delete(path);
                }
            }
        }

        private static ExcelWorkbookProfile NewProfile(string path)
        {
            ExcelWorkbookProfile profile = new ExcelWorkbookProfile
            {
                FilePath = path,
                WorksheetName = "MTO",
                AnchorColumn = "MDR_UNIQUE_ID",
            };
            profile.ColumnMappings.Add(new GoogleSheetColumnMapping
            {
                SheetColumn = "Comments",
                RevitParameter = "Comments",
                IsEditable = true,
            });
            return profile;
        }
    }
}
