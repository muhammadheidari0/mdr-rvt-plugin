using System.Reflection;
using Mdr.Revit.Addin.Commands;
using Mdr.Revit.Core.Models;
using Xunit;

namespace Mdr.Revit.Addin.Tests
{
    public sealed class ExcelSyncCommandProfileTests
    {
        [Fact]
        public void BuildProfile_EnforcesProtectedSystemColumns()
        {
            ExcelSyncCommandRequest request = new ExcelSyncCommandRequest
            {
                FilePath = "C:\\Temp\\mto.xlsx",
                WorksheetName = "MTO",
                AnchorColumn = "MDR_UNIQUE_ID",
            };
            request.ColumnMappings.Add(new GoogleSheetColumnMapping
            {
                SheetColumn = "MDR_UNIQUE_ID",
                RevitParameter = "MDR_UNIQUE_ID",
                IsEditable = true,
            });
            request.ColumnMappings.Add(new GoogleSheetColumnMapping
            {
                SheetColumn = "Comments",
                RevitParameter = "Comments",
                IsEditable = true,
            });
            request.ProtectedColumns.Add("Comments");

            MethodInfo? method = typeof(ExcelSyncCommand).GetMethod(
                "BuildProfile",
                BindingFlags.NonPublic | BindingFlags.Static);

            Assert.NotNull(method);
            ExcelWorkbookProfile profile =
                Assert.IsType<ExcelWorkbookProfile>(method!.Invoke(null, new object[] { request }));

            Assert.Contains(profile.ColumnMappings, x =>
                x.SheetColumn == "MDR_UNIQUE_ID" &&
                !x.IsEditable);
            Assert.Contains(profile.ColumnMappings, x =>
                x.SheetColumn == "MDR_ELEMENT_ID" &&
                !x.IsEditable);
            Assert.Contains(profile.ColumnMappings, x =>
                x.SheetColumn == "Comments" &&
                !x.IsEditable);
        }
    }
}
