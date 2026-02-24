using System.Reflection;
using Mdr.Revit.Addin.Commands;
using Mdr.Revit.Core.Models;
using Xunit;

namespace Mdr.Revit.Addin.Tests
{
    public sealed class GoogleSyncCommandProfileTests
    {
        [Fact]
        public void BuildProfile_EnforcesProtectedSystemColumns()
        {
            GoogleSyncCommandRequest request = new GoogleSyncCommandRequest
            {
                SpreadsheetId = "sheet-id",
                WorksheetName = "Sheet1",
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

            MethodInfo? method = typeof(GoogleSyncCommand).GetMethod(
                "BuildProfile",
                BindingFlags.NonPublic | BindingFlags.Static);

            Assert.NotNull(method);
            GoogleSheetSyncProfile profile =
                Assert.IsType<GoogleSheetSyncProfile>(method!.Invoke(null, new object[] { request }));

            Assert.Contains(profile.ColumnMappings, x =>
                x.SheetColumn == "MDR_UNIQUE_ID" &&
                !x.IsEditable);
            Assert.Contains(profile.ColumnMappings, x =>
                x.SheetColumn == "MDR_ELEMENT_ID" &&
                !x.IsEditable);
            Assert.Contains(profile.ColumnMappings, x =>
                x.SheetColumn == "Comments" &&
                !x.IsEditable);

            Assert.Contains(profile.ProtectedColumns, x => x == "MDR_UNIQUE_ID");
            Assert.Contains(profile.ProtectedColumns, x => x == "MDR_ELEMENT_ID");
        }
    }
}
