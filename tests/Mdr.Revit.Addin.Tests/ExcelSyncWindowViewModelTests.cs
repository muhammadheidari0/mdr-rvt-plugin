using Mdr.Revit.Addin.UI;
using Mdr.Revit.Core.Models;
using Xunit;

namespace Mdr.Revit.Addin.Tests
{
    public sealed class ExcelSyncWindowViewModelTests
    {
        [Fact]
        public void SetScheduleNames_SelectsFirstScheduleAndBuildsDefaultFile()
        {
            ExcelSyncWindowViewModel vm = new ExcelSyncWindowViewModel();
            vm.SetDefaults("C:\\Temp\\Excel", string.Empty, "MDR_UNIQUE_ID");

            vm.SetScheduleNames(new[] { "Walls", "Doors" });

            Assert.Equal("Walls", vm.SelectedScheduleName);
            Assert.Equal("Walls", vm.WorksheetName);
            Assert.EndsWith("Walls.xlsx", vm.FilePath);
        }

        [Fact]
        public void ApplyProtectedColumns_MarksMappingsAsReadOnly()
        {
            ExcelSyncWindowViewModel vm = new ExcelSyncWindowViewModel();

            vm.ApplyProtectedColumns(new[] { "MDR_UNIQUE_ID", "LockedColumn" });

            Assert.Contains(vm.ColumnMappings, x => x.ExcelColumn == "MDR_UNIQUE_ID" && !x.IsEditable);
            Assert.Contains(vm.ColumnMappings, x => x.ExcelColumn == "LockedColumn" && !x.IsEditable);
        }

        [Fact]
        public void BuildRequest_MapsFieldsAndMappings()
        {
            ExcelSyncWindowViewModel vm = new ExcelSyncWindowViewModel
            {
                Direction = "IMPORT",
                SelectedScheduleName = "MTO Main",
                FilePath = "C:\\Temp\\mto.xlsx",
                WorksheetName = "MTO",
                AnchorColumn = "MDR_UNIQUE_ID",
                PreviewOnly = true,
            };
            vm.SetColumnMappings(new[]
            {
                new GoogleSheetColumnMapping
                {
                    SheetColumn = "Comments",
                    RevitParameter = "Comments",
                    IsEditable = true,
                },
            });

            ExcelSyncFromAppRequest request = vm.BuildRequest();

            Assert.Equal(GoogleSyncDirections.Import, request.Direction);
            Assert.Equal("MTO Main", request.ScheduleName);
            Assert.Equal("C:\\Temp\\mto.xlsx", request.FilePath);
            Assert.Equal("MTO", request.WorksheetName);
            Assert.True(request.PreviewOnly);
            Assert.Contains(request.ColumnMappings, x => x.SheetColumn == "Comments" && x.RevitParameter == "Comments");
        }
    }
}
