using Mdr.Revit.Addin.UI;
using Mdr.Revit.Core.Models;
using Xunit;

namespace Mdr.Revit.Addin.Tests
{
    public sealed class GoogleSyncWindowViewModelTests
    {
        [Fact]
        public void SetScheduleNames_WhenSelectionMissing_SelectsFirstSchedule()
        {
            GoogleSyncWindowViewModel vm = new GoogleSyncWindowViewModel
            {
                SelectedScheduleName = "Unknown",
            };

            vm.SetScheduleNames(new[] { "Doors", "Walls" });

            Assert.Equal("Doors", vm.SelectedScheduleName);
            Assert.Equal(2, vm.ScheduleNames.Count);
        }

        [Fact]
        public void ApplyProtectedColumns_MarksMappingsAsReadOnly()
        {
            GoogleSyncWindowViewModel vm = new GoogleSyncWindowViewModel();

            vm.ApplyProtectedColumns(new[] { "MDR_UNIQUE_ID", "CustomLocked" });

            Assert.Contains(vm.ColumnMappings, x => x.SheetColumn == "MDR_UNIQUE_ID" && !x.IsEditable);
            Assert.Contains(vm.ColumnMappings, x => x.SheetColumn == "CustomLocked" && !x.IsEditable);
        }

        [Fact]
        public void BuildRequest_MapsFieldsAndMappings()
        {
            GoogleSyncWindowViewModel vm = new GoogleSyncWindowViewModel
            {
                Direction = "IMPORT",
                SelectedScheduleName = "MTO Main",
                SpreadsheetId = "spreadsheet-1",
                WorksheetName = "SheetA",
                AnchorColumn = "MDR_UNIQUE_ID",
                AuthorizeInteractively = true,
                PreviewOnly = true,
            };

            vm.SetScheduleNames(new[] { "MTO Main" });
            vm.ApplyProtectedColumns(new[] { "MDR_UNIQUE_ID" });

            GoogleSyncFromAppRequest request = vm.BuildRequest();

            Assert.Equal(GoogleSyncDirections.Import, request.Direction);
            Assert.Equal("MTO Main", request.ScheduleName);
            Assert.Equal("spreadsheet-1", request.SpreadsheetId);
            Assert.Equal("SheetA", request.WorksheetName);
            Assert.True(request.AuthorizeInteractively);
            Assert.True(request.PreviewOnly);
            Assert.Contains(request.ColumnMappings, x => x.SheetColumn == "MDR_UNIQUE_ID");
        }
    }
}
