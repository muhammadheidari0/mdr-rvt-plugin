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
        public void SetColumnMappings_UsesScheduleDefinitionMappingsAndKeepsProtectedSystemColumns()
        {
            GoogleSyncWindowViewModel vm = new GoogleSyncWindowViewModel();

            vm.SetColumnMappings(new[]
            {
                new GoogleSheetColumnMapping
                {
                    SheetColumn = "Room Name",
                    RevitParameter = "Name",
                    IsEditable = true,
                },
                new GoogleSheetColumnMapping
                {
                    SheetColumn = "Room Number",
                    RevitParameter = "Number",
                    IsEditable = true,
                },
            });
            vm.ApplyProtectedColumns(new[] { "MDR_UNIQUE_ID", "MDR_ELEMENT_ID" });

            Assert.Contains(vm.ColumnMappings, x => x.SheetColumn == "Room Name" && x.RevitParameter == "Name");
            Assert.Contains(vm.ColumnMappings, x => x.SheetColumn == "Room Number" && x.RevitParameter == "Number");
            Assert.Contains(vm.ColumnMappings, x => x.SheetColumn == "MDR_UNIQUE_ID" && !x.IsEditable);
            Assert.Contains(vm.ColumnMappings, x => x.SheetColumn == "MDR_ELEMENT_ID" && !x.IsEditable);
        }

        [Fact]
        public void SetColumnMappings_WhenEmpty_FallsBackToDefaultSystemMappings()
        {
            GoogleSyncWindowViewModel vm = new GoogleSyncWindowViewModel();

            vm.SetColumnMappings(System.Array.Empty<GoogleSheetColumnMapping>());

            Assert.Contains(vm.ColumnMappings, x => x.SheetColumn == "MDR_UNIQUE_ID");
            Assert.Contains(vm.ColumnMappings, x => x.SheetColumn == "MDR_ELEMENT_ID");
        }

        [Fact]
        public void BuildRequest_MapsFieldsAndMappings()
        {
            GoogleSyncWindowViewModel vm = new GoogleSyncWindowViewModel
            {
                Direction = "IMPORT",
                SelectedScheduleName = "MTO Main",
                SpreadsheetId = "https://docs.google.com/spreadsheets/d/spreadsheet-1/edit?usp=sharing",
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

        [Theory]
        [InlineData("spreadsheet-only-id", "spreadsheet-only-id")]
        [InlineData(" https://docs.google.com/spreadsheets/d/AbCDef_123-XYZ/edit#gid=0 ", "AbCDef_123-XYZ")]
        [InlineData("https://docs.google.com/spreadsheets/d/AbCDef_123-XYZ", "AbCDef_123-XYZ")]
        public void NormalizeSpreadsheetId_HandlesIdAndUrl(string input, string expected)
        {
            string actual = GoogleSyncWindowViewModel.NormalizeSpreadsheetId(input);
            Assert.Equal(expected, actual);
        }
    }
}
