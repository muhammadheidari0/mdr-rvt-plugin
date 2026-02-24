using Mdr.Revit.Addin.UI;
using Mdr.Revit.Core.Models;
using Xunit;

namespace Mdr.Revit.Addin.Tests
{
    public sealed class PublishWindowViewModelTests
    {
        [Fact]
        public void BuildSelectedItems_ReturnsOnlyCheckedRows_WithSequentialIndexes()
        {
            PublishWindowViewModel vm = new PublishWindowViewModel();
            vm.SetSheets(new[]
            {
                new PublishSheetItem { SheetUniqueId = "sheet-1", SheetNumber = "A-101", SheetName = "Plan 1" },
                new PublishSheetItem { SheetUniqueId = "sheet-2", SheetNumber = "A-102", SheetName = "Plan 2" },
                new PublishSheetItem { SheetUniqueId = "sheet-3", SheetNumber = "A-103", SheetName = "Plan 3" },
            });

            vm.SheetRows[1].IsSelected = false;

            var selected = vm.BuildSelectedItems();

            Assert.Equal(2, selected.Count);
            Assert.Equal("sheet-1", selected[0].SheetUniqueId);
            Assert.Equal("sheet-3", selected[1].SheetUniqueId);
            Assert.Equal(0, selected[0].ItemIndex);
            Assert.Equal(1, selected[1].ItemIndex);
        }

        [Fact]
        public void BuildSelectedItems_AppliesIncludeNativeFlag()
        {
            PublishWindowViewModel vm = new PublishWindowViewModel
            {
                IncludeNative = true,
            };
            vm.SetSheets(new[]
            {
                new PublishSheetItem { SheetUniqueId = "sheet-1", IncludeNative = false },
            });

            var selected = vm.BuildSelectedItems();

            Assert.Single(selected);
            Assert.True(selected[0].IncludeNative);
        }

        [Fact]
        public void ApplyPublishResult_MapsItemStatusBackToRows()
        {
            PublishWindowViewModel vm = new PublishWindowViewModel();
            vm.SetSheets(new[]
            {
                new PublishSheetItem { SheetUniqueId = "sheet-1", SheetNumber = "A-101", SheetName = "Plan 1" },
                new PublishSheetItem { SheetUniqueId = "sheet-2", SheetNumber = "A-102", SheetName = "Plan 2" },
            });

            var selected = vm.BuildSelectedItems();
            Assert.Equal(2, selected.Count);
            Assert.Equal("pending", vm.SheetRows[0].LastState);
            Assert.Equal("pending", vm.SheetRows[1].LastState);

            PublishBatchResponse response = new PublishBatchResponse();
            response.Items.Add(new PublishItemResult { ItemIndex = 0, State = "completed" });
            response.Items.Add(new PublishItemResult
            {
                ItemIndex = 1,
                State = "failed",
                ErrorCode = "export_pdf_failed",
                ErrorMessage = "PDF export output was not produced.",
            });

            vm.ApplyPublishResult(response);

            Assert.Equal("completed", vm.SheetRows[0].LastState);
            Assert.Equal("failed", vm.SheetRows[1].LastState);
            Assert.Equal("export_pdf_failed", vm.SheetRows[1].LastErrorCode);
        }
    }
}
