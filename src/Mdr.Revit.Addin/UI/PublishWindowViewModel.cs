using System;
using System.Collections.Generic;
using Mdr.Revit.Core.Models;

namespace Mdr.Revit.Addin.UI
{
    public sealed class PublishWindowViewModel
    {
        private readonly List<PublishSheetSelectionItem> _sheetRows = new List<PublishSheetSelectionItem>();

        public IReadOnlyList<PublishSheetSelectionItem> SheetRows => _sheetRows;

        public bool IncludeNative { get; set; }

        public bool RetryFailedItems { get; set; } = true;

        public string Username { get; set; } = string.Empty;

        public string Password { get; set; } = string.Empty;

        public string ProjectCode { get; set; } = string.Empty;

        public string OutputDirectory { get; set; } = string.Empty;

        public int SelectedCount
        {
            get
            {
                int count = 0;
                for (int i = 0; i < _sheetRows.Count; i++)
                {
                    if (_sheetRows[i].IsSelected)
                    {
                        count++;
                    }
                }

                return count;
            }
        }

        public void SetSheets(IReadOnlyList<PublishSheetItem> sheets)
        {
            _sheetRows.Clear();
            if (sheets == null || sheets.Count == 0)
            {
                return;
            }

            for (int i = 0; i < sheets.Count; i++)
            {
                PublishSheetItem source = sheets[i];
                if (source == null)
                {
                    continue;
                }

                _sheetRows.Add(PublishSheetSelectionItem.FromPublishItem(source));
            }
        }

        public IReadOnlyList<PublishSheetItem> BuildSelectedItems()
        {
            List<PublishSheetItem> selected = new List<PublishSheetItem>();

            for (int i = 0; i < _sheetRows.Count; i++)
            {
                PublishSheetSelectionItem row = _sheetRows[i];
                if (!row.IsSelected)
                {
                    continue;
                }

                PublishSheetItem item = row.ToPublishSheetItem();
                item.ItemIndex = selected.Count;
                item.IncludeNative = IncludeNative || item.IncludeNative;
                selected.Add(item);
            }

            return selected;
        }
    }

    public sealed class PublishSheetSelectionItem
    {
        public bool IsSelected { get; set; } = true;

        public string SheetUniqueId { get; set; } = string.Empty;

        public string SheetNumber { get; set; } = string.Empty;

        public string SheetName { get; set; } = string.Empty;

        public string RequestedRevision { get; set; } = string.Empty;

        public string StatusCode { get; set; } = string.Empty;

        public bool IncludeNative { get; set; }

        public string DocNumber { get; set; } = string.Empty;

        public string FileSha256 { get; set; } = string.Empty;

        public DocumentMetadata Metadata { get; set; } = new DocumentMetadata();

        public static PublishSheetSelectionItem FromPublishItem(PublishSheetItem item)
        {
            if (item == null)
            {
                throw new ArgumentNullException(nameof(item));
            }

            return new PublishSheetSelectionItem
            {
                IsSelected = true,
                SheetUniqueId = item.SheetUniqueId ?? string.Empty,
                SheetNumber = item.SheetNumber ?? string.Empty,
                SheetName = item.SheetName ?? string.Empty,
                RequestedRevision = item.RequestedRevision ?? string.Empty,
                StatusCode = item.StatusCode ?? string.Empty,
                IncludeNative = item.IncludeNative,
                DocNumber = item.DocNumber ?? string.Empty,
                FileSha256 = item.FileSha256 ?? string.Empty,
                Metadata = item.Metadata?.Clone() ?? new DocumentMetadata(),
            };
        }

        public PublishSheetItem ToPublishSheetItem()
        {
            return new PublishSheetItem
            {
                SheetUniqueId = SheetUniqueId ?? string.Empty,
                SheetNumber = SheetNumber ?? string.Empty,
                SheetName = SheetName ?? string.Empty,
                RequestedRevision = RequestedRevision ?? string.Empty,
                StatusCode = StatusCode ?? string.Empty,
                IncludeNative = IncludeNative,
                DocNumber = DocNumber ?? string.Empty,
                FileSha256 = FileSha256 ?? string.Empty,
                Metadata = Metadata?.Clone() ?? new DocumentMetadata(),
            };
        }
    }
}
