using System;
using System.Collections.Generic;
using Mdr.Revit.Core.Models;

namespace Mdr.Revit.RevitAdapter.Extractors
{
    public sealed class SheetExtractor
    {
        private readonly Func<IReadOnlyList<PublishSheetItem>> _provider;

        public SheetExtractor()
            : this(() => Array.Empty<PublishSheetItem>())
        {
        }

        public SheetExtractor(Func<IReadOnlyList<PublishSheetItem>> provider)
        {
            _provider = provider ?? throw new ArgumentNullException(nameof(provider));
        }

        public IReadOnlyList<PublishSheetItem> ExtractSelectedSheets()
        {
            return _provider();
        }
    }
}
