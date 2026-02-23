using System;
using System.Collections.Generic;
using Mdr.Revit.Core.Contracts;
using Mdr.Revit.Core.Models;

namespace Mdr.Revit.RevitAdapter.Extractors
{
    public sealed class RevitExtractorAdapter : IRevitExtractor
    {
        private readonly SheetExtractor _sheetExtractor;
        private readonly ScheduleExtractor _scheduleExtractor;

        public RevitExtractorAdapter(
            SheetExtractor sheetExtractor,
            ScheduleExtractor scheduleExtractor)
        {
            _sheetExtractor = sheetExtractor ?? throw new ArgumentNullException(nameof(sheetExtractor));
            _scheduleExtractor = scheduleExtractor ?? throw new ArgumentNullException(nameof(scheduleExtractor));
        }

        public IReadOnlyList<PublishSheetItem> GetSelectedSheets()
        {
            return _sheetExtractor.ExtractSelectedSheets();
        }

        public IReadOnlyList<ScheduleRow> GetScheduleRows(string profileCode)
        {
            return _scheduleExtractor.ExtractRows(profileCode);
        }
    }
}
