using System.Collections.Generic;
using Mdr.Revit.Core.Models;

namespace Mdr.Revit.Core.Contracts
{
    public interface IRevitExtractor
    {
        IReadOnlyList<PublishSheetItem> GetSelectedSheets();

        IReadOnlyList<ScheduleRow> GetScheduleRows(string profileCode);
    }
}
