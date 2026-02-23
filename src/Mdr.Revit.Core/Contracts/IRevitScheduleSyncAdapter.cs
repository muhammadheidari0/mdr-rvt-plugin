using System.Collections.Generic;
using Mdr.Revit.Core.Models;

namespace Mdr.Revit.Core.Contracts
{
    public interface IRevitScheduleSyncAdapter
    {
        IReadOnlyList<ScheduleSyncRow> ExtractRows(
            string scheduleName,
            GoogleSheetSyncProfile profile);

        ScheduleSyncDiffResult BuildDiff(
            IReadOnlyList<ScheduleSyncRow> incomingRows,
            GoogleSheetSyncProfile profile);

        ScheduleSyncApplyResult ApplyDiff(
            ScheduleSyncDiffResult diff,
            GoogleSheetSyncProfile profile);
    }
}
