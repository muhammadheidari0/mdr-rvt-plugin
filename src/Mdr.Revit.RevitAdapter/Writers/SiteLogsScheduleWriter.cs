using Mdr.Revit.Core.Contracts;
using Mdr.Revit.Core.Models;

namespace Mdr.Revit.RevitAdapter.Writers
{
    public sealed class SiteLogsScheduleWriter : IRevitWriter
    {
        public SiteLogApplyResult ApplySiteLogRows(SiteLogPullResponse pullResponse)
        {
            int totalRows = pullResponse.ManpowerRows.Count + pullResponse.EquipmentRows.Count + pullResponse.ActivityRows.Count;

            return new SiteLogApplyResult
            {
                RunId = pullResponse.RunId,
                AppliedCount = totalRows,
                FailedCount = 0,
            };
        }
    }
}
