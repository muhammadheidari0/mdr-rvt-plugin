using Mdr.Revit.Core.Models;
using Mdr.Revit.RevitAdapter.Writers;
using Xunit;

namespace Mdr.Revit.RevitAdapter.Tests
{
    public sealed class SiteLogsScheduleWriterTests
    {
        [Fact]
        public void ApplySiteLogRows_ReturnsCombinedAppliedCount()
        {
            SiteLogsScheduleWriter writer = new SiteLogsScheduleWriter();
            SiteLogPullResponse response = new SiteLogPullResponse
            {
                RunId = "run-sync",
            };
            response.ManpowerRows.Add(new SiteLogRow { SyncKey = "1" });
            response.EquipmentRows.Add(new SiteLogRow { SyncKey = "2" });
            response.ActivityRows.Add(new SiteLogRow { SyncKey = "3" });

            SiteLogApplyResult result = writer.ApplySiteLogRows(response);

            Assert.Equal(3, result.AppliedCount);
            Assert.Equal("run-sync", result.RunId);
        }
    }
}
