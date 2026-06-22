using System.Linq;
using Mdr.Revit.Core.Models;
using Mdr.Revit.Core.Validation;
using Xunit;

namespace Mdr.Revit.Core.Tests
{
    public sealed class ArcaSerialNumberingPlannerTests
    {
        [Fact]
        public void Plan_ExistingSerialNo_WritesSeriesFromTrailingDigits()
        {
            ArcaSerialNumberingPlanner planner = new ArcaSerialNumberingPlanner();
            ArcaSerialNumberingPlanRequest request = CreateRequest();
            request.Elements.Add(new ArcaSerialNumberingElementSnapshot
            {
                ElementKey = "el-1",
                ElementId = 1,
                Block = "A",
                LevelCode = "01",
                TypeMark = "TM",
                SerialNo = "A01-TM005",
            });

            ArcaSerialNumberingPlan plan = planner.Plan(request);

            ArcaSerialNumberingWrite write = Assert.Single(plan.Writes, x => x.Target == "Series");
            Assert.Equal("005", write.Value);
            Assert.Contains(plan.Preview, x => x.Target == "Serial No" && x.Status == SmartNumberingPreviewStates.Skipped);
        }

        [Fact]
        public void Plan_NewElements_CreatesSeriesByScopeTypeMarkAndElementId()
        {
            ArcaSerialNumberingPlanner planner = new ArcaSerialNumberingPlanner();
            ArcaSerialNumberingPlanRequest request = CreateRequest();
            request.Elements.Add(new ArcaSerialNumberingElementSnapshot
            {
                ElementKey = "scope-2",
                ElementId = 10,
                Block = "A",
                LevelCode = "01",
                TypeMark = "TM",
                ScopeSort = 2,
            });
            request.Elements.Add(new ArcaSerialNumberingElementSnapshot
            {
                ElementKey = "scope-1",
                ElementId = 20,
                Block = "A",
                LevelCode = "01",
                TypeMark = "TM",
                ScopeSort = 1,
            });

            ArcaSerialNumberingPlan plan = planner.Plan(request);

            Assert.Contains(plan.Writes, x => x.ElementKey == "scope-1" && x.Target == "Series" && x.Value == "001");
            Assert.Contains(plan.Writes, x => x.ElementKey == "scope-2" && x.Target == "Series" && x.Value == "002");
            Assert.Contains(plan.Writes, x => x.ElementKey == "scope-1" && x.Target == "Serial No" && x.Value == "A01-TM001");
            Assert.Contains(plan.Writes, x => x.ElementKey == "scope-2" && x.Target == "Serial No" && x.Value == "A01-TM002");
        }

        [Fact]
        public void Plan_SkipsSerialNo_WhenTypeMarkOrNumericSeriesIsMissing()
        {
            ArcaSerialNumberingPlanner planner = new ArcaSerialNumberingPlanner();
            ArcaSerialNumberingPlanRequest request = CreateRequest();
            request.Elements.Add(new ArcaSerialNumberingElementSnapshot
            {
                ElementKey = "no-type",
                ElementId = 1,
                Block = "A",
                LevelCode = "01",
            });
            request.Elements.Add(new ArcaSerialNumberingElementSnapshot
            {
                ElementKey = "bad-series",
                ElementId = 2,
                Block = "A",
                LevelCode = "01",
                TypeMark = "TM",
                Series = "ABC",
            });

            ArcaSerialNumberingPlan plan = planner.Plan(request);

            Assert.Empty(plan.Writes);
            Assert.Contains(plan.Preview, x => x.ElementKey == "no-type" && x.Status == SmartNumberingPreviewStates.Skipped);
            Assert.Contains(plan.Preview, x => x.ElementKey == "bad-series" && x.Status == SmartNumberingPreviewStates.Skipped);
        }

        [Theory]
        [InlineData("L01-Level 01", "L01")]
        [InlineData("L02", "L02")]
        [InlineData("", "")]
        public void NormalizeLevelCode_UsesTextBeforeDash(string value, string expected)
        {
            Assert.Equal(expected, ArcaSerialNumberingPlanner.NormalizeLevelCode(value));
        }

        private static ArcaSerialNumberingPlanRequest CreateRequest()
        {
            return new ArcaSerialNumberingPlanRequest
            {
                SelectedBlock = "A",
                SelectedLevelCode = "01",
            };
        }
    }
}
