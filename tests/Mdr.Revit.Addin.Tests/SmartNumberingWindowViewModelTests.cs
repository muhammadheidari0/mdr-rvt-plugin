using Mdr.Revit.Addin.UI;
using Mdr.Revit.Core.Models;
using Xunit;

namespace Mdr.Revit.Addin.Tests
{
    public sealed class SmartNumberingWindowViewModelTests
    {
        [Fact]
        public void SetRules_SelectsDefaultAndAppliesTemplate()
        {
            SmartNumberingWindowViewModel vm = new SmartNumberingWindowViewModel();
            SmartNumberingRule defaultRule = new SmartNumberingRule
            {
                RuleId = "default",
                Formula = "{Mark}-{Sequence:5}",
                StartAt = 5,
                SequenceWidth = 5,
            };
            defaultRule.Targets.Add("Serial No");

            SmartNumberingRule secondaryRule = new SmartNumberingRule
            {
                RuleId = "doors",
                Formula = "{Comments}-{Sequence:4}",
                StartAt = 1,
                SequenceWidth = 4,
            };
            secondaryRule.Targets.Add("Type Mark");

            vm.SetRules(new[] { defaultRule, secondaryRule }, defaultRuleId: "doors");

            Assert.Equal("doors", vm.SelectedRuleId);
            Assert.Equal("{Comments}-{Sequence:4}", vm.Formula);
            Assert.Equal(4, vm.SequenceWidth);
            Assert.Equal("Type Mark", vm.TargetsText);
        }

        [Fact]
        public void BuildRule_ParsesTargetsAndRemovesDuplicates()
        {
            SmartNumberingWindowViewModel vm = new SmartNumberingWindowViewModel
            {
                SelectedRuleId = "runtime",
                Formula = "{Mark}-{Sequence:5}",
                StartAt = 2,
                SequenceWidth = 5,
                TargetsText = "Serial No, Type Mark, Serial No",
            };

            SmartNumberingRule rule = vm.BuildRule();

            Assert.Equal("runtime", rule.RuleId);
            Assert.Equal(2, rule.StartAt);
            Assert.Equal(5, rule.SequenceWidth);
            Assert.Equal(2, rule.Targets.Count);
            Assert.Contains("Serial No", rule.Targets);
            Assert.Contains("Type Mark", rule.Targets);
        }

        [Fact]
        public void BuildRule_ArcaMode_CapturesCategoryBlockAndLevel()
        {
            SmartNumberingWindowViewModel vm = new SmartNumberingWindowViewModel
            {
                SelectedRuleId = "arca-serial",
                Mode = SmartNumberingModes.Arca,
                CategoryBuiltInName = "OST_Walls",
                SelectedBlock = "A",
                SelectedLevel = "01",
            };

            SmartNumberingRule rule = vm.BuildRule();

            Assert.Equal(SmartNumberingModes.Arca, rule.Mode);
            Assert.Equal("OST_Walls", rule.CategoryBuiltInName);
            Assert.Equal("A", rule.SelectedBlock);
            Assert.Equal("01", rule.SelectedLevel);
            Assert.Contains("Serial No", rule.Targets);
        }

        [Fact]
        public void SetMetadata_SelectsFirstArcaBlockAndLevelWhenEmpty()
        {
            SmartNumberingWindowViewModel vm = new SmartNumberingWindowViewModel
            {
                Mode = SmartNumberingModes.Arca,
            };
            SmartNumberingMetadata metadata = new SmartNumberingMetadata();
            metadata.Blocks.Add("A");
            metadata.Levels.Add("01");

            vm.SetMetadata(metadata);

            Assert.Equal("A", vm.SelectedBlock);
            Assert.Equal("01", vm.SelectedLevel);
        }

        [Fact]
        public void UpdatePreview_WithErrors_DisablesApply()
        {
            SmartNumberingWindowViewModel vm = new SmartNumberingWindowViewModel();
            SmartNumberingResult result = new SmartNumberingResult();
            result.Preview.Add(new SmartNumberingPreviewItem
            {
                ElementKey = "el-1",
                Target = "Serial No",
                ProposedValue = "A0001",
                Status = SmartNumberingPreviewStates.Error,
                ErrorCode = "placeholder_missing",
                ErrorMessage = "Mark is missing.",
            });
            result.FailedCount = 1;
            result.SkippedCount = 1;

            vm.UpdatePreview(result);

            Assert.True(vm.HasPreviewErrors);
            Assert.False(vm.CanApply);
            Assert.Single(vm.PreviewRows);
        }
    }
}
