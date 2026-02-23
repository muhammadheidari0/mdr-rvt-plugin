using Mdr.Revit.Core.Contracts;
using Mdr.Revit.Core.Models;
using Mdr.Revit.Core.UseCases;
using Mdr.Revit.Core.Validation;
using Xunit;

namespace Mdr.Revit.Core.Tests
{
    public sealed class ApplySmartNumberingUseCaseTests
    {
        [Fact]
        public void Execute_ValidRule_DelegatesToEngine()
        {
            FakeEngine engine = new FakeEngine();
            ApplySmartNumberingUseCase useCase = new ApplySmartNumberingUseCase(engine, new SmartNumberingFormulaParser());

            SmartNumberingRule rule = new SmartNumberingRule
            {
                RuleId = "default",
                Formula = "{Block}{Level}-{Sequence:5}",
            };
            rule.Targets.Add("Serial No");

            SmartNumberingResult result = useCase.Execute(rule, previewOnly: true);

            Assert.Equal(1, engine.CallCount);
            Assert.Equal(3, result.Preview.Count);
        }

        private sealed class FakeEngine : ISmartNumberingEngine
        {
            public int CallCount { get; private set; }

            public SmartNumberingResult Apply(SmartNumberingRule rule, bool previewOnly)
            {
                _ = rule;
                _ = previewOnly;
                CallCount++;

                SmartNumberingResult result = new SmartNumberingResult();
                result.Preview.Add(new SmartNumberingPreviewItem { ElementKey = "1", Value = "A00001" });
                result.Preview.Add(new SmartNumberingPreviewItem { ElementKey = "2", Value = "A00002" });
                result.Preview.Add(new SmartNumberingPreviewItem { ElementKey = "3", Value = "A00003" });
                return result;
            }
        }
    }
}
