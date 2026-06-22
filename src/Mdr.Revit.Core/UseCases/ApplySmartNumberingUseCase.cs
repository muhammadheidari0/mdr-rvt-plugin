using System;
using Mdr.Revit.Core.Contracts;
using Mdr.Revit.Core.Models;
using Mdr.Revit.Core.Validation;

namespace Mdr.Revit.Core.UseCases
{
    public sealed class ApplySmartNumberingUseCase
    {
        private readonly ISmartNumberingEngine _smartNumberingEngine;
        private readonly SmartNumberingFormulaParser _formulaParser;

        public ApplySmartNumberingUseCase(
            ISmartNumberingEngine smartNumberingEngine,
            SmartNumberingFormulaParser formulaParser)
        {
            _smartNumberingEngine = smartNumberingEngine ?? throw new ArgumentNullException(nameof(smartNumberingEngine));
            _formulaParser = formulaParser ?? throw new ArgumentNullException(nameof(formulaParser));
        }

        public SmartNumberingResult Execute(SmartNumberingRule rule, bool previewOnly)
        {
            if (rule == null)
            {
                throw new ArgumentNullException(nameof(rule));
            }

            if (IsArcaRule(rule))
            {
                if (string.IsNullOrWhiteSpace(rule.CategoryBuiltInName))
                {
                    throw new InvalidOperationException("ARCA smart numbering category is required.");
                }

                if (string.IsNullOrWhiteSpace(rule.SelectedBlock))
                {
                    throw new InvalidOperationException("ARCA smart numbering block is required.");
                }

                if (string.IsNullOrWhiteSpace(rule.SelectedLevel))
                {
                    throw new InvalidOperationException("ARCA smart numbering level is required.");
                }

                return _smartNumberingEngine.Apply(rule, previewOnly);
            }

            if (string.IsNullOrWhiteSpace(rule.Formula))
            {
                throw new InvalidOperationException("Smart numbering formula is required.");
            }

            if (rule.Targets.Count == 0)
            {
                throw new InvalidOperationException("At least one target parameter is required.");
            }

            _formulaParser.Parse(rule.Formula, rule.SequenceWidth);
            return _smartNumberingEngine.Apply(rule, previewOnly);
        }

        private static bool IsArcaRule(SmartNumberingRule rule)
        {
            return string.Equals(rule.Mode, SmartNumberingModes.Arca, StringComparison.OrdinalIgnoreCase);
        }
    }
}
