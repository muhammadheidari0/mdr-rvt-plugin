using System;
using System.IO;
using Autodesk.Revit.UI;
using Mdr.Revit.Core.Contracts;
using Mdr.Revit.Core.Models;
using Mdr.Revit.Core.UseCases;
using Mdr.Revit.Core.Validation;
using Mdr.Revit.Infra.Logging;
using Mdr.Revit.RevitAdapter.Extractors;
using Mdr.Revit.RevitAdapter.Writers;

namespace Mdr.Revit.Addin.Commands
{
    public sealed class SmartNumberingCommand
    {
        private readonly ISmartNumberingEngine _smartNumberingEngine;
        private readonly SmartNumberingFormulaParser _formulaParser;
        private readonly PluginLogger _logger;

        public SmartNumberingCommand()
            : this(
                new NullSmartNumberingEngine(),
                new SmartNumberingFormulaParser(),
                new PluginLogger(DefaultLogDirectory()))
        {
        }

        public SmartNumberingCommand(UIDocument uiDocument)
            : this(
                RevitApiExtractors.CreateSmartNumberingEngine(uiDocument),
                new SmartNumberingFormulaParser(),
                new PluginLogger(DefaultLogDirectory()))
        {
        }

        internal SmartNumberingCommand(
            ISmartNumberingEngine smartNumberingEngine,
            SmartNumberingFormulaParser formulaParser,
            PluginLogger logger)
        {
            _smartNumberingEngine = smartNumberingEngine ?? throw new ArgumentNullException(nameof(smartNumberingEngine));
            _formulaParser = formulaParser ?? throw new ArgumentNullException(nameof(formulaParser));
            _logger = logger ?? throw new ArgumentNullException(nameof(logger));
        }

        public string Id => "mdr.smartNumbering";

        public SmartNumberingResult Execute(SmartNumberingCommandRequest request)
        {
            if (request == null)
            {
                throw new ArgumentNullException(nameof(request));
            }

            if (request.Rule == null)
            {
                throw new InvalidOperationException("Smart numbering rule is required.");
            }

            _logger.Info("Running smart numbering rule_id=" + request.Rule.RuleId);
            ApplySmartNumberingUseCase useCase = new ApplySmartNumberingUseCase(_smartNumberingEngine, _formulaParser);
            return useCase.Execute(request.Rule, request.PreviewOnly);
        }

        private static string DefaultLogDirectory()
        {
            return Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "MDR",
                "RevitPlugin",
                "logs");
        }
    }

    public sealed class SmartNumberingCommandRequest
    {
        public SmartNumberingRule Rule { get; set; } = new SmartNumberingRule();

        public bool PreviewOnly { get; set; }
    }

    internal sealed class NullSmartNumberingEngine : ISmartNumberingEngine
    {
        public SmartNumberingResult Apply(SmartNumberingRule rule, bool previewOnly)
        {
            _ = rule;
            _ = previewOnly;
            return new SmartNumberingResult();
        }
    }
}
