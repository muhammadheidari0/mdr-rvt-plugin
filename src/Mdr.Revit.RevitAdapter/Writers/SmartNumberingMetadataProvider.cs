using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Mdr.Revit.Core.Contracts;
using Mdr.Revit.Core.Models;
using Mdr.Revit.RevitAdapter.Helpers;

namespace Mdr.Revit.RevitAdapter.Writers
{
    public sealed class SmartNumberingMetadataProvider : ISmartNumberingMetadataProvider
    {
        private readonly UIDocument? _uiDocument;
        private readonly ParameterAccessor _parameterAccessor;
        private readonly ArcaLevelCodeResolver _levelCodeResolver;

        public SmartNumberingMetadataProvider()
            : this(null, new ParameterAccessor(), new ArcaLevelCodeResolver())
        {
        }

        public SmartNumberingMetadataProvider(UIDocument uiDocument)
            : this(uiDocument, new ParameterAccessor(), new ArcaLevelCodeResolver())
        {
        }

        internal SmartNumberingMetadataProvider(
            UIDocument? uiDocument,
            ParameterAccessor parameterAccessor,
            ArcaLevelCodeResolver levelCodeResolver)
        {
            _uiDocument = uiDocument;
            _parameterAccessor = parameterAccessor ?? throw new ArgumentNullException(nameof(parameterAccessor));
            _levelCodeResolver = levelCodeResolver ?? throw new ArgumentNullException(nameof(levelCodeResolver));
        }

        public SmartNumberingMetadata GetMetadata(SmartNumberingRule rule)
        {
            SmartNumberingMetadata metadata = new SmartNumberingMetadata();
            Document? document = _uiDocument?.Document;
            if (document == null)
            {
                return metadata;
            }

            if (!TryResolveBuiltInCategory(rule?.CategoryBuiltInName, out BuiltInCategory category))
            {
                return metadata;
            }

            List<Element> elements = new FilteredElementCollector(document)
                .OfCategory(category)
                .WhereElementIsNotElementType()
                .Cast<Element>()
                .ToList();

            AddDistinctSorted(metadata.Blocks, elements
                .Select(x => _parameterAccessor.ReadValue(x, "Block")));
            AddDistinctSorted(metadata.Levels, elements
                .Select(x => _levelCodeResolver.GetEffectiveLevelCode(x)));

            return metadata;
        }

        private static bool TryResolveBuiltInCategory(string? categoryName, out BuiltInCategory category)
        {
            string value = string.IsNullOrWhiteSpace(categoryName)
                ? "OST_Walls"
                : categoryName.Trim();
            return Enum.TryParse(value, out category);
        }

        private static void AddDistinctSorted(List<string> target, IEnumerable<string> values)
        {
            List<string> sorted = values
                .Select(x => (x ?? string.Empty).Trim())
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(x => x, StringComparer.OrdinalIgnoreCase)
                .ToList();
            for (int i = 0; i < sorted.Count; i++)
            {
                target.Add(sorted[i]);
            }
        }
    }
}
