using System;
using Autodesk.Revit.DB;
using Mdr.Revit.Core.Validation;

namespace Mdr.Revit.RevitAdapter.Helpers
{
    public sealed class ArcaLevelCodeResolver
    {
        private static readonly string[] LevelCodeParameterNames =
        {
            "Level Code",
            "LevelCode",
            "Level_Code",
        };

        private static readonly string[] BaseLevelParameterNames =
        {
            "Base Constraint",
            "Base Level",
            "Reference Level",
            "Level",
        };

        private static readonly string[] BuiltInLevelParameterNames =
        {
            "WALL_BASE_CONSTRAINT",
            "INSTANCE_REFERENCE_LEVEL_PARAM",
            "FAMILY_LEVEL_PARAM",
            "SCHEDULE_LEVEL_PARAM",
        };

        public string GetEffectiveLevelCode(Element element)
        {
            if (element == null)
            {
                return string.Empty;
            }

            string baseLevel = GetBaseConstraintCode(element);
            if (!string.IsNullOrWhiteSpace(baseLevel))
            {
                return baseLevel;
            }

            return GetLevelCodeParameter(element);
        }

        private static string GetBaseConstraintCode(Element element)
        {
            for (int i = 0; i < BuiltInLevelParameterNames.Length; i++)
            {
                if (!Enum.TryParse(BuiltInLevelParameterNames[i], out BuiltInParameter builtInParameter))
                {
                    continue;
                }

                string value = ReadParameter(element, builtInParameter);
                if (!string.IsNullOrWhiteSpace(value))
                {
                    return ArcaSerialNumberingPlanner.NormalizeLevelCode(value);
                }
            }

            for (int i = 0; i < BaseLevelParameterNames.Length; i++)
            {
                string value = ReadParameter(element, BaseLevelParameterNames[i], includeType: false);
                if (!string.IsNullOrWhiteSpace(value))
                {
                    return ArcaSerialNumberingPlanner.NormalizeLevelCode(value);
                }
            }

            return string.Empty;
        }

        private static string GetLevelCodeParameter(Element element)
        {
            for (int i = 0; i < LevelCodeParameterNames.Length; i++)
            {
                string value = ReadParameter(element, LevelCodeParameterNames[i], includeType: false);
                if (!string.IsNullOrWhiteSpace(value))
                {
                    return ArcaSerialNumberingPlanner.NormalizeLevelCode(value);
                }
            }

            for (int i = 0; i < LevelCodeParameterNames.Length; i++)
            {
                string value = ReadParameter(element, LevelCodeParameterNames[i], includeType: true, typeOnly: true);
                if (!string.IsNullOrWhiteSpace(value))
                {
                    return ArcaSerialNumberingPlanner.NormalizeLevelCode(value);
                }
            }

            return string.Empty;
        }

        private static string ReadParameter(Element element, BuiltInParameter builtInParameter)
        {
            try
            {
                Parameter? parameter = element.get_Parameter(builtInParameter);
                return ReadParameterValue(element.Document, parameter);
            }
            catch
            {
                return string.Empty;
            }
        }

        private static string ReadParameter(
            Element element,
            string parameterName,
            bool includeType,
            bool typeOnly = false)
        {
            if (element == null || string.IsNullOrWhiteSpace(parameterName))
            {
                return string.Empty;
            }

            if (!typeOnly)
            {
                string value = ReadParameterValue(element.Document, element.LookupParameter(parameterName.Trim()));
                if (!string.IsNullOrWhiteSpace(value))
                {
                    return value;
                }
            }

            if (!includeType)
            {
                return string.Empty;
            }

            ElementId typeId = element.GetTypeId();
            if (typeId == ElementId.InvalidElementId)
            {
                return string.Empty;
            }

            Element? typeElement = element.Document.GetElement(typeId);
            return ReadParameterValue(element.Document, typeElement?.LookupParameter(parameterName.Trim()));
        }

        private static string ReadParameterValue(Document document, Parameter? parameter)
        {
            if (parameter == null)
            {
                return string.Empty;
            }

            try
            {
                if (parameter.StorageType == StorageType.ElementId)
                {
                    ElementId id = parameter.AsElementId();
                    if (id != ElementId.InvalidElementId)
                    {
                        Element? referenced = document.GetElement(id);
                        if (!string.IsNullOrWhiteSpace(referenced?.Name))
                        {
                            return referenced.Name.Trim();
                        }
                    }
                }

                string value = string.Empty;
                try
                {
                    value = parameter.AsString() ?? string.Empty;
                }
                catch
                {
                    value = string.Empty;
                }

                if (!string.IsNullOrWhiteSpace(value))
                {
                    return value.Trim();
                }

                try
                {
                    value = parameter.AsValueString() ?? string.Empty;
                }
                catch
                {
                    value = string.Empty;
                }

                if (!string.IsNullOrWhiteSpace(value))
                {
                    return value.Trim();
                }

                if (parameter.StorageType == StorageType.Integer)
                {
                    return parameter.AsInteger().ToString();
                }
            }
            catch
            {
                return string.Empty;
            }

            return string.Empty;
        }
    }
}
