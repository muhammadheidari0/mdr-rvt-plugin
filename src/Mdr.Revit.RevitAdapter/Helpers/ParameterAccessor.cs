using System;
using System.Globalization;
using Autodesk.Revit.DB;

namespace Mdr.Revit.RevitAdapter.Helpers
{
    public sealed class ParameterAccessor
    {
        public string ReadValue(Element element, string parameterName)
        {
            Parameter? parameter = ResolveParameter(element, parameterName, requireWritable: false);
            if (parameter == null)
            {
                return string.Empty;
            }

            string text = parameter.AsString();
            if (!string.IsNullOrWhiteSpace(text))
            {
                return text.Trim();
            }

            string valueText = parameter.AsValueString();
            if (!string.IsNullOrWhiteSpace(valueText))
            {
                return valueText.Trim();
            }

            if (parameter.StorageType == StorageType.Integer)
            {
                return parameter.AsInteger().ToString(CultureInfo.InvariantCulture);
            }

            if (parameter.StorageType == StorageType.Double)
            {
                return parameter.AsDouble().ToString("G17", CultureInfo.InvariantCulture);
            }

            if (parameter.StorageType == StorageType.ElementId)
            {
                return parameter.AsElementId().Value.ToString(CultureInfo.InvariantCulture);
            }

            return string.Empty;
        }

        public bool CanWriteValue(
            Element element,
            string parameterName,
            string candidateValue,
            out string errorCode,
            out string errorMessage)
        {
            errorCode = string.Empty;
            errorMessage = string.Empty;

            Parameter? parameter = ResolveParameter(element, parameterName, requireWritable: true);
            if (parameter == null)
            {
                errorCode = "parameter_read_only";
                errorMessage = "Writable parameter was not found: " + parameterName;
                return false;
            }

            return ValidateStorageType(parameter, candidateValue ?? string.Empty, out errorCode, out errorMessage);
        }

        public bool TryWriteValue(
            Element element,
            string parameterName,
            string value,
            out string errorCode,
            out string errorMessage)
        {
            errorCode = string.Empty;
            errorMessage = string.Empty;

            Parameter? parameter = ResolveParameter(element, parameterName, requireWritable: true);
            if (parameter == null)
            {
                errorCode = "parameter_read_only";
                errorMessage = "Writable parameter was not found: " + parameterName;
                return false;
            }

            if (!ValidateStorageType(parameter, value ?? string.Empty, out errorCode, out errorMessage))
            {
                return false;
            }

            try
            {
                if (parameter.StorageType == StorageType.String)
                {
                    return parameter.Set(value ?? string.Empty);
                }

                if (parameter.StorageType == StorageType.Integer)
                {
                    int intValue = ParseInteger(value);
                    return parameter.Set(intValue);
                }

                if (parameter.StorageType == StorageType.Double)
                {
                    double doubleValue = double.Parse(value, NumberStyles.Float, CultureInfo.InvariantCulture);
                    return parameter.Set(doubleValue);
                }

                if (parameter.StorageType == StorageType.ElementId)
                {
                    int elementId = ParseInteger(value);
                    return parameter.Set(new ElementId(elementId));
                }

                errorCode = "type_mismatch";
                errorMessage = "Unsupported parameter storage type.";
                return false;
            }
            catch (Exception ex)
            {
                errorCode = "apply_failed";
                errorMessage = ex.Message;
                return false;
            }
        }

        private static bool ValidateStorageType(
            Parameter parameter,
            string candidateValue,
            out string errorCode,
            out string errorMessage)
        {
            errorCode = string.Empty;
            errorMessage = string.Empty;

            if (parameter.StorageType == StorageType.String)
            {
                return true;
            }

            if (string.IsNullOrWhiteSpace(candidateValue))
            {
                errorCode = "type_mismatch";
                errorMessage = "Empty value is not valid for non-string parameter.";
                return false;
            }

            if (parameter.StorageType == StorageType.Integer || parameter.StorageType == StorageType.ElementId)
            {
                if (int.TryParse(candidateValue, NumberStyles.Integer, CultureInfo.InvariantCulture, out _))
                {
                    return true;
                }

                errorCode = "type_mismatch";
                errorMessage = "Value must be an integer.";
                return false;
            }

            if (parameter.StorageType == StorageType.Double)
            {
                if (double.TryParse(candidateValue, NumberStyles.Float, CultureInfo.InvariantCulture, out _))
                {
                    return true;
                }

                errorCode = "type_mismatch";
                errorMessage = "Value must be a number.";
                return false;
            }

            errorCode = "type_mismatch";
            errorMessage = "Unsupported parameter storage type.";
            return false;
        }

        private static int ParseInteger(string raw)
        {
            string normalized = (raw ?? string.Empty).Trim();
            if (normalized.Equals("true", StringComparison.OrdinalIgnoreCase))
            {
                return 1;
            }

            if (normalized.Equals("false", StringComparison.OrdinalIgnoreCase))
            {
                return 0;
            }

            return int.Parse(normalized, NumberStyles.Integer, CultureInfo.InvariantCulture);
        }

        private static Parameter? ResolveParameter(Element element, string parameterName, bool requireWritable)
        {
            if (element == null || string.IsNullOrWhiteSpace(parameterName))
            {
                return null;
            }

            string normalized = parameterName.Trim();
            Parameter? instance = element.LookupParameter(normalized);
            if (instance != null && (!requireWritable || !instance.IsReadOnly))
            {
                return instance;
            }

            ElementId typeId = element.GetTypeId();
            if (typeId == ElementId.InvalidElementId)
            {
                return null;
            }

            Element? typeElement = element.Document.GetElement(typeId);
            Parameter? typeParameter = typeElement?.LookupParameter(normalized);
            if (typeParameter != null && (!requireWritable || !typeParameter.IsReadOnly))
            {
                return typeParameter;
            }

            return null;
        }
    }
}
