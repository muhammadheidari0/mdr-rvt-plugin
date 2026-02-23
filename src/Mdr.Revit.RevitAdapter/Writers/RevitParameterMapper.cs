using System;
using System.Collections.Generic;

namespace Mdr.Revit.RevitAdapter.Writers
{
    public sealed class RevitParameterMapper
    {
        public Dictionary<string, string> BuildParameterMap(IReadOnlyDictionary<string, string> source)
        {
            if (source == null)
            {
                throw new ArgumentNullException(nameof(source));
            }

            Dictionary<string, string> map = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            foreach (KeyValuePair<string, string> item in source)
            {
                map[item.Key] = item.Value;
            }

            return map;
        }
    }
}
