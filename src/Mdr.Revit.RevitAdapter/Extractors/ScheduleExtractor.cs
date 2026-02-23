using System;
using System.Collections.Generic;
using Mdr.Revit.Core.Models;

namespace Mdr.Revit.RevitAdapter.Extractors
{
    public sealed class ScheduleExtractor
    {
        private readonly Func<string, IReadOnlyList<ScheduleRow>> _provider;

        public ScheduleExtractor()
            : this(_ => Array.Empty<ScheduleRow>())
        {
        }

        public ScheduleExtractor(Func<string, IReadOnlyList<ScheduleRow>> provider)
        {
            _provider = provider ?? throw new ArgumentNullException(nameof(provider));
        }

        public IReadOnlyList<ScheduleRow> ExtractRows(string profileCode)
        {
            if (string.IsNullOrWhiteSpace(profileCode))
            {
                throw new ArgumentException("Profile code is required.", nameof(profileCode));
            }

            return _provider(profileCode);
        }
    }
}
