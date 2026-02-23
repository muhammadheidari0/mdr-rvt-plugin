using System.Collections.Generic;

namespace Mdr.Revit.Addin.Ribbon
{
    public sealed class RibbonBuilder
    {
        public IReadOnlyList<RibbonCommandDescriptor> Build()
        {
            return new[]
            {
                new RibbonCommandDescriptor("mdr.login", "Login to MDR", "Authenticate plugin session."),
                new RibbonCommandDescriptor("mdr.publishSheets", "Publish Selected Sheets", "Export and publish selected Revit sheets."),
                new RibbonCommandDescriptor("mdr.pushSchedules", "Push Schedules", "Extract mapped schedules and ingest to MDR."),
                new RibbonCommandDescriptor("mdr.syncSiteLogs", "Sync Contractor Reports", "Pull and apply verified site-log updates."),
                new RibbonCommandDescriptor("mdr.googleSync", "Google Sheets Sync", "Sync Revit schedules with Google Sheets."),
                new RibbonCommandDescriptor("mdr.smartNumbering", "Smart Numbering", "Generate rule-based element numbering."),
                new RibbonCommandDescriptor("mdr.checkUpdates", "Check Updates", "Check and prepare plugin updates."),
                new RibbonCommandDescriptor("mdr.settings", "Settings", "Edit plugin runtime configuration."),
            };
        }
    }

    public sealed class RibbonCommandDescriptor
    {
        public RibbonCommandDescriptor(string id, string title, string hint)
        {
            Id = id ?? string.Empty;
            Title = title ?? string.Empty;
            Hint = hint ?? string.Empty;
        }

        public string Id { get; }

        public string Title { get; }

        public string Hint { get; }
    }
}
