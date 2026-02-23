using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using Autodesk.Revit.UI;
using Mdr.Revit.Addin.Commands;

namespace Mdr.Revit.Addin
{
    public sealed class RevitExternalApplication : IExternalApplication
    {
        private const string RibbonTabName = "MDR";
        private const string RibbonPanelName = "BIM";

        public Result OnStartup(UIControlledApplication application)
        {
            if (application == null)
            {
                return Result.Failed;
            }

            try
            {
                EnsureRibbonTab(application, RibbonTabName);
                RibbonPanel panel = EnsureRibbonPanel(application, RibbonTabName, RibbonPanelName);
                AddRibbonButton(
                    panel,
                    "mdr.googleSync",
                    "Google Sheets Sync",
                    "Open Google Sheets schedule sync dialog.",
                    typeof(GoogleSyncExternalCommand));
                AddRibbonButton(
                    panel,
                    "mdr.smartNumbering",
                    "Smart Numbering",
                    "Generate and apply rule-based numbering with live preview.",
                    typeof(SmartNumberingExternalCommand));
                return Result.Succeeded;
            }
            catch
            {
                return Result.Failed;
            }
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            _ = application;
            return Result.Succeeded;
        }

        private static void EnsureRibbonTab(UIControlledApplication application, string tabName)
        {
            try
            {
                application.CreateRibbonTab(tabName);
            }
            catch
            {
                // Revit throws when tab already exists.
            }
        }

        private static RibbonPanel EnsureRibbonPanel(
            UIControlledApplication application,
            string tabName,
            string panelName)
        {
            IList<RibbonPanel> panels = application.GetRibbonPanels(tabName);
            RibbonPanel? existing = panels.FirstOrDefault(x =>
                string.Equals(x.Name, panelName, StringComparison.OrdinalIgnoreCase));
            if (existing != null)
            {
                return existing;
            }

            return application.CreateRibbonPanel(tabName, panelName);
        }

        private static void AddRibbonButton(
            RibbonPanel panel,
            string buttonId,
            string buttonText,
            string tooltip,
            Type commandType)
        {
            bool exists = panel.GetItems()
                .Any(x => string.Equals(x.Name, buttonId, StringComparison.OrdinalIgnoreCase));
            if (exists)
            {
                return;
            }

            string assemblyPath = Assembly.GetExecutingAssembly().Location;
            PushButtonData data = new PushButtonData(
                buttonId,
                buttonText,
                assemblyPath,
                commandType.FullName);
            PushButton? button = panel.AddItem(data) as PushButton;
            if (button != null)
            {
                button.ToolTip = tooltip;
            }
        }
    }
}
