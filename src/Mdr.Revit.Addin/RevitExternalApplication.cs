using System;
using System.Collections.Generic;
using System.IO;
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
        private static readonly string[] PreloadAssemblyFiles =
        {
            "System.Text.Json.dll",
            "System.Text.Encodings.Web.dll",
            "Microsoft.Bcl.AsyncInterfaces.dll",
            "System.Memory.dll",
            "System.Buffers.dll",
            "System.Runtime.CompilerServices.Unsafe.dll",
            "System.Numerics.Vectors.dll",
            "System.Threading.Tasks.Extensions.dll",
            "System.ValueTuple.dll",
            "Mdr.Revit.Core.dll",
            "Mdr.Revit.Infra.dll",
            "Mdr.Revit.Client.dll",
            "Mdr.Revit.RevitAdapter.dll",
        };
        private static readonly object ResolverLock = new object();
        private static bool _resolverRegistered;

        public Result OnStartup(UIControlledApplication application)
        {
            if (application == null)
            {
                return Result.Failed;
            }

            try
            {
                EnsureAssemblyResolver();
                PreloadPluginDependencies();
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

        private static void EnsureAssemblyResolver()
        {
            lock (ResolverLock)
            {
                if (_resolverRegistered)
                {
                    return;
                }

                AppDomain.CurrentDomain.AssemblyResolve += ResolveAssemblyFromPluginDirectory;
                _resolverRegistered = true;
            }
        }

        private static Assembly? ResolveAssemblyFromPluginDirectory(object sender, ResolveEventArgs args)
        {
            _ = sender;
            if (args == null || string.IsNullOrWhiteSpace(args.Name))
            {
                return null;
            }

            try
            {
                string? pluginDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                if (string.IsNullOrWhiteSpace(pluginDirectory))
                {
                    return null;
                }

                AssemblyName requested = new AssemblyName(args.Name);
                if (string.IsNullOrWhiteSpace(requested.Name))
                {
                    return null;
                }

                string candidatePath = Path.Combine(pluginDirectory, requested.Name + ".dll");
                if (!File.Exists(candidatePath))
                {
                    return null;
                }

                return Assembly.LoadFrom(candidatePath);
            }
            catch (Exception ex)
            {
                WriteBootstrapLog("AssemblyResolve failed for " + args.Name, ex);
                return null;
            }
        }

        private static void PreloadPluginDependencies()
        {
            string? pluginDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            if (string.IsNullOrWhiteSpace(pluginDirectory))
            {
                return;
            }

            for (int i = 0; i < PreloadAssemblyFiles.Length; i++)
            {
                string fileName = PreloadAssemblyFiles[i];
                string path = Path.Combine(pluginDirectory, fileName);
                if (!File.Exists(path))
                {
                    continue;
                }

                try
                {
                    AssemblyName name = AssemblyName.GetAssemblyName(path);
                    bool alreadyLoaded = AppDomain.CurrentDomain
                        .GetAssemblies()
                        .Any(x => string.Equals(x.FullName, name.FullName, StringComparison.OrdinalIgnoreCase));
                    if (alreadyLoaded)
                    {
                        continue;
                    }

                    Assembly.LoadFrom(path);
                }
                catch (Exception ex)
                {
                    WriteBootstrapLog("Preload failed for " + fileName, ex);
                }
            }
        }

        private static void WriteBootstrapLog(string message, Exception? ex = null)
        {
            try
            {
                string logDirectory = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "MDR",
                    "RevitPlugin",
                    "logs");
                Directory.CreateDirectory(logDirectory);

                string logPath = Path.Combine(logDirectory, "bootstrap.log");
                string line = DateTime.UtcNow.ToString("o") + " " + message;
                if (ex != null)
                {
                    line += " :: " + ex;
                }

                File.AppendAllText(logPath, line + Environment.NewLine);
            }
            catch
            {
                // Ignore bootstrap log failures.
            }
        }
    }
}
