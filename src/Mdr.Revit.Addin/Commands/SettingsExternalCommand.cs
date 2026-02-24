using System;
using System.IO;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Mdr.Revit.Addin.UI;

namespace Mdr.Revit.Addin.Commands
{
    [Transaction(TransactionMode.Manual)]
    public sealed class SettingsExternalCommand : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            _ = elements;

            try
            {
                App app = commandData?.Application?.ActiveUIDocument != null
                    ? new App(commandData.Application.ActiveUIDocument)
                    : new App();

                SettingsAccessWorkflow workflow = new SettingsAccessWorkflow(app);
                bool updated = workflow.OpenSettings();
                return updated ? Result.Succeeded : Result.Cancelled;
            }
            catch (OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                WriteCommandError("settings", ex);
                return Result.Failed;
            }
        }

        private static void WriteCommandError(string commandName, Exception ex)
        {
            try
            {
                string logDirectory = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "MDR",
                    "RevitPlugin",
                    "logs");
                Directory.CreateDirectory(logDirectory);

                string logPath = Path.Combine(logDirectory, "command-errors.log");
                string line = DateTime.UtcNow.ToString("o") + " " + commandName + " " + ex;
                File.AppendAllText(logPath, line + Environment.NewLine);
            }
            catch
            {
                // Ignore logging failures in command exception path.
            }
        }
    }
}
