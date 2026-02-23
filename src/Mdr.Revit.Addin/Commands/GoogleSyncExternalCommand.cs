using System;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Mdr.Revit.Addin.UI;

namespace Mdr.Revit.Addin.Commands
{
    [Transaction(TransactionMode.Manual)]
    public sealed class GoogleSyncExternalCommand : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            _ = elements;
            if (commandData?.Application?.ActiveUIDocument == null)
            {
                message = "No active Revit document is open.";
                return Result.Failed;
            }

            try
            {
                App app = new App(commandData.Application.ActiveUIDocument);
                GoogleSyncWindow window = new GoogleSyncWindow(app);
                bool? dialogResult = window.ShowDialog();
                return dialogResult == false ? Result.Cancelled : Result.Succeeded;
            }
            catch (OperationCanceledException)
            {
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                message = ex.Message;
                return Result.Failed;
            }
        }
    }
}
