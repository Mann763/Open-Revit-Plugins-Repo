using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI;
using System;
using System.Reflection;

namespace SharedCoordExporter
{
    [Transaction(TransactionMode.Manual)]
    public class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication application)
        {
            try
            {
                string tabName = "MEP Tools";
                try { application.CreateRibbonTab(tabName); } catch { }

                RibbonPanel panel = application.CreateRibbonPanel(tabName, "Export");

                string dllPath = Assembly.GetExecutingAssembly().Location;

                PushButtonData sharedBtn = new PushButtonData(
                    "SharedCoords",
                    "Export\nShared Coords",
                    dllPath,
                    "SharedCoordExporter.Command"
                );

                PushButtonData propsBtn = new PushButtonData(
                    "AllProps",
                    "Export\nAll Properties",
                    dllPath,
                    "SharedCoordExporter.GetAllProperties"
                );

                PushButtonData uuidBtn = new PushButtonData(
                    "SelectUUID",
                    "Select\nBy UUID",
                    dllPath,
                    "SelectByUUID.SearchUsingUUID"
                );

                PushButtonData rotateBtn = new PushButtonData(
                    "RotateModel",
                    "Rotate\nModel",
                    dllPath,
                    "SharedCoordExporter.RotateModelCommand"
                );

                panel.AddItem(sharedBtn);
                panel.AddItem(propsBtn);
                panel.AddItem(uuidBtn);
                panel.AddItem(rotateBtn);

                return Result.Succeeded;
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Ribbon Error", ex.Message);
                return Result.Failed;
            }
        }

        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Succeeded;
        }
    }
}
