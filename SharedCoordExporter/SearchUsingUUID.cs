using Autodesk.Revit.UI;
using Autodesk.Revit.DB;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.UI.Selection;
using Microsoft.VisualBasic;
using System.Collections.Generic;

namespace SelectByUUID
{
    [Transaction(TransactionMode.Manual)]
    public class SearchUsingUUID : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            string uuid = Interaction.InputBox("Enter Element UUID (UniqueId)", "Select Element");

            if (string.IsNullOrWhiteSpace(uuid))
                return Result.Cancelled;

            Element element = doc.GetElement(uuid);

            if (element == null)
            {
                TaskDialog.Show("Result", "Element not found");
                return Result.Failed;
            }

            uidoc.Selection.SetElementIds(new List<ElementId> { element.Id });
            uidoc.ShowElements(element.Id);

            return Result.Succeeded;
        }
    }
}
