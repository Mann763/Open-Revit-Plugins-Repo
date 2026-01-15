using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Linq;

namespace SharedCoordExporter
{
    [Transaction(TransactionMode.Manual)]
    public class RotateModelCommand : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            string input = Microsoft.VisualBasic.Interaction.InputBox(
                "Enter rotation angle in degrees (positive = CCW, negative = CW)",
                "Rotate Model",
                "0");

            if (!double.TryParse(input, out double angleDeg))
                return Result.Cancelled;

            double angleRad = angleDeg * Math.PI / 180.0;

            BasePoint pbp = new FilteredElementCollector(doc)
                .OfCategory(BuiltInCategory.OST_ProjectBasePoint)
                .FirstElement() as BasePoint;

            if (pbp == null)
                return Result.Failed;

            XYZ origin = new XYZ(pbp.Position.X, pbp.Position.Y, pbp.Position.Z);
            Line axis = Line.CreateBound(origin, origin + XYZ.BasisZ);

            var elems = new FilteredElementCollector(doc)
                .WhereElementIsNotElementType()
                .Where(e => e.Category != null && e.Category.CategoryType == CategoryType.Model)
                .Select(e => e.Id)
                .ToList();

            using (Transaction t = new Transaction(doc, "Rotate Model"))
            {
                t.Start();
                ElementTransformUtils.RotateElements(doc, elems, axis, angleRad);
                t.Commit();
            }

            TaskDialog.Show("Rotate Model", $"Model rotated by {angleDeg}°");
            return Result.Succeeded;
        }
    }
}
