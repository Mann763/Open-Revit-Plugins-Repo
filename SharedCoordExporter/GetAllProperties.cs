using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace SharedCoordExporter
{
    [Transaction(TransactionMode.Manual)]
    public class GetAllProperties : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;
            View view = doc.ActiveView;

            string path = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.Desktop),
                "Element_Properties_Matrix.csv"
            );

            var collector = new FilteredElementCollector(doc, view.Id)
                .WhereElementIsNotElementType()
                .ToList();

            HashSet<string> allProps = new HashSet<string>();
            Dictionary<ElementId, Dictionary<string, string>> data =
                new Dictionary<ElementId, Dictionary<string, string>>();

            foreach (Element el in collector)
            {
                var dict = new Dictionary<string, string>();
                foreach (Parameter p in el.Parameters)
                {
                    if (p.Definition == null) continue;

                    string name = p.Definition.Name;
                    string val = GetValue(p);

                    dict[name] = val;
                    allProps.Add(name);
                }
                data[el.Id] = dict;
            }

            var headers = allProps.OrderBy(x => x).ToList();

            using (StreamWriter sw = new StreamWriter(path, false, Encoding.UTF8))
            {
                sw.Write("ElementId,UniqueId,Category,Name");
                foreach (string h in headers)
                    sw.Write("," + h.Replace(",", ";"));
                sw.WriteLine();

                foreach (Element el in collector)
                {
                    sw.Write($"{el.Id.Value},{el.UniqueId},{el.Category?.Name},{el.Name.Replace(",", ";")}");

                    var values = data[el.Id];
                    foreach (string h in headers)
                        sw.Write("," + (values.ContainsKey(h) ? values[h] : ""));

                    sw.WriteLine();
                }
            }

            TaskDialog.Show("Done", "Clean matrix exported (no repetition).");
            return Result.Succeeded;
        }

        private string GetValue(Parameter p)
        {
            switch (p.StorageType)
            {
                case StorageType.String: return p.AsString();
                case StorageType.Double: return p.AsValueString();
                case StorageType.Integer: return p.AsInteger().ToString();
                case StorageType.ElementId: return p.AsElementId().Value.ToString();
                default: return "";
            }
        }
    }
}
