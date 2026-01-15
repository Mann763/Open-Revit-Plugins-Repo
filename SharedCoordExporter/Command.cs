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
    public class Command : IExternalCommand
    {
        class FlowData
        {
            public List<string> In = new List<string>();
            public List<string> Out = new List<string>();
        }

        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            TaskDialog td = new TaskDialog("Export Options");
            td.MainInstruction = "Flow Export";
            td.MainContent = "Choose how connections are resolved.";
            td.VerificationText = "Skip pipe fittings & accessories (direct connections only)";
            td.CommonButtons = TaskDialogCommonButtons.Ok | TaskDialogCommonButtons.Cancel;

            if (td.Show() != TaskDialogResult.Ok)
                return Result.Cancelled;

            bool skipAccessories = td.WasVerificationChecked();

            string desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string filePath = Path.Combine(desktop, "MEP_Geo_Saudi_Corrected.csv");

            ProjectLocation projLoc = doc.ActiveProjectLocation;
            SiteLocation site = doc.SiteLocation;
            Transform sharedTransform = projLoc.GetTransform();
            double earthRadius = 6378137.0;

            List<BuiltInCategory> categories = new List<BuiltInCategory>
            {
                BuiltInCategory.OST_PipeCurves,
                BuiltInCategory.OST_MechanicalEquipment,
                BuiltInCategory.OST_PlumbingFixtures
            };

            if (!skipAccessories)
            {
                categories.Add(BuiltInCategory.OST_PipeFitting);
                categories.Add(BuiltInCategory.OST_PipeAccessory);
            }

            ElementMulticategoryFilter catFilter = new ElementMulticategoryFilter(categories);
            FilteredElementCollector collector = new FilteredElementCollector(doc)
                .WherePasses(catFilter)
                .WhereElementIsNotElementType();

            using (StreamWriter sw = new StreamWriter(filePath, false, Encoding.UTF8))
            {
                sw.WriteLine(
                    "UniqueID,Category,Name,PointLabel,Easting_M,Northing_M,Elevation_M,Latitude,Longitude,Altitude_M," +
                    "Pipe_IN,Pipe_OUT,Valve_IN,Valve_OUT,Pump_IN,Pump_OUT,Tank_IN,Tank_OUT,FlowMeter_IN,FlowMeter_OUT,Chiller_IN,Chiller_OUT");

                foreach (Element el in collector)
                {
                    try
                    {
                        string elUid = el.UniqueId;
                        string elCat = el.Category?.Name ?? "";
                        string elName = el.Name.Replace(',', ';');

                        List<(string, XYZ)> points = new List<(string, XYZ)>();

                        if (el.Location is LocationCurve lc)
                        {
                            points.Add(("Start", lc.Curve.GetEndPoint(0)));
                            points.Add(("End", lc.Curve.GetEndPoint(1)));
                        }
                        else if (el.Location is LocationPoint lp)
                        {
                            points.Add(("Location", lp.Point));
                        }

                        var con = GetFlowConnections(el, skipAccessories);

                        foreach (var p in points)
                        {
                            XYZ sharedPt = sharedTransform.OfPoint(p.Item2);

                            double eastM = UnitUtils.ConvertFromInternalUnits(sharedPt.X, UnitTypeId.Meters);
                            double northM = UnitUtils.ConvertFromInternalUnits(sharedPt.Y, UnitTypeId.Meters);
                            double elevM = UnitUtils.ConvertFromInternalUnits(sharedPt.Z, UnitTypeId.Meters);

                            double internalNorthM = UnitUtils.ConvertFromInternalUnits(p.Item2.Y, UnitTypeId.Meters);
                            double internalEastM = UnitUtils.ConvertFromInternalUnits(p.Item2.X, UnitTypeId.Meters);

                            double latOffset = internalNorthM / earthRadius;
                            double lonOffset = internalEastM / (earthRadius * Math.Cos(site.Latitude));

                            double latDeg = (site.Latitude + latOffset) * (180.0 / Math.PI);
                            double lonDeg = (site.Longitude + lonOffset) * (180.0 / Math.PI);

                            sw.WriteLine(
                                $"{elUid},{elCat},{elName},{p.Item1},{eastM:F4},{northM:F4},{elevM:F4},{latDeg:F8},{lonDeg:F8},{elevM:F4}," +
                                $"{string.Join("|", con["Pipe"].In)},{string.Join("|", con["Pipe"].Out)}," +
                                $"{string.Join("|", con["Valve"].In)},{string.Join("|", con["Valve"].Out)}," +
                                $"{string.Join("|", con["Pump"].In)},{string.Join("|", con["Pump"].Out)}," +
                                $"{string.Join("|", con["Tank"].In)},{string.Join("|", con["Tank"].Out)}," +
                                $"{string.Join("|", con["FlowMeter"].In)},{string.Join("|", con["FlowMeter"].Out)}," +
                                $"{string.Join("|", con["Chiller"].In)},{string.Join("|", con["Chiller"].Out)}");
                        }
                    }
                    catch { }
                }
            }

            TaskDialog.Show("Export Complete", "Flow-aware connections exported.");
            return Result.Succeeded;
        }

        private Dictionary<string, FlowData> GetFlowConnections(Element el, bool skipAccessories)
        {
            var map = new Dictionary<string, FlowData>
            {
                { "Pipe", new FlowData() },
                { "Valve", new FlowData() },
                { "Pump", new FlowData() },
                { "Tank", new FlowData() },
                { "FlowMeter", new FlowData() },
                { "Chiller", new FlowData() }
            };

            ConnectorSet connectors = null;

            if (el is Autodesk.Revit.DB.Plumbing.Pipe p)
                connectors = p.ConnectorManager.Connectors;
            else if (el is FamilyInstance fi && fi.MEPModel != null)
                connectors = fi.MEPModel.ConnectorManager.Connectors;

            if (connectors == null || connectors.IsEmpty)
                return map;

            foreach (Connector c in connectors)
            {
                foreach (Connector rc in c.AllRefs)
                {
                    if (rc.Owner.Id == el.Id) continue;

                    Element target = rc.Owner;

                    if (skipAccessories && IsAccessoryOrFitting(target))
                    {
                        target = ResolveThroughAccessory(target, el);
                        if (target == null) continue;
                    }

                    string key = GetElementKey(target);
                    if (key == null) continue;

                    string uid = target.UniqueId;

                    if (c.Direction == FlowDirectionType.In)
                        map[key].In.Add(uid);
                    else if (c.Direction == FlowDirectionType.Out)
                        map[key].Out.Add(uid);
                    else
                    {
                        map[key].In.Add(uid);
                        map[key].Out.Add(uid);
                    }
                }
            }

            foreach (var k in map.Keys.ToList())
            {
                map[k].In = map[k].In.Distinct().ToList();
                map[k].Out = map[k].Out.Distinct().ToList();
            }

            return map;
        }

        private bool IsAccessoryOrFitting(Element el)
        {
            if (el.Category == null) return false;

            long id = el.Category.Id.Value;
            return id == (int)BuiltInCategory.OST_PipeFitting ||
                   id == (int)BuiltInCategory.OST_PipeAccessory;
        }

        private Element ResolveThroughAccessory(Element accessory, Element source)
        {
            ConnectorSet conns = null;

            if (accessory is FamilyInstance fi && fi.MEPModel != null)
                conns = fi.MEPModel.ConnectorManager.Connectors;

            if (conns == null) return null;

            foreach (Connector c in conns)
            {
                foreach (Connector rc in c.AllRefs)
                {
                    if (rc.Owner.Id == accessory.Id) continue;
                    if (rc.Owner.Id == source.Id) continue;

                    return rc.Owner;
                }
            }
            return null;
        }

        private string GetElementKey(Element el)
        {
            if (el is Autodesk.Revit.DB.Plumbing.Pipe)
                return "Pipe";

            if (el is FamilyInstance fi)
            {
                string n = fi.Symbol.Family.Name.ToLower();
                if (n.Contains("valve")) return "Valve";
                if (n.Contains("pump")) return "Pump";
                if (n.Contains("tank")) return "Tank";
                if (n.Contains("flow")) return "FlowMeter";
                if (n.Contains("chiller")) return "Chiller";
            }
            return null;
        }
    }
}
