using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Linq;

namespace SharedCoordExporter
{
    [Transaction(TransactionMode.Manual)]
    public class Command : IExternalCommand
    {
        public Result Execute(
            ExternalCommandData commandData,
            ref string message,
            ElementSet elements)
        {
            Document doc = commandData.Application.ActiveUIDocument.Document;

            string desktop = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
            string filePath = Path.Combine(desktop, "MEP_Geo_Saudi_Corrected.csv");

            // 1. Get Project Locations
            ProjectLocation projLoc = doc.ActiveProjectLocation;
            SiteLocation site = doc.SiteLocation;

            // 2. This transform converts from Internal Revit coordinates (feet)
            // to Shared Coordinates (Easting/Northing in feet)
            Transform sharedTransform = projLoc.GetTransform();

            // 3. To get Geo-coordinates properly, we need the distance from the 
            // Internal Origin, NOT the Easting/Northing value.
            // We use the SiteLocation's position as the 0,0 geographic anchor.
            double earthRadius = 6378137.0;

            List<BuiltInCategory> categories = new List<BuiltInCategory>
            {
                BuiltInCategory.OST_PipeFitting,
                BuiltInCategory.OST_PipeAccessory,
                BuiltInCategory.OST_MechanicalEquipment,
                BuiltInCategory.OST_PlumbingFixtures,
                BuiltInCategory.OST_PipeCurves
            };

            ElementMulticategoryFilter catFilter = new ElementMulticategoryFilter(categories);
            FilteredElementCollector collector = new FilteredElementCollector(doc)
                .WherePasses(catFilter)
                .WhereElementIsNotElementType();

            using (StreamWriter sw = new StreamWriter(filePath, false, Encoding.UTF8))
            {
                sw.WriteLine("UniqueID,Category,Name,PointLabel,Easting_M,Northing_M,Elevation_M,Latitude,Longitude,Altitude_M,Connected_UniqueIDs");

                foreach (Element el in collector)
                {
                    try
                    {
                        string elUid = el.UniqueId;
                        string elCat = el.Category.Name;
                        string elName = el.Name.Replace(',', ';');

                        List<(string label, XYZ point)> points = new List<(string, XYZ)>();
                        if (el.Location is LocationCurve lc)
                        {
                            points.Add(("Start", lc.Curve.GetEndPoint(0)));
                            points.Add(("End", lc.Curve.GetEndPoint(1)));
                        }
                        else if (el.Location is LocationPoint lp)
                        {
                            points.Add(("Location", lp.Point));
                        }

                        string connectedString = GetConnectedIDs(el);

                        foreach (var (label, pt) in points)
                        {
                            // A. Get Easting/Northing/Elevation (Shared Coordinates)
                            XYZ sharedPt = sharedTransform.OfPoint(pt);
                            double eastM = UnitUtils.ConvertFromInternalUnits(sharedPt.X, UnitTypeId.Meters);
                            double northM = UnitUtils.ConvertFromInternalUnits(sharedPt.Y, UnitTypeId.Meters);
                            double elevM = UnitUtils.ConvertFromInternalUnits(sharedPt.Z, UnitTypeId.Meters);

                            // B. Calculate Geographic Offset
                            // We use the internal Revit distance (pt) for the offset calculation
                            // to avoid the "Africa" shift caused by using UTM Easting/Northing values.
                            double internalNorthM = UnitUtils.ConvertFromInternalUnits(pt.Y, UnitTypeId.Meters);
                            double internalEastM = UnitUtils.ConvertFromInternalUnits(pt.X, UnitTypeId.Meters);

                            double latOffset = internalNorthM / earthRadius;
                            double lonOffset = internalEastM / (earthRadius * Math.Cos(site.Latitude));

                            double latDeg = (site.Latitude + latOffset) * (180.0 / Math.PI);
                            double lonDeg = (site.Longitude + lonOffset) * (180.0 / Math.PI);

                            sw.WriteLine($"{elUid},{elCat},{elName},{label},{eastM:F4},{northM:F4},{elevM:F4},{latDeg:F8},{lonDeg:F8},{elevM:F4},{connectedString}");
                        }
                    }
                    catch { continue; }
                }
            }

            TaskDialog.Show("Export Complete", "Coordinates calculated for Saudi Arabia UTM-38N.");
            return Result.Succeeded;
        }

        private string GetConnectedIDs(Element el)
        {
            ConnectorSet connectors = null;
            if (el is Autodesk.Revit.DB.Plumbing.Pipe pipe)
                connectors = pipe.ConnectorManager.Connectors;
            else if (el is FamilyInstance fi && fi.MEPModel != null)
                connectors = fi.MEPModel.ConnectorManager?.Connectors;

            if (connectors == null || connectors.IsEmpty) return "None";

            List<string> ids = new List<string>();
            foreach (Connector conn in connectors)
            {
                foreach (Connector refConn in conn.AllRefs)
                {
                    if (refConn.Owner.Id != el.Id)
                        ids.Add(refConn.Owner.UniqueId);
                }
            }
            return ids.Count > 0 ? string.Join("|", ids.Distinct()) : "None";
        }
    }
}