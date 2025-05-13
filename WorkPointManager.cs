using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;

namespace TakeoffBridge
{
    /// <summary>
    /// Manages work points in part files and fabrication tickets
    /// </summary>
    public class WorkPointManager
    {
        // Constants
        private const string WorkPointDictName = "WORKPOINTS";
        private const string WorkPointName = "PRIMARY";

        /// <summary>
        /// Command to set a work point in the current drawing
        /// </summary>
        [CommandMethod("SETWORKPOINT")]
        public static void SetWorkPointCommand()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // Prompt for point selection
            PromptPointOptions ppo = new PromptPointOptions("\nSelect work point location: ");
            PromptPointResult ppr = ed.GetPoint(ppo);

            if (ppr.Status != PromptStatus.OK)
                return;

            Point3d workPoint = ppr.Value;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                try
                {
                    // Store the work point
                    StoreWorkPoint(db, tr, workPoint);
                    tr.Commit();
                    ed.WriteMessage($"\nWork point set at: ({workPoint.X}, {workPoint.Y}, {workPoint.Z})");
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nError setting work point: {ex.Message}");
                    tr.Abort();
                }
            }
        }

        /// <summary>
        /// Stores a work point in the drawing's dictionary
        /// </summary>
        public static void StoreWorkPoint(Database db, Transaction tr, Point3d workPoint)
        {
            // Get the named objects dictionary
            DBDictionary nod = (DBDictionary)tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForRead);

            // Create or open the work points dictionary
            DBDictionary workPointsDict;
            if (!nod.Contains(WorkPointDictName))
            {
                // Create new dictionary for work points
                nod.UpgradeOpen();
                workPointsDict = new DBDictionary();
                nod.SetAt(WorkPointDictName, workPointsDict);
                tr.AddNewlyCreatedDBObject(workPointsDict, true);
            }
            else
            {
                // Open existing dictionary
                workPointsDict = (DBDictionary)tr.GetObject(nod.GetAt(WorkPointDictName), OpenMode.ForWrite);
            }

            // Create or update the work point entry
            Xrecord xrec;
            if (!workPointsDict.Contains(WorkPointName))
            {
                xrec = new Xrecord();
                workPointsDict.SetAt(WorkPointName, xrec);
                tr.AddNewlyCreatedDBObject(xrec, true);
            }
            else
            {
                xrec = (Xrecord)tr.GetObject(workPointsDict.GetAt(WorkPointName), OpenMode.ForWrite);
            }

            // Store the point data in the Xrecord
            xrec.Data = new ResultBuffer(
                new TypedValue((int)DxfCode.Real, workPoint.X),
                new TypedValue((int)DxfCode.Real, workPoint.Y),
                new TypedValue((int)DxfCode.Real, workPoint.Z)
            );
        }

        /// <summary>
        /// Retrieves a work point from the drawing's dictionary
        /// </summary>
        public static Point3d? GetWorkPoint(Database db, Transaction tr)
        {
            // Get the named objects dictionary
            DBDictionary nod = (DBDictionary)tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForRead);

            // Check if work points dictionary exists
            if (!nod.Contains(WorkPointDictName))
                return null;

            // Open the work points dictionary
            DBDictionary workPointsDict = (DBDictionary)tr.GetObject(nod.GetAt(WorkPointDictName), OpenMode.ForRead);

            // Check if the primary work point exists
            if (!workPointsDict.Contains(WorkPointName))
                return null;

            // Get the work point record
            Xrecord xrec = (Xrecord)tr.GetObject(workPointsDict.GetAt(WorkPointName), OpenMode.ForRead);
            if (xrec.Data == null)
                return null;

            // Parse the point data
            TypedValue[] values = xrec.Data.AsArray();
            if (values.Length >= 3 &&
                values[0].TypeCode == (int)DxfCode.Real &&
                values[1].TypeCode == (int)DxfCode.Real &&
                values[2].TypeCode == (int)DxfCode.Real)
            {
                double x = (double)values[0].Value;
                double y = (double)values[1].Value;
                double z = (double)values[2].Value;
                return new Point3d(x, y, z);
            }

            return null;
        }

        /// <summary>
        /// Copies a work point from one database to another
        /// </summary>
        public static void CopyWorkPoint(Database sourceDb, Database targetDb, Transaction sourceTr, Transaction targetTr)
        {
            // Get the work point from source database
            Point3d? workPoint = GetWorkPoint(sourceDb, sourceTr);
            if (!workPoint.HasValue)
                return;

            // Store the work point in target database
            StoreWorkPoint(targetDb, targetTr, workPoint.Value);
        }

        /// <summary>
        /// Creates a marker to visualize the work point
        /// </summary>
        public static void CreateWorkPointMarker(Database db, Transaction tr, Point3d workPoint)
        {
            // Get the current space (model or paper)
            BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
            BlockTableRecord space = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

            // Create a work point marker (crosshair)
            double size = 0.25; // Size of crosshair

            // Horizontal line
            using (Line line1 = new Line())
            {
                line1.StartPoint = new Point3d(workPoint.X - size, workPoint.Y, workPoint.Z);
                line1.EndPoint = new Point3d(workPoint.X + size, workPoint.Y, workPoint.Z);
                line1.Layer = "WORKPOINTS";
                space.AppendEntity(line1);
                tr.AddNewlyCreatedDBObject(line1, true);
            }

            // Vertical line
            using (Line line2 = new Line())
            {
                line2.StartPoint = new Point3d(workPoint.X, workPoint.Y - size, workPoint.Z);
                line2.EndPoint = new Point3d(workPoint.X, workPoint.Y + size, workPoint.Z);
                line2.Layer = "WORKPOINTS";
                space.AppendEntity(line2);
                tr.AddNewlyCreatedDBObject(line2, true);
            }

            // Vertical Z line
            using (Line line3 = new Line())
            {
                line3.StartPoint = new Point3d(workPoint.X, workPoint.Y, workPoint.Z - size);
                line3.EndPoint = new Point3d(workPoint.X, workPoint.Y, workPoint.Z + size);
                line3.Layer = "WORKPOINTS";
                space.AppendEntity(line3);
                tr.AddNewlyCreatedDBObject(line3, true);
            }
        }

        /// <summary>
        /// Ensures the WORKPOINTS layer exists
        /// </summary>
        public static void EnsureWorkpointsLayer(Database db, Transaction tr)
        {
            // Get the layer table
            LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);

            // Create WORKPOINTS layer if it doesn't exist
            if (!lt.Has("WORKPOINTS"))
            {
                lt.UpgradeOpen();
                LayerTableRecord ltr = new LayerTableRecord();
                ltr.Name = "WORKPOINTS";
                ltr.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByAci, 1); // Red
                ltr.IsPlottable = false; // Make it non-plottable
                ltr.ViewportVisibilityDefault = false; // Hide in new viewports by default
                lt.Add(ltr);
                tr.AddNewlyCreatedDBObject(ltr, true);
            }
        }

        /// <summary>
        /// Command to visualize work points in the current drawing
        /// </summary>
        [CommandMethod("SHOWWORKPOINT")]
        public static void ShowWorkPointCommand()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                try
                {
                    // Get the work point
                    Point3d? workPoint = GetWorkPoint(db, tr);
                    if (!workPoint.HasValue)
                    {
                        ed.WriteMessage("\nNo work point found in this drawing.");
                        return;
                    }

                    // Ensure layer exists
                    EnsureWorkpointsLayer(db, tr);

                    // Create a marker
                    CreateWorkPointMarker(db, tr, workPoint.Value);

                    tr.Commit();
                    ed.WriteMessage($"\nWork point visualized at: ({workPoint.Value.X}, {workPoint.Value.Y}, {workPoint.Value.Z})");
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nError visualizing work point: {ex.Message}");
                    tr.Abort();
                }
            }
        }
    }
}