using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.Geometry;
using System;
using System.IO;
using System.Collections.Generic;

namespace TakeoffBridge
{
    public class DrawfabCommands
    {
        [CommandMethod("CREATEFABPART")]
        public void CreateFabPart()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            try
            {
                // 1. Ask for the part drawing file
                PromptOpenFileOptions fileOpts = new PromptOpenFileOptions("\nSelect part drawing file: ")
                {
                    Filter = "Drawing Files (*.dwg)|*.dwg|All files (*.*)|*.*"
                };
                PromptFileNameResult fileRes = ed.GetFileNameForOpen(fileOpts);
                if (fileRes.Status != PromptStatus.OK) return;

                string partDrawingPath = fileRes.StringResult;
                string partFileName = Path.GetFileNameWithoutExtension(partDrawingPath);

                // 2. Ask for length
                PromptDoubleOptions lengthOpts = new PromptDoubleOptions("\nEnter extrusion length: ")
                {
                    DefaultValue = 48.0,
                    AllowNegative = false,
                    AllowZero = false
                };
                PromptDoubleResult lengthRes = ed.GetDouble(lengthOpts);
                if (lengthRes.Status != PromptStatus.OK) return;

                double length = lengthRes.Value;

                // 3. Ask for miter cuts
                double miterLeft = 90.0, tiltLeft = 90.0, miterRight = 90.0, tiltRight = 90.0;

                PromptKeywordOptions miterOpts = new PromptKeywordOptions("\nApply miter cuts? ");
                miterOpts.Keywords.Add("Yes");
                miterOpts.Keywords.Add("No");
                miterOpts.Keywords.Default = "No";
                miterOpts.AllowNone = true;

                PromptResult miterRes = ed.GetKeywords(miterOpts);
                if (miterRes.Status != PromptStatus.OK) return;

                bool applyMiter = miterRes.StringResult == "Yes";

                if (applyMiter)
                {
                    // Get miter and tilt angles
                    PromptDoubleOptions angOpts = new PromptDoubleOptions("\nEnter left miter angle (90 = straight cut): ")
                    {
                        DefaultValue = 90.0,
                        AllowNegative = false
                    };

                    PromptDoubleResult res = ed.GetDouble(angOpts);
                    if (res.Status != PromptStatus.OK) return;
                    miterLeft = res.Value;

                    angOpts.Message = "\nEnter left tilt angle (90 = straight cut): ";
                    res = ed.GetDouble(angOpts);
                    if (res.Status != PromptStatus.OK) return;
                    tiltLeft = res.Value;

                    angOpts.Message = "\nEnter right miter angle (90 = straight cut): ";
                    res = ed.GetDouble(angOpts);
                    if (res.Status != PromptStatus.OK) return;
                    miterRight = res.Value;

                    angOpts.Message = "\nEnter right tilt angle (90 = straight cut): ";
                    res = ed.GetDouble(angOpts);
                    if (res.Status != PromptStatus.OK) return;
                    tiltRight = res.Value;
                }

                // 4. Ask for insertion point and orientation
                PromptPointOptions ptOpts = new PromptPointOptions("\nSpecify insertion point: ");
                PromptPointResult ptRes = ed.GetPoint(ptOpts);
                if (ptRes.Status != PromptStatus.OK) return;

                Point3d insertionPoint = ptRes.Value;

                // 5. Create matrix for final position
                Matrix3d finalTransform = Matrix3d.Displacement(
                    new Vector3d(insertionPoint.X, insertionPoint.Y, insertionPoint.Z));

                // 6. Create the part
                ed.WriteMessage("\nCreating extruded part...");

                string blockName = "FAB_" + partFileName + "_" + DateTime.Now.Ticks.ToString();

                Drawfab drawfab = new Drawfab();
                List<Solid3d> solids = drawfab.CreateExtrudedPart(
                    db,
                    partDrawingPath,
                    blockName,
                    length,
                    miterLeft,
                    tiltLeft,
                    miterRight,
                    tiltRight,
                    false,
                    string.Empty,
                    finalTransform);

                ed.WriteMessage($"\nSuccessfully created {solids.Count} solid(s).");
                ed.WriteMessage($"\nPart dimensions: Width = {drawfab.Width:F3}, Height = {drawfab.Height:F3}");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError: {ex.Message}");
                ed.WriteMessage($"\nStack trace: {ex.StackTrace}");
            }
        }

        [CommandMethod("QF")]
        public void QuickFabPart()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            try
            {
                ed.WriteMessage("\nCreating test part with custom angles...");

                // Get the current directory where the DWG is located
                string dwgPath = doc.Name;
                string currentDirectory = Path.GetDirectoryName(dwgPath);

                // Create path to the template file
                string templatePath = Path.Combine(currentDirectory, "601TVG005.dwg");

                // Check if the template file exists
                if (!File.Exists(templatePath))
                {
                    ed.WriteMessage($"\nError: Template file not found at {templatePath}");
                    return;
                }

                // Get insertion point
                

                Point3d insertionPoint = new Point3d(0, 0, 0);

                double length = 30.0;

                // Ask for left miter angle
                PromptDoubleOptions leftMiterOpts = new PromptDoubleOptions("\nEnter left miter angle (90 = straight): ");
                leftMiterOpts.DefaultValue = 90.0;
                PromptDoubleResult leftMiterRes = ed.GetDouble(leftMiterOpts);
                if (leftMiterRes.Status != PromptStatus.OK) return;
                double miterLeft = leftMiterRes.Value;

                // Ask for left tilt angle
                PromptDoubleOptions leftTiltOpts = new PromptDoubleOptions("\nEnter left tilt angle (90 = straight): ");
                leftTiltOpts.DefaultValue = 90.0;
                PromptDoubleResult leftTiltRes = ed.GetDouble(leftTiltOpts);
                if (leftTiltRes.Status != PromptStatus.OK) return;
                double tiltLeft = leftTiltRes.Value;

                // Ask for right miter angle
                PromptDoubleOptions rightMiterOpts = new PromptDoubleOptions("\nEnter right miter angle (90 = straight): ");
                rightMiterOpts.DefaultValue = 90.0;
                PromptDoubleResult rightMiterRes = ed.GetDouble(rightMiterOpts);
                if (rightMiterRes.Status != PromptStatus.OK) return;
                double miterRight = rightMiterRes.Value;

                // Ask for right tilt angle
                PromptDoubleOptions rightTiltOpts = new PromptDoubleOptions("\nEnter right tilt angle (90 = straight): ");
                rightTiltOpts.DefaultValue = 90.0;
                PromptDoubleResult rightTiltRes = ed.GetDouble(rightTiltOpts);
                if (rightTiltRes.Status != PromptStatus.OK) return;
                double tiltRight = rightTiltRes.Value;


                bool preserveOriginalOrientation = false;

                // Create block name
                string blockName = "QUICKFAB_" + DateTime.Now.Ticks.ToString();

                // Create transformation
                Matrix3d finalTransform = Matrix3d.Displacement(
                    new Vector3d(insertionPoint.X, insertionPoint.Y, insertionPoint.Z));

                ed.WriteMessage($"\nCreating part with - Left: {miterLeft}°/{tiltLeft}°, Right: {miterRight}°/{tiltRight}°");
                ed.WriteMessage($"\nPreserve original orientation: {preserveOriginalOrientation}");

                bool handed = false; //setting this to true will mirror the part front to back for verticals
                bool addDimensions = true;

                // Create part
                Drawfab drawfab = new Drawfab();
                List<Solid3d> solids = drawfab.CreateExtrudedPart(
                    db,
                    templatePath,
                    blockName,
                    length,
                    miterLeft,
                    tiltLeft,
                    miterRight,
                    tiltRight,
                    handed,
                    string.Empty,
                    finalTransform,
                    preserveOriginalOrientation,
                    addDimensions);

                ed.WriteMessage($"\nSuccessfully created {solids.Count} solid(s).");
                ed.WriteMessage($"\nPart dimensions: Width = {drawfab.Width:F3}, Height = {drawfab.Height:F3}");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError: {ex.Message}");
                ed.WriteMessage($"\nStack trace: {ex.StackTrace}");
            }
        }
    }
}