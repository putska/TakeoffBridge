using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using System;
using System.IO;

namespace TakeoffBridge
{
    public class PartDrawingGenerator
    {
        // Mirror type enum to specify how a part should be mirrored
        public enum MirrorType
        {
            None = 0,
            VerticalMullion = 1,
            ShearBlockClip = 2,
            Both = 3
        }

        /// <summary>
        /// Draws a part using the existing AutoLISP drwprt function
        /// </summary>
        /// <param name="partNo">Part number/name</param>
        /// <param name="length">Length of the part</param>
        /// <param name="dorl">Door Opening Rotation Left</param>
        /// <param name="dotl">Door Opening Tilt Left</param>
        /// <param name="dorr">Door Opening Rotation Right</param>
        /// <param name="dotr">Door Opening Tilt Right</param>
        /// <param name="mir">Mirror type (0=None, 1=Vertical, 2=Shear)</param>
        /// <param name="location">Insert location</param>
        /// <param name="side">Side (L, R, B or null)</param>
        /// <param name="height">Height for attachments</param>
        /// <param name="outputPath">Output file path (optional)</param>
        /// <returns>True if successful</returns>
        public bool DrwPrt(
            string partNo,
            double length,
            double dorl,
            double dotl,
            double dorr,
            double dotr,
            int mir,
            Point3d location,
            string side,
            double height,
            string outputPath = null)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;

            try
            {

                string fullPartPath = Path.Combine(@"C:\CSE\Takeoff\fabs\dies", partNo);

                // Store current document
                Document originalDoc = Application.DocumentManager.MdiActiveDocument;

                // Prepare for lisp call
                string lispCommand = CreateLispCommand(
                    fullPartPath, length, dorl, dotl, dorr, dotr, mir,
                    location, side, height);

                // Execute the LISP command to draw the part
                doc.SendStringToExecute(lispCommand, true, false, true);

                // If an output path is provided, save the drawing
                if (!string.IsNullOrEmpty(outputPath))
                {
                    // Ensure directory exists
                    string directory = Path.GetDirectoryName(outputPath);
                    if (!Directory.Exists(directory))
                    {
                        Directory.CreateDirectory(directory);
                    }

                    // Save the drawing
                    SaveDrawingAs(doc.Database, outputPath);
                }

                return true;
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in DrwPrt: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Alternative method signature that matches your existing method in TemplateGenerator
        /// </summary>
        public bool DrawPart(
            string sourcePath,
            string outputPath,
            double length,
            double dorl,
            double dotl,
            double dorr,
            double dotr,
            MirrorType mirrorType,
            bool isAttached,
            string side,
            double height)
        {
            // Extract part number from source path
            string partNo = Path.GetFileName(sourcePath);

            // Create location based on attachment status
            Point3d location = isAttached ? new Point3d(0, 0, 0) : Point3d.Origin;

            // Call the main method
            return DrwPrt(
                partNo,
                length,
                dorl,
                dotl,
                dorr,
                dotr,
                (int)mirrorType,
                location,
                side,
                height,
                outputPath
            );
        }

        /// <summary>
        /// Creates the LISP command string to call drwprt
        /// </summary>
        private string CreateLispCommand(
            string partNo,
            double length,
            double dorl,
            double dotl,
            double dorr,
            double dotr,
            int mir,
            Point3d location,
            string side,
            double height)
        {
            // Make sure the part number has properly escaped backslashes for LISP
            string escapedPartNo = partNo.Replace("\\", "\\\\");

            // Format for LISP: (drwprt "PARTNO" LENGTH DORL DOTL DORR DOTR MIR LOCATION SIDE HEIGHT)
            string locationStr = location != Point3d.Origin ?
                $"(list {location.X} {location.Y} {location.Z})" :
                "nil";

            string sideStr = !string.IsNullOrEmpty(side) ? $"\"{side}\"" : "nil";

            return $"(drwprt \"{partNo}\" {length} {dorl} {dotl} {dorr} {dotr} {mir} {locationStr} {sideStr} {height})";
        }

        /// <summary>
        /// Save a drawing to a new file
        /// </summary>
        private void SaveDrawingAs(Database db, string outputPath)
        {
            try
            {
                // Direct approach: just save the input database to the output path
                db.SaveAs(outputPath, DwgVersion.Current);
            }
            catch (System.Exception ex)
            {
                // Log the error but don't throw
                System.Diagnostics.Debug.WriteLine($"Error saving drawing: {ex.Message}");
            }
        }

        /// <summary>
        /// Command method that can be called directly from AutoCAD to test the DrwPrt function
        /// </summary>
        [CommandMethod("TESTDRWPRT")]
        public void TestDrawPart()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor editor = doc.Editor;

            try
            {
                // Example parameters
                string partNo = "601TVG005";  // Replace with an actual part number
                double length = 30.0;
                string outputPath = Path.Combine(
                    @"C:\CSE\Takeoff\fabs\dies",
                    $"{partNo}-test.dwg"
                );

                // Call DrwPrt
                bool success = DrwPrt(
                    partNo,
                    length,
                    45,   // dorl
                    90,   // dotl
                    45,   // dorr
                    90,   // dotr
                    0,    // mir
                    Point3d.Origin,
                    null, // side
                    0,    // height
                    outputPath
                );

                if (success)
                {
                    editor.WriteMessage($"\nSuccessfully created part drawing: {outputPath}");
                }
                else
                {
                    editor.WriteMessage($"\nFailed to create part drawing");
                }
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage($"\nError in TestDrawPart: {ex.Message}");
            }
        }
    }
}