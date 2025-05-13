using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TakeoffBridge
{
    // Define a class to store mullion placement data
    public class MullionPlacementData
    {
        public double Width { get; set; }
        public double GlassPocketOffset { get; set; }
        public double OverallWidth { get; set; }


        // Command to set and store mullion data
        [CommandMethod("StoreMullionData")]
        public void StoreMullionData()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            // Prompt for width
            PromptDoubleOptions widthOpts = new PromptDoubleOptions("\nEnter mullion width: ");
            widthOpts.DefaultValue = 2.5; // Default to 2.5 inches
            widthOpts.UseDefaultValue = true;
            PromptDoubleResult widthRes = ed.GetDouble(widthOpts);
            if (widthRes.Status != PromptStatus.OK)
                return;

            // Prompt for glass pocket offset
            PromptDoubleOptions offsetOpts = new PromptDoubleOptions("\nEnter offset from face to back of glass pocket: ");
            offsetOpts.DefaultValue = 1.25; // Default to 1.25 inches
            offsetOpts.UseDefaultValue = true;
            PromptDoubleResult offsetRes = ed.GetDouble(offsetOpts);
            if (offsetRes.Status != PromptStatus.OK)
                return;

            // Store the data
            MullionPlacementData data = new MullionPlacementData
            {
                Width = widthRes.Value,
                GlassPocketOffset = offsetRes.Value
            };

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                try
                {
                    // Store the data in the drawing dictionary
                    StoreMullionData(db, tr, data);
                    ed.WriteMessage($"\nMullion data stored: Width = {data.Width}, Glass Pocket Offset = {data.GlassPocketOffset}");
                    tr.Commit();
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nError storing mullion data: {ex.Message}");
                    tr.Abort();
                }
            }
        }

        // Method to store mullion data in the drawing dictionary
        public static void StoreMullionData(Database db, Transaction tr, MullionPlacementData data)
        {
            // Get the named objects dictionary
            DBDictionary nod = (DBDictionary)tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForRead);

            // Create or get our mullion data dictionary
            DBDictionary mullionDict;
            if (!nod.Contains("MULLION_DATA"))
            {
                nod.UpgradeOpen();
                mullionDict = new DBDictionary();
                nod.SetAt("MULLION_DATA", mullionDict);
                tr.AddNewlyCreatedDBObject(mullionDict, true);
            }
            else
            {
                ObjectId dictId = nod.GetAt("MULLION_DATA");
                mullionDict = (DBDictionary)tr.GetObject(dictId, OpenMode.ForWrite);
            }

            // Create or update the mullion data record
            Xrecord xrec;
            if (!mullionDict.Contains("PLACEMENT_DATA"))
            {
                xrec = new Xrecord();
                mullionDict.SetAt("PLACEMENT_DATA", xrec);
                tr.AddNewlyCreatedDBObject(xrec, true);
            }
            else
            {
                ObjectId xrecId = mullionDict.GetAt("PLACEMENT_DATA");
                xrec = (Xrecord)tr.GetObject(xrecId, OpenMode.ForWrite);
            }

            // Store the data in the Xrecord
            ResultBuffer rb = new ResultBuffer(
                new TypedValue((int)DxfCode.Real, data.Width),
                new TypedValue((int)DxfCode.Real, data.GlassPocketOffset)
            );
            xrec.Data = rb;
        }

        // Method to retrieve mullion data from the drawing dictionary
        public static MullionPlacementData GetMullionData(Database db, Transaction tr)
        {
            try
            {
                // Get the named objects dictionary
                DBDictionary nod = (DBDictionary)tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForRead);

                // Check if our dictionary exists
                if (!nod.Contains("MULLION_DATA"))
                    return null;

                // Get our custom dictionary
                ObjectId dictId = nod.GetAt("MULLION_DATA");
                DBDictionary mullionDict = (DBDictionary)tr.GetObject(dictId, OpenMode.ForRead);

                // Check if the placement data exists
                if (!mullionDict.Contains("PLACEMENT_DATA"))
                    return null;

                // Get the data Xrecord
                ObjectId xrecId = mullionDict.GetAt("PLACEMENT_DATA");
                Xrecord xrec = (Xrecord)tr.GetObject(xrecId, OpenMode.ForRead);

                // Read the data from the Xrecord
                ResultBuffer rb = xrec.Data;
                if (rb == null)
                    return null;

                TypedValue[] tvs = rb.AsArray();
                if (tvs.Length < 2)
                    return null;

                // Create the data object
                MullionPlacementData data = new MullionPlacementData
                {
                    Width = (double)tvs[0].Value,
                    GlassPocketOffset = (double)tvs[1].Value
                };

                return data;
            }
            catch
            {
                return null;
            }
        }

        // Command to show the stored mullion data
        [CommandMethod("ShowMullionData")]
        public void ShowMullionData()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                MullionPlacementData data = GetMullionData(db, tr);
                if (data != null)
                {
                    ed.WriteMessage($"\nMullion data: Width = {data.Width}, Glass Pocket Offset = {data.GlassPocketOffset}");
                }
                else
                {
                    ed.WriteMessage("\nNo mullion data found in this drawing.");
                }
                tr.Commit();
            }
        }
    }
}
