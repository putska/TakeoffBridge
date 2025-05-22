using System;
using System.Collections.Generic;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Windows;
using System.Linq;
using System.Net.Mail;
using System.Linq.Expressions;
using Newtonsoft.Json;
using System.Security.Cryptography;
using System.Drawing;
using System.Windows.Forms;

// This file contains our custom MetalComponent entity definition

namespace TakeoffBridge
{
   





    // Commands for working with metal components
    public class MetalComponentCommands
    {
        // Use the same naming throughout
        public static string PendingHandle = null; // Note: no underscore, and public access
        public static List<ChildPart> PendingParts = null; // Note: no underscore, and public access
        public static ParentComponentData PendingParentData;
        public static CopyPropertyData PendingCopyData;

        public class OpeningComponentData
        {
            public string Floor { get; set; }
            public string Elevation { get; set; }

            public string TopType { get; set; }
            public string BottomType { get; set; }
            public string LeftType { get; set; }
            public string RightType { get; set; }

            public bool CreateTop { get; set; } = true;
            public bool CreateBottom { get; set; } = true;
            public bool CreateLeft { get; set; } = true;
            public bool CreateRight { get; set; } = true;

            // Add this new property
            public bool ForceDialogToOpen { get; set; } = true;
        }

        // Add static field to hold the opening data
        public static OpeningComponentData PendingOpeningData;

        // Add these classes
        public class ParentComponentData
        {
            public string ComponentType { get; set; }
            public string Floor { get; set; }
            public string Elevation { get; set; }
        }

        public class CopyPropertyData
        {
            public string SourceHandle { get; set; }
            public List<string> TargetHandles { get; set; }
            public string ComponentType { get; set; }
            public string Floor { get; set; }
            public string Elevation { get; set; }
            public List<ChildPart> Parts { get; set; }
        }



        // Static palette set to maintain a single instance
        private static PaletteSet _paletteSet;


        // New CommandMethod to update parent data
        [CommandMethod("UPDATEPARENT")]
        public void UpdateParent()
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            try
            {
                // Check if we have pending data
                if (PendingHandle == null || PendingParentData == null)
                {
                    ed.WriteMessage("\nNo pending parent data update.");
                    return;
                }

                // Get the object ID from handle
                long longHandle = Convert.ToInt64(PendingHandle, 16);
                Handle h = new Handle(longHandle);
                ObjectId objId = db.GetObjectId(false, h, 0);

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    // Get the entity
                    Entity ent = tr.GetObject(objId, OpenMode.ForWrite) as Entity;

                    // Create component Xdata
                    ResultBuffer rb = new ResultBuffer(
                        new TypedValue((int)DxfCode.ExtendedDataRegAppName, "METALCOMP"),
                        new TypedValue((int)DxfCode.ExtendedDataAsciiString, PendingParentData.ComponentType),
                        new TypedValue((int)DxfCode.ExtendedDataAsciiString, PendingParentData.Floor),
                        new TypedValue((int)DxfCode.ExtendedDataAsciiString, PendingParentData.Elevation)
                    );

                    ent.XData = rb;
                    tr.Commit();
                }

                // Clear pending data
                //PendingHandle = null;
                PendingParentData = null;
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in UPDATEPARENT: {ex.Message}");
            }
        }

        [CommandMethod("EXECUTECOPY")]
        public void ExecuteCopy()
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            try
            {
                // Check if we have pending data
                if (PendingCopyData == null)
                {
                    ed.WriteMessage("\nNo pending copy operation.");
                    return;
                }

                // Process each target
                foreach (string targetHandle in PendingCopyData.TargetHandles)
                {
                    // Get the object ID from handle
                    long longHandle = Convert.ToInt64(targetHandle, 16);
                    Handle h = new Handle(longHandle);
                    ObjectId objId = db.GetObjectId(false, h, 0);

                    // Update parent data if needed
                    if (PendingCopyData.ComponentType != null)
                    {
                        using (Transaction tr = db.TransactionManager.StartTransaction())
                        {
                            Entity ent = tr.GetObject(objId, OpenMode.ForWrite) as Entity;

                            ResultBuffer rb = new ResultBuffer(
                                new TypedValue((int)DxfCode.ExtendedDataRegAppName, "METALCOMP"),
                                new TypedValue((int)DxfCode.ExtendedDataAsciiString, PendingCopyData.ComponentType),
                                new TypedValue((int)DxfCode.ExtendedDataAsciiString, PendingCopyData.Floor),
                                new TypedValue((int)DxfCode.ExtendedDataAsciiString, PendingCopyData.Elevation)
                            );

                            ent.XData = rb;
                            tr.Commit();
                        }
                    }

                    // Update child parts if needed
                    if (PendingCopyData.Parts != null)
                    {
                        // Store in static fields for the parts update command
                        PendingHandle = targetHandle;
                        PendingParts = new List<ChildPart>(PendingCopyData.Parts);

                        // Execute the update parts command
                        UpdateMetalXdata(targetHandle, PendingCopyData.Parts);
                    }
                }

                // Clear pending data
                PendingCopyData = null;
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in EXECUTECOPY: {ex.Message}");
            }
        }

        [CommandMethod("ADDMETALPART")]
        public void AddMetalPart()
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            try
            {
                // Check if we have pending data from the panel
                string componentType = "Horizontal";
                string floor = "01";
                string elevation = "A";
                List<ChildPart> customParts = null;

                // If we have pending data from the panel, use it
                if (PendingParentData != null)
                {
                    componentType = PendingParentData.ComponentType ?? componentType;
                    floor = PendingParentData.Floor ?? floor;
                    elevation = PendingParentData.Elevation ?? elevation;
                }

                // Use custom parts if they were provided
                if (PendingParts != null && PendingParts.Count > 0)
                {
                    customParts = new List<ChildPart>(PendingParts);
                }

                // Prompt for start point (we still need user input for this)
                PromptPointOptions pPtOpts = new PromptPointOptions("\nEnter start point: ");
                PromptPointResult pPtRes = ed.GetPoint(pPtOpts);
                if (pPtRes.Status != PromptStatus.OK) return;
                Point3d startPt = pPtRes.Value;

                // Prompt for end point
                pPtOpts = new PromptPointOptions("\nEnter end point: ");
                pPtOpts.UseBasePoint = true;
                pPtOpts.BasePoint = startPt;
                pPtRes = ed.GetPoint(pPtOpts);
                if (pPtRes.Status != PromptStatus.OK) return;
                Point3d endPt = pPtRes.Value;

                // Register all required application names first
                using (Transaction trReg = db.TransactionManager.StartTransaction())
                {
                    RegAppTable regTable = (RegAppTable)trReg.GetObject(db.RegAppTableId, OpenMode.ForRead);

                    // Register basic apps
                    RegisterApp(regTable, "METALCOMP", trReg);
                    RegisterApp(regTable, "METALPARTSINFO", trReg);

                    // Register chunk apps preemptively
                    for (int i = 0; i < 10; i++) // Assume we won't need more than 10 chunks
                    {
                        RegisterApp(regTable, $"METALPARTS{i}", trReg);
                    }

                    trReg.Commit();
                }

                // Now create everything in a single transaction
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    // Try to get the TAG layer
                    LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                    ObjectId tagLayerId = ObjectId.Null;

                    if (lt.Has("TAG"))
                    {
                        tagLayerId = lt["TAG"];
                    }

                    BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                    // Create polyline
                    Polyline pline = new Polyline();
                    pline.SetDatabaseDefaults();
                    pline.ColorIndex = 1; // Red color to make it visible
                    pline.ConstantWidth = 0.1; // Give it some width to be more visible
                    pline.AddVertexAt(0, new Point2d(startPt.X, startPt.Y), 0, 0, 0);
                    pline.AddVertexAt(1, new Point2d(endPt.X, endPt.Y), 0, 0, 0);

                    // If TAG layer exists, set the polyline's layer to TAG
                    if (tagLayerId != ObjectId.Null)
                    {
                        pline.LayerId = tagLayerId;
                    }

                    // Add polyline to the database
                    ObjectId plineId = btr.AppendEntity(pline);
                    tr.AddNewlyCreatedDBObject(pline, true);

                    // Add component Xdata
                    ResultBuffer rbComp = new ResultBuffer(
                        new TypedValue((int)DxfCode.ExtendedDataRegAppName, "METALCOMP"),
                        new TypedValue((int)DxfCode.ExtendedDataAsciiString, componentType),
                        new TypedValue((int)DxfCode.ExtendedDataAsciiString, floor),
                        new TypedValue((int)DxfCode.ExtendedDataAsciiString, elevation)
                    );
                    pline.XData = rbComp;

                    // Create child parts based on the component type if no custom parts provided
                    List<ChildPart> childParts = customParts;
                    if (childParts == null || childParts.Count == 0)
                    {
                        childParts = new List<ChildPart>();
                        if (componentType.ToUpper().Contains("HORIZONTAL"))
                        {
                            // Create parts with specific start and end adjustments
                            childParts.Add(new ChildPart("Horizontal Body", "HB", 0.0, 0.0, "Aluminum"));
                            childParts.Add(new ChildPart("Flat Filler", "FF", -0.03125, 0.0, "Aluminum"));
                            childParts.Add(new ChildPart("Face Cap", "FC", 0.0, 0.0, "Aluminum"));

                            // Left side attachments
                            ChildPart sbLeft = new ChildPart("Shear Block Left", "SBL", -1.25, 0.0, "Aluminum");
                            sbLeft.Attach = "L"; // Set attachment side
                            childParts.Add(sbLeft);

                            // Right side attachments
                            ChildPart sbRight = new ChildPart("Shear Block Right", "SBR", 0.0, -1.25, "Aluminum");
                            sbRight.Attach = "R"; // Set attachment side
                            childParts.Add(sbRight);
                        }
                        else if (componentType.ToUpper().Contains("VERTICAL"))
                        {
                            // Create vertical parts with bottom/top adjustments
                            ChildPart vb = new ChildPart("Vertical Body", "VB", 0.0, 0.0, "Aluminum");
                            vb.Clips = true; // Enable clips for this part
                            childParts.Add(vb);

                            // Bottom-adjusted part
                            childParts.Add(new ChildPart("Pressure Plate", "PP", -0.0625, 0.0, "Aluminum"));

                            // Top-adjusted part
                            childParts.Add(new ChildPart("Snap Cover", "SC", 0.0, -0.125, "Aluminum"));
                        }
                    }

                    // Store child parts as JSON in additional Xdata
                    string partsJson = Newtonsoft.Json.JsonConvert.SerializeObject(childParts);

                    // Add info Xdata
                    const int maxChunkSize = 250;
                    int numChunks = (int)Math.Ceiling((double)partsJson.Length / maxChunkSize);

                    ResultBuffer rbInfo = new ResultBuffer(
                        new TypedValue((int)DxfCode.ExtendedDataRegAppName, "METALPARTSINFO"),
                        new TypedValue((int)DxfCode.ExtendedDataInteger32, numChunks)
                    );
                    pline.XData = rbInfo;

                    // Add chunk Xdata
                    for (int i = 0; i < numChunks; i++)
                    {
                        int startIndex = i * maxChunkSize;
                        int length = Math.Min(maxChunkSize, partsJson.Length - startIndex);
                        string chunk = partsJson.Substring(startIndex, length);

                        ResultBuffer rbChunk = new ResultBuffer(
                            new TypedValue((int)DxfCode.ExtendedDataRegAppName, $"METALPARTS{i}"),
                            new TypedValue((int)DxfCode.ExtendedDataAsciiString, chunk)
                        );
                        pline.XData = rbChunk;
                    }

                    tr.Commit();

                    using (Transaction trx = doc.Database.TransactionManager.StartTransaction())
                    {
                        MarkNumberManager.Instance.ProcessComponentMarkNumbers(plineId, trx, forceProcess: true);
                        trx.Commit();
                    }

                    // Check if the object was created successfully
                    ed.WriteMessage($"\nMetal component created with handle: {pline.Handle}");
                    ed.WriteMessage($"\nLocation: ({pline.GetPoint2dAt(0).X}, {pline.GetPoint2dAt(0).Y}) to ({pline.GetPoint2dAt(1).X}, {pline.GetPoint2dAt(1).Y})");
                    ed.WriteMessage($"\nComponent has {childParts.Count} child parts");

                    // Clear the pending data after use
                    PendingParentData = null;
                    PendingParts = null;
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError: {ex.Message}");
                ed.WriteMessage($"\nStack trace: {ex.StackTrace}");
            }
        }

        // Add this to your MetalComponentCommands class:
        public void CreateMetalPartWithValues(string componentType, string floor, string elevation, List<ChildPart> customParts = null)
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            try
            {
                // Prompt for start point (we still need user input for this)
                PromptPointOptions pPtOpts = new PromptPointOptions("\nEnter start point: ");
                PromptPointResult pPtRes = ed.GetPoint(pPtOpts);
                if (pPtRes.Status != PromptStatus.OK) return;
                Point3d startPt = pPtRes.Value;

                // Prompt for end point
                pPtOpts = new PromptPointOptions("\nEnter end point: ");
                pPtOpts.UseBasePoint = true;
                pPtOpts.BasePoint = startPt;
                pPtRes = ed.GetPoint(pPtOpts);
                if (pPtRes.Status != PromptStatus.OK) return;
                Point3d endPt = pPtRes.Value;

                // Register all required application names first
                using (Transaction trReg = db.TransactionManager.StartTransaction())
                {
                    RegAppTable regTable = (RegAppTable)trReg.GetObject(db.RegAppTableId, OpenMode.ForRead);

                    // Register basic apps
                    RegisterApp(regTable, "METALCOMP", trReg);
                    RegisterApp(regTable, "METALPARTSINFO", trReg);

                    // Register chunk apps preemptively
                    for (int i = 0; i < 10; i++) // Assume we won't need more than 10 chunks
                    {
                        RegisterApp(regTable, $"METALPARTS{i}", trReg);
                    }

                    trReg.Commit();
                }

                LayerTable lt = (LayerTable)db.LayerTableId.GetObject(OpenMode.ForRead);
                ObjectId tagLayerId = ObjectId.Null;

                if (lt.Has("TAG"))
                {
                    // If TAG layer exists, set the polyline's layer to TAG
                    tagLayerId = lt["TAG"];
                }

                // Now create everything in a single transaction
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                    // Create polyline
                    Polyline pline = new Polyline();
                    pline.SetDatabaseDefaults();
                    pline.ColorIndex = 1; // Red color to make it visible
                    pline.ConstantWidth = 0.1; // Give it some width to be more visible
                    pline.AddVertexAt(0, new Point2d(startPt.X, startPt.Y), 0, 0, 0);
                    pline.AddVertexAt(1, new Point2d(endPt.X, endPt.Y), 0, 0, 0);

                    // Add polyline to the database
                    ObjectId plineId = btr.AppendEntity(pline);
                    tr.AddNewlyCreatedDBObject(pline, true);

                    if (tagLayerId != ObjectId.Null)
                    {
                        // If TAG layer exists, set the polyline's layer to TAG
                        pline.LayerId = tagLayerId;
                    }

                    // Add component Xdata
                    ResultBuffer rbComp = new ResultBuffer(
                        new TypedValue((int)DxfCode.ExtendedDataRegAppName, "METALCOMP"),
                        new TypedValue((int)DxfCode.ExtendedDataAsciiString, componentType),
                        new TypedValue((int)DxfCode.ExtendedDataAsciiString, floor),
                        new TypedValue((int)DxfCode.ExtendedDataAsciiString, elevation)
                    );
                    pline.XData = rbComp;

                    // Use custom parts if provided, otherwise create default parts
                    List<ChildPart> childParts = customParts;

                    if (childParts == null || childParts.Count == 0)
                    {
                        childParts = new List<ChildPart>();
                        if (componentType.ToUpper().Contains("HORIZONTAL"))
                        {
                            // Create parts with specific start and end adjustments
                            childParts.Add(new ChildPart("Horizontal Body", "HB", 0.0, 0.0, "Aluminum"));
                            childParts.Add(new ChildPart("Flat Filler", "FF", -0.03125, 0.0, "Aluminum"));
                            childParts.Add(new ChildPart("Face Cap", "FC", 0.0, 0.0, "Aluminum"));

                            // Left side attachments
                            ChildPart sbLeft = new ChildPart("Shear Block Left", "SBL", -1.25, 0.0, "Aluminum");
                            sbLeft.Attach = "L"; // Set attachment side
                            childParts.Add(sbLeft);

                            // Right side attachments
                            ChildPart sbRight = new ChildPart("Shear Block Right", "SBR", 0.0, -1.25, "Aluminum");
                            sbRight.Attach = "R"; // Set attachment side
                            childParts.Add(sbRight);
                        }
                        else if (componentType.ToUpper().Contains("VERTICAL"))
                        {
                            // Create vertical parts with bottom/top adjustments
                            ChildPart vb = new ChildPart("Vertical Body", "VB", 0.0, 0.0, "Aluminum");
                            vb.Clips = true; // Enable clips for this part
                            childParts.Add(vb);

                            // Bottom-adjusted part
                            childParts.Add(new ChildPart("Pressure Plate", "PP", -0.0625, 0.0, "Aluminum"));

                            // Top-adjusted part
                            childParts.Add(new ChildPart("Snap Cover", "SC", 0.0, -0.125, "Aluminum"));
                        }
                    }

                    // Store child parts as JSON in additional Xdata
                    string partsJson = Newtonsoft.Json.JsonConvert.SerializeObject(childParts);

                    // Add info Xdata
                    const int maxChunkSize = 250;
                    int numChunks = (int)Math.Ceiling((double)partsJson.Length / maxChunkSize);

                    ResultBuffer rbInfo = new ResultBuffer(
                        new TypedValue((int)DxfCode.ExtendedDataRegAppName, "METALPARTSINFO"),
                        new TypedValue((int)DxfCode.ExtendedDataInteger32, numChunks)
                    );
                    pline.XData = rbInfo;

                    // Add chunk Xdata
                    for (int i = 0; i < numChunks; i++)
                    {
                        int startIndex = i * maxChunkSize;
                        int length = Math.Min(maxChunkSize, partsJson.Length - startIndex);
                        string chunk = partsJson.Substring(startIndex, length);

                        ResultBuffer rbChunk = new ResultBuffer(
                            new TypedValue((int)DxfCode.ExtendedDataRegAppName, $"METALPARTS{i}"),
                            new TypedValue((int)DxfCode.ExtendedDataAsciiString, chunk)
                        );
                        pline.XData = rbChunk;
                    }

                    tr.Commit();

                    using (Transaction trx = doc.Database.TransactionManager.StartTransaction())
                    {
                        MarkNumberManager.Instance.ProcessComponentMarkNumbers(plineId, trx, forceProcess: true);
                        trx.Commit();
                    }

                    // Check if the object was created successfully
                    ed.WriteMessage($"\nMetal component created with handle: {pline.Handle}");
                    ed.WriteMessage($"\nLocation: ({pline.GetPoint2dAt(0).X}, {pline.GetPoint2dAt(0).Y}) to ({pline.GetPoint2dAt(1).X}, {pline.GetPoint2dAt(1).Y})");
                    ed.WriteMessage($"\nComponent has {childParts.Count} child parts");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError: {ex.Message}");
                ed.WriteMessage($"\nStack trace: {ex.StackTrace}");
            }
        }

        // Helper method to register an application
        private void RegisterApp(RegAppTable regTable, string appName, Transaction tr)
        {
            if (!regTable.Has(appName))
            {
                regTable.UpgradeOpen();
                RegAppTableRecord record = new RegAppTableRecord();
                record.Name = appName;
                regTable.Add(record);
                tr.AddNewlyCreatedDBObject(record, true);
            }
        }

        [CommandMethod("UPDATEMETAL")]
        public void UpdateMetalCommand()
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            try
            {
                // Check if we have pending data
                if (string.IsNullOrEmpty(PendingHandle) || PendingParts == null)
                {
                    ed.WriteMessage("\nNo pending metal part updates.");
                    return;
                }

                // Store the handle before it gets cleared
                string handle = PendingHandle;

                // Process the update
                UpdateMetalXdata(PendingHandle, PendingParts);

                // Clear the pending data
                PendingHandle = null;
                PendingParts = null;

                // Now process mark numbers after the update
                try
                {
                    using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
                    {
                        // Convert handle to ObjectId
                        long longHandle = Convert.ToInt64(handle, 16);
                        Handle h = new Handle(longHandle);
                        ObjectId objId = doc.Database.GetObjectId(false, h, 0);

                        // Process mark numbers
                        MarkNumberManager.Instance.ProcessComponentMarkNumbers(objId, tr, forceProcess: true);

                        tr.Commit();
                    }

                    ed.WriteMessage("\nMark numbers processed successfully.");
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nError processing mark numbers: {ex.Message}");
                }

                ed.WriteMessage("\nMetal part updated successfully.");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in UPDATEMETAL: {ex.Message}");
            }
        }


        // Method to update the Xdata
        private void UpdateMetalXdata(string handle, List<ChildPart> parts)
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                try
                {
                    // Convert handle to ObjectId
                    long longHandle = Convert.ToInt64(handle, 16);
                    Handle h = new Handle(longHandle);
                    ObjectId objId = db.GetObjectId(false, h, 0);

                    // Serialize parts to JSON
                    string jsonData = Newtonsoft.Json.JsonConvert.SerializeObject(parts);

                    // Register all required applications
                    RegAppTable regTable = (RegAppTable)tr.GetObject(db.RegAppTableId, OpenMode.ForRead);

                    // Register METALPARTSINFO
                    RegisterApp(regTable, "METALPARTSINFO", tr);

                    // Calculate chunks and register all chunk apps
                    const int maxChunkSize = 250;
                    int numChunks = (int)Math.Ceiling((double)jsonData.Length / maxChunkSize);

                    for (int i = 0; i < numChunks; i++)
                    {
                        RegisterApp(regTable, $"METALPARTS{i}", tr);
                    }

                    // Get the entity
                    Entity ent = tr.GetObject(objId, OpenMode.ForWrite) as Entity;

                    // Add METALPARTSINFO xdata
                    ResultBuffer rbInfo = new ResultBuffer(
                        new TypedValue((int)DxfCode.ExtendedDataRegAppName, "METALPARTSINFO"),
                        new TypedValue((int)DxfCode.ExtendedDataInteger32, numChunks)
                    );
                    ent.XData = rbInfo;

                    // Add each chunk
                    for (int i = 0; i < numChunks; i++)
                    {
                        int startIndex = i * maxChunkSize;
                        int length = Math.Min(maxChunkSize, jsonData.Length - startIndex);
                        string chunk = jsonData.Substring(startIndex, length);

                        ResultBuffer rbChunk = new ResultBuffer(
                            new TypedValue((int)DxfCode.ExtendedDataRegAppName, $"METALPARTS{i}"),
                            new TypedValue((int)DxfCode.ExtendedDataAsciiString, chunk)
                        );
                        ent.XData = rbChunk;
                    }

                    tr.Commit();
                    using (Transaction trx = doc.Database.TransactionManager.StartTransaction())
                    {
                        MarkNumberManager.Instance.ProcessComponentMarkNumbers(objId, trx);
                        trx.Commit();
                    }
                }
                catch
                {
                    tr.Abort();
                    throw;
                }
            }
        }


        [CommandMethod("DETECTATTACHMENTS")]
        public void DetectAttachments()
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            ed.WriteMessage("\nDetecting part attachments with end-specific adjustments...");

            // Declare attachments list outside the using block so it stays in scope
            List<Attachment> attachments = new List<Attachment>();

            try
            {
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    // Get all components using the central manager
                    List<DrawingComponentManager.Component> components =
                        DrawingComponentManager.Instance.GetAllComponents(tr);

                    // Group components by elevation
                    var componentsByElevation = components.GroupBy(c => c.Elevation);

                    // Process each elevation group
                    foreach (var elevGroup in componentsByElevation)
                    {
                        ed.WriteMessage($"\nProcessing elevation {elevGroup.Key}");

                        // Get verticals with parts that allow clips
                        var verticals = elevGroup.Where(c => c.Type.ToUpper() == "VERTICAL" &&
                                                          c.Parts.Any(p => p.Clips)).ToList();

                        // Get horizontals with parts that attach
                        var horizontals = elevGroup.Where(c => c.Type.ToUpper() == "HORIZONTAL" &&
                                                           c.Parts.Any(p => !string.IsNullOrEmpty(p.Attach))).ToList();

                        ed.WriteMessage($"\n  Found {verticals.Count} verticals and {horizontals.Count} horizontals");

                        // Process each vertical
                        foreach (var vertical in verticals)
                        {
                            // Calculate vertical bottom and top points
                            Point3d verticalBottom, verticalTop;
                            if (vertical.StartPoint.Y < vertical.EndPoint.Y)
                            {
                                verticalBottom = vertical.StartPoint;
                                verticalTop = vertical.EndPoint;
                            }
                            else
                            {
                                verticalBottom = vertical.EndPoint;
                                verticalTop = vertical.StartPoint;
                            }

                            // Find parts in vertical that allow clips
                            var clipParts = vertical.Parts.Where(p => p.Clips).ToList();

                            foreach (var clipPart in clipParts)
                            {
                                // Calculate adjusted vertical points for this specific part
                                Vector3d direction = verticalTop - verticalBottom;
                                double length = direction.Length;
                                if (length > 0)
                                {
                                    direction = direction / length;
                                }

                                Point3d adjustedBottom = verticalBottom - (direction * clipPart.StartAdjustment);
                                Point3d adjustedTop = verticalTop + (direction * clipPart.EndAdjustment);

                                // Find horizontals that might intersect
                                foreach (var horizontal in horizontals)
                                {
                                    // Find parts in horizontal that have attachments
                                    var attachingParts = horizontal.Parts.Where(p => !string.IsNullOrEmpty(p.Attach)).ToList();

                                    foreach (var hPart in attachingParts)
                                    {
                                        // Calculate adjusted horizontal points
                                        Point3d horizontalStart, horizontalEnd;
                                        if (horizontal.StartPoint.X < horizontal.EndPoint.X)
                                        {
                                            horizontalStart = horizontal.StartPoint;
                                            horizontalEnd = horizontal.EndPoint;
                                        }
                                        else
                                        {
                                            horizontalStart = horizontal.EndPoint;
                                            horizontalEnd = horizontal.StartPoint;
                                        }

                                        Vector3d hDirection = horizontalEnd - horizontalStart;
                                        double hLength = hDirection.Length;
                                        if (hLength > 0)
                                        {
                                            hDirection = hDirection / hLength;
                                        }

                                        Point3d adjustedStart = horizontalStart + (hDirection * hPart.StartAdjustment);
                                        Point3d adjustedEnd = horizontalEnd - (hDirection * hPart.EndAdjustment);

                                        // Check for intersection
                                        if (LinesIntersect(adjustedBottom, adjustedTop,
                                                          adjustedStart, adjustedEnd,
                                                          out Point3d intersectionPt,
                                                          proximityThreshold: 6.0))
                                        {
                                            double height = intersectionPt.Y - adjustedBottom.Y + hPart.Adjust;
                                            string side = DetermineSide(adjustedStart, adjustedEnd,
                                                                     adjustedBottom, adjustedTop);

                                            // If part attaches on the determined side
                                            if (hPart.Attach == side)
                                            {
                                                // Calculate position along horizontal
                                                double position = CalculatePosition(horizontal.StartPoint, horizontal.EndPoint, intersectionPt);

                                                Attachment attachment = new Attachment
                                                {
                                                    HorizontalHandle = horizontal.Handle,
                                                    VerticalHandle = vertical.Handle,
                                                    HorizontalPartType = hPart.PartType,
                                                    VerticalPartType = clipPart.PartType,
                                                    Side = side,
                                                    Position = position,
                                                    Height = height,
                                                    Invert = hPart.Invert,
                                                    Adjust = hPart.Adjust
                                                };

                                                attachments.Add(attachment);

                                                ed.WriteMessage($"\n    Attachment: {hPart.PartType} to {clipPart.PartType} on {side} side at height {height:F3}");
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }

                    // Save attachments to the drawing using the central manager
                    DrawingComponentManager.SaveAttachmentsToDrawing(attachments, tr);

                    tr.Commit();

                }

                // Now attachments is still in scope here
                ed.WriteMessage($"\nDetected and saved {attachments.Count} attachments with end-specific adjustments");

                // Update mark numbers
                ed.WriteMessage("\nUpdating mark numbers based on new attachments...");
                //MarkNumberDisplay.Instance.UpdateAllMarkNumbers();
                MarkNumberManager.GenerateMarkNumbers(); // Call the fixed method directly

                ed.WriteMessage("\nMark numbers updated successfully.");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in GenerateMarkNumbers: {ex.Message}");
                ed.WriteMessage($"\n{ex.StackTrace}");
            }
        }

        // Helper classes
        private class MetalComponent
        {
            public string Handle { get; set; }
            public string Type { get; set; }
            public string Elevation { get; set; }
            public string Floor { get; set; }
            public Point3d StartPoint { get; set; }
            public Point3d EndPoint { get; set; }
            public List<ChildPart> Parts { get; set; }
        }

       

        private bool LinesIntersect(Point3d line1Start, Point3d line1End,
                           Point3d line2Start, Point3d line2End,
                           out Point3d intersectionPoint,
                           double proximityThreshold = 6.0)
        {
            // First try exact intersection
            if (ExactIntersection(line1Start, line1End, line2Start, line2End, out intersectionPoint))
            {
                return true;
            }



            // Check if vertical line is near horizontal line's endpoints
            // Get bounding box of vertical line
            double minY = Math.Min(line1Start.Y, line1End.Y);
            double maxY = Math.Max(line1Start.Y, line1End.Y);
            double x1 = line1Start.X; // Assuming vertical line has nearly constant X

            // Get bounding box of horizontal line
            double minX = Math.Min(line2Start.X, line2End.X);
            double maxX = Math.Max(line2Start.X, line2End.X);
            double y2 = line2Start.Y; // Assuming horizontal line has nearly constant Y

            // Check if horizontal's Y is within vertical's height range
            if (minY <= y2 && y2 <= maxY)
            {
                // Check if left end of horizontal is close to vertical
                if (Math.Abs(minX - x1) <= proximityThreshold)
                {
                    intersectionPoint = new Point3d(x1, y2, 0);
                    return true;
                }

                // Check if right end of horizontal is close to vertical
                if (Math.Abs(maxX - x1) <= proximityThreshold)
                {
                    intersectionPoint = new Point3d(x1, y2, 0);
                    return true;
                }
            }

            // No intersection found
            return false;
        }

        // Helper to calculate distance from point to line segment
        private double PointToLineDistance(Point3d point, Point3d lineStart, Point3d lineEnd)
        {
            // Create vectors
            Vector3d v = lineEnd - lineStart;
            Vector3d w = point - lineStart;

            // Calculate projection ratio
            double c1 = w.DotProduct(v);
            double c2 = v.DotProduct(v);

            // Calculate the parameter of the closest point
            double b = (c2 > 0) ? c1 / c2 : 0;

            // Clamp to line segment
            b = Math.Max(0, Math.Min(1, b));

            // Get the closest point on the line
            Point3d closestPoint = lineStart + b * v;

            // Return the distance
            return point.DistanceTo(closestPoint);
        }

        // Method for exact intersection calculation
        private bool ExactIntersection(Point3d line1Start, Point3d line1End,
                                      Point3d line2Start, Point3d line2End,
                                      out Point3d intersectionPoint)
        {
            // Simple 2D line intersection algorithm
            double x1 = line1Start.X;
            double y1 = line1Start.Y;
            double x2 = line1End.X;
            double y2 = line1End.Y;

            double x3 = line2Start.X;
            double y3 = line2Start.Y;
            double x4 = line2End.X;
            double y4 = line2End.Y;

            double denominator = ((y4 - y3) * (x2 - x1) - (x4 - x3) * (y2 - y1));

            if (Math.Abs(denominator) < 0.0001)
            {
                // Lines are parallel
                intersectionPoint = new Point3d();
                return false;
            }

            double ua = ((x4 - x3) * (y1 - y3) - (y4 - y3) * (x1 - x3)) / denominator;
            double ub = ((x2 - x1) * (y1 - y3) - (y2 - y1) * (x1 - x3)) / denominator;

            // If ua and ub are between 0 and 1, lines intersect
            if (ua >= 0 && ua <= 1 && ub >= 0 && ub <= 1)
            {
                // Calculate intersection point
                intersectionPoint = new Point3d(
                    x1 + ua * (x2 - x1),
                    y1 + ua * (y2 - y1),
                    0
                );
                return true;
            }

            intersectionPoint = new Point3d();
            return false;
        }

        private string DetermineSide(Point3d hStart, Point3d hEnd, Point3d vStart, Point3d vEnd)
        {
            // Ensure horizontal vector always goes left to right
            Point3d hLeft, hRight;
            if (hStart.X <= hEnd.X)
            {
                hLeft = hStart;
                hRight = hEnd;
            }
            else
            {
                hLeft = hEnd;
                hRight = hStart;
            }

            // Get the X coordinate of the vertical
            double vX = (vStart.X + vEnd.X) / 2.0;  // Midpoint X

            // Debug output
            Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(
                $"\nHorizontal X range: {hLeft.X:F3} to {hRight.X:F3}, Vertical X: {vX:F3}");

            // UPDATED LOGIC: Now returning the side from the vertical's perspective

            // If vertical's X is less than horizontal's left X, horizontal is on vertical's right
            if (vX < hLeft.X)
            {
                return "R";  // Changed from "L" to "R"
            }

            // If vertical's X is greater than horizontal's right X, horizontal is on vertical's left
            if (vX > hRight.X)
            {
                return "L";  // Changed from "R" to "L"
            }

            // If vertical's X is between horizontal's endpoints, determine based on which end it's closer to
            double distToLeft = Math.Abs(vX - hLeft.X);
            double distToRight = Math.Abs(vX - hRight.X);

            // Return the side from vertical's perspective
            return distToLeft < distToRight ? "R" : "L";  // Swapped "L" and "R"
        }

        private double CalculatePosition(Point3d start, Point3d end, Point3d point)
        {
            // Calculate distance from start to intersection point
            double totalLength = start.DistanceTo(end);
            double distanceFromStart = start.DistanceTo(point);

            // Return position as percentage along the line
            return distanceFromStart / totalLength;
        }

        private void SaveAttachmentsToDrawing(List<Attachment> attachments)
        {
            // Serialize attachments to JSON
            string json = JsonConvert.SerializeObject(attachments);
            System.Diagnostics.Debug.WriteLine($"Saving {attachments.Count} attachments to drawing");

            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                // Get named objects dictionary
                DBDictionary nod = (DBDictionary)tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForWrite);

                // Create or update dictionary entry
                const string dictName = "METALATTACHMENTS";

                if (nod.Contains(dictName))
                {
                    // Update existing
                    DBObject obj = tr.GetObject(nod.GetAt(dictName), OpenMode.ForWrite);
                    if (obj is Xrecord)
                    {
                        Xrecord xrec = obj as Xrecord;
                        ResultBuffer rb = new ResultBuffer(new TypedValue((int)DxfCode.Text, json));
                        xrec.Data = rb;
                        System.Diagnostics.Debug.WriteLine("Updated existing METALATTACHMENTS record");
                    }
                }
                else
                {
                    // Create new
                    Xrecord xrec = new Xrecord();
                    ResultBuffer rb = new ResultBuffer(new TypedValue((int)DxfCode.Text, json));
                    xrec.Data = rb;

                    nod.SetAt(dictName, xrec);
                    tr.AddNewlyCreatedDBObject(xrec, true);
                    System.Diagnostics.Debug.WriteLine("Created new METALATTACHMENTS record");
                }

                tr.Commit();
            }
        }

        // Method to get all parts data from a polyline
        public static string GetPartsJsonFromEntity(Polyline pline)
        {
            // First check if we have the chunk count info
            ResultBuffer rbInfo = pline.GetXDataForApplication("METALPARTSINFO");
            if (rbInfo == null) return "";

            int chunkCount = 0;
            TypedValue[] xdataInfo = rbInfo.AsArray();
            if (xdataInfo.Length >= 2)
            {
                chunkCount = (int)xdataInfo[1].Value;
            }

            // Read each chunk using its specific RegApp name
            string partsJson = "";
            for (int i = 0; i < chunkCount; i++)
            {
                string chunkAppName = $"METALPARTS{i}";
                ResultBuffer rbChunk = pline.GetXDataForApplication(chunkAppName);

                if (rbChunk != null)
                {
                    TypedValue[] xdataChunk = rbChunk.AsArray();
                    if (xdataChunk.Length >= 2)
                    {
                        partsJson += xdataChunk[1].Value.ToString();
                    }
                }
            }

            return partsJson;
        }



        [CommandMethod("LISTATTACHMENTS")]
        public void ListAttachments()
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            try
            {
                // Use the central manager to get attachments
                List<Attachment> attachments =
                    DrawingComponentManager.Instance.GetAllAttachments();

                if (attachments.Count == 0)
                {
                    ed.WriteMessage("\nNo attachments found in drawing. Run DETECTATTACHMENTS first.");
                    return;
                }

                ed.WriteMessage($"\nFound {attachments.Count} attachments:");

                // Group by vertical handle
                var groupedAttachments = attachments.GroupBy(a => a.VerticalHandle);

                foreach (var group in groupedAttachments)
                {
                    ed.WriteMessage($"\n\nVertical {group.Key}:");

                    // Group by side (left/right)
                    var bySide = group.GroupBy(a => a.Side);

                    foreach (var sideGroup in bySide)
                    {
                        ed.WriteMessage($"\n  {sideGroup.Key} side attachments:");

                        foreach (var attachment in sideGroup)
                        {
                            ed.WriteMessage($"\n    Horizontal {attachment.HorizontalHandle}: {attachment.HorizontalPartType} to {attachment.VerticalPartType}");
                            ed.WriteMessage($"\n      Position: {attachment.Position:F3}, Height: {attachment.Height:F3}, Invert: {attachment.Invert}, Adjust: {attachment.Adjust:F3}");
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in LISTATTACHMENTS: {ex.Message}");
            }
        }




        [CommandMethod("METALEDITOR", CommandFlags.Modal)]
        public void ShowMetalEditor()
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            try
            {
                ed.WriteMessage("\nStarting Enhanced Metal Editor...");

                // Check if the palette set already exists
                if (_paletteSet == null)
                {
                    ed.WriteMessage("\nCreating palette set...");

                    // Create new palette set with a unique GUID
                    _paletteSet = new PaletteSet("Enhanced Metal Component Editor",
                        new Guid("1D8F97D4-C5E8-4A59-B1F3-53431D97C9A2"));

                    ed.WriteMessage("\nCreating standard editing panel...");

                    // Create the enhanced panel explicitly first
                    EnhancedMetalComponentPanel editPanel = new EnhancedMetalComponentPanel(false);

                    // Add editing panel to palette set
                    _paletteSet.Add("Edit Parts", editPanel);

                    ed.WriteMessage("\nCreating new component panel...");

                    // Create the panel for new component creation
                    EnhancedMetalComponentPanel newPanel = new EnhancedMetalComponentPanel(true);

                    // Add new component panel to palette set
                    _paletteSet.Add("Create Parts", newPanel);

                    ed.WriteMessage("\nCreating elevations panel...");
                    // Create the elevations panel
                    ElevationPanel elevationsPanel = new ElevationPanel();
                    // Add elevations panel to palette set
                    _paletteSet.Add("Elevations", elevationsPanel);

                    ed.WriteMessage("\nPanels added successfully.");
                }

                // Show the palette set
                _paletteSet.Visible = true;
                ed.WriteMessage("\nPalette set shown successfully.");

                // Start the stretch monitor
                ed.WriteMessage("\nStarting Stretch Monitor...");
                doc.SendStringToExecute("STRETCHMONITOR ", true, false, false);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in ShowMetalEditor: {ex.Message}");
                ed.WriteMessage($"\nStack trace: {ex.StackTrace}");
            }
        }


        [CommandMethod("COPYTOMULTIPLECOMPONENTS", CommandFlags.Modal)]
        public void CopyToMultipleComponents()
        {
            // Call the implementation
            CopyProperty();
        }


        [CommandMethod("COPYPROPERTY")]
        public void CopyProperty()
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            // Prompt for source component
            PromptEntityOptions peoSource = new PromptEntityOptions("\nSelect source component: ");
            peoSource.SetRejectMessage("\nOnly polylines with metal component data can be selected.");
            peoSource.AddAllowedClass(typeof(Polyline), false);

            PromptEntityResult perSource = ed.GetEntity(peoSource);
            if (perSource.Status != PromptStatus.OK) return;

            ObjectId sourceId = perSource.ObjectId;

            // Get source data
            string componentType = "", floor = "", elevation = "";
            List<ChildPart> sourceParts = new List<ChildPart>();

            using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
            {
                Entity ent = tr.GetObject(sourceId, OpenMode.ForRead) as Entity;
                if (ent is Polyline)
                {
                    // Get component properties
                    ResultBuffer rbComp = ent.GetXDataForApplication("METALCOMP");
                    if (rbComp != null)
                    {
                        TypedValue[] xdataComp = rbComp.AsArray();
                        for (int i = 1; i < xdataComp.Length; i++)
                        {
                            if (i == 1) componentType = xdataComp[i].Value.ToString();
                            if (i == 2) floor = xdataComp[i].Value.ToString();
                            if (i == 3) elevation = xdataComp[i].Value.ToString();
                        }
                    }

                    // Get parts
                    string partsJson = GetPartsJsonFromEntity(ent as Polyline);
                    if (!string.IsNullOrEmpty(partsJson))
                    {
                        try
                        {
                            sourceParts = Newtonsoft.Json.JsonConvert.DeserializeObject<List<ChildPart>>(partsJson);
                        }
                        catch { /* Handle error */ }
                    }
                }

                tr.Commit();
            }

            if (sourceParts.Count == 0)
            {
                ed.WriteMessage("\nNo parts found in source component.");
                return;
            }

            // Ask what user wants to copy with a more detailed dialog
            using (CopyPropertyDialog dialog = new CopyPropertyDialog(componentType, floor, elevation, sourceParts))
            {
                if (dialog.ShowDialog() == DialogResult.OK)
                {
                    // Get selected components
                    PromptSelectionOptions pso = new PromptSelectionOptions();
                    pso.MessageForAdding = "\nSelect target components to receive properties: ";

                    PromptSelectionResult psr = ed.GetSelection(pso);
                    if (psr.Status != PromptStatus.OK) return;

                    SelectionSet ss = psr.Value;

                    // Store for command-based update
                    List<string> targetHandles = new List<string>();

                    // Create a transaction to access the objects
                    using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
                    {
                        foreach (SelectedObject so in ss)
                        {
                            if (so.ObjectId == sourceId) continue; // Skip source

                            Entity ent = tr.GetObject(so.ObjectId, OpenMode.ForRead) as Entity;
                            if (!(ent is Polyline)) continue;

                            targetHandles.Add(ent.Handle.ToString());
                        }

                        tr.Commit();
                    }

                    // Use command-based update for all targets
                    if (targetHandles.Count > 0)
                    {
                        // Store data for command
                        PendingCopyData = new CopyPropertyData
                        {
                            SourceHandle = sourceId.Handle.ToString(),
                            TargetHandles = targetHandles,
                            ComponentType = dialog.CopyComponentType ? componentType : null,
                            Floor = dialog.CopyFloor ? floor : null,
                            Elevation = dialog.CopyElevation ? elevation : null,
                            Parts = dialog.CopyParts ? sourceParts : null
                        };

                        // Execute copy command
                        doc.SendStringToExecute("EXECUTECOPY ", true, false, false);

                        ed.WriteMessage($"\nProperties copied to {targetHandles.Count} components.");
                    }
                    else
                    {
                        ed.WriteMessage("\nNo valid target components selected.");
                    }
                }
            }
        }

        [CommandMethod("CREATEOPENINGCOMPONENTS")]
        public void CreateOpeningComponents()
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            try
            {
                // Debug data
                System.Diagnostics.Debug.WriteLine("CREATEOPENINGCOMPONENTS command starting");

                // Check if we have pending data from the panel
                string floor = "01";
                string elevation = "A";

                // Component types for each side
                string topType = "Head";
                string bottomType = "Sill";
                string leftType = "JambL";
                string rightType = "JambR";

                // Which sides to create
                bool createTop = true;
                bool createBottom = true;
                bool createLeft = true;
                bool createRight = true;

                // Child parts for each side - This is the key addition
                List<ChildPart> topParts = null;
                List<ChildPart> bottomParts = null;
                List<ChildPart> leftParts = null;
                List<ChildPart> rightParts = null;

                // Flag to control dialog display
                bool showDialog = true;

                // Check if we have pending opening data
                if (PendingOpeningData != null)
                {
                    System.Diagnostics.Debug.WriteLine("Found PendingOpeningData");

                    floor = PendingOpeningData.Floor ?? floor;
                    elevation = PendingOpeningData.Elevation ?? elevation;

                    topType = PendingOpeningData.TopType ?? topType;
                    bottomType = PendingOpeningData.BottomType ?? bottomType;
                    leftType = PendingOpeningData.LeftType ?? leftType;
                    rightType = PendingOpeningData.RightType ?? rightType;

                    createTop = PendingOpeningData.CreateTop;
                    createBottom = PendingOpeningData.CreateBottom;
                    createLeft = PendingOpeningData.CreateLeft;
                    createRight = PendingOpeningData.CreateRight;

                    // Always show dialog when ForceDialogToOpen is true
                    showDialog = PendingOpeningData.GetType().GetProperty("ForceDialogToOpen") != null ?
                                (bool)PendingOpeningData.GetType().GetProperty("ForceDialogToOpen").GetValue(PendingOpeningData, null) :
                                true;

                    // Get child parts from the panel if available
                    if (PendingParts != null && PendingParts.Count > 0)
                    {
                        // Clone the parts to avoid reference issues
                        List<ChildPart> allParts = new List<ChildPart>(PendingParts);

                        // We'll use the same parts for all sides, but in a real implementation
                        // you might want to filter or create specific parts for each side
                        topParts = allParts;
                        bottomParts = allParts;
                        leftParts = allParts;
                        rightParts = allParts;

                        System.Diagnostics.Debug.WriteLine($"Found {allParts.Count} pending parts to use for components");
                    }
                }

                // If we need to show the dialog, do it
                if (showDialog)
                {
                    System.Diagnostics.Debug.WriteLine("Showing ComponentPropertiesDialog");

                    using (ComponentPropertiesDialog dialog = new ComponentPropertiesDialog())
                    {
                        // Initialize dialog with pending data values
                        dialog.InitializeValues(
                            floor, elevation,
                            topType, bottomType, leftType, rightType,
                            createTop, createBottom, createLeft, createRight);

                        if (dialog.ShowDialog() != DialogResult.OK)
                        {
                            PendingOpeningData = null; // Clear pending data if dialog is canceled
                            return;
                        }

                        // Get values from dialog
                        floor = dialog.Floor;
                        elevation = dialog.Elevation;

                        topType = dialog.TopType;
                        bottomType = dialog.BottomType;
                        leftType = dialog.LeftType;
                        rightType = dialog.RightType;

                        createTop = dialog.CreateTop;
                        createBottom = dialog.CreateBottom;
                        createLeft = dialog.CreateLeft;
                        createRight = dialog.CreateRight;
                    }
                }

                // Debug output
                System.Diagnostics.Debug.WriteLine($"Using values: Floor={floor}, Elevation={elevation}");
                System.Diagnostics.Debug.WriteLine($"Types: Top={topType}, Bottom={bottomType}, Left={leftType}, Right={rightType}");
                System.Diagnostics.Debug.WriteLine($"Create: Top={createTop}, Bottom={createBottom}, Left={createLeft}, Right={createRight}");

                // Ensure TAG layer exists
                EnsureLayerExists("TAG");

                // IMPORTANT: Explicitly clear any running command or selection state
                doc.SendStringToExecute("'_.CANCEL ", true, false, true);

                // Wait a moment to ensure the command is canceled
                System.Threading.Thread.Sleep(100);

                // Get the center point using proper prompt
                PromptPointOptions ppo = new PromptPointOptions("\nSelect center of lite: ");
                ppo.AllowNone = false;

                Point3d centerPoint;
                while (true)
                {
                    PromptPointResult ppr = ed.GetPoint(ppo);
                    if (ppr.Status != PromptStatus.OK)
                    {
                        // User canceled - abort the command
                        ed.WriteMessage("\nCommand canceled.");
                        PendingOpeningData = null;
                        return;
                    }

                    centerPoint = ppr.Value;
                    ed.WriteMessage($"\nSelected center point: ({centerPoint.X}, {centerPoint.Y})");

                    // Use boundary detection to find the perimeter
                    List<Point3d> boundaryPoints = DetectBoundary(centerPoint);
                    if (boundaryPoints.Count < 4)
                    {
                        ed.WriteMessage("\nCould not detect boundary. Try again.");
                        continue;
                    }

                    // Create components for each side of the boundary
                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        // Sort points to ensure proper order (bottom-left, bottom-right, top-right, top-left)
                        boundaryPoints = SortPoints(boundaryPoints);

                        // Check which sides to create - Pass the child parts
                        if (createBottom)
                            CreateSideComponent(tr, boundaryPoints[0], boundaryPoints[1], bottomType, floor, elevation, bottomParts);

                        if (createRight)
                            CreateSideComponent(tr, boundaryPoints[1], boundaryPoints[2], rightType, floor, elevation, rightParts);

                        if (createTop)
                            CreateSideComponent(tr, boundaryPoints[2], boundaryPoints[3], topType, floor, elevation, topParts);

                        if (createLeft)
                            CreateSideComponent(tr, boundaryPoints[3], boundaryPoints[0], leftType, floor, elevation, leftParts);

                        tr.Commit();
                    }

                    ed.WriteMessage("\nComponents created successfully.");

                    // Ask if user wants to continue
                    PromptResult promptRes = ed.GetString("\nCreate more components? [Yes/No] <Yes>: ");
                    if (promptRes.StringResult.ToUpper() != "Y" &&
                        promptRes.StringResult.ToUpper() != "YES" &&
                        !string.IsNullOrEmpty(promptRes.StringResult))
                        break;
                }

                // Clear the pending data after use
                PendingOpeningData = null;
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError: {ex.Message}");
                ed.WriteMessage($"\nStack trace: {ex.StackTrace}");

                // Make sure we clear the pending data even on error
                PendingOpeningData = null;
            }
        }

        private List<Point3d> DetectBoundary(Point3d centerPoint)
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            List<Point3d> boundaryPoints = new List<Point3d>();

            try
            {
                // Save current layer for later restoration
                ObjectId originalLayerId = db.Clayer;
                string tempLayer = "TEMP_BOUNDARY";

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    // Ensure temp layer exists
                    EnsureLayerExists(tempLayer);

                    // Get the layer ID for the temp layer
                    LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                    if (lt.Has(tempLayer))
                    {
                        // Set active layer to the temp layer
                        db.Clayer = lt[tempLayer];

                        // IMPORTANT: Use a slightly different approach to create boundary
                        // First, make sure no other command or selection is active
                        doc.SendStringToExecute("'_.CANCEL ", true, false, true);

                        // Use the command-based approach, but with explicit command options
                        // We need to ensure AutoCAD's command sequence is not confused
                        ed.WriteMessage("\nCreating boundary...");

                        ed.Command(
                            "_boundary",
                            new Point3d(centerPoint.X, centerPoint.Y, 0),  // Specify center point directly
                            "_a",                     // All visible layers
                            "i",                      // Island options (i for islands)
                            "n",                      // No for island detection to only get Outer boundary
                            "",                       // Default Ray Cast
                            "",                       // Accept default boundary name
                            ""                        // End command
                        );

                        // To find the most recently created entity, we'll use a selection approach
                        TypedValue[] filterList = new TypedValue[]
                        {
                    new TypedValue((int)DxfCode.LayerName, tempLayer)
                        };

                        SelectionFilter filter = new SelectionFilter(filterList);
                        PromptSelectionResult selRes = ed.SelectAll(filter);

                        if (selRes.Status == PromptStatus.OK && selRes.Value.Count > 0)
                        {
                            // Get the most recently created entity - typically the last one in the selection set
                            ObjectId boundaryId = selRes.Value[selRes.Value.Count - 1].ObjectId;

                            // Extract vertices from the boundary
                            Polyline boundary = (Polyline)tr.GetObject(boundaryId, OpenMode.ForRead);

                            for (int i = 0; i < boundary.NumberOfVertices; i++)
                            {
                                boundaryPoints.Add(boundary.GetPoint3dAt(i));
                            }

                            // Erase the temporary boundary
                            boundary.UpgradeOpen();
                            boundary.Erase();
                        }
                        else
                        {
                            ed.WriteMessage("\nNo boundary created. Try a different point.");
                        }

                        // Restore original layer
                        db.Clayer = originalLayerId;
                    }

                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError detecting boundary: {ex.Message}");
                ed.WriteMessage($"\nStack trace: {ex.StackTrace}");
            }

            return boundaryPoints;
        }

        

        private List<Point3d> SortPoints(List<Point3d> points)
        {
            // Sort points to ensure they are in clockwise order starting from bottom-left
            // First, find centroid
            Point3d centroid = new Point3d(
                points.Average(p => p.X),
                points.Average(p => p.Y),
                points.Average(p => p.Z)
            );

            // Sort by angle from centroid
            return points.OrderBy(p => Math.Atan2(p.Y - centroid.Y, p.X - centroid.X)).ToList();
        }

        // Modify the CreateSideComponent method to accept custom parts
        private void CreateSideComponent(Transaction tr, Point3d startPoint, Point3d endPoint,
                                       string componentType, string floor, string elevation,
                                       List<ChildPart> customParts = null)
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            // Save current layer for later restoration
            ObjectId originalLayerId = db.Clayer;

            try
            {
                // Set layer to TAG
                LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                if (lt.Has("TAG"))
                {
                    // Get the TAG layer without changing current layer (avoids eLock violations)
                    ObjectId tagLayerId = lt["TAG"];

                    // Create our polyline
                    BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                    Polyline pline = new Polyline();
                    pline.SetDatabaseDefaults();
                    pline.ColorIndex = 1;
                    pline.ConstantWidth = 0.1;
                    pline.AddVertexAt(0, new Point2d(startPoint.X, startPoint.Y), 0, 0, 0);
                    pline.AddVertexAt(1, new Point2d(endPoint.X, endPoint.Y), 0, 0, 0);

                    // Set the layer directly on the polyline
                    pline.LayerId = tagLayerId;

                    // Add polyline to database
                    ObjectId plineId = btr.AppendEntity(pline);
                    tr.AddNewlyCreatedDBObject(pline, true);

                    // Add component data
                    ResultBuffer rbComp = new ResultBuffer(
                        new TypedValue((int)DxfCode.ExtendedDataRegAppName, "METALCOMP"),
                        new TypedValue((int)DxfCode.ExtendedDataAsciiString, componentType),
                        new TypedValue((int)DxfCode.ExtendedDataAsciiString, floor),
                        new TypedValue((int)DxfCode.ExtendedDataAsciiString, elevation)
                    );
                    pline.XData = rbComp;

                    // Use custom parts if provided, otherwise get default parts
                    List<ChildPart> parts = customParts ?? GetDefaultPartsForType(componentType);

                    // Register all required application names
                    RegAppTable regTable = (RegAppTable)tr.GetObject(db.RegAppTableId, OpenMode.ForRead);
                    RegisterApp(regTable, "METALCOMP", tr);
                    RegisterApp(regTable, "METALPARTSINFO", tr);

                    // Register chunk apps preemptively
                    for (int i = 0; i < 10; i++)
                    {
                        RegisterApp(regTable, $"METALPARTS{i}", tr);
                    }

                    // Store child parts as JSON in additional Xdata
                    string partsJson = Newtonsoft.Json.JsonConvert.SerializeObject(parts);

                    // Add info Xdata
                    const int maxChunkSize = 250;
                    int numChunks = (int)Math.Ceiling((double)partsJson.Length / maxChunkSize);

                    ResultBuffer rbInfo = new ResultBuffer(
                        new TypedValue((int)DxfCode.ExtendedDataRegAppName, "METALPARTSINFO"),
                        new TypedValue((int)DxfCode.ExtendedDataInteger32, numChunks)
                    );
                    pline.XData = rbInfo;

                    // Add chunk Xdata
                    for (int i = 0; i < numChunks; i++)
                    {
                        int startIndex = i * maxChunkSize;
                        int length = Math.Min(maxChunkSize, partsJson.Length - startIndex);
                        string chunk = partsJson.Substring(startIndex, length);

                        ResultBuffer rbChunk = new ResultBuffer(
                            new TypedValue((int)DxfCode.ExtendedDataRegAppName, $"METALPARTS{i}"),
                            new TypedValue((int)DxfCode.ExtendedDataAsciiString, chunk)
                        );
                        pline.XData = rbChunk;
                    }

                    // Process mark numbers
                    MarkNumberManager.Instance.ProcessComponentMarkNumbers(plineId, tr, forceProcess: true);
                }
            }
            catch
            {
                throw;
            }
        }

        private void EnsureLayerExists(string layerName)
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                // Open the Layer table
                LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);

                if (!lt.Has(layerName))
                {
                    // Create the layer if it doesn't exist
                    lt.UpgradeOpen();
                    LayerTableRecord ltr = new LayerTableRecord();
                    ltr.Name = layerName;
                    ltr.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByLayer, 1); // Red

                    lt.Add(ltr);
                    tr.AddNewlyCreatedDBObject(ltr, true);
                }

                tr.Commit();
            }
        }

        // Helper method to get the last created entity
        private ObjectId GetLastCreatedEntityId()
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                // Get the most recently added entity
                using (BlockTableRecordEnumerator enumerator = btr.GetEnumerator())
                {
                    ObjectId lastId = ObjectId.Null;

                    while (enumerator.MoveNext())
                    {
                        lastId = enumerator.Current;
                    }

                    return lastId;
                }
            }
        }

        private List<ChildPart> GetDefaultPartsForType(string componentType)
        {
            List<ChildPart> parts = new List<ChildPart>();

            // Add default parts based on component type
            if (componentType.ToUpper().Contains("HORIZONTAL") ||
                componentType.ToUpper().Contains("SILL") ||
                componentType.ToUpper().Contains("HEAD"))
            {
                // Add default horizontal parts
                parts.Add(new ChildPart("Horizontal Body", "HB", 0.0, 0.0, "Aluminum"));
                parts.Add(new ChildPart("Flat Filler", "FF", -0.03125, 0.0, "Aluminum"));
                parts.Add(new ChildPart("Face Cap", "FC", 0.0, 0.0, "Aluminum"));

                // For certain types, add specific parts
                if (componentType.ToUpper().Contains("SILL"))
                {
                    parts.Add(new ChildPart("Sill Flashing", "SF", 0.0, -0.5, "Aluminum"));
                }
                else if (componentType.ToUpper().Contains("HEAD"))
                {
                    parts.Add(new ChildPart("Head Flashing", "HF", -0.5, 0.0, "Aluminum"));
                }

                // Left side attachments
                ChildPart sbLeft = new ChildPart("Shear Block Left", "SBL", -1.25, 0.0, "Aluminum");
                sbLeft.Attach = "L";
                parts.Add(sbLeft);

                // Right side attachments
                ChildPart sbRight = new ChildPart("Shear Block Right", "SBR", 0.0, -1.25, "Aluminum");
                sbRight.Attach = "R";
                parts.Add(sbRight);
            }
            else if (componentType.ToUpper().Contains("VERTICAL") ||
                     componentType.ToUpper().Contains("JAMB"))
            {
                // Add default vertical parts
                ChildPart vb = new ChildPart("Vertical Body", "VB", 0.0, 0.0, "Aluminum");
                vb.Clips = true;
                parts.Add(vb);

                parts.Add(new ChildPart("Pressure Plate", "PP", -0.0625, 0.0, "Aluminum"));
                parts.Add(new ChildPart("Snap Cover", "SC", 0.0, -0.125, "Aluminum"));

                // For jambs, add specific parts
                if (componentType.ToUpper().Contains("JAMB"))
                {
                    ChildPart jf = new ChildPart("Jamb Filler", "JF", 0.0, 0.0, "Aluminum");
                    jf.Attach = componentType.ToUpper().Contains("JAMBL") ? "L" : "R";
                    parts.Add(jf);
                }
            }

            return parts;
        }







    }
    }