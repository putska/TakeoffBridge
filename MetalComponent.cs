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

// This file contains our custom MetalComponent entity definition

namespace TakeoffBridge
{
    // Define our child part class
    public class ChildPart
    {
        public string Name { get; set; }
        public string PartType { get; set; }

        // Original length adjustment (keeping for backward compatibility)
        public double LengthAdjustment { get; set; }

        // New end-specific adjustments
        public double StartAdjustment { get; set; } // Left for horizontal, Bottom for vertical
        public double EndAdjustment { get; set; }   // Right for horizontal, Top for vertical

        // New fixed length property
        public bool IsFixedLength { get; set; }
        public double FixedLength { get; set; }

        public string MarkNumber { get; set; }
        public string Material { get; set; }

        // Attachment properties
        public string Attach { get; set; } // "L", "R", or null/empty
        public bool Invert { get; set; }   // Whether the part should be inverted
        public double Adjust { get; set; } // Vertical adjustment (inches)
        public bool Clips { get; set; }    // For vertical parts - whether clips attach to this part

        // Default constructor for JSON deserialization
        public ChildPart()
        {
            // Initialize default values
            Name = "";
            PartType = "";
            LengthAdjustment = 0.0;
            StartAdjustment = 0.0;
            EndAdjustment = 0.0;
            IsFixedLength = false;
            FixedLength = 0.0;
            Material = "";
            MarkNumber = "";
            Attach = "";
            Invert = false;
            Adjust = 0.0;
            Clips = false;
        }

        public ChildPart(string name, string partType, double lengthAdjustment, string material)
        {
            Name = name;
            PartType = partType;
            LengthAdjustment = lengthAdjustment;
            Material = material;
            MarkNumber = ""; // Will be assigned later

            // Initialize new end-specific adjustments
            // Initially set them to half the total adjustment on each end
            StartAdjustment = lengthAdjustment / 2;
            EndAdjustment = lengthAdjustment / 2;

            // Default to non-fixed length
            IsFixedLength = false;
            FixedLength = 0.0;

            // Initialize attachment properties with defaults
            Attach = "";
            Invert = false;
            Adjust = 0.0;
            Clips = false;
        }

        // Additional constructor for direct end adjustments
        public ChildPart(string name, string partType, double startAdjustment, double endAdjustment, string material)
        {
            Name = name;
            PartType = partType;
            StartAdjustment = startAdjustment;
            EndAdjustment = endAdjustment;
            LengthAdjustment = startAdjustment + endAdjustment; // For backward compatibility
            Material = material;
            MarkNumber = "";

            // Default to non-fixed length
            IsFixedLength = false;
            FixedLength = 0.0;

            // Initialize attachment properties with defaults
            Attach = "";
            Invert = false;
            Adjust = 0.0;
            Clips = false;
        }

        // Constructor for fixed length parts
        public ChildPart(string name, string partType, double fixedLength, string material, bool isFixed)
        {
            if (!isFixed) throw new ArgumentException("This constructor should only be used for fixed length parts");

            Name = name;
            PartType = partType;
            FixedLength = fixedLength;
            IsFixedLength = true;
            Material = material;
            MarkNumber = "";

            // Set adjustments to 0 for fixed length parts
            StartAdjustment = 0;
            EndAdjustment = 0;
            LengthAdjustment = 0;

            // Initialize attachment properties with defaults
            Attach = "";
            Invert = false;
            Adjust = 0.0;
            Clips = false;
        }

        // Calculate the actual length based on parent component length
        public double GetActualLength(double parentLength)
        {
            // If fixed length, just return that value
            if (IsFixedLength)
            {
                return FixedLength;
            }

            // Otherwise use the sum of both adjustments to get the total length adjustment
            return parentLength + StartAdjustment + EndAdjustment;
        }

        // Calculate the actual start point offset
        public double GetStartOffset()
        {
            return StartAdjustment;
        }

        // Calculate the actual end point offset
        public double GetEndOffset()
        {
            return EndAdjustment;
        }
    }

    // Commands for working with metal components
    public class MetalComponentCommands
    {
        // Use the same naming throughout
        public static string PendingHandle = null; // Note: no underscore, and public access
        public static List<ChildPart> PendingParts = null; // Note: no underscore, and public access
        public static ParentComponentData PendingParentData;
        public static CopyPropertyData PendingCopyData;

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
                        doc.SendStringToExecute("UPDATEMETAL ", true, false, false);
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
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            try
            {
                // Prompt for component type
                PromptStringOptions pStrOpts = new PromptStringOptions("\nEnter component type (Horizontal/Vertical): ");
                pStrOpts.DefaultValue = "Horizontal";
                PromptResult pStrRes = ed.GetString(pStrOpts);
                if (pStrRes.Status != PromptStatus.OK) return;
                string componentType = pStrRes.StringResult;

                // Prompt for floor
                pStrOpts = new PromptStringOptions("\nEnter floor: ");
                pStrOpts.DefaultValue = "01";
                pStrRes = ed.GetString(pStrOpts);
                if (pStrRes.Status != PromptStatus.OK) return;
                string floor = pStrRes.StringResult;

                // Prompt for elevation
                pStrOpts = new PromptStringOptions("\nEnter elevation: ");
                pStrOpts.DefaultValue = "A";
                pStrRes = ed.GetString(pStrOpts);
                if (pStrRes.Status != PromptStatus.OK) return;
                string elevation = pStrRes.StringResult;

                // Prompt for start point
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

                    // Add component Xdata
                    ResultBuffer rbComp = new ResultBuffer(
                        new TypedValue((int)DxfCode.ExtendedDataRegAppName, "METALCOMP"),
                        new TypedValue((int)DxfCode.ExtendedDataAsciiString, componentType),
                        new TypedValue((int)DxfCode.ExtendedDataAsciiString, floor),
                        new TypedValue((int)DxfCode.ExtendedDataAsciiString, elevation)
                    );
                    pline.XData = rbComp;

                    // Create child parts with end-specific adjustments
                    List<ChildPart> childParts = new List<ChildPart>();
                    if (componentType.ToUpper() == "HORIZONTAL")
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
                    else if (componentType.ToUpper() == "VERTICAL")
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

                    // Check if the object was created successfully
                    ed.WriteMessage($"\nMetal component created with handle: {pline.Handle}");
                    ed.WriteMessage($"\nLocation: ({pline.GetPoint2dAt(0).X}, {pline.GetPoint2dAt(0).Y}) to ({pline.GetPoint2dAt(1).X}, {pline.GetPoint2dAt(1).Y})");
                    ed.WriteMessage($"\nComponent has {childParts.Count} child parts with end-specific adjustments");
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

                // Process the update
                UpdateMetalXdata(PendingHandle, PendingParts);

                // Clear the pending data
                PendingHandle = null;
                PendingParts = null;

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
                }
                catch
                {
                    tr.Abort();
                    throw;
                }
            }
        }


        private Point3d CalculateAdjustedPointForPart(Point3d startPoint, Point3d endPoint, ChildPart part, string orientation)
        {
            // For horizontal components
            if (orientation.ToUpper() == "HORIZONTAL")
            {
                // Get the direction vector
                Vector3d direction = endPoint - startPoint;
                // Alternative approach if GetNormalized() is not available:
                double length = direction.Length;
                if (length > 0)
                {
                     direction = direction / length;
                 }

                // Calculate the adjusted start point (accounting for StartAdjustment)
                Point3d adjustedStartPoint = startPoint + (direction * part.StartAdjustment);

                // Calculate the adjusted end point (accounting for EndAdjustment)
                Point3d adjustedEndPoint = endPoint - (direction * part.EndAdjustment);

                // Return the appropriate point based on the part's attachment side
                if (part.Attach == "L")
                {
                    return adjustedStartPoint;
                }
                else if (part.Attach == "R")
                {
                    return adjustedEndPoint;
                }
                else
                {
                    // If no specific side, return the midpoint
                    return new Point3d(
                        (adjustedStartPoint.X + adjustedEndPoint.X) / 2,
                        (adjustedStartPoint.Y + adjustedEndPoint.Y) / 2,
                        (adjustedStartPoint.Z + adjustedEndPoint.Z) / 2
                    );
                }
            }
            // For vertical components
            else if (orientation.ToUpper() == "VERTICAL")
            {
                // For vertical components, assume Y is up/down
                // Determine which point is bottom and which is top
                Point3d bottomPoint, topPoint;
                if (startPoint.Y < endPoint.Y)
                {
                    bottomPoint = startPoint;
                    topPoint = endPoint;
                }
                else
                {
                    bottomPoint = endPoint;
                    topPoint = startPoint;
                }

                // Get the direction vector
                Vector3d direction = topPoint - bottomPoint;
                double length = direction.Length;
                if (length > 0)
                {
                    direction = direction / length;
                }

                // Calculate the adjusted bottom point (accounting for StartAdjustment)
                Point3d adjustedBottomPoint = bottomPoint + (direction * part.StartAdjustment);

                // Calculate the adjusted top point (accounting for EndAdjustment)
                Point3d adjustedTopPoint = topPoint - (direction * part.EndAdjustment);

                // Return the appropriate point based on the clips property
                if (part.Clips)
                {
                    // For parts with clips, you may want to return the whole adjusted line segment
                    // But for now, let's return the midpoint as an example
                    return new Point3d(
                        (adjustedBottomPoint.X + adjustedTopPoint.X) / 2,
                        (adjustedBottomPoint.Y + adjustedTopPoint.Y) / 2,
                        (adjustedBottomPoint.Z + adjustedTopPoint.Z) / 2
                    );
                }
                else
                {
                    // For parts without clips, return the midpoint as default
                    return new Point3d(
                        (adjustedBottomPoint.X + adjustedTopPoint.X) / 2,
                        (adjustedBottomPoint.Y + adjustedTopPoint.Y) / 2,
                        (adjustedBottomPoint.Z + adjustedTopPoint.Z) / 2
                    );
                }
            }

            // Default case - return the midpoint of the original line
            return new Point3d(
                (startPoint.X + endPoint.X) / 2,
                (startPoint.Y + endPoint.Y) / 2,
                (startPoint.Z + endPoint.Z) / 2
            );
        }

        // Modified method for detecting attachments
        [CommandMethod("DETECTATTACHMENTS")]
        public void DetectAttachments()
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            ed.WriteMessage("\nDetecting part attachments with end-specific adjustments...");

            try
            {
                // Collect all metal components with their parts
                List<MetalComponent> components = new List<MetalComponent>();

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    // Get all entities in model space
                    BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                    foreach (ObjectId objId in btr)
                    {
                        Entity ent = tr.GetObject(objId, OpenMode.ForRead) as Entity;
                        if (ent is Polyline)
                        {
                            Polyline pline = ent as Polyline;

                            // Check if it has METALCOMP Xdata
                            ResultBuffer rbComp = ent.GetXDataForApplication("METALCOMP");
                            if (rbComp != null)
                            {
                                string componentType = "";
                                string floor = "";
                                string elevation = "";

                                // Extract basic component data
                                TypedValue[] xdataComp = rbComp.AsArray();
                                for (int i = 1; i < xdataComp.Length; i++) // Skip app name
                                {
                                    if (i == 1) componentType = xdataComp[i].Value.ToString();
                                    if (i == 2) floor = xdataComp[i].Value.ToString();
                                    if (i == 3) elevation = xdataComp[i].Value.ToString();
                                }

                                // Create metal component
                                MetalComponent comp = new MetalComponent
                                {
                                    Handle = ent.Handle.ToString(),
                                    Type = componentType,
                                    Elevation = elevation,
                                    Floor = floor,
                                    StartPoint = pline.GetPoint3dAt(0),
                                    EndPoint = pline.GetPoint3dAt(1),
                                    Parts = new List<ChildPart>()
                                };

                                // Get parts data
                                string partsJson = GetPartsJsonFromEntity(pline);
                                if (!string.IsNullOrEmpty(partsJson))
                                {
                                    try
                                    {
                                        comp.Parts = Newtonsoft.Json.JsonConvert.DeserializeObject<List<ChildPart>>(partsJson);
                                    }
                                    catch (System.Exception ex)
                                    {
                                        ed.WriteMessage($"\nError deserializing parts: {ex.Message}");
                                    }
                                }

                                components.Add(comp);
                            }
                        }
                    }

                    tr.Commit();
                }

                // Group components by elevation
                var componentsByElevation = components.GroupBy(c => c.Elevation);

                // List to store attachments
                List<Attachment> attachments = new List<Attachment>();

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
                        // Calculate vertical bottom and top points, accounting for parts with adjustments
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
                            // For vertical parts, StartAdjustment affects bottom, EndAdjustment affects top
                            Vector3d direction = verticalTop - verticalBottom;
                            double length = direction.Length;
                            if (length > 0)
                            {
                                direction = direction / length;
                            }

                            Point3d adjustedBottom = verticalBottom + (direction * clipPart.StartAdjustment);
                            Point3d adjustedTop = verticalTop - (direction * clipPart.EndAdjustment);

                            // Find horizontals that might intersect with this adjusted vertical
                            foreach (var horizontal in horizontals)
                            {
                                // Find parts in horizontal that have attachments
                                var attachingParts = horizontal.Parts.Where(p => !string.IsNullOrEmpty(p.Attach)).ToList();

                                foreach (var hPart in attachingParts)
                                {
                                    // Calculate the adjusted horizontal points for this specific part
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
                                    //hDirection = hDirection.GetNormalizedVector();
                                    double hLength = hDirection.Length;
                                    if (hLength > 0)
                                    {
                                        hDirection = hDirection / hLength;
                                    }

                                    Point3d adjustedStart = horizontalStart + (hDirection * hPart.StartAdjustment);
                                    Point3d adjustedEnd = horizontalEnd - (hDirection * hPart.EndAdjustment);

                                    // Check for intersection between adjusted vertical and horizontal parts
                                    if (LinesIntersect(adjustedBottom, adjustedTop,
                                                      adjustedStart, adjustedEnd,
                                                      out Point3d intersectionPt,
                                                      proximityThreshold: 6.0))
                                    {
                                        // Determine which side (L or R) the vertical is on
                                        string side = DetermineSide(adjustedStart, adjustedEnd,
                                                                 adjustedBottom, adjustedTop);

                                        // If part attaches on the determined side
                                        if (hPart.Attach == side)
                                        {
                                            // Calculate height from vertical bottom
                                            double height = intersectionPt.Y - verticalBottom.Y;

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

                // Save attachments to the drawing
                SaveAttachmentsToDrawing(attachments);

                ed.WriteMessage($"\nDetected {attachments.Count} attachments with end-specific adjustments");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in DETECTATTACHMENTS: {ex.Message}");
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

        private class Attachment
        {
            public string HorizontalHandle { get; set; }
            public string VerticalHandle { get; set; }
            public string HorizontalPartType { get; set; }
            public string VerticalPartType { get; set; }
            public string Side { get; set; }
            public double Position { get; set; }      // Position along horizontal
            public double Height { get; set; }        // Height from bottom of vertical
            public bool Invert { get; set; }
            public double Adjust { get; set; }
        }

        // Helper methods
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

            // If no exact intersection, check for proximity
            // Assuming line1 is vertical and line2 is horizontal

            // Get bounding box of vertical line
            double minY = Math.Min(line1Start.Y, line1End.Y);
            double maxY = Math.Max(line1Start.Y, line1End.Y);
            double x1 = line1Start.X;

            // Get bounding box of horizontal line
            double minX = Math.Min(line2Start.X, line2End.X);
            double maxX = Math.Max(line2Start.X, line2End.X);
            double y2 = line2Start.Y;

            // Check if horizontal end is close to vertical
            if (minY <= y2 && y2 <= maxY)  // Horizontal is at same height as vertical
            {
                // Check if left end of horizontal is close to vertical
                if (Math.Abs(minX - x1) <= proximityThreshold)
                {
                    // Create intersection point at the projection
                    intersectionPoint = new Point3d(x1, y2, 0);
                    return true;
                }

                // Check if right end of horizontal is close to vertical
                if (Math.Abs(maxX - x1) <= proximityThreshold)
                {
                    // Create intersection point at the projection
                    intersectionPoint = new Point3d(x1, y2, 0);
                    return true;
                }
            }

            // No intersection found
            return false;
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
            string json = Newtonsoft.Json.JsonConvert.SerializeObject(attachments);

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
            Database db = doc.Database;

            try
            {
                List<Attachment> attachments = LoadAttachmentsFromDrawing();

                if (attachments.Count == 0)
                {
                    ed.WriteMessage("\nNo attachments found in drawing. Run DETECTATTACHMENTS first.");
                    return;
                }

                ed.WriteMessage($"\nFound {attachments.Count} attachments:");

                // Group by vertical handle for better organization
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

        private List<Attachment> LoadAttachmentsFromDrawing()
        {
            List<Attachment> attachments = new List<Attachment>();

            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                // Get named objects dictionary
                DBDictionary nod = (DBDictionary)tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForRead);

                // Check if entry exists
                const string dictName = "METALATTACHMENTS";

                if (nod.Contains(dictName))
                {
                    DBObject obj = tr.GetObject(nod.GetAt(dictName), OpenMode.ForRead);
                    if (obj is Xrecord)
                    {
                        Xrecord xrec = obj as Xrecord;
                        ResultBuffer rb = xrec.Data;

                        if (rb != null)
                        {
                            TypedValue[] values = rb.AsArray();
                            if (values.Length > 0 && values[0].TypeCode == (int)DxfCode.Text)
                            {
                                string json = values[0].Value.ToString();
                                attachments = Newtonsoft.Json.JsonConvert.DeserializeObject<List<Attachment>>(json);
                            }
                        }
                    }
                }

                tr.Commit();
            }

            return attachments;
        }

        [CommandMethod("METALTAKEOFF")]
        public async void MetalTakeoff()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            ed.WriteMessage("\nMetal Takeoff Bridge started.");

            try
            {
                // Get drawing name for reference
                string drawingName = doc.Name;
                ed.WriteMessage($"\nProcessing drawing: {drawingName}");

                // Collect metal components
                List<MetalComponentData> metalCompDataList = new List<MetalComponentData>();

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                    foreach (ObjectId objId in btr)
                    {
                        Entity ent = tr.GetObject(objId, OpenMode.ForRead) as Entity;

                        if (ent is Polyline)
                        {
                            Polyline pline = ent as Polyline;

                            // Check if it has METALCOMP Xdata
                            ResultBuffer rbComp = ent.GetXDataForApplication("METALCOMP");
                            if (rbComp != null)
                            {
                                MetalComponentData compData = new MetalComponentData
                                {
                                    Handle = ent.Handle.ToString(),
                                    ChildParts = new List<ChildPartData>()
                                };

                                // Extract basic component data
                                TypedValue[] xdataComp = rbComp.AsArray();
                                for (int i = 1; i < xdataComp.Length; i++) // Skip app name
                                {
                                    if (i == 1) compData.ComponentType = xdataComp[i].Value.ToString();
                                    if (i == 2) compData.Floor = xdataComp[i].Value.ToString();
                                    if (i == 3) compData.Elevation = xdataComp[i].Value.ToString();
                                }

                                // Get component length
                                if (pline.NumberOfVertices >= 2)
                                {
                                    Point3d startPt = pline.GetPoint3dAt(0);
                                    Point3d endPt = pline.GetPoint3dAt(1);
                                    compData.Length = startPt.DistanceTo(endPt);

                                    compData.StartPoint = new double[] { startPt.X, startPt.Y, startPt.Z };
                                    compData.EndPoint = new double[] { endPt.X, endPt.Y, endPt.Z };
                                }

                                // Extract child parts data
                                string partsJson = "";
                                bool hasMoreChunks = true;
                                int chunkIndex = 0;
                                await System.Threading.Tasks.Task.Delay(1);
                                while (hasMoreChunks)
                                {
                                    ResultBuffer rbParts = ent.GetXDataForApplication("METALPARTS");
                                    if (rbParts == null)
                                    {
                                        // No more chunks or no parts data
                                        hasMoreChunks = false;
                                        continue;
                                    }

                                    TypedValue[] xdataParts = rbParts.AsArray();
                                    if (xdataParts.Length < 3) continue;

                                    // Check if this is the chunk we're looking for
                                    string indexStr = xdataParts[1].Value.ToString();
                                    if (int.Parse(indexStr) == chunkIndex)
                                    {
                                        partsJson += xdataParts[2].Value.ToString();
                                        chunkIndex += xdataParts[2].Value.ToString().Length;
                                    }
                                    else
                                    {
                                        // We've read all chunks
                                        hasMoreChunks = false;
                                    }
                                }

                                // If we have parts JSON, deserialize it
                                if (!string.IsNullOrEmpty(partsJson))
                                {
                                    try
                                    {
                                        List<ChildPart> childParts = Newtonsoft.Json.JsonConvert.DeserializeObject<List<ChildPart>>(partsJson);

                                        // Convert to ChildPartData for API
                                        foreach (ChildPart part in childParts)
                                        {
                                            ChildPartData partData = new ChildPartData
                                            {
                                                Name = part.Name,
                                                PartType = part.PartType,
                                                LengthAdjustment = part.LengthAdjustment,
                                                Material = part.Material,
                                                MarkNumber = part.MarkNumber,
                                                ActualLength = part.GetActualLength(compData.Length)
                                            };

                                            compData.ChildParts.Add(partData);
                                        }
                                    }
                                    catch (System.Exception ex)
                                    {
                                        ed.WriteMessage($"\nError deserializing parts: {ex.Message}");
                                    }
                                }

                                metalCompDataList.Add(compData);
                                ed.WriteMessage($"\nCaptured metal component: Type={compData.ComponentType}, Floor={compData.Floor}, Parts={compData.ChildParts.Count}");
                            }
                        }
                    }

                    tr.Commit();
                }

                ed.WriteMessage($"\nCollected {metalCompDataList.Count} metal components.");

                // Send data to web API
                if (metalCompDataList.Count > 0)
                {
                    ed.WriteMessage("\nSending data to web application...");

                    // Similar to your existing glass takeoff code
                    // Implementation of SendMetalDataToWebApp would be similar to SendDataToWebApp

                    ed.WriteMessage("\nMetal data sent for processing.");
                }

                ed.WriteMessage("\nMetal Takeoff completed successfully.");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError: {ex.Message}");
                ed.WriteMessage($"\nStack trace: {ex.StackTrace}");
            }
        }

        [CommandMethod("METALEDITOR")]
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

                    ed.WriteMessage("\nCreating enhanced panel...");

                    // Create the enhanced panel explicitly first
                    EnhancedMetalComponentPanel panel = new EnhancedMetalComponentPanel();

                    // Add panel to palette set
                    _paletteSet.Add("Metal Parts", panel);

                    ed.WriteMessage("\nEnhanced panel added successfully.");
                }

                // Show the palette set
                _paletteSet.Visible = true;
                ed.WriteMessage("\nPalette set shown successfully.");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in ShowMetalEditor: {ex.Message}");
                ed.WriteMessage($"\nStack trace: {ex.StackTrace}");
            }
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

            // Ask if user wants to copy parent data, child data, or both
            PromptKeywordOptions pKeyOpts = new PromptKeywordOptions("\nWhat to copy? ");
            pKeyOpts.Keywords.Add("Parent");
            pKeyOpts.Keywords.Add("Child");
            pKeyOpts.Keywords.Add("Both");
            pKeyOpts.AllowNone = false;
            pKeyOpts.Keywords.Default = "Both";

            PromptResult pKeyRes = ed.GetKeywords(pKeyOpts);
            if (pKeyRes.Status != PromptStatus.OK) return;

            bool copyParent = pKeyRes.StringResult == "Parent" || pKeyRes.StringResult == "Both";
            bool copyChild = pKeyRes.StringResult == "Child" || pKeyRes.StringResult == "Both";

            // Prompt for selection set of targets
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
                    ComponentType = copyParent ? componentType : null,
                    Floor = copyParent ? floor : null,
                    Elevation = copyParent ? elevation : null,
                    Parts = copyChild ? sourceParts : null
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


        // Static palette set to maintain a single instance
        private static PaletteSet _paletteSet;

        [CommandMethod("SIMULATETAKEOFF")]
        public void SimulateTakeoff()
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            ed.WriteMessage("\nSimulating Metal Takeoff...");

            try
            {
                // Get drawing name for reference
                string drawingName = doc.Name;
                ed.WriteMessage($"\nProcessing drawing: {drawingName}");

                // Collect metal components
                List<MetalComponentData> metalCompDataList = new List<MetalComponentData>();
                Dictionary<string, int> partTypeCounts = new Dictionary<string, int>();

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                    foreach (ObjectId objId in btr)
                    {
                        Entity ent = tr.GetObject(objId, OpenMode.ForRead) as Entity;

                        if (ent is Polyline)
                        {
                            Polyline pline = ent as Polyline;

                            // Check if it has METALCOMP Xdata
                            ResultBuffer rbComp = ent.GetXDataForApplication("METALCOMP");
                            if (rbComp != null)
                            {
                                // Create a new component data object
                                MetalComponentData compData = new MetalComponentData
                                {
                                    Handle = ent.Handle.ToString(),
                                    ChildParts = new List<ChildPartData>()
                                };

                                // Extract basic component data
                                TypedValue[] xdataComp = rbComp.AsArray();
                                for (int i = 1; i < xdataComp.Length; i++) // Skip app name
                                {
                                    if (i == 1) compData.ComponentType = xdataComp[i].Value.ToString();
                                    if (i == 2) compData.Floor = xdataComp[i].Value.ToString();
                                    if (i == 3) compData.Elevation = xdataComp[i].Value.ToString();
                                }

                                // Get component length
                                if (pline.NumberOfVertices >= 2)
                                {
                                    Point3d startPt = pline.GetPoint3dAt(0);
                                    Point3d endPt = pline.GetPoint3dAt(1);
                                    compData.Length = startPt.DistanceTo(endPt);

                                    compData.StartPoint = new double[] { startPt.X, startPt.Y, startPt.Z };
                                    compData.EndPoint = new double[] { endPt.X, endPt.Y, endPt.Z };
                                }

                                // Extract child parts data
                                string partsJson = "";
                                bool hasMoreChunks = true;
                                int chunkIndex = 0;

                                while (hasMoreChunks)
                                {
                                    ResultBuffer rbParts = ent.GetXDataForApplication("METALPARTS");
                                    if (rbParts == null)
                                    {
                                        // No more chunks or no parts data
                                        hasMoreChunks = false;
                                        continue;
                                    }

                                    TypedValue[] xdataParts = rbParts.AsArray();
                                    if (xdataParts.Length < 3) continue;

                                    // Check if this is the chunk we're looking for
                                    string indexStr = xdataParts[1].Value.ToString();
                                    if (int.Parse(indexStr) == chunkIndex)
                                    {
                                        partsJson += xdataParts[2].Value.ToString();
                                        chunkIndex += xdataParts[2].Value.ToString().Length;
                                    }
                                    else
                                    {
                                        // We've read all chunks
                                        hasMoreChunks = false;
                                    }
                                }

                                // If we have parts JSON, deserialize it
                                if (!string.IsNullOrEmpty(partsJson))
                                {
                                    try
                                    {
                                        List<ChildPart> childParts = Newtonsoft.Json.JsonConvert.DeserializeObject<List<ChildPart>>(partsJson);

                                        // Convert to ChildPartData for API
                                        foreach (ChildPart part in childParts)
                                        {
                                            ChildPartData partData = new ChildPartData
                                            {
                                                Name = part.Name,
                                                PartType = part.PartType,
                                                LengthAdjustment = part.LengthAdjustment,
                                                Material = part.Material,
                                                MarkNumber = part.MarkNumber,
                                                ActualLength = part.GetActualLength(compData.Length)
                                            };

                                            // Assign a simulated mark number
                                            string partTypeKey = partData.PartType;
                                            if (!partTypeCounts.ContainsKey(partTypeKey))
                                            {
                                                partTypeCounts[partTypeKey] = 1;
                                            }
                                            else
                                            {
                                                partTypeCounts[partTypeKey]++;
                                            }

                                            partData.MarkNumber = $"{partData.PartType}-{partTypeCounts[partTypeKey]:D3}";
                                            compData.ChildParts.Add(partData);
                                        }
                                    }
                                    catch (System.Exception ex)
                                    {
                                        ed.WriteMessage($"\nError deserializing parts: {ex.Message}");
                                    }
                                }

                                metalCompDataList.Add(compData);
                            }
                        }
                    }

                    tr.Commit();
                }

                // Print takeoff summary
                ed.WriteMessage($"\n\n=== METAL TAKEOFF SUMMARY ===");
                ed.WriteMessage($"\nDrawing: {drawingName}");
                ed.WriteMessage($"\nComponents: {metalCompDataList.Count}");

                // Print part totals by type
                ed.WriteMessage($"\n\n--- Part Counts by Type ---");
                foreach (var pair in partTypeCounts)
                {
                    ed.WriteMessage($"\n{pair.Key}: {pair.Value}");
                }

                // Print detailed component info
                ed.WriteMessage($"\n\n--- Component Details ---");
                foreach (var comp in metalCompDataList)
                {
                    ed.WriteMessage($"\n\nComponent: {comp.ComponentType} (Floor {comp.Floor}, Elev {comp.Elevation})");
                    ed.WriteMessage($"\nLength: {comp.Length:F4}");
                    ed.WriteMessage($"\nParts:");

                    foreach (var part in comp.ChildParts)
                    {
                        ed.WriteMessage($"\n  {part.MarkNumber}: {part.Name}, {part.ActualLength:F4} units, {part.Material}");
                    }
                }

                ed.WriteMessage($"\n\n=== END OF TAKEOFF ===");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError: {ex.Message}");
            }
        }
    }

    // Data classes for API communication
    public class MetalComponentData
    {
        public string Handle { get; set; }
        public string ComponentType { get; set; }
        public string Floor { get; set; }
        public string Elevation { get; set; }
        public double Length { get; set; }
        public double[] StartPoint { get; set; }
        public double[] EndPoint { get; set; }
        public List<ChildPartData> ChildParts { get; set; }
    }

    public class ChildPartData
    {
        public string Name { get; set; }
        public string PartType { get; set; }
        public double LengthAdjustment { get; set; }
        public double ActualLength { get; set; }
        public string Material { get; set; }
        public string MarkNumber { get; set; }
    }
}