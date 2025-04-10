using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.PlottingServices;
using Autodesk.AutoCAD.Runtime;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using static TakeoffBridge.FabTicketGenerator;

namespace TakeoffBridge
{
    public class FabTicketGenerator
    {
        private readonly string templatePath = @"C:\CSE\Takeoff\fabs\dies";
        private readonly string outputPath = @"C:\CSE\Takeoff\fabs\tickets";

        // Main class to store part information for fabrication
        public class FabPartInfo
        {
            public string PartNumber { get; set; }
            public string PartType { get; set; }
            public string Fab { get; set; }
            public string Finish { get; set; }
            public string MarkNumber { get; set; }
            public double Length { get; set; }
            public bool IsVertical { get; set; }
            public double LeftTilt { get; set; } = 90;
            public double LeftRotation { get; set; } = 45;
            public double RightTilt { get; set; } = 90;
            public double RightRotation { get; set; } = 45;
            public List<AttachmentInfo> Attachments { get; set; } = new List<AttachmentInfo>();
        }

        // Reused from TemplateGenerator
        public class AttachmentInfo
        {
            public string Side { get; set; }
            public double Position { get; set; }
            public double Height { get; set; }
            public bool Invert { get; set; }
            public string AttachedPartNumber { get; set; }
            public string AttachedFab { get; set; }
        }

        public FabTicketGenerator()
        {
            // Ensure output directory exists
            if (!Directory.Exists(outputPath))
            {
                Directory.CreateDirectory(outputPath);
            }
        }

        public bool GenerateFabricationTicket(FabPartInfo partInfo)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor editor = doc.Editor;

            try
            {
                // Check if the template exists
                string templateFile = Path.Combine(templatePath, $"{partInfo.PartNumber}-{partInfo.Fab}.dwg");
                if (!File.Exists(templateFile))
                {
                    editor.WriteMessage($"\nError: Template file not found: {templateFile}");
                    return false;
                }
                // Create the output filename
                string ticketFile = Path.Combine(outputPath, $"{partInfo.MarkNumber}_{partInfo.PartNumber}_{partInfo.Length}.dwg");

                // Open the template and create the fabrication ticket
                return CreateTicketFromTemplate(templateFile, ticketFile, partInfo, editor);
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage($"\nError checking template file: {ex.Message}");
                return false;
            }
        }

        // Helper method to batch generate multiple fabrication tickets
        public void BatchGenerateFabricationTickets(List<FabPartInfo> partInfos)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor editor = doc.Editor;

            editor.WriteMessage($"\nStarting batch generation of {partInfos.Count} fabrication tickets...");

            int successCount = 0;
            int failureCount = 0;

            foreach (var partInfo in partInfos)
            {
                if (GenerateFabricationTicket(partInfo))
                {
                    successCount++;
                }
                else
                {
                    failureCount++;
                    editor.WriteMessage($"\nFailed to generate ticket for {partInfo.MarkNumber} {partInfo.PartNumber}-{partInfo.Fab}");
                }
            }

            editor.WriteMessage($"\nBatch generation complete. Success: {successCount}, Failed: {failureCount}");
        }

        // Method to gather part information from the drawing and create fabrication tickets
        public void GenerateTicketsFromDrawing()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor editor = doc.Editor;

            // First, ensure all templates exist by running the template generator
            TemplateGenerator templateGen = new TemplateGenerator();
            templateGen.GenerateTemplatesFromDrawing();

            // Now collect all parts that need fabrication tickets
            List<FabPartInfo> partsToProcess = new List<FabPartInfo>();

            // Get component and attachment data similar to TemplateGenerator
            var attachments = LoadAttachmentsFromDrawing();
            Dictionary<string, string> componentTypesByHandle = new Dictionary<string, string>();
            Dictionary<string, List<TemplateGenerator.ChildPart>> partsByComponentHandle = new Dictionary<string, List<TemplateGenerator.ChildPart>>();

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                // Collect data from all components (similar to TemplateGenerator)
                // This code would be similar to the component data collection in TemplateGenerator
                // but instead of just identifying unique part-fab combinations, we'd collect
                // all instances with their specific mark numbers, lengths, and angles

                // For example:
                var components = GetAllMetalComponents(db, tr, editor);

                foreach (var component in components)
                {
                    string componentType = GetComponentType(component);
                    if (componentType != null)
                    {
                        componentTypesByHandle[component.Handle.ToString()] = componentType;

                        List<TemplateGenerator.ChildPart> childParts = GetChildParts(component);
                        partsByComponentHandle[component.Handle.ToString()] = childParts;

                        // Process each actual part for fabrication
                        foreach (var part in childParts)
                        {
                            // Skip parts without required info
                            if (string.IsNullOrEmpty(part.Fab) || string.IsNullOrEmpty(part.Name))
                                continue;

                            // Create fabrication info
                            FabPartInfo fabInfo = new FabPartInfo
                            {
                                PartNumber = part.Name,
                                PartType = part.PartType,
                                Fab = part.Fab,
                                Finish = part.Finish,
                                MarkNumber = part.MarkNumber,
                                IsVertical = componentType == "Vertical",
                                // If it's a fixed length part, use that, otherwise calculate
                                Length = part.IsFixedLength ? part.FixedLength : CalculatePartLength(component, part)
                            };

                            // Add to collection
                            partsToProcess.Add(fabInfo);
                        }
                    }
                }

                // Now process attachments and associate them with vertical parts
                foreach (var attachment in attachments)
                {
                    string vertHandle = attachment.VerticalHandle;

                    // Check if we have data for this vertical component
                    if (partsByComponentHandle.ContainsKey(vertHandle))
                    {
                        List<TemplateGenerator.ChildPart> verticalParts = partsByComponentHandle[vertHandle];

                        // Find the matching vertical part
                        foreach (var part in verticalParts)
                        {
                            if (part.PartType == attachment.VerticalPartType)
                            {
                                // Find the corresponding FabPartInfo
                                var fabPart = partsToProcess.FirstOrDefault(fp =>
                                    fp.PartNumber == part.Name &&
                                    fp.Fab == part.Fab &&
                                    fp.MarkNumber == part.MarkNumber);

                                if (fabPart != null)
                                {
                                    // Find the horizontal part info
                                    if (partsByComponentHandle.TryGetValue(attachment.HorizontalHandle, out var horizontalParts))
                                    {
                                        var horizontalPart = horizontalParts.FirstOrDefault(hp =>
                                            hp.PartType == attachment.HorizontalPartType);

                                        if (horizontalPart != null)
                                        {
                                            // Add attachment information
                                            fabPart.Attachments.Add(new AttachmentInfo
                                            {
                                                Side = attachment.Side,
                                                Position = attachment.Position,
                                                Height = attachment.Height,
                                                Invert = attachment.Invert,
                                                AttachedPartNumber = horizontalPart.Name,
                                                AttachedFab = horizontalPart.Fab
                                            });
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                tr.Commit();
            }

            // Now generate all fabrication tickets
            editor.WriteMessage($"\nFound {partsToProcess.Count} parts to generate fabrication tickets for.");
            BatchGenerateFabricationTickets(partsToProcess);
        }

        // Helper to calculate part length - in a real implementation this would consider
        // component geometry, start/end adjustments, etc.
        private double CalculatePartLength(Entity component, TemplateGenerator.ChildPart part)
        {
            // Placeholder implementation - in reality would calculate based on:
            // - Component geometry length
            // - Part's length adjustment value
            // - Start/end adjustments

            // For now, return a default or mock calculation
            return 60.0; // Example default length
        }

        // Helper method to get all metal components
        private List<Entity> GetAllMetalComponents(Database db, Transaction tr, Editor editor)
        {
            List<Entity> components = new List<Entity>();

            // Get all polylines in the drawing that might be metal components
            TypedValue[] filterList = new TypedValue[] {
        new TypedValue((int)DxfCode.Start, "LWPOLYLINE")
    };

            // Create filter without using a 'using' statement
            SelectionFilter filter = new SelectionFilter(filterList);
            PromptSelectionResult selRes = editor.SelectAll(filter);

            if (selRes.Status == PromptStatus.OK)
            {
                foreach (ObjectId id in selRes.Value.GetObjectIds())
                {
                    Entity ent = (Entity)tr.GetObject(id, OpenMode.ForRead);
                    // Check if this has metal component xdata
                    if (ent.GetXDataForApplication("METALCOMP") != null)
                    {
                        components.Add(ent);
                    }
                }
            }

            return components;
        }

        // Helper methods for XData (reused from TemplateGenerator)
        private string GetComponentType(Entity ent)
        {
            // Get component type from Xdata
            ResultBuffer rbComp = ent.GetXDataForApplication("METALCOMP");
            if (rbComp != null)
            {
                TypedValue[] tvs = rbComp.AsArray();
                if (tvs.Length > 1 && tvs[1].TypeCode == (int)DxfCode.ExtendedDataAsciiString)
                {
                    return tvs[1].Value.ToString();
                }
            }
            return null;
        }

        private List<TemplateGenerator.ChildPart> GetChildParts(Entity ent)
        {
            // Implementation identical to the one in TemplateGenerator
            List<TemplateGenerator.ChildPart> result = new List<TemplateGenerator.ChildPart>();
            // Get info about parts chunks
            int numChunks = 0;
            ResultBuffer rbInfo = ent.GetXDataForApplication("METALPARTSINFO");
            if (rbInfo != null)
            {
                TypedValue[] infoValues = rbInfo.AsArray();
                foreach (TypedValue tv in infoValues)
                {
                    if (tv.TypeCode == (int)DxfCode.ExtendedDataInteger32)
                    {
                        numChunks = Convert.ToInt32(tv.Value);
                        break;
                    }
                }
            }
            // Build JSON string from chunks if chunks exist
            if (numChunks > 0)
            {
                StringBuilder jsonBuilder = new StringBuilder();
                for (int i = 0; i < numChunks; i++)
                {
                    string appName = $"METALPARTS{i}";
                    ResultBuffer rbChunk = ent.GetXDataForApplication(appName);
                    if (rbChunk != null)
                    {
                        TypedValue[] chunkValues = rbChunk.AsArray();
                        foreach (TypedValue tv in chunkValues)
                        {
                            if (tv.TypeCode == (int)DxfCode.ExtendedDataAsciiString)
                            {
                                jsonBuilder.Append(tv.Value.ToString());
                            }
                        }
                    }
                }
                // Parse the complete JSON
                if (jsonBuilder.Length > 0)
                {
                    try
                    {
                        result = JsonConvert.DeserializeObject<List<TemplateGenerator.ChildPart>>(jsonBuilder.ToString());
                    }
                    catch (System.Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error parsing child parts JSON: {ex.Message}");
                    }
                }
            }
            return result;
        }

        private List<TemplateGenerator.Attachment> LoadAttachmentsFromDrawing()
        {
            // Implementation identical to the one in TemplateGenerator
            List<TemplateGenerator.Attachment> loadedAttachments = new List<TemplateGenerator.Attachment>();
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor editor = doc.Editor;

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
                                loadedAttachments = JsonConvert.DeserializeObject<List<TemplateGenerator.Attachment>>(json);
                            }
                        }
                    }
                }
                tr.Commit();
            }
            return loadedAttachments;
        }

        private bool CreateTicketFromTemplate(string templateFile, string outputFile, FabPartInfo partInfo, Editor editor)
        {
            try
            {
                // First just copy the template file
                File.Copy(templateFile, outputFile, true);

                // Then open the copied file for modification
                using (Database db = new Database(false, true))
                {
                    db.ReadDwgFile(outputFile, FileOpenMode.OpenForReadAndWriteNoShare, false, null);

                    // Modify the drawing
                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        try
                        {
                            // 1. Modify the main part to the correct length
                            ModifyMainPart(db, tr, partInfo);

                            // 2. Add attachments if they exist
                            if (partInfo.Attachments != null && partInfo.Attachments.Count > 0)
                            {
                                AddAttachments(db, tr, partInfo);
                            }

                            // 3. Update tables with part information
                            UpdateTables(db, tr, partInfo);

                            // Commit changes
                            tr.Commit();

                            // Save the modified drawing
                            db.SaveAs(outputFile, DwgVersion.Current);

                            editor.WriteMessage($"\nGenerated fabrication ticket: {outputFile}");
                            return true;
                        }
                        catch (System.Exception ex)
                        {
                            editor.WriteMessage($"\nError modifying ticket from template: {ex.Message}");
                            tr.Abort();
                            return false;
                        }
                    }
                }
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage($"\nError creating ticket from template: {ex.Message}");
                return false;
            }
        }

        private void ModifyMainPart(Database db, Transaction tr, FabPartInfo partInfo)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor editor = doc.Editor;

            try
            {
                // In a real implementation, you would find the main part
                // and modify its properties. Since we're working with a template,
                // we'll use the LISP drwprt function to update the main part.

                // Find the model space
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord modelSpace = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                // We need to commit the transaction before calling any external functions
                tr.Commit();

                // Call the drwprt function to replace the main part with updated dimensions
                PartDrawingGenerator generator = new PartDrawingGenerator();
                TakeoffBridge.PartDrawingGenerator.MirrorType mirrorType = partInfo.IsVertical ?
                    TakeoffBridge.PartDrawingGenerator.MirrorType.VerticalMullion :
                    TakeoffBridge.PartDrawingGenerator.MirrorType.None;

                // Note: This is a simplified approach. A more complete implementation would
                // modify the existing entities rather than replacing them.

                // Since we're modifying an existing database rather than creating a new file,
                // we'll pass null for the outputPath parameter
                generator.DrwPrt(
                    partInfo.PartNumber,
                    partInfo.Length,
                    partInfo.LeftRotation,
                    partInfo.LeftTilt,
                    partInfo.RightRotation,
                    partInfo.RightTilt,
                    (int)mirrorType,
                    Point3d.Origin,
                    null, // No side designation for main template
                    0,    // Default height
                    null  // No output file, modifying current database
                );
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage($"\nError modifying main part: {ex.Message}");
                throw; // Re-throw to be caught by the calling method
            }
        }

        private void AddAttachments(Database db, Transaction tr, FabPartInfo partInfo)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor editor = doc.Editor;

            try
            {
                // For each attachment in partInfo.Attachments, we'll use DrwPrt to create
                // the attachment at the right position

                // We need to commit the transaction before calling any external functions
                tr.Commit();

                // Create a new part drawing generator
                PartDrawingGenerator generator = new PartDrawingGenerator();

                foreach (var attachment in partInfo.Attachments)
                {
                    // Set mirror type based on whether attachment is inverted
                    TakeoffBridge.PartDrawingGenerator.MirrorType mirrorType = attachment.Invert ?
                        TakeoffBridge.PartDrawingGenerator.MirrorType.ShearBlockClip :
                        TakeoffBridge.PartDrawingGenerator.MirrorType.None;

                    // Create a point at the attachment position/height
                    Point3d attachPosition = new Point3d(attachment.Position, 0, attachment.Height);

                    // Call DrwPrt to create the attachment
                    generator.DrwPrt(
                        attachment.AttachedPartNumber,
                        0, // Length 0 for attachments (shear blocks)
                        45, // Standard angles
                        90,
                        45,
                        90,
                        (int)mirrorType,
                        attachPosition,
                        attachment.Side,
                        attachment.Height,
                        null // No output file, adding to current database
                    );
                }

                // After adding all attachments, we need to start a new transaction
                tr = db.TransactionManager.StartTransaction();
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage($"\nError adding attachments: {ex.Message}");
                throw; // Re-throw to be caught by the calling method
            }
        }

        private void UpdateTables(Database db, Transaction tr, FabPartInfo partInfo)
        {
            // Find all tables in the drawing
            BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
            BlockTableRecord modelSpace = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

            foreach (ObjectId objId in modelSpace)
            {
                if (tr.GetObject(objId, OpenMode.ForRead) is Table table)
                {
                    // Open the table for write
                    table = (Table)tr.GetObject(objId, OpenMode.ForWrite);

                    //// Determine which type of table we're dealing with
                    //if (table.Rows >= 1 && table.Columns >= 7) // Main parts table (assuming it has at least 7 columns)
                    //{
                    //    // Update the main parts table with our information
                    //    // Assuming standard column layout:
                    //    // Col 0: Mark Number
                    //    // Col 1: Part Number
                    //    // Col 2: Fab
                    //    // Col 3: Length
                    //    // Col 4: Left Angle
                    //    // Col 5: Right Angle
                    //    // Col 6: Finish

                    //    // Find first empty row or create one
                    //    int rowIndex = 1; // Skip header row
                    //    bool rowFound = false;

                    //    while (rowIndex < table.Rows)
                    //    {
                    //        // Check if this row is empty
                    //        string markText = table.GetCellValue(rowIndex, 0)?.ToString() ?? "";
                    //        if (string.IsNullOrWhiteSpace(markText))
                    //        {
                    //            rowFound = true;
                    //            break;
                    //        }
                    //        rowIndex++;
                    //    }

                    //    if (!rowFound && rowIndex >= table.Rows)
                    //    {
                    //        // Add a new row
                    //        table.InsertRows(rowIndex, 1, table.GetRowHeight(1));
                    //    }

                    //    // Update cell values
                    //    table.SetCellValue(rowIndex, 0, partInfo.MarkNumber ?? "");
                    //    table.SetCellValue(rowIndex, 1, partInfo.PartNumber ?? "");
                    //    table.SetCellValue(rowIndex, 2, partInfo.Fab ?? "");
                    //    table.SetCellValue(rowIndex, 3, partInfo.Length.ToString("F3"));

                    //    // Format angles for left and right sides
                    //    string leftAngle = $"{partInfo.LeftRotation}° / {partInfo.LeftTilt}°";
                    //    string rightAngle = $"{partInfo.RightRotation}° / {partInfo.RightTilt}°";

                    //    table.SetCellValue(rowIndex, 4, leftAngle);
                    //    table.SetCellValue(rowIndex, 5, rightAngle);
                    //    table.SetCellValue(rowIndex, 6, partInfo.Finish ?? "");
                    //}
                    //else if (table.Rows >= 1 && table.Columns >= 5) // Attachments table (typically has fewer columns)
                    //{
                    //    // Clear any existing attachment rows (keep header)
                    //    while (table.Rows > 1)
                    //    {
                    //        table.DeleteRows(1, 1);
                    //    }

                    //    // Add a row for each attachment
                    //    for (int i = 0; i < partInfo.Attachments.Count; i++)
                    //    {
                    //        var attachment = partInfo.Attachments[i];

                    //        // Add a new row
                    //        if (i > 0 || table.Rows <= 1)
                    //        {
                    //            table.InsertRows(table.Rows, 1, table.GetRowHeight(0));
                    //        }

                    //        // Set cell values for the attachment
                    //        // Col 0: Part Number
                    //        // Col 1: Fab
                    //        // Col 2: Side
                    //        // Col 3: Position
                    //        // Col 4: Height

                    //        table.SetCellValue(table.Rows - 1, 0, attachment.AttachedPartNumber ?? "");
                    //        table.SetCellValue(table.Rows - 1, 1, attachment.AttachedFab ?? "");
                    //        table.SetCellValue(table.Rows - 1, 2, attachment.Side ?? "");
                    //        table.SetCellValue(table.Rows - 1, 3, attachment.Position.ToString("F3"));
                    //        table.SetCellValue(table.Rows - 1, 4, attachment.Height.ToString("F3"));

                    //        // Add Invert flag if necessary
                    //        if (table.Columns >= 6 && attachment.Invert)
                    //        {
                    //            table.SetCellValue(table.Rows - 1, 5, "Yes");
                    //        }
                    //    }
                    //}
                }
            }
        }
    }
}