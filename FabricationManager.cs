using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Newtonsoft.Json.Linq;

namespace TakeoffBridge
{
    /// <summary>
    /// Consolidated class that manages both template generation and fabrication ticket creation.
    /// </summary>
    public partial class FabricationManager
    {
        #region Properties and Fields

        // Path configuration
        private readonly string RootPath;
        private readonly string PartsPath;
        private readonly string FabsPath;
        private readonly string DiesPath;
        private readonly string ApPath;
        private readonly string TemplateFile;

        // Data structures for parts
        public class PartInfo
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
            public bool IsShopUse { get; set; }
            public List<AttachmentInfo> Attachments { get; set; } = new List<AttachmentInfo>();

            // For tracking uniqueness in templates
            public override bool Equals(object obj)
            {
                if (obj is PartInfo other)
                {
                    return PartNumber == other.PartNumber && Fab == other.Fab;
                }
                return false;
            }

            public override int GetHashCode()
            {
                return $"{PartNumber}-{Fab}".GetHashCode();
            }
        }

        public class AttachmentInfo
        {
            public string Side { get; set; }
            public double Position { get; set; }
            public double Height { get; set; }
            public bool Invert { get; set; }
            public string AttachedPartNumber { get; set; }
            public string AttachedFab { get; set; }
        }

        

        // Text locations from VBA code
        private readonly double DefaultTextHeight = 0.125;
        private readonly double PartNoX = 0.55; // ptx - 0.3 (ptx = 0.85)
        private readonly double PartNoY = 4.76954;
        private readonly double FabX = 2.43832; // ptx + 1.58832
        private readonly double FabY = 4.76954;
        private readonly double PageNoX = 3.75053; // ptx + 2.90053
        private readonly double PageNoY = 4.76954;
        private readonly double FinishX = 4.99192; // ptx + 4.14192
        private readonly double FinishY = 4.76954;
        private readonly double DescriptionX = 0.65; // ptx - 0.2
        private readonly double DescriptionY = 4.4681; // pty - 0.30144
        private readonly double JobX = 1.31993; // ptx + 0.46993
        private readonly double JobY = 4.13015; // pty - 0.63939

        // Table information
        private readonly double TableStartX = 0.85;
        private readonly double TableStartY = 3.22875; // pty - 1.54079
        private readonly double TableRowHeight = 0.33;
        private readonly int MaxRowsPerTicket = 9;

        #endregion

        #region Constructor

        /// <summary>
        /// Initializes a new instance of the FabricationManager class.
        /// </summary>
        public FabricationManager()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            string docFolder = Path.GetDirectoryName(doc.Name);

            // Initialize paths based on the current document's folder
            RootPath = docFolder;
            PartsPath = Path.Combine(RootPath, "parts");
            FabsPath = Path.Combine(RootPath, "fabs");
            DiesPath = Path.Combine(FabsPath, "dies");
            ApPath = Path.Combine(DiesPath, "ap");
            TemplateFile = Path.Combine(DiesPath, "fabtemplate.dwg");

            // Ensure directories exist
            EnsureDirectoriesExist();
        }

        /// <summary>
        /// Creates required directories if they don't exist
        /// </summary>
        private void EnsureDirectoriesExist()
        {
            try
            {
                if (!Directory.Exists(PartsPath))
                    Directory.CreateDirectory(PartsPath);

                if (!Directory.Exists(FabsPath))
                    Directory.CreateDirectory(FabsPath);

                if (!Directory.Exists(DiesPath))
                    Directory.CreateDirectory(DiesPath);

                if (!Directory.Exists(ApPath))
                    Directory.CreateDirectory(ApPath);
            }
            catch (System.Exception ex)
            {
                Document doc = Application.DocumentManager.MdiActiveDocument;
                System.Diagnostics.Debug.WriteLine($"Error Creating Directory Structure: {ex.Message}");
            }
        }

        #endregion

        #region Main Process Methods

        /// <summary>
        /// Main entry point to process the current drawing and generate fabrication tickets
        /// </summary>
        public void ProcessDrawingAndGenerateTickets()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            ed.WriteMessage("\n--- Starting Fabrication Process ---");

            try
            {
                // 1. Get data from drawing
                ed.WriteMessage("\nStep 1: Collecting data from drawing...");
                List<PartInfo> allParts = CollectPartsDataFromDrawing();
                ed.WriteMessage($"\nFound {allParts.Count} parts in the drawing.");

                // 2. Ensure all templates exist
                ed.WriteMessage("\nStep 2: Ensuring templates exist...");
                EnsureTemplatesExist(allParts);

                // 3. Generate fabrication tickets and attachments
                ed.WriteMessage("\nStep 3: Generating fabrication tickets...");
                GenerateFabricationTickets(allParts);

                ed.WriteMessage("\n--- Fabrication Process Completed Successfully ---");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in fabrication process: {ex.Message}");
                if (ex.InnerException != null)
                {
                    ed.WriteMessage($"\nInner exception: {ex.InnerException.Message}");
                }
            }
        }

        #endregion

        #region Data Collection Methods

        private List<PartInfo> CollectPartsDataFromDrawing()
        {
            List<PartInfo> allParts = new List<PartInfo>();
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            try
            {
                // Get all components with attachments
                List<DrawingComponentManager.Component> components =
                    DrawingComponentManager.Instance.GetComponentsWithAttachments();

                ed.WriteMessage($"\nFound {components.Count} metal components in drawing.");

                // Convert to PartInfo objects
                foreach (var comp in components)
                {
                    foreach (var part in comp.Parts)
                    {
                        // Skip parts with missing essential information
                        if (string.IsNullOrEmpty(part.Fab) || string.IsNullOrEmpty(part.Name))
                            continue;

                        // Calculate part length correctly
                        double partLength;
                        if (part.IsFixedLength)
                        {
                            partLength = part.FixedLength;
                        }
                        else
                        {
                            partLength = comp.Length + part.StartAdjustment + part.EndAdjustment;
                        }

                        // Create part info
                        PartInfo partInfo = new PartInfo
                        {
                            PartNumber = part.Name,
                            PartType = part.PartType,
                            Fab = part.Fab,
                            Finish = part.Finish,
                            MarkNumber = part.MarkNumber,
                            Length = partLength,
                            IsVertical = comp.Type == "Vertical",
                            IsShopUse = part.IsShopUse
                        };

                        // Add attachments
                        foreach (var attachment in part.Attachments)
                        {
                            partInfo.Attachments.Add(new AttachmentInfo
                            {
                                Side = attachment.Side,
                                Position = attachment.Position,
                                Height = attachment.Height,
                                Invert = attachment.Invert,
                                AttachedPartNumber = attachment.AttachedPartNumber,
                                AttachedFab = attachment.AttachedFab
                            });
                        }

                        allParts.Add(partInfo);
                    }
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError collecting part data: {ex.Message}");
                throw;
            }

            return allParts;
        }



        #endregion

        #region Template Management

        /// <summary>
        /// Ensures all necessary templates exist for the given parts
        /// </summary>
        private void EnsureTemplatesExist(List<PartInfo> parts)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            // Get unique part numbers and fabs
            var uniquePartNos = parts.Select(p => p.PartNumber).Distinct().ToList();
            var allFabs = parts.Select(p => p.Fab).Distinct().ToList();

            // Ensure template.dwg exists
            if (!File.Exists(TemplateFile))
            {
                ed.WriteMessage($"\nError: Base template file not found: {TemplateFile}");
                return;
            }

            // Process each unique part number
            foreach (string partNo in uniquePartNos)
            {
                // First check/create Fab 1 template (base template)
                string baseFabPath = Path.Combine(DiesPath, $"{partNo}-1.dwg");
                bool baseExists = File.Exists(baseFabPath);

                if (!baseExists)
                {
                    // Create base template
                    ed.WriteMessage($"\nCreating base template for part {partNo}-1...");

                    // Get vertical flag from matching parts
                    bool isVertical = parts.Any(p => p.PartNumber == partNo && p.IsVertical);

                    if (!GenerateBaseTemplate(partNo, isVertical))
                    {
                        ed.WriteMessage($"\nFailed to create base template for {partNo}-1");
                        continue;
                    }

                    baseExists = true;
                }

                // Now handle other fabs if base template exists
                if (baseExists)
                {
                    foreach (string fab in allFabs)
                    {
                        if (fab == "1") continue; // Skip fab 1, already handled

                        string fabPath = Path.Combine(DiesPath, $"{partNo}-{fab}.dwg");
                        if (!File.Exists(fabPath))
                        {
                            ed.WriteMessage($"\nCopying base template to create {partNo}-{fab}...");
                            if (!CopyBaseTemplate(partNo, fab))
                            {
                                ed.WriteMessage($"\nFailed to copy base template for {partNo}-{fab}");
                            }
                        }
                    }
                }
            }
        }

        /// <summary>
        /// Generates a base template (Fab 1) for a part
        /// </summary>
        private bool GenerateBaseTemplate(string partNumber, bool isVertical)
        {
            Document currentDoc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = currentDoc.Editor;

            string templatePath = Path.Combine(DiesPath, $"{partNumber}-1.dwg");
            string partPath = Path.Combine(PartsPath, $"{partNumber}.dwg");

            try
            {
                // Create a new database from scratch
                using (Database newDb = new Database(true, false))
                {
                    // Copy the template file's contents into our new database
                    newDb.ReadDwgFile(TemplateFile, FileOpenMode.OpenForReadAndAllShare, true, "");

                    // Add the part using Drawfab
                    using (Transaction tr = newDb.TransactionManager.StartTransaction())
                    {
                        try
                        {
                            // Check if part file exists
                            if (File.Exists(partPath))
                            {
                                // Create Drawfab instance
                                Drawfab drawfab = new Drawfab(partNumber);

                                // Use a standard length (30 inches)
                                double standardLength = 30.0;

                                // Determine handed parameters based on whether it's vertical
                                bool handed = false;
                                string handedSide = "";

                                // Create transformation for insertion point
                                Matrix3d transform = Matrix3d.Displacement(new Vector3d(0, 0, 0));

                                // Create the part in the database
                                drawfab.CreateExtrudedPart(
                                    newDb,
                                    partPath,
                                    partNumber,
                                    standardLength,
                                    90, // Left miter 
                                    90, // Left tilt
                                    90, // Right miter
                                    90, // Right tilt
                                    handed,
                                    handedSide,
                                    transform,
                                    false,
                                    false,
                                    true,
                                    isVertical);
                            }
                            else
                            {
                                ed.WriteMessage($"\nPart file not found: {partPath}. Created blank template.");
                            }

                            tr.Commit();
                        }
                        catch (System.Exception ex)
                        {
                            ed.WriteMessage($"\nError creating part: {ex.Message}");
                            return false;
                        }
                    }

                    using (Transaction tr = newDb.TransactionManager.StartTransaction())
                    {
                        try
                        {
                            // ... existing code for part insertion ...

                            // Update layer visibility based on part type
                            SetLayerVisibility(newDb, tr, isVertical);

                            // Update viewport visibility based on part type
                            SetViewportVisibility(newDb, tr, isVertical);

                            tr.Commit();
                        }
                        catch (System.Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Error: {ex.Message}");
                            tr.Abort();
                        }
                    }

                    // Save to the output file
                    newDb.SaveAs(templatePath, DwgVersion.Current);
                    ed.WriteMessage($"\nTemplate created successfully for {partNumber}");
                    return true;
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError creating template: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Controls layer visibility in the drawing
        /// </summary>
        private void SetLayerVisibility(Database db, Transaction tr, bool isVertical)
        {
            // Get the layer table
            LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);

            // Check if our layers exist
            if (lt.Has("x-title|BORDER4") && lt.Has("x-title|BORDER1"))
            {
                // Get layer records for write
                LayerTableRecord ltrBorder4 = (LayerTableRecord)tr.GetObject(lt["x-title|BORDER4"], OpenMode.ForWrite);
                LayerTableRecord ltrBorder1 = (LayerTableRecord)tr.GetObject(lt["x-title|BORDER1"], OpenMode.ForWrite);

                if (isVertical)
                {
                    // Turn on Border4, turn off Border1
                    ltrBorder4.IsOff = false;
                    ltrBorder1.IsOff = true;
                }
                else
                {
                    // Turn off Border4, turn on Border1
                    ltrBorder4.IsOff = true;
                    ltrBorder1.IsOff = false;
                }
            }
            else
            {
                System.Diagnostics.Debug.WriteLine("Required layers not found in template");
            }
        }

        /// <summary>
        /// Manages viewport visibility - keeps first 4 viewports on, toggles last 2 based on part type
        /// </summary>
        private void SetViewportVisibility(Database db, Transaction tr, bool isVertical)
        {

            // Get all layouts
            DBDictionary layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);

            // Find the layout we're working with (not Model)
            Layout layout = null;
            foreach (DBDictionaryEntry entry in layoutDict)
            {
                if (entry.Key != "Model")
                {
                    layout = (Layout)tr.GetObject(entry.Value, OpenMode.ForRead);
                    break;
                }
            }

            if (layout == null)
            {
                System.Diagnostics.Debug.WriteLine("No paper space layout found");
                return;
            }

            // Get the block table record for this layout
            BlockTableRecord btr = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForRead);

            // Track which viewports we've modified
            bool found534 = false;
            bool found536 = false;

            // First pass: Set all viewports as needed
            foreach (ObjectId id in btr)
            {
                if (id.ObjectClass.DxfName == "VIEWPORT")
                {
                    Viewport vp = (Viewport)tr.GetObject(id, OpenMode.ForWrite);
                    string handle = vp.Handle.ToString();

                    // Turn on first 4 viewports for all parts
                    //if (handle == "531" || handle == "52D" || handle == "52F" || handle == "52B")
                    if (handle == "531" || handle == "52D")
                    {
                        vp.On = true;
                        System.Diagnostics.Debug.WriteLine($"Set viewport with handle {handle} to ON (always)");
                    }
                    if (handle == "52F" || handle == "52B")
                    {
                        vp.On = false;
                        System.Diagnostics.Debug.WriteLine($"Set viewport with handle {handle} to OFF (always)");
                    }
                    // Special handling for paperspace viewport (handle 23)
                    if (handle == "23")
                    {
                        // Move to VPORTS layer and turn it off
                        vp.Layer = "VPORTS";
                        vp.On = false;
                        System.Diagnostics.Debug.WriteLine($"Moved paperspace viewport (handle 23) to VPORTS layer and turned it OFF");
                        continue;
                    }
                    // Handle the last 2 viewports based on part type
                    else if (handle == "534")
                    {
                        vp.On = isVertical;
                        found534 = true;
                        System.Diagnostics.Debug.WriteLine($"Set viewport with handle 534 to {(isVertical ? "ON" : "OFF")}");
                    }
                    else if (handle == "536")
                    {
                        vp.On = isVertical;
                        found536 = true;
                        System.Diagnostics.Debug.WriteLine($"Set viewport with handle 536 to {(isVertical ? "ON" : "OFF")}");
                    }
                    else
                    {
                        // For any other viewport, log that we found it but didn't modify it
                        System.Diagnostics.Debug.WriteLine($"Found unrecognized viewport with handle {handle}, didn't modify");
                    }
                }
            }

            // Second pass: If we couldn't find viewports by handle, try by position
            if (!found534 || !found536)
            {
                foreach (ObjectId id in btr)
                {
                    if (id.ObjectClass.DxfName == "VIEWPORT")
                    {
                        Viewport vp = (Viewport)tr.GetObject(id, OpenMode.ForWrite);

                        // Only process if we haven't found one of our target viewports
                        if ((!found534 || !found536) && vp.Handle.ToString() != "531" &&
                            vp.Handle.ToString() != "52D" && vp.Handle.ToString() != "52F" &&
                            vp.Handle.ToString() != "52B")
                        {
                            // Try to identify by position for the last 2 viewports
                            Point3d center = vp.CenterPoint;

                            // The bottom-left viewport (around position 2.16, 3.59)
                            if (!found536 && Math.Abs(center.X - 2.16382) < 0.1 && Math.Abs(center.Y - 3.59155) < 0.1)
                            {
                                vp.On = isVertical;
                                found536 = true;
                                System.Diagnostics.Debug.WriteLine($"Set viewport at ({center.X}, {center.Y}) to {(isVertical ? "ON" : "OFF")} (as 536)");
                            }
                            // The bottom-right viewport (around position 11.06, 3.58)
                            else if (!found534 && Math.Abs(center.X - 11.0607) < 0.1 && Math.Abs(center.Y - 3.58284) < 0.1)
                            {
                                vp.On = isVertical;
                                found534 = true;
                                System.Diagnostics.Debug.WriteLine($"Set viewport at ({center.X}, {center.Y}) to {(isVertical ? "ON" : "OFF")} (as 534)");
                            }
                        }
                    }
                }
            }

            // Report if we still couldn't find our specific viewports
            if (!found534 || !found536)
            {
                System.Diagnostics.Debug.WriteLine($"Warning: Could not find all target viewports. Missing: {(!found534 ? "534 " : "")}{(!found536 ? "536" : "")}");
            }
        }


        /// <summary>
        /// Copies a base template to create other Fab templates
        /// </summary>
        private bool CopyBaseTemplate(string partNumber, string fabNumber)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            string basePath = Path.Combine(DiesPath, $"{partNumber}-1.dwg");
            string targetPath = Path.Combine(DiesPath, $"{partNumber}-{fabNumber}.dwg");

            try
            {
                // Simple file copy
                File.Copy(basePath, targetPath, true);
                return true;
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError copying template: {ex.Message}");
                return false;
            }
        }

        #endregion

        #region Fabrication Ticket Generation

        /// <summary>
        /// Generates fabrication tickets for all parts
        /// </summary>
        private void GenerateFabricationTickets(List<PartInfo> parts)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            // Group parts by part number and fab
            var partGroups = parts.GroupBy(p => new { p.PartNumber, p.Fab, p.Finish });

            foreach (var group in partGroups)
            {
                string partNo = group.Key.PartNumber;
                string fab = group.Key.Fab;
                string finish = group.Key.Finish;

                ed.WriteMessage($"\nProcessing part group: {partNo}-{fab} ({finish})");

                // Get template path
                string templatePath = Path.Combine(DiesPath, $"{partNo}-{fab}.dwg");
                if (!File.Exists(templatePath))
                {
                    ed.WriteMessage($"\nError: Template not found for {partNo}-{fab}");
                    continue;
                }

                // Process parts in chunks of MAX_ROWS_PER_TICKET
                int pageNo = 1;
                List<PartInfo> currentPage = new List<PartInfo>();

                foreach (var part in group)
                {
                    // If we're at row limit or the part has attachments (one part per sheet for attachments)
                    if (currentPage.Count >= MaxRowsPerTicket || part.Attachments.Count > 0)
                    {
                        // If we have parts in the current page, create a ticket
                        if (currentPage.Count > 0)
                        {
                            string ticketFile = Path.Combine(FabsPath, $"{partNo}_{pageNo:D2}.dwg");
                            CreateFabricationTicket(templatePath, ticketFile, currentPage, pageNo);
                            pageNo++;
                            currentPage.Clear();
                        }

                        // If this part has attachments, process it individually
                        if (part.Attachments.Count > 0)
                        {
                            string ticketFile = Path.Combine(FabsPath, $"{partNo}_{pageNo:D2}.dwg");
                            CreateFabricationTicket(templatePath, ticketFile, new List<PartInfo> { part }, pageNo);
                            pageNo++;
                        }
                        else
                        {
                            // Otherwise add it to a new page
                            currentPage.Add(part);
                        }
                    }
                    else
                    {
                        // Add to current page
                        currentPage.Add(part);
                    }
                }

                // Process remaining parts if any
                if (currentPage.Count > 0)
                {
                    string ticketFile = Path.Combine(FabsPath, $"{partNo}_{pageNo:D2}.dwg");
                    CreateFabricationTicket(templatePath, ticketFile, currentPage, pageNo);
                }
            }
        }

        /// <summary>
        /// Creates a fabrication ticket for a group of parts using the database-only approach
        /// </summary>
        private void CreateFabricationTicket(string templatePath, string outputPath, List<PartInfo> parts, int pageNo)
        {
            if (parts.Count == 0) return;
            Document currentDoc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = currentDoc.Editor;
            ed.WriteMessage($"\nCreating fabrication ticket: {outputPath}");

            // Determine if this is a vertical part (has attachments)
            // Using the representative part (first in the list)
            var representativePart = parts.First();
            bool isVertical = representativePart.Attachments != null && representativePart.Attachments.Count > 0;

            try
            {
                // Work directly with databases instead of documents
                using (Database db = new Database(false, true))
                {
                    // Read the template file
                    db.ReadDwgFile(templatePath, FileOpenMode.OpenForReadAndAllShare, true, "");
                    // Modify the database
                    using (Transaction tr = db.TransactionManager.StartTransaction())
                    {
                        // Get the layout dictionary
                        DBDictionary layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);
                        // Find the first layout (paper space) - typically it's named "Layout1"
                        // You can modify this to target a specific layout by name if needed
                        Layout layout = null;
                        // Try to find "Layout1" first
                        if (layoutDict.Contains("Layout1"))
                        {
                            layout = (Layout)tr.GetObject(layoutDict.GetAt("Layout1"), OpenMode.ForRead);
                        }
                        else
                        {
                            // Otherwise, get the first layout that's not "Model"
                            foreach (DBDictionaryEntry entry in layoutDict)
                            {
                                if (entry.Key != "Model")
                                {
                                    layout = (Layout)tr.GetObject(entry.Value, OpenMode.ForRead);
                                    break;
                                }
                            }
                        }
                        if (layout == null)
                        {
                            throw new System.Exception("No paper space layout found in template");
                        }
                        // Get the block table record for this layout
                        BlockTableRecord paperSpace = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForWrite);

                        // Add header information to paper space
                        AddHeaderText(paperSpace, tr, representativePart, pageNo, db);

                        // Add parts to the table in paper space - now with isVertical parameter
                        AddPartsToTable(paperSpace, tr, parts, db, isVertical);

                        // Set layer visibility based on vertical status
                        SetLayerVisibility(db, tr, isVertical);

                        // Set viewport visibility based on vertical status
                        SetViewportVisibility(db, tr, isVertical);

                        // Process attachments if this is a vertical part
                        if (isVertical)
                        {
                            Point3d? templateWorkPoint = WorkPointManager.GetWorkPoint(db, tr);
                            if (templateWorkPoint.HasValue)
                            {
                                ed.WriteMessage($"\nTemplate workpoint found: ({templateWorkPoint.Value.X}, {templateWorkPoint.Value.Y}, {templateWorkPoint.Value.Z})");
                            }
                            // Process each part's attachments
                            foreach (var part in parts)
                            {
                                if (part.Attachments != null && part.Attachments.Count > 0)
                                {
                                    ProcessAttachmentsForTicket(part, db, tr);
                                }
                            }
                        }

                        // Commit changes
                        tr.Commit();
                    }
                    // Save directly to the output path
                    db.SaveAs(outputPath, DwgVersion.Current);
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError creating fabrication ticket: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Adds header text to the fabrication ticket
        /// </summary>
        private void AddHeaderText(BlockTableRecord paperSpace, Transaction tr, PartInfo part, int pageNo, Database db)
        {
            try
            {
                // Determine if this is a vertical part drawing
                bool isVertical = part.IsVertical;

                // Adjust text Y position based on vertical flag
                double pty = isVertical ? 1.799 : PartNoY;
                double jobY = isVertical ? 1.15961 : JobY;
                double descriptionY = isVertical ? 1.497 : DescriptionY;

                // Add part number
                AddDText(paperSpace, tr, part.PartNumber.ToUpper(), PartNoX, pty, DefaultTextHeight, TextHorizontalMode.TextLeft, db);

                // Add fab number
                AddDText(paperSpace, tr, part.Fab.ToUpper(), FabX, pty, DefaultTextHeight, TextHorizontalMode.TextLeft, db);

                // Add page number
                AddDText(paperSpace, tr, pageNo.ToString(), PageNoX, pty, DefaultTextHeight, TextHorizontalMode.TextLeft, db);

                // Add finish
                AddDText(paperSpace, tr, part.Finish.ToUpper(), FinishX, pty, DefaultTextHeight, TextHorizontalMode.TextLeft, db);

                // Add description (if available)
                string description = GetPartDescription(part.PartNumber);
                if (!string.IsNullOrEmpty(description))
                {
                    AddDText(paperSpace, tr, description.ToUpper(), DescriptionX, descriptionY, DefaultTextHeight, TextHorizontalMode.TextLeft, db);
                }

                // Add job number
                string jobNumber = GetJobNumber();
                AddDText(paperSpace, tr, jobNumber, JobX, jobY, DefaultTextHeight, TextHorizontalMode.TextLeft, db);
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error adding header text: {ex.Message}");
            }
        }

        /// <summary>
        /// Adds a single-line text entity to the drawing
        /// </summary>
        private void AddDText(BlockTableRecord paperSpace, Transaction tr, string text, double x, double y, double height, TextHorizontalMode horizontalMode, Database db)
        {
            // Create a new text object
            DBText textObj = new DBText();
            textObj.SetDatabaseDefaults(db);

            // Set properties
            textObj.Position = new Point3d(x, y, 0);
            textObj.Height = height;
            textObj.TextString = text;
            textObj.HorizontalMode = horizontalMode;

            // Apply horizontal justification
            if (horizontalMode != TextHorizontalMode.TextLeft)
            {
                textObj.AdjustAlignment(db);
            }

            // Add to modelspace
            paperSpace.AppendEntity(textObj);
            tr.AddNewlyCreatedDBObject(textObj, true);
        }

        /// <summary>
        /// Gets the job number from the current drawing
        /// </summary>
        private string GetJobNumber()
        {
            // In a real implementation, this would extract job information from drawing
            // For now, return a placeholder
            return "JOB#12345";
        }

        /// <summary>
        /// Gets the part description from part database
        /// </summary>
        private string GetPartDescription(string partNumber)
        {
            // In a real implementation, this would lookup the description in a database
            // For now, return a generic description
            return $"{partNumber} PART";
        }

        /// <summary>
        /// Adds parts information to the ticket table, with special handling for vertical tickets
        /// </summary>
        private void AddPartsToTable(BlockTableRecord paperSpace, Transaction tr, List<PartInfo> parts, Database db, bool isVertical)
        {
            // For vertical tickets, adjust the starting position and max rows
            double startY = isVertical ? TableStartY - (TableRowHeight * (MaxRowsPerTicket)) : TableStartY;
            int maxRows = isVertical ? 1 : MaxRowsPerTicket;
            double currentY = startY;

            // Group parts by mark number and sum quantities
            var groupedParts = parts
                .GroupBy(p => new {
                    p.MarkNumber,
                    p.Length,
                    p.IsShopUse
                    // Add any other properties that should be identical for grouping
                })
                .Select(group => new {
                    MarkNumber = group.Key.MarkNumber,
                    Length = group.Key.Length,
                    IsShopUse = group.Key.IsShopUse,
                    Quantity = group.Count(), // Count of identical parts
                                              // You could also sum other numeric properties here if needed
                                              // Example: TotalWeight = group.Sum(p => p.Weight)


                })
                .ToList();

            // Sort the grouped parts if needed
            // groupedParts = groupedParts.OrderBy(p => p.MarkNumber).ToList();

            int row = 1;
            foreach (var part in groupedParts)
            {
                // Don't exceed max rows per ticket
                if (row > maxRows)
                    break;

                // Mark number
                AddDText(paperSpace, tr, part.MarkNumber, TableStartX + 0.3, currentY, DefaultTextHeight, TextHorizontalMode.TextLeft, db);

                // Length
                string formattedLength = FormatInches(part.Length);
                AddDText(paperSpace, tr, formattedLength, TableStartX + 0.98, currentY, DefaultTextHeight, TextHorizontalMode.TextLeft, db);

                // Quantity - now using the count from grouping
                AddDText(paperSpace, tr, part.Quantity.ToString(), TableStartX + 2.214, currentY, DefaultTextHeight, TextHorizontalMode.TextLeft, db);



                // Shop use indicator
                if (part.IsShopUse)
                {
                    AddDText(paperSpace, tr, "SHOP USE", TableStartX + 2.214 + 1.39625, currentY, DefaultTextHeight, TextHorizontalMode.TextLeft, db);
                }
                else
                {
                    // Distribution information would go here
                    AddDText(paperSpace, tr, "FIELD", TableStartX + 2.214 + 1.39625, currentY, DefaultTextHeight, TextHorizontalMode.TextLeft, db);
                }

                // Move to next row
                currentY -= TableRowHeight;
                row++;
            }
        }

        /// <summary>
        /// Formats a length value into fractional inches without converting to feet
        /// </summary>
        private string FormatInches(double inches)
        {
            // Round to nearest 1/16 inch
            int sixteenths = (int)Math.Round(inches * 16);
            int wholeInches = sixteenths / 16;
            int numerator = sixteenths % 16;

            // Simplify fraction
            int gcd = GCD(numerator, 16);
            if (gcd > 0) // Avoid division by zero
            {
                numerator /= gcd;
                int denominator = 16 / gcd;

                // Format the result
                if (numerator == 0)
                    return $"{wholeInches}\"";
                else
                    return $"{wholeInches} {numerator}/{denominator}\"";
            }
            else
            {
                // No fraction part
                return $"{wholeInches}\"";
            }
        }

        /// <summary>
        /// Formats a length value into a readable format (feet-inches or decimal)
        /// </summary>
        private string FormatLength(double length)
        {
            // Format as feet and inches with fraction
            int feet = (int)(length / 12);
            double inches = length % 12;

            // Round to nearest 1/16 inch
            int sixteenths = (int)Math.Round(inches * 16) / 12;
            int whole = (int)inches;
            int numerator = sixteenths % 16;

            // Simplify fraction
            int gcd = GCD(numerator, 16);
            numerator /= gcd;
            int denominator = 16 / gcd;

            // Format the result
            if (feet > 0)
            {
                if (numerator == 0)
                    return $"{feet}'-{whole}\"";
                else
                    return $"{feet}'-{whole} {numerator}/{denominator}\"";
            }
            else
            {
                if (numerator == 0)
                    return $"{whole}\"";
                else
                    return $"{whole} {numerator}/{denominator}\"";
            }
        }

        /// <summary>
        /// Calculates the greatest common divisor (for fraction simplification)
        /// </summary>
        private int GCD(int a, int b)
        {
            while (b != 0)
            {
                int temp = b;
                b = a % b;
                a = temp;
            }
            return a;
        }



        #endregion

        #region Attachment Processing

        /// <summary>
        /// Processes attachments for a ticket during the initial fabrication ticket creation
        /// </summary>
        private void ProcessAttachmentsForTicket(PartInfo part, Database db, Transaction tr)
        {
            try
            {
                // Get model space for 3D parts
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord modelSpace = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                // Find paper space layout
                DBDictionary layoutDict = (DBDictionary)tr.GetObject(db.LayoutDictionaryId, OpenMode.ForRead);

                // Find the first layout (paper space) that's not Model
                Layout layout = null;
                foreach (DBDictionaryEntry entry in layoutDict)
                {
                    if (entry.Key != "Model")
                    {
                        layout = (Layout)tr.GetObject(entry.Value, OpenMode.ForRead);
                        break;
                    }
                }

                if (layout == null)
                {
                    throw new System.Exception("No paper space layout found in ticket");
                }

                // Get the block table record for this layout
                BlockTableRecord paperSpace = (BlockTableRecord)tr.GetObject(layout.BlockTableRecordId, OpenMode.ForWrite);

                // Process each attachment
                foreach (var attachment in part.Attachments)
                {
                    // Determine attachment drawing path
                    string attachmentPartPath = Path.Combine(ApPath, $"{attachment.AttachedPartNumber}.dwg");
                    if (!File.Exists(attachmentPartPath))
                    {
                        System.Diagnostics.Debug.WriteLine($"Attachment part file not found: {attachmentPartPath}");
                        continue;
                    }

                    // Calculate position (workpoint in db will be origin, we use transforms for positioning)
                    Point3d insertionPoint = Point3d.Origin;

                    // Store the workpoint in the database (where the parts will be attached)
                    WorkPointManager.StoreWorkPoint(db, tr, insertionPoint);

                    // Insert the attachment part
                    InsertAttachmentPart(
                        db,
                        tr,
                        modelSpace,
                        attachmentPartPath,
                        attachment.AttachedPartNumber,
                        insertionPoint,
                        attachment.Side,
                        attachment.Invert,
                        attachment.Position,
                        attachment.Height,
                        part.Length);
                }

                // Add attachment information to attachment table (in paper space)
                AddAttachmentTableInfo(paperSpace, tr, part.Attachments, db);
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error processing attachments for ticket: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Inserts an attachment part drawing
        /// </summary>
        private void InsertAttachmentPart(
            Database db,
            Transaction tr,
            BlockTableRecord modelSpace,
            string partPath,
            string partNumber,
            Point3d insertionPoint,
            string side,
            bool invert,
            double position,
            double height,
            double mainPartLength)
        {
            try
            {
                // Get mullion data from the database
                MullionPlacementData calculatedData = CalculateMullionDimensions(db, tr);
                MullionPlacementData storedData = MullionPlacementData.GetMullionData(db, tr);
                // Use calculated dimensions if available, otherwise fall back to stored or default values
                double mullionWidth = calculatedData.Width > 0 ? calculatedData.Width :
                              (storedData != null ? storedData.Width : 2.5);

                double overallWidth = calculatedData.OverallWidth > 0 ? calculatedData.OverallWidth :
                                      (storedData != null ? storedData.OverallWidth : 4.5);

                // For glass pocket offset, prefer stored value as it can't be reliably calculated
                double glassPocketOffset = storedData != null ? storedData.GlassPocketOffset : 1.25;

                System.Diagnostics.Debug.WriteLine($"Using mullion data: Width = {mullionWidth}, Glass Pocket Offset = {glassPocketOffset}, Overall Width = {overallWidth}");

                // Create Drawfab instance for this part
                Drawfab drawfab = new Drawfab(partNumber);

                // Determine handed parameters based on side and invert
                bool handed = (side == "L" || side == "R");
                string handedSide = "R";

                // Calculate the scaled position along the vertical
                double templateLength = 30.0;
                double relativePosition = (height / mainPartLength) * templateLength;

                // Special handling for parts within 3 inches of the bottom
                if (height <= 3.0)
                {
                    // Position at the bottom of the vertical
                    relativePosition = mullionWidth;
                }

                // Special handling for parts within 3 inches of the top
                if (height >= (mainPartLength - 3.0))
                {
                    // Position at the top of the vertical
                    relativePosition = templateLength;
                }

                // Calculate offsets based on side and mullion data
                double yOffset = overallWidth - glassPocketOffset; // Same Y offset for both sides
                double zOffset = 0;

                // Z offset depends on side
                if (side == "R")
                {
                    zOffset = -mullionWidth / 2; // Right side
                }
                else if (side == "L")
                {
                    zOffset = mullionWidth / 2; // Left side
                }

                // For vertical orientation, X is position along vertical, Y is glass pocket offset, Z is width offset
                Matrix3d transform = Matrix3d.Displacement(
                    new Vector3d(-relativePosition, yOffset, zOffset));

                // Create a matrix for potential mirroring
                Matrix3d mirrorTransform = Matrix3d.Identity;

                // If it's a left-side part, create a mirror transformation
                if (side == "L")
                {
                    // Create a mirror line that runs along the X axis through the center
                    // of the part (at 0,0,0 in local coordinates)
                    Point3d mirrorBase = new Point3d(0, 0, mullionWidth / 2);
                    Vector3d mirrorNormal = Vector3d.ZAxis; // Mirror across the XZ plane
                    mirrorTransform = Matrix3d.Mirroring(new Plane(mirrorBase, mirrorNormal));

                    // Apply mirror transform to our positioning transform
                    transform = transform.PreMultiplyBy(mirrorTransform);

                    System.Diagnostics.Debug.WriteLine("Applied mirror transformation for left side");
                }

                // Calculate the middle of the part after initial positioning
                // This would be after your initial transform calculation but before applying inversion
                Point3d partMiddle = new Point3d(-relativePosition, yOffset, zOffset);

                // After your existing transform and mirror code, before creating the part
                // Apply inversion if requested (mirror around Y axis)
                if (invert)
                {
                    // Create a mirror line that runs along the Y axis through the center
                    // of the part in world coordinate space
                    Point3d mirrorBase = new Point3d(1, 0, 0);
                    Vector3d mirrorNormal = Vector3d.XAxis; // Mirror across the YZ plane (around Y axis)
                    Matrix3d invertTransform = Matrix3d.Mirroring(new Plane(mirrorBase, mirrorNormal));

                    // Apply invert transform to our positioning transform
                    transform = transform.PreMultiplyBy(invertTransform);

                    System.Diagnostics.Debug.WriteLine("Applied inversion transformation");
                }

                // Create the attachment part
                drawfab.CreateExtrudedPart(
                    db,                // Use the current database
                    partPath,          // Path to the part drawing
                    partNumber,        // Block name
                    0,                 // Length 0 for attachments
                    90, 90, 90, 90,    // Standard angles
                    handed,            // Whether it's a handed part
                    handedSide,        // Which side (L/R)
                    transform,         // Position transformation
                    false,             // preserveOriginalOrientation
                    false,             // visualizeOnly
                    false,             // no dimensions for attachments
                    true);             // isVertical = true for attachments
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error inserting attached part: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Adds attachment information to the attachment table in the ticket
        /// </summary>
        private void AddAttachmentTableInfo(BlockTableRecord paperSpace, Transaction tr, List<AttachmentInfo> attachments, Database db)
        {
            // Based on VBA code, attachment table is at the top with:
            // Column 1: Part Number
            // Column 2: Fab
            // Column 3: Side
            // Column 4: Position
            // Column 5: Height
            // Column 6: Invert Flag

            // Table starting position from VBA code
            double tableX = 0.2086; // vInsNew1
            double tableY = 11.4687;
            double rowHeight = 0.21479;

            // Add each attachment to the table
            for (int i = 0; i < attachments.Count; i++)
            {
                var attachment = attachments[i];
                double currentY = tableY - (i * rowHeight);

                // Part number
                AddDText(paperSpace, tr, attachment.AttachedPartNumber, tableX, currentY, DefaultTextHeight, TextHorizontalMode.TextLeft, db);

                // Side
                AddDText(paperSpace, tr, attachment.Side, tableX + 1.25, currentY, DefaultTextHeight, TextHorizontalMode.TextLeft, db);

                // Position
                string formattedPosition = FormatInches(attachment.Height);
                AddDText(paperSpace, tr, formattedPosition, tableX + 1.6875, currentY, DefaultTextHeight, TextHorizontalMode.TextLeft, db);

                // Height
                string fab = attachment.AttachedFab;
                AddDText(paperSpace, tr, fab, tableX + 2.5625, currentY, DefaultTextHeight, TextHorizontalMode.TextLeft, db);

                // Invert indicator
                if (attachment.Invert)
                {
                    AddDText(paperSpace, tr, "X", tableX + 3.625, currentY, DefaultTextHeight, TextHorizontalMode.TextLeft, db);
                }
            }
        }

        /// <summary>
        /// Calculates the actual dimensions of the vertical mullion from the model space
        /// </summary>
        private MullionPlacementData CalculateMullionDimensions(Database db, Transaction tr)
        {
            MullionPlacementData data = new MullionPlacementData();

            try
            {
                // Get the model space
                BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                BlockTableRecord ms = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                // Look for the extruded solids (the vertical mullion)
                Extents3d extents = new Extents3d();
                bool foundExtrusion = false;

                foreach (ObjectId id in ms)
                {
                    DBObject obj = tr.GetObject(id, OpenMode.ForRead);

                    // Check for 3D solids or regions
                    if (obj is Solid3d solid)
                    {
                        extents.AddExtents(solid.GeometricExtents);
                        foundExtrusion = true;
                    }
                    else if (obj is Region region)
                    {
                        extents.AddExtents(region.GeometricExtents);
                        foundExtrusion = true;
                    }
                }

                if (!foundExtrusion)
                {
                    // No extrusions found, check for named block references (the vertical mullion)
                    foreach (ObjectId id in ms)
                    {
                        DBObject obj = tr.GetObject(id, OpenMode.ForRead);
                        if (obj is BlockReference blockRef)
                        {
                            // Check if this is our vertical mullion block
                            BlockTableRecord btr = (BlockTableRecord)tr.GetObject(blockRef.BlockTableRecord, OpenMode.ForRead);
                            if (btr.Name.Contains("VG") || btr.Name.Contains("VM")) // Adjust the pattern as needed
                            {
                                extents.AddExtents(blockRef.GeometricExtents);
                                foundExtrusion = true;
                            }
                        }
                    }
                }

                if (foundExtrusion)
                {
                    // Calculate width (Z dimension in world coords)
                    data.Width = Math.Abs(extents.MaxPoint.Z - extents.MinPoint.Z);

                    // Calculate overall width (Y dimension in world coords)
                    data.OverallWidth = Math.Abs(extents.MaxPoint.Y - extents.MinPoint.Y);

                    // For glass pocket offset, we'll use our defaults or any stored data
                    // since we can't directly determine it from the geometry
                    MullionPlacementData storedData = MullionPlacementData.GetMullionData(db, tr);
                    if (storedData != null)
                    {
                        data.GlassPocketOffset = storedData.GlassPocketOffset;
                    }

                    System.Diagnostics.Debug.WriteLine($"Calculated mullion dimensions: Width={data.Width}, OverallWidth={data.OverallWidth}, GlassPocketOffset={data.GlassPocketOffset}");
                }
                else
                {
                    System.Diagnostics.Debug.WriteLine("No vertical mullion found in model space, using default dimensions");
                }
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error calculating mullion dimensions: {ex.Message}");
            }

            return data;
        }

        #endregion

    }



    /// <summary>
    /// Extension methods for AutoCAD entities
    /// </summary>
    public static class Extensions
    {
        /// <summary>
        /// Sets up the alignment for the text object
        /// </summary>
        public static void AdjustAlignment(this DBText text, Database db)
        {
            switch (text.HorizontalMode)
            {
                case TextHorizontalMode.TextCenter:
                    text.AlignmentPoint = new Point3d(text.Position.X, text.Position.Y, 0);
                    break;

                case TextHorizontalMode.TextRight:
                    text.AlignmentPoint = new Point3d(text.Position.X, text.Position.Y, 0);
                    break;

                default:
                    // No adjustment needed for left alignment
                    break;
            }
        }
    }
}