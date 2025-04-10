using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.DatabaseServices.Filters;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;

namespace TakeoffBridge
{
    public class TemplateGenerator
    {
        // Class to store information needed for templates
        public class PartTemplateInfo
        {
            public string PartNumber { get; set; }
            public string PartType { get; set; }
            public string Fab { get; set; }
            public string Finish { get; set; }
            public bool IsVertical { get; set; }
            public List<AttachmentInfo> Attachments { get; set; } = new List<AttachmentInfo>();

            // For tracking uniqueness
            public override bool Equals(object obj)
            {
                if (obj is PartTemplateInfo other)
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
        }

        // Using the ChildPart class structure from your code
        public class ChildPart
        {
            public string Name { get; set; }
            public string PartType { get; set; }
            public double LengthAdjustment { get; set; }
            public bool IsShopUse { get; set; }
            public double StartAdjustment { get; set; }
            public double EndAdjustment { get; set; }
            public bool IsFixedLength { get; set; }
            public double FixedLength { get; set; }
            public string MarkNumber { get; set; }
            public string Material { get; set; }
            public string Attach { get; set; }
            public bool Invert { get; set; }
            public double Adjust { get; set; }
            public bool Clips { get; set; }
            public string Finish { get; set; }
            public string Fab { get; set; }
        }

        public class Attachment
        {
            // Change these from Handle type to string
            public string HorizontalHandle { get; set; }
            public string VerticalHandle { get; set; }

            // Keep the rest as they are
            public string HorizontalPartType { get; set; }
            public string VerticalPartType { get; set; }
            public string Side { get; set; }
            public double Position { get; set; }
            public double Height { get; set; }
            public bool Invert { get; set; }
            public double Adjust { get; set; }
        }

        private readonly string templatesPath = @"C:\CSE\Takeoff\fabs\dies";
        private readonly string baseTemplatePath;

        public TemplateGenerator()
        {
            baseTemplatePath = Path.Combine(templatesPath, "fabtemplate.dwg");
        }

        public void GenerateTemplatesFromDrawing()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor editor = doc.Editor;

            // Store the original document
            Document originalDoc = Application.DocumentManager.MdiActiveDocument;

            // Dictionary to store unique part-fab combinations
            Dictionary<string, PartTemplateInfo> partTemplates = new Dictionary<string, PartTemplateInfo>();

            // Load attachments
            List<Attachment> allAttachments = LoadAttachmentsFromDrawing();
            editor.WriteMessage($"\nLoaded {allAttachments.Count} attachments from drawing");

            // First, collect all components and store their data
            Dictionary<string, List<ChildPart>> partsByComponentHandle = new Dictionary<string, List<ChildPart>>();
            Dictionary<string, string> componentTypesByHandle = new Dictionary<string, string>();

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                // Get all polylines in the drawing
                TypedValue[] filterList = new TypedValue[] {
                    new TypedValue((int)DxfCode.Start, "LWPOLYLINE")
                };

                SelectionFilter filter = new SelectionFilter(filterList);
                PromptSelectionResult selRes = editor.SelectAll(filter);

                if (selRes.Status == PromptStatus.OK)
                {
                    foreach (ObjectId id in selRes.Value.GetObjectIds())
                    {
                        Entity ent = (Entity)tr.GetObject(id, OpenMode.ForRead);

                        // Check if this is a metal component
                        string componentType = GetComponentType(ent);
                        if (componentType != null)
                        {
                            // Store component type
                            componentTypesByHandle[ent.Handle.ToString()] = componentType;

                            // Get child parts and store them
                            List<ChildPart> childParts = GetChildParts(ent);
                            partsByComponentHandle[ent.Handle.ToString()] = childParts;

                            // Process each part for templates
                            foreach (ChildPart part in childParts)
                            {
                                // Skip parts with no fab or no part type
                                if (string.IsNullOrEmpty(part.Fab) || string.IsNullOrEmpty(part.PartType) || string.IsNullOrEmpty(part.Name))
                                    continue;

                                // Create unique key for part-fab combination
                                string key = $"{part.Name}-{part.Fab}";

                                if (!partTemplates.ContainsKey(key))
                                {
                                    partTemplates[key] = new PartTemplateInfo
                                    {
                                        PartNumber = part.Name,
                                        PartType = part.PartType,
                                        Fab = part.Fab,
                                        Finish = part.Finish,
                                        IsVertical = componentType == "Vertical"
                                    };

                                    editor.WriteMessage($"\nAdded template for {key}, IsVertical={partTemplates[key].IsVertical}");
                                }
                            }
                        }
                    }
                }

                tr.Commit();
            }

            // Now process attachments to associate them with templates
            foreach (var attachment in allAttachments)
            {
                string vertHandle = attachment.VerticalHandle;

                // Check if we have data for this vertical component
                if (partsByComponentHandle.ContainsKey(vertHandle))
                {
                    List<ChildPart> verticalParts = partsByComponentHandle[vertHandle];

                    // Find the part matching the vertical part type
                    foreach (ChildPart part in verticalParts)
                    {
                        if (part.PartType == attachment.VerticalPartType)
                        {
                            // Create key for template
                            string key = $"{part.Name}-{part.Fab}";

                            // Make sure template exists
                            if (!partTemplates.ContainsKey(key))
                            {
                                partTemplates[key] = new PartTemplateInfo
                                {
                                    PartNumber = part.Name,
                                    PartType = part.PartType,
                                    Fab = part.Fab,
                                    Finish = part.Finish,
                                    IsVertical = true
                                };
                            }

                            // Add attachment to template
                            partTemplates[key].Attachments.Add(new AttachmentInfo
                            {
                                Side = attachment.Side,
                                Position = attachment.Position,
                                Height = attachment.Height,
                                Invert = attachment.Invert
                            });

                            editor.WriteMessage($"\nAdded attachment to template {key}: Side={attachment.Side}, Height={attachment.Height}");
                        }
                    }
                }
            }

            // Generate templates for each unique part-fab combination
            int generatedCount = 0;
            int skippedCount = 0;
            int copiedCount = 0;

            // Create a list of all fab numbers in the drawing
            HashSet<string> allFabs = new HashSet<string>();
            foreach (var template in partTemplates.Values)
            {
                allFabs.Add(template.Fab);
            }

            // Process each unique part number
            var uniquePartNumbers = partTemplates.Values
                .Select(t => t.PartNumber)
                .Distinct()
                .ToList();

            foreach (var partNumber in uniquePartNumbers)
            {
                // Generate base template (Fab 1) if it doesn't exist
                string baseFabKey = $"{partNumber}-1";
                bool baseTemplateExists = false;

                if (partTemplates.ContainsKey(baseFabKey))
                {
                    // We have a Fab 1 in the drawing, generate it if needed
                    var templateInfo = partTemplates[baseFabKey];
                    string outputPath = Path.Combine(templatesPath, $"{partNumber}-1.dwg");

                    if (File.Exists(outputPath))
                    {
                        editor.WriteMessage($"\nBase template already exists: {outputPath}");
                        baseTemplateExists = true;
                        skippedCount++;
                    }
                    else
                    {
                        // Generate base template
                        if (GenerateBaseTemplate(templateInfo, editor))
                        {
                            generatedCount++;
                            baseTemplateExists = true;
                            editor.WriteMessage($"\nGenerated base template: {outputPath}");
                        }
                        else
                        {
                            editor.WriteMessage($"\nFailed to generate base template: {outputPath}");
                        }
                    }
                }
                else
                {
                    // Create a default Fab 1 template if any other fab exists for this part
                    var partFabs = partTemplates.Values
                        .Where(t => t.PartNumber == partNumber)
                        .Select(t => t.Fab)
                        .ToList();

                    if (partFabs.Count > 0)
                    {
                        var defaultTemplate = new PartTemplateInfo
                        {
                            PartNumber = partNumber,
                            Fab = "1",
                            IsVertical = partTemplates.Values
                                .First(t => t.PartNumber == partNumber)
                                .IsVertical
                        };

                        string outputPath = Path.Combine(templatesPath, $"{partNumber}-1.dwg");

                        if (File.Exists(outputPath))
                        {
                            editor.WriteMessage($"\nBase template already exists: {outputPath}");
                            baseTemplateExists = true;
                            skippedCount++;
                        }
                        else
                        {
                            // Generate default base template
                            if (GenerateBaseTemplate(defaultTemplate, editor))
                            {
                                generatedCount++;
                                baseTemplateExists = true;
                                editor.WriteMessage($"\nGenerated default base template: {outputPath}");
                            }
                            else
                            {
                                editor.WriteMessage($"\nFailed to generate default base template: {outputPath}");
                            }
                        }
                    }
                }

                // If we have a base template, generate copies for other fabs
                if (baseTemplateExists)
                {
                    foreach (string fab in allFabs)
                    {
                        if (fab == "1") continue; // Skip fab 1, already handled

                        string fabPath = Path.Combine(templatesPath, $"{partNumber}-{fab}.dwg");

                        if (File.Exists(fabPath))
                        {
                            editor.WriteMessage($"\nFab template already exists: {fabPath}");
                            skippedCount++;
                        }
                        else
                        {
                            // Copy the base template
                            if (CopyBaseTemplate(partNumber, fab, editor))
                            {
                                copiedCount++;
                                editor.WriteMessage($"\nCopied base template to create: {fabPath}");
                            }
                            else
                            {
                                editor.WriteMessage($"\nFailed to copy base template for: {fabPath}");
                            }
                        }
                    }
                }
            }

            editor.WriteMessage($"\nGenerated {generatedCount} base templates, copied {copiedCount} additional templates.");
            editor.WriteMessage($"\nSkipped {skippedCount} existing templates.");
            editor.WriteMessage($"\nTotal unique part-fab combinations: {partTemplates.Count}");
        }

        private bool GenerateBaseTemplate(PartTemplateInfo templateInfo, Editor editor)
        {
            string outputPath = Path.Combine(templatesPath, $"{templateInfo.PartNumber}-{templateInfo.Fab}.dwg");
            Document currentDoc = Application.DocumentManager.MdiActiveDocument;

            try
            {
                // 1. Copy the base template
                string baseTemplatePath = Path.Combine(templatesPath, "fabtemplate.dwg");
                if (!File.Exists(baseTemplatePath))
                {
                    editor.WriteMessage($"\nError: Base template not found: {baseTemplatePath}");
                    return false;
                }
                File.Copy(baseTemplatePath, outputPath, true);

                // 2. Open the new document
                using (Document templateDoc = Application.DocumentManager.Open(outputPath, false))
                using (templateDoc.LockDocument()) // Lock to prevent conflicts
                {
                    // Switch to the new document
                    Application.DocumentManager.MdiActiveDocument = templateDoc;

                    // Force AutoCAD to process the switch (critical!)
                    //Application.DocumentManager.ExecuteInCommandContextAsync(new Action<object>(o => { }), null).Wait();

                    // 3. Execute the LISP command
                    string escapedPath = templateInfo.PartNumber.Replace("\\", "\\\\");
                    int mirrorType = templateInfo.IsVertical ? 1 : 0;
                    string drwprtCmd = $"(drwprt \"{escapedPath}\" 30.0 45 90 45 90 {mirrorType} nil nil 0)\n";

                    if (!ExecuteLispWithWait(drwprtCmd, "DRWPRT", 30000))
                    {
                        editor.WriteMessage("\nDRWPRT command timed out or failed.");
                        return false;
                    }

                    // 4. Save the document
                    templateDoc.Database.SaveAs(outputPath, DwgVersion.Current);

                    // 5. Switch back to the original document
                    //Application.DocumentManager.MdiActiveDocument = currentDoc;
                    ////Application.DocumentManager.ExecuteInCommandContextAsync(
                    ////    new Action<object>(o => { }),
                    ////    null
                    //).Wait();
                } // Document auto-closes when exiting the 'using' block

                return true;
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage($"\nError: {ex.Message}");
                // Ensure we revert to the original document on failure
                Application.DocumentManager.MdiActiveDocument = currentDoc;
                return false;
            }
        }

        // Helper method to execute LISP and wait for completion
        private bool ExecuteLispWithWait(string lispCode, string commandName, int timeout)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            using (ManualResetEvent waitHandle = new ManualResetEvent(false))
            {
                // Track command completion
                EventHandler<CommandEventArgs> handler = (s, e) =>
                {
                    if (e.GlobalCommandName.Equals(commandName, StringComparison.OrdinalIgnoreCase))
                    {
                        //Application.DocumentManager.DocumentLockModeChanged -= handler;
                        //waitHandle.Set();
                    }
                };

                //Application.DocumentManager.DocumentLockModeChanged += handler;
                doc.SendStringToExecute(lispCode, true, false, true);

                return waitHandle.WaitOne(timeout);
            }
        }


        private bool CopyBaseTemplate(string partNumber, string fabNumber, Editor editor)
        {
            string basePath = Path.Combine(templatesPath, $"{partNumber}-1.dwg");
            string outputPath = Path.Combine(templatesPath, $"{partNumber}-{fabNumber}.dwg");

            try
            {
                // Use standard Database clone approach
                using (Database baseDb = new Database(false, true))
                {
                    baseDb.ReadDwgFile(basePath, FileOpenMode.OpenForReadAndAllShare, false, null);

                    using (Database newDb = (Database)baseDb.Clone())
                    {
                        newDb.SaveAs(outputPath, DwgVersion.Current);
                        return true;
                    }
                }
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage($"\nError copying template {partNumber}-1 to {partNumber}-{fabNumber}: {ex.Message}");
                return false;
            }
        }

        // Helper method to load attachments from the drawing
        private List<Attachment> LoadAttachmentsFromDrawing()
        {
            List<Attachment> loadedAttachments = new List<Attachment>();
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
                                editor.WriteMessage($"\nJSON Value: {json}");
                                loadedAttachments = JsonConvert.DeserializeObject<List<Attachment>>(json);
                            }
                        }
                    }
                }
                else
                {
                    editor.WriteMessage("\nNo METALATTACHMENTS dictionary found in drawing");
                }
                tr.Commit();
            }
            return loadedAttachments;
        }

        // Your existing GetComponentType method
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

        // Your existing GetChildParts method
        private List<ChildPart> GetChildParts(Entity ent)
        {
            List<ChildPart> result = new List<ChildPart>();
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
                        result = JsonConvert.DeserializeObject<List<ChildPart>>(jsonBuilder.ToString());
                    }
                    catch (System.Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error parsing child parts JSON: {ex.Message}");
                    }
                }
            }
            return result;
        }
    }


}