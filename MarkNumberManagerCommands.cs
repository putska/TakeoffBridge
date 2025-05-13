using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using System.Text;
using Newtonsoft.Json;

namespace TakeoffBridge
{
    // This should be a separate class file
    public class MarkNumberManagerCommands
    {

        [CommandMethod("GENERATEMARKNUMBERS")]
        public void GenerateMarkNumbers()
        {
            // This calls the static method in MarkNumberManager
            MarkNumberManager.GenerateMarkNumbers();
        }


        [CommandMethod("MARKNUMBERREPORT")]
        public void GenerateMarkNumberReport()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            // Get the mark number manager
            var manager = GlassTakeoffBridge.GlassTakeoffApp.MarkNumberManager;

            // Generate a simple report
            ed.WriteMessage("\n--- Mark Number Report ---");

            // Get counts from the manager (you'll need to add these methods)
            // int horizontalCount = manager.GetHorizontalPartCount();
            // int verticalCount = manager.GetVerticalPartCount();

            ed.WriteMessage($"\nHorizontal Mark Count: {manager.GetHorizontalPartCount()}");
            ed.WriteMessage($"\nVertical Mark Count: {manager.GetVerticalPartCount()}");

            ed.WriteMessage("\n--- End of Report ---");
        }

        [CommandMethod("LISTMARKNUMBERS")]
        public void ListMarkNumbers()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            try
            {
                ed.WriteMessage("\n--- Mark Number Listing ---");

                // Get the MarkNumberManager instance
                MarkNumberManager manager = MarkNumberManager.Instance;

                // Reflect on the manager to access private dictionaries
                Type managerType = manager.GetType();

                // Get the horizontal mark dictionary
                var horizontalMarksField = managerType.GetField("_horizontalPartMarks",
                    System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);

                // Get the vertical mark dictionary  
                var verticalMarksField = managerType.GetField("_verticalPartMarks",
                    System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);

                // Get the next mark numbers dictionary
                var nextMarkNumbersField = managerType.GetField("_nextMarkNumbers",
                    System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);

                if (horizontalMarksField != null && verticalMarksField != null && nextMarkNumbersField != null)
                {
                    var horizontalMarks = horizontalMarksField.GetValue(manager) as Dictionary<string, string>;
                    var verticalMarks = verticalMarksField.GetValue(manager) as Dictionary<string, string>;
                    var nextMarkNumbers = nextMarkNumbersField.GetValue(manager) as Dictionary<string, int>;

                    // Display horizontal marks
                    ed.WriteMessage($"\n\nHorizontal Part Marks ({horizontalMarks.Count}):");
                    int i = 0;
                    foreach (var kvp in horizontalMarks)
                    {
                        if (i < 20) // Limit to 20 entries to avoid flooding the console
                        {
                            ed.WriteMessage($"\n  Key: {kvp.Key}");
                            ed.WriteMessage($"\n  Mark: {kvp.Value}");
                            ed.WriteMessage("\n  ------------------");
                        }
                        i++;
                    }

                    if (i > 20)
                    {
                        ed.WriteMessage($"\n  ... and {i - 20} more entries");
                    }

                    // Display vertical marks
                    ed.WriteMessage($"\n\nVertical Part Marks ({verticalMarks.Count}):");
                    i = 0;
                    foreach (var kvp in verticalMarks)
                    {
                        if (i < 20) // Limit to 20 entries
                        {
                            ed.WriteMessage($"\n  Key: {kvp.Key}");
                            ed.WriteMessage($"\n  Mark: {kvp.Value}");
                            ed.WriteMessage("\n  ------------------");
                        }
                        i++;
                    }

                    if (i > 20)
                    {
                        ed.WriteMessage($"\n  ... and {i - 20} more entries");
                    }

                    // Display next mark numbers
                    ed.WriteMessage($"\n\nNext Mark Numbers ({nextMarkNumbers.Count}):");
                    foreach (var kvp in nextMarkNumbers)
                    {
                        ed.WriteMessage($"\n  Part Type: {kvp.Key}, Next Number: {kvp.Value}");
                    }

                    ed.WriteMessage("\n\n--- End of Mark Number Listing ---");
                }
                else
                {
                    ed.WriteMessage("\nCould not access mark number dictionaries via reflection.");
                }

                // Now, let's scan the drawing to find components with mark numbers
                ed.WriteMessage("\n\n--- Scanning Drawing for Mark Numbers ---");

                using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
                {
                    int componentsWithMarks = 0;
                    int totalPartsWithMarks = 0;
                    Dictionary<string, int> partTypeCount = new Dictionary<string, int>();

                    // Get all polylines in the drawing
                    TypedValue[] tvs = new TypedValue[] {
                new TypedValue((int)DxfCode.Start, "LWPOLYLINE")
            };

                    SelectionFilter filter = new SelectionFilter(tvs);
                    PromptSelectionResult selRes = ed.SelectAll(filter);

                    if (selRes.Status == PromptStatus.OK)
                    {
                        ed.WriteMessage($"\nFound {selRes.Value.Count} polylines in drawing.");

                        foreach (SelectedObject selObj in selRes.Value)
                        {
                            ObjectId id = selObj.ObjectId;
                            Entity ent = tr.GetObject(id, OpenMode.ForRead) as Entity;

                            if (ent != null)
                            {
                                // Check if this is a metal component
                                ResultBuffer rbComp = ent.GetXDataForApplication("METALCOMP");
                                if (rbComp != null)
                                {
                                    string handle = ent.Handle.ToString();

                                    // Extract component type
                                    TypedValue[] xdataComp = rbComp.AsArray();
                                    string componentType = "";
                                    for (int i = 1; i < xdataComp.Length; i++) // Skip app name
                                    {
                                        if (i == 1)
                                        {
                                            componentType = xdataComp[i].Value.ToString();
                                            break;
                                        }
                                    }

                                    // Get parts
                                    List<ChildPart> parts = null;

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

                                        // Parse the JSON
                                        if (jsonBuilder.Length > 0)
                                        {
                                            try
                                            {
                                                parts = JsonConvert.DeserializeObject<List<ChildPart>>(jsonBuilder.ToString());
                                            }
                                            catch (System.Exception ex)
                                            {
                                                ed.WriteMessage($"\nError parsing parts JSON for {handle}: {ex.Message}");
                                            }
                                        }
                                    }

                                    if (parts != null && parts.Count > 0)
                                    {
                                        int partsWithMarks = 0;
                                        foreach (var part in parts)
                                        {
                                            if (!string.IsNullOrEmpty(part.MarkNumber))
                                            {
                                                partsWithMarks++;
                                                totalPartsWithMarks++;

                                                // Track by part type
                                                string key = part.PartType ?? "Unknown";
                                                if (!partTypeCount.ContainsKey(key))
                                                    partTypeCount[key] = 0;

                                                partTypeCount[key]++;
                                            }
                                        }

                                        if (partsWithMarks > 0)
                                        {
                                            componentsWithMarks++;
                                            ed.WriteMessage($"\nComponent {handle} ({componentType}) has {partsWithMarks}/{parts.Count} parts with mark numbers");

                                            // Show the first few parts with mark numbers
                                            int shown = 0;
                                            foreach (var part in parts)
                                            {
                                                if (!string.IsNullOrEmpty(part.MarkNumber) && shown < 3)
                                                {
                                                    ed.WriteMessage($"\n  Part: {part.Name}, Type: {part.PartType}, Mark: {part.MarkNumber}");
                                                    shown++;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }

                        ed.WriteMessage($"\n\nFound {totalPartsWithMarks} parts with mark numbers in {componentsWithMarks} components.");

                        // Show part type counts
                        ed.WriteMessage("\nMark numbers by part type:");
                        foreach (var kvp in partTypeCount)
                        {
                            ed.WriteMessage($"\n  {kvp.Key}: {kvp.Value} parts");
                        }
                    }

                    tr.Commit();
                }

                ed.WriteMessage("\n\n--- End of Scan ---");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError listing mark numbers: {ex.Message}");
                ed.WriteMessage($"\nStack trace: {ex.StackTrace}");
            }
        }

        [CommandMethod("MARKSELECTED")]
        public void MarkSelected()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            // Prompt for selection
            PromptEntityResult result = ed.GetEntity("\nSelect component to mark: ");
            if (result.Status != PromptStatus.OK)
                return;

            // Process the selected component
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                MarkNumberManager.Instance.ProcessComponentMarkNumbers(result.ObjectId, tr);
                tr.Commit();
            }

            ed.WriteMessage("\nComponent marked.");
        }

    }
}