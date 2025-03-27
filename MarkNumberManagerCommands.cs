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
            MarkNumberManager.GenerateAllMarkNumbers();
        }

        [CommandMethod("REGENERATEALLMARKNUMBERS")]
        public void RegenerateAllMarkNumbers()
        {
            // This calls the static method in MarkNumberManager
            MarkNumberManager.RegenerateAllMarkNumbers();
        }

        // Add this command class if it doesn't exist
        [CommandMethod("FORCEPROCESS")]
        public void ForceProcessCommand()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            // Prompt for component selection
            PromptEntityOptions options = new PromptEntityOptions("\nSelect component to force process: ");
            options.SetRejectMessage("\nMust select a polyline");
            options.AddAllowedClass(typeof(Polyline), false);

            PromptEntityResult result = ed.GetEntity(options);
            if (result.Status == PromptStatus.OK)
            {
                // Force process the selected component
                MarkNumberManager.ForceProcess(result.ObjectId);

                ed.WriteMessage("\nComponent processed. Check output window for details.");
            }
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

        [CommandMethod("CLEANUPMETALXDATA")]
        public void CleanupMetalXdata()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            ed.WriteMessage("\nCleaning up metal component XData...");

            try
            {
                // First register all apps we need
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    RegAppTable regTable = (RegAppTable)tr.GetObject(db.RegAppTableId, OpenMode.ForWrite);

                    // Register apps
                    RegisterApp(regTable, "METALCOMP", tr);
                    RegisterApp(regTable, "METALPARTSINFO", tr);

                    for (int i = 0; i < 20; i++)  // Register more than needed to be safe
                    {
                        RegisterApp(regTable, $"METALPARTS{i}", tr);
                    }

                    tr.Commit();
                }

                // Get all polylines in the drawing
                TypedValue[] tvs = new TypedValue[] {
            new TypedValue((int)DxfCode.Start, "LWPOLYLINE")
        };

                SelectionFilter filter = new SelectionFilter(tvs);
                PromptSelectionResult selRes = ed.SelectAll(filter);

                if (selRes.Status == PromptStatus.OK)
                {
                    SelectionSet ss = selRes.Value;
                    int cleanedCount = 0;

                    // Process each component in its own transaction
                    foreach (SelectedObject selObj in ss)
                    {
                        ObjectId id = selObj.ObjectId;

                        try
                        {
                            using (Transaction tr = db.TransactionManager.StartTransaction())
                            {
                                Entity ent = tr.GetObject(id, OpenMode.ForRead) as Entity;

                                if (ent != null)
                                {
                                    // Check if it has METALCOMP Xdata
                                    ResultBuffer rbComp = ent.GetXDataForApplication("METALCOMP");
                                    if (rbComp != null)
                                    {
                                        // Build parts data
                                        List<ChildPart> parts = new List<ChildPart>();
                                        bool hasPartsData = false;

                                        // Get chunk count
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

                                        // If we have parts data, read it all
                                        if (numChunks > 0)
                                        {
                                            // Build all chunks
                                            StringBuilder jsonBuilder = new StringBuilder();

                                            // Collect data from all chunks in order
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
                                                            hasPartsData = true;
                                                        }
                                                    }
                                                }
                                            }

                                            // Try to parse the data
                                            if (jsonBuilder.Length > 0)
                                            {
                                                try
                                                {
                                                    parts = JsonConvert.DeserializeObject<List<ChildPart>>(jsonBuilder.ToString());
                                                    ed.WriteMessage($"\nLoaded {parts.Count} parts from component {ent.Handle}");
                                                }
                                                catch (System.Exception ex)
                                                {
                                                    ed.WriteMessage($"\nError parsing JSON for {ent.Handle}: {ex.Message}");
                                                    tr.Commit();
                                                    continue;
                                                }
                                            }
                                        }

                                        // If we have parts data, we'll rewrite it
                                        if (hasPartsData && parts.Count > 0)
                                        {
                                            // Open entity for write
                                            Entity entForWrite = tr.GetObject(id, OpenMode.ForWrite) as Entity;

                                            // Serialize to JSON
                                            string json = JsonConvert.SerializeObject(parts);

                                            // Calculate new chunks
                                            const int maxChunkSize = 1000;
                                            int newChunkCount = (int)Math.Ceiling((double)json.Length / maxChunkSize);

                                            // First, update the chunk count
                                            ResultBuffer rbNewInfo = new ResultBuffer(
                                                new TypedValue((int)DxfCode.ExtendedDataRegAppName, "METALPARTSINFO"),
                                                new TypedValue((int)DxfCode.ExtendedDataInteger32, newChunkCount)
                                            );
                                            entForWrite.XData = rbNewInfo;

                                            // Then write each chunk one by one
                                            for (int i = 0; i < newChunkCount; i++)
                                            {
                                                int startIndex = i * maxChunkSize;
                                                int length = Math.Min(maxChunkSize, json.Length - startIndex);
                                                string chunk = json.Substring(startIndex, length);

                                                ResultBuffer rbNewChunk = new ResultBuffer(
                                                    new TypedValue((int)DxfCode.ExtendedDataRegAppName, $"METALPARTS{i}"),
                                                    new TypedValue((int)DxfCode.ExtendedDataAsciiString, chunk)
                                                );
                                                entForWrite.XData = rbNewChunk;
                                            }

                                            // Now remove any old chunks that might be leftover
                                            for (int i = newChunkCount; i < 20; i++)
                                            {
                                                string appName = $"METALPARTS{i}";
                                                // Check if this app exists before trying to write to it
                                                using (Transaction regTr = db.TransactionManager.StartTransaction())
                                                {
                                                    RegAppTable regTable = (RegAppTable)regTr.GetObject(db.RegAppTableId, OpenMode.ForRead);
                                                    if (regTable.Has(appName))
                                                    {
                                                        // Only write if app exists
                                                        ResultBuffer rbEmpty = new ResultBuffer(
                                                            new TypedValue((int)DxfCode.ExtendedDataRegAppName, appName)
                                                        );
                                                        entForWrite.XData = rbEmpty;
                                                    }
                                                    regTr.Commit();
                                                }
                                            }

                                            cleanedCount++;
                                        }
                                    }
                                }

                                tr.Commit();
                            }
                        }
                        catch (System.Exception ex)
                        {
                            ed.WriteMessage($"\nError processing entity {id}: {ex.Message}");
                        }
                    }

                    ed.WriteMessage($"\nCleaned up XData for {cleanedCount} components.");
                }

                ed.WriteMessage("\nCleanup operation completed.");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError: {ex.Message}");
                ed.WriteMessage($"\nStack trace: {ex.StackTrace}");
            }
        }

        private void RegisterApp(RegAppTable regTable, string appName, Transaction tr)
        {
            if (!regTable.Has(appName))
            {
                RegAppTableRecord record = new RegAppTableRecord();
                record.Name = appName;
                regTable.Add(record);
                tr.AddNewlyCreatedDBObject(record, true);
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