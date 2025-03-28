using System;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using System.Linq;

namespace TakeoffBridge
{
    /// <summary>
    /// Commands for managing mark number displays
    /// </summary>
    public class MarkNumberDisplayCommands
    {

        // Static flag to track if the monitor is active
        private static bool _monitorActive = false;

        // Store event handlers so we can remove them later
        private static CommandEventHandler _commandStartHandler;
        private static CommandEventHandler _commandEndHandler;


        [CommandMethod("SHOWMARKNUMBERS")]
        public void ShowMarkNumbers()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            try
            {
                ed.WriteMessage("\nUpdating all mark number displays...");

                // Get the mark number display manager
                MarkNumberDisplay.Instance.UpdateAllMarkNumbers();

                ed.WriteMessage("\nMark number display update complete.");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError: {ex.Message}");
            }
        }

        [CommandMethod("CLEARMARKNUMBERS")]
        public void ClearMarkNumbers()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            try
            {
                ed.WriteMessage("\nClearing all mark number displays...");

                // Get the mark number display manager
                MarkNumberDisplay.Instance.ClearAllMarkTexts();

                ed.WriteMessage("\nMark number displays cleared.");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError: {ex.Message}");
            }
        }

        [CommandMethod("UPDATESELECTEDMARKNUMBERS")]
        public void UpdateSelectedMarkNumbers()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            try
            {
                // Prompt for selection of components
                PromptSelectionOptions opts = new PromptSelectionOptions();
                opts.MessageForAdding = "\nSelect components to update mark numbers: ";
                opts.AllowDuplicates = false;

                PromptSelectionResult result = ed.GetSelection(opts);

                if (result.Status != PromptStatus.OK)
                {
                    return;
                }

                SelectionSet ss = result.Value;

                using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
                {
                    int count = 0;

                    foreach (SelectedObject selObj in ss)
                    {
                        Entity ent = tr.GetObject(selObj.ObjectId, OpenMode.ForRead) as Entity;

                        if (ent is Polyline)
                        {
                            // Check if it has metal component data
                            ResultBuffer rbComp = ent.GetXDataForApplication("METALCOMP");
                            if (rbComp != null)
                            {
                                // Update this component's mark numbers
                                MarkNumberDisplay.Instance.UpdateMarkNumbersForComponent(selObj.ObjectId);
                                count++;
                            }
                        }
                    }

                    tr.Commit();

                    ed.WriteMessage($"\nUpdated mark numbers for {count} components.");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError: {ex.Message}");
            }
        }

        [CommandMethod("COPYANDUPDATECOMPONENT")]
        public void CopyAndUpdateComponent()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            try
            {
                // Prompt for source component
                PromptEntityOptions peoSource = new PromptEntityOptions("\nSelect source component: ");
                peoSource.SetRejectMessage("\nOnly polylines with metal component data can be selected.");
                peoSource.AddAllowedClass(typeof(Polyline), false);

                PromptEntityResult perSource = ed.GetEntity(peoSource);
                if (perSource.Status != PromptStatus.OK) return;

                ObjectId sourceId = perSource.ObjectId;

                // Ensure the source has mark numbers displayed
                MarkNumberDisplay.Instance.UpdateMarkNumbersForComponent(sourceId);

                // Get the source handle
                string sourceHandle;
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    Entity ent = tr.GetObject(sourceId, OpenMode.ForRead) as Entity;
                    sourceHandle = ent.Handle.ToString();
                    tr.Commit();
                }

                // Use standard AutoCAD COPY command
                ed.WriteMessage("\nUsing AutoCAD COPY command. Press ESC when done.");
                doc.SendStringToExecute("_COPY ", true, false, false);

                // After the copy, we need to find the newly created components
                // This is tricky, but we can use the idle event to check after the command completes
                Application.Idle += delegate
                {
                    Application.Idle -= delegate { };

                    try
                    {
                        // Find newly created components that weren't processed
                        using (Transaction tr = db.TransactionManager.StartTransaction())
                        {
                            // Get all polylines in the drawing
                            TypedValue[] tvs = new TypedValue[] {
                                new TypedValue((int)DxfCode.Start, "LWPOLYLINE")
                            };

                            SelectionFilter filter = new SelectionFilter(tvs);
                            PromptSelectionResult selRes = ed.SelectAll(filter);

                            if (selRes.Status == PromptStatus.OK)
                            {
                                SelectionSet ss = selRes.Value;
                                int count = 0;

                                // Get the source component
                                Entity sourceEnt = tr.GetObject(sourceId, OpenMode.ForRead) as Entity;

                                foreach (SelectedObject selObj in ss)
                                {
                                    ObjectId id = selObj.ObjectId;

                                    // Skip the source
                                    if (id == sourceId) continue;

                                    Entity ent = tr.GetObject(id, OpenMode.ForRead) as Entity;

                                    if (ent != null)
                                    {
                                        string componentType = MarkNumberDisplay.Instance.GetComponentType(ent);

                                        if (!string.IsNullOrEmpty(componentType))
                                        {
                                            // First update mark numbers for this component
                                            MarkNumberDisplay.Instance.UpdateMarkNumbersForComponent(id);

                                            // Calculate transformation from source to this component
                                            Matrix3d transform = CalculateTransformation(sourceEnt, ent);

                                            // Copy custom positions from source to this component
                                            MarkNumberDisplay.Instance.CopyCustomPositions(
                                                sourceHandle,
                                                ent.Handle.ToString(),
                                                transform);

                                            // Update again to apply the custom positions
                                            MarkNumberDisplay.Instance.UpdateMarkNumbersForComponent(id);

                                            count++;
                                        }
                                    }
                                }

                                ed.WriteMessage($"\nProcessed {count} newly created components.");
                            }

                            tr.Commit();
                        }
                    }
                    catch (System.Exception ex)
                    {
                        ed.WriteMessage($"\nError processing copied components: {ex.Message}");
                    }
                };
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError: {ex.Message}");
            }
        }

        /// <summary>
        /// Calculates a transformation matrix from source to target component
        /// </summary>
        private Matrix3d CalculateTransformation(Entity source, Entity target)
        {
            if (!(source is Polyline) || !(target is Polyline))
            {
                return Matrix3d.Identity;
            }

            Polyline sourcePline = source as Polyline;
            Polyline targetPline = target as Polyline;

            // Get component points
            Point3d sourceStart = sourcePline.GetPoint3dAt(0);
            Point3d sourceEnd = sourcePline.GetPoint3dAt(1);

            Point3d targetStart = targetPline.GetPoint3dAt(0);
            Point3d targetEnd = targetPline.GetPoint3dAt(1);

            // Calculate translation vector from source to target
            Vector3d translation = targetStart - sourceStart;

            // Calculate rotation angle if needed
            Vector3d sourceDir = sourceEnd - sourceStart;
            Vector3d targetDir = targetEnd - targetStart;

            double sourceAngle = Math.Atan2(sourceDir.Y, sourceDir.X);
            double targetAngle = Math.Atan2(targetDir.Y, targetDir.X);
            double rotation = targetAngle - sourceAngle;

            // Create transformation matrices
            Matrix3d translateToOrigin = Matrix3d.Displacement(-sourceStart.GetAsVector());
            Matrix3d rotate = Matrix3d.Rotation(rotation, Vector3d.ZAxis, Point3d.Origin);
            Matrix3d translateToTarget = Matrix3d.Displacement(targetStart.GetAsVector());

            // Combine transformations
            Matrix3d transform = translateToOrigin * rotate * translateToTarget;

            return transform;
        }


            [CommandMethod("STRETCHMONITOR", CommandFlags.Transparent)]
            public void StretchMonitor()
            {
                Document doc = Application.DocumentManager.MdiActiveDocument;
                Editor ed = doc.Editor;

                if (_monitorActive)
                {
                    ed.WriteMessage("\nStretch Monitor is already active.");
                    return;
                }

                ed.WriteMessage("\nStretch Monitor active. Mark numbers will update after stretch operations.");

                // Set up command start handler
                _commandStartHandler = (sender, e) => {
                    if (IsModificationCommand(e.GlobalCommandName))
                    {
                        ed.WriteMessage($"\n{e.GlobalCommandName} command detected. Will update mark numbers when completed.");

                        // Set up a one-time handler for this specific command
                        CommandEventHandler endHandler = null;
                        endHandler = (s, args) => {
                            if (args.GlobalCommandName == e.GlobalCommandName)
                            {
                                // Unregister this one-time handler
                                doc.CommandEnded -= endHandler;

                                ed.WriteMessage($"\n{args.GlobalCommandName} completed. Updating mark numbers...");

                                // First refresh all mark numbers
                                RefreshAllMarkNumbers();

                                // Then update all displays
                                MarkNumberDisplay.Instance.UpdateAllMarkNumbers();
                            }
                        };

                        doc.CommandEnded += endHandler;
                    }
                };

                // General command end handler that will clean up if needed
                _commandEndHandler = (sender, e) => {
                    // Nothing specific needed here, but we'll keep it for future use
                };

                // Register the handlers
                doc.CommandWillStart += _commandStartHandler;
                doc.CommandEnded += _commandEndHandler;

                _monitorActive = true;
            }

            [CommandMethod("STOPMONITOR", CommandFlags.Transparent)]
            public void StopMonitor()
            {
                Document doc = Application.DocumentManager.MdiActiveDocument;
                Editor ed = doc.Editor;

                if (!_monitorActive)
                {
                    ed.WriteMessage("\nStretch Monitor is not currently active.");
                    return;
                }

                // Unregister the handlers
                doc.CommandWillStart -= _commandStartHandler;
                doc.CommandEnded -= _commandEndHandler;

                _monitorActive = false;

                ed.WriteMessage("\nStretch Monitor deactivated.");
            }

            // Helper method to check if a command is one we want to monitor
            private bool IsModificationCommand(string commandName)
            {
                string[] commandsToMonitor = new string[] {
            "STRETCH", "MOVE", "COPY", "SCALE", "ROTATE", "LENGTHEN", "GRIP_STRETCH"
        };

                return commandsToMonitor.Contains(commandName.ToUpper());
            }

            // Helper method to refresh all mark numbers
            private void RefreshAllMarkNumbers()
            {
                Document doc = Application.DocumentManager.MdiActiveDocument;

                using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
                {
                    // Get all polylines in the drawing
                    TypedValue[] tvs = new TypedValue[] {
                new TypedValue((int)DxfCode.Start, "LWPOLYLINE")
            };

                    SelectionFilter filter = new SelectionFilter(tvs);
                    PromptSelectionResult selRes = doc.Editor.SelectAll(filter);

                    if (selRes.Status == PromptStatus.OK)
                    {
                        SelectionSet ss = selRes.Value;

                        foreach (SelectedObject selObj in ss)
                        {
                            ObjectId id = selObj.ObjectId;
                            Entity ent = tr.GetObject(id, OpenMode.ForRead) as Entity;

                            if (ent is Polyline)
                            {
                                Polyline pline = ent as Polyline;
                                string componentType = MarkNumberDisplay.Instance.GetComponentType(pline);

                                if (!string.IsNullOrEmpty(componentType))
                                {
                                    // Process mark numbers for this component
                                    MarkNumberManager.Instance.ProcessComponentMarkNumbers(id, tr, forceProcess: true);
                                }
                            }
                        }
                    }

                    tr.Commit();
                }
            }
        }

}