using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace TakeoffBridge
{
    /// <summary>
    /// Central manager for retrieving metal components, parts, and attachments from the drawing
    /// </summary>
    public class DrawingComponentManager
    {
        #region Singleton Implementation

        private static DrawingComponentManager _instance;
        private Database _db;

        private DrawingComponentManager()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            _db = doc.Database;
        }

        public static DrawingComponentManager Instance
        {
            get
            {
                Document doc = Application.DocumentManager.MdiActiveDocument;
                if (_instance == null || _instance._db != doc.Database)
                {
                    _instance = new DrawingComponentManager();
                }
                return _instance;
            }
        }

        #endregion

        #region Common Data Structures

        /// <summary>
        /// Represents a metal component in the drawing
        /// </summary>
        public class Component
        {
            public string Handle { get; set; }
            public string Type { get; set; } // "Horizontal" or "Vertical"
            public string Floor { get; set; }
            public string Elevation { get; set; }
            public Point3d StartPoint { get; set; }
            public Point3d EndPoint { get; set; }
            public double Length { get; set; }
            public List<ChildPart> Parts { get; set; } = new List<ChildPart>();
        }

       

        #endregion

        #region Public Methods

        /// <summary>
        /// Gets all components from the current drawing
        /// </summary>
        public List<Component> GetAllComponents(Transaction tr = null)
        {
            bool ownsTransaction = (tr == null);
            List<Component> components = new List<Component>();

            try
            {
                if (ownsTransaction)
                {
                    tr = _db.TransactionManager.StartTransaction();
                }

                // Get all polylines that might be metal components
                TypedValue[] filterList = new TypedValue[] {
                    new TypedValue((int)DxfCode.Start, "LWPOLYLINE")
                };

                SelectionFilter filter = new SelectionFilter(filterList);
                PromptSelectionResult selRes = Application.DocumentManager.MdiActiveDocument.Editor.SelectAll(filter);

                if (selRes.Status == PromptStatus.OK)
                {
                    foreach (ObjectId id in selRes.Value.GetObjectIds())
                    {
                        Entity ent = (Entity)tr.GetObject(id, OpenMode.ForRead);

                        // Check if this has metal component xdata
                        ResultBuffer rbComp = ent.GetXDataForApplication("METALCOMP");
                        if (rbComp != null)
                        {
                            Polyline pline = ent as Polyline;
                            if (pline == null) continue;

                            Component comp = new Component
                            {
                                Handle = ent.Handle.ToString(),
                                Parts = new List<ChildPart>()
                            };

                            // Extract basic component data
                            TypedValue[] xdataComp = rbComp.AsArray();
                            for (int i = 1; i < xdataComp.Length; i++) // Skip app name
                            {
                                if (i == 1) comp.Type = xdataComp[i].Value.ToString();
                                if (i == 2) comp.Floor = xdataComp[i].Value.ToString();
                                if (i == 3) comp.Elevation = xdataComp[i].Value.ToString();
                            }

                            // Get geometry data
                            if (pline.NumberOfVertices >= 2)
                            {
                                comp.StartPoint = pline.GetPoint3dAt(0);
                                comp.EndPoint = pline.GetPoint3dAt(1);
                                comp.Length = comp.StartPoint.DistanceTo(comp.EndPoint);
                            }

                            // Get child parts
                            comp.Parts = GetPartsFromEntity(pline, tr);

                            components.Add(comp);
                        }
                    }
                }

                if (ownsTransaction)
                {
                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage($"\nError getting components: {ex.Message}");

                if (ownsTransaction && tr != null)
                {
                    tr.Abort();
                }

                throw;
            }
            finally
            {
                if (ownsTransaction && tr != null)
                {
                    tr.Dispose();
                }
            }

            return components;
        }

        /// <summary>
        /// Gets all attachments from the current drawing
        /// </summary>
        public List<Attachment> GetAllAttachments(Transaction tr = null)
        {
            bool ownsTransaction = (tr == null);
            List<Attachment> attachments = new List<Attachment>();

            try
            {
                if (ownsTransaction)
                {
                    tr = _db.TransactionManager.StartTransaction();
                }

                // Get named objects dictionary
                DBDictionary nod = (DBDictionary)tr.GetObject(_db.NamedObjectsDictionaryId, OpenMode.ForRead);

                // Check if attachment entry exists
                const string dictName = "METALATTACHMENTS";
                if (nod.Contains(dictName))
                {
                    DBObject obj = tr.GetObject(nod.GetAt(dictName), OpenMode.ForRead);
                    if (obj is Xrecord xrec && xrec.Data != null)
                    {
                        TypedValue[] values = xrec.Data.AsArray();
                        if (values.Length > 0 && values[0].TypeCode == (int)DxfCode.Text)
                        {
                            string json = values[0].Value.ToString();
                            attachments = JsonConvert.DeserializeObject<List<Attachment>>(json);
                        }
                    }
                }

                if (ownsTransaction)
                {
                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage($"\nError loading attachments: {ex.Message}");

                if (ownsTransaction && tr != null)
                {
                    tr.Abort();
                }

                throw;
            }
            finally
            {
                if (ownsTransaction && tr != null)
                {
                    tr.Dispose();
                }
            }

            return attachments;
        }

        /// <summary>
        /// Saves attachments to the drawing
        /// </summary>
        public void SaveAttachmentsToDrawing(List<Attachment> attachments, Transaction tr = null)
        {
            bool ownsTransaction = (tr == null);

            try
            {
                if (ownsTransaction)
                {
                    tr = _db.TransactionManager.StartTransaction();
                }

                // Serialize attachments to JSON
                string json = JsonConvert.SerializeObject(attachments);

                // Get named objects dictionary
                DBDictionary nod = (DBDictionary)tr.GetObject(_db.NamedObjectsDictionaryId, OpenMode.ForWrite);

                // Create or update dictionary entry
                const string dictName = "METALATTACHMENTS";

                if (nod.Contains(dictName))
                {
                    // Update existing
                    DBObject obj = tr.GetObject(nod.GetAt(dictName), OpenMode.ForWrite);
                    if (obj is Xrecord xrec)
                    {
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

                if (ownsTransaction)
                {
                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage($"\nError saving attachments: {ex.Message}");

                if (ownsTransaction && tr != null)
                {
                    tr.Abort();
                }

                throw;
            }
            finally
            {
                if (ownsTransaction && tr != null)
                {
                    tr.Dispose();
                }
            }
        }

        /// <summary>
        /// Processes components and attachments together to match them properly
        /// </summary>
        public List<Component> GetComponentsWithAttachments(Transaction tr = null)
        {
            bool ownsTransaction = (tr == null);

            try
            {
                if (ownsTransaction)
                {
                    tr = _db.TransactionManager.StartTransaction();
                }

                // Get all components and attachments
                List<Component> components = GetAllComponents(tr);
                List<Attachment> attachments = GetAllAttachments(tr);

                // Create lookup dictionaries
                Dictionary<string, Component> componentsByHandle = components.ToDictionary(c => c.Handle);
                Dictionary<string, List<Attachment>> attachmentsByVertical = new Dictionary<string, List<Attachment>>();

                // Group attachments by vertical handle
                foreach (var attachment in attachments)
                {
                    if (!attachmentsByVertical.ContainsKey(attachment.VerticalHandle))
                    {
                        attachmentsByVertical[attachment.VerticalHandle] = new List<Attachment>();
                    }
                    attachmentsByVertical[attachment.VerticalHandle].Add(attachment);
                }

                // Associate attachments with their components
                foreach (var component in components)
                {
                    if (component.Type == "Vertical" && attachmentsByVertical.ContainsKey(component.Handle))
                    {
                        // Process each vertical part to associate attachments
                        foreach (var attachment in attachmentsByVertical[component.Handle])
                        {
                            // Find the matching vertical part
                            ChildPart verticalPart = component.Parts.FirstOrDefault(p => p.PartType == attachment.VerticalPartType);

                            // Find the matching horizontal component and part
                            if (verticalPart != null && componentsByHandle.TryGetValue(attachment.HorizontalHandle, out Component horizontalComp))
                            {
                                ChildPart horizontalPart = horizontalComp.Parts.FirstOrDefault(p => p.PartType == attachment.HorizontalPartType);

                                if (horizontalPart != null)
                                {
                                    // Update the vertical part with attachment information
                                    if (verticalPart.Attachments == null)
                                    {
                                        verticalPart.Attachments = new List<PartAttachment>();
                                    }

                                    verticalPart.Attachments.Add(new PartAttachment
                                    {
                                        Side = attachment.Side,
                                        Position = attachment.Position,
                                        Height = attachment.Height,
                                        Invert = attachment.Invert,
                                        Adjust = attachment.Adjust,
                                        AttachedPartNumber = horizontalPart.Name,
                                        AttachedPartType = horizontalPart.PartType,
                                        AttachedFab = horizontalPart.Fab
                                    });
                                }
                            }
                        }
                    }
                }

                if (ownsTransaction)
                {
                    tr.Commit();
                }

                return components;
            }
            catch (System.Exception ex)
            {
                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage($"\nError processing components with attachments: {ex.Message}");

                if (ownsTransaction && tr != null)
                {
                    tr.Abort();
                }

                throw;
            }
            finally
            {
                if (ownsTransaction && tr != null)
                {
                    tr.Dispose();
                }
            }
        }

        /// <summary>
        /// Gets a component by its handle
        /// </summary>
        public Component GetComponentByHandle(string handle, Transaction tr = null)
        {
            bool ownsTransaction = (tr == null);

            try
            {
                if (ownsTransaction)
                {
                    tr = _db.TransactionManager.StartTransaction();
                }

                // Get the entity by handle
                long longHandle = Convert.ToInt64(handle, 16);
                Handle h = new Handle(longHandle);
                ObjectId objId = _db.GetObjectId(false, h, 0);

                Entity ent = tr.GetObject(objId, OpenMode.ForRead) as Entity;
                if (ent == null || !(ent is Polyline))
                {
                    return null;
                }

                // Check if it's a metal component
                ResultBuffer rbComp = ent.GetXDataForApplication("METALCOMP");
                if (rbComp == null)
                {
                    return null;
                }

                Polyline pline = ent as Polyline;
                Component comp = new Component
                {
                    Handle = handle,
                    Parts = new List<ChildPart>()
                };

                // Extract basic component data
                TypedValue[] xdataComp = rbComp.AsArray();
                for (int i = 1; i < xdataComp.Length; i++) // Skip app name
                {
                    if (i == 1) comp.Type = xdataComp[i].Value.ToString();
                    if (i == 2) comp.Floor = xdataComp[i].Value.ToString();
                    if (i == 3) comp.Elevation = xdataComp[i].Value.ToString();
                }

                // Get geometry data
                if (pline.NumberOfVertices >= 2)
                {
                    comp.StartPoint = pline.GetPoint3dAt(0);
                    comp.EndPoint = pline.GetPoint3dAt(1);
                    comp.Length = comp.StartPoint.DistanceTo(comp.EndPoint);
                }

                // Get child parts
                comp.Parts = GetPartsFromEntity(pline, tr);

                if (ownsTransaction)
                {
                    tr.Commit();
                }

                return comp;
            }
            catch (System.Exception ex)
            {
                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage($"\nError getting component by handle: {ex.Message}");

                if (ownsTransaction && tr != null)
                {
                    tr.Abort();
                }

                return null;
            }
            finally
            {
                if (ownsTransaction && tr != null)
                {
                    tr.Dispose();
                }
            }
        }

        #endregion

        #region Helper Methods

        /// <summary>
        /// Gets child parts from an entity - public for compatibility with existing code
        /// </summary>
        public List<ChildPart> GetPartsFromEntity(Entity ent, Transaction tr = null)
        {
            bool ownsTransaction = (tr == null);

            try
            {
                if (ownsTransaction)
                {
                    tr = _db.TransactionManager.StartTransaction();
                }

                List<ChildPart> result = GetPartsFromEntityInternal(ent, tr);

                if (ownsTransaction)
                {
                    tr.Commit();
                }

                return result;
            }
            catch (System.Exception ex)
            {
                Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage($"\nError getting parts from entity: {ex.Message}");

                if (ownsTransaction && tr != null)
                {
                    tr.Abort();
                }

                return new List<ChildPart>();
            }
            finally
            {
                if (ownsTransaction && tr != null)
                {
                    tr.Dispose();
                }
            }
        }



        /// <summary>
        /// Gets child parts from an entity
        /// </summary>
        private List<ChildPart> GetPartsFromEntityInternal(Entity ent, Transaction tr)
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

                        // Ensure every part has an Attachments collection
                        foreach (var part in result)
                        {
                            if (part.Attachments == null)
                            {
                                part.Attachments = new List<PartAttachment>();
                            }
                        }
                    }
                    catch (System.Exception ex)
                    {
                        Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage($"\nError parsing parts JSON: {ex.Message}");
                    }
                }
            }

            return result;
        }

        #endregion
    }
}