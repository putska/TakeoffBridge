using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using Newtonsoft.Json;

namespace TakeoffBridge
{
    public class MarkNumberManager
    {
        // Singleton instance
        private static MarkNumberManager _instance;

        // Database reference
        private Database _db;

        // Cache of unique part configurations and their assigned mark numbers
        private Dictionary<string, string> _horizontalPartMarks = new Dictionary<string, string>();
        private Dictionary<string, string> _verticalPartMarks = new Dictionary<string, string>();

        // Track the next available number for each part type
        private Dictionary<string, int> _nextMarkNumbers = new Dictionary<string, int>();

        // Cache of entity handles to avoid redundant processing
        private HashSet<string> _processedEntities = new HashSet<string>();

        // Constructor is private for singleton pattern
        private MarkNumberManager()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            _db = doc.Database;

            // Initialize by loading existing mark numbers from the drawing
            InitializeFromDrawing();
        }

        // Singleton access
        public static MarkNumberManager Instance
        {
            get
            {
                if (_instance == null)
                {
                    _instance = new MarkNumberManager();
                }
                return _instance;
            }
        }

        #region Event Handlers

        // Called when a new component is created
        public void OnComponentCreated(ObjectId entityId)
        {
            using (Transaction tr = _db.TransactionManager.StartTransaction())
            {
                Entity ent = tr.GetObject(entityId, OpenMode.ForRead) as Entity;
                if (ent != null)
                {
                    string handle = ent.Handle.ToString();

                    // Skip if already processed
                    if (_processedEntities.Contains(handle))
                    {
                        tr.Commit();
                        return;
                    }

                    // Process the new component
                    ProcessComponentMarkNumbers(entityId, tr);

                    // Add to processed set
                    _processedEntities.Add(handle);
                }
                tr.Commit();
            }
        }

        // Called when a component is modified
        public void OnComponentModified(ObjectId entityId)
        {
            using (Transaction tr = _db.TransactionManager.StartTransaction())
            {
                // Update this component and any connected components
                ProcessComponentMarkNumbers(entityId, tr);

                // Check if this modification affects any attachments
                UpdateAttachedComponents(entityId, tr);

                tr.Commit();
            }
        }

        // Called when a component is stretched (changing its length)
        public void OnComponentStretched(ObjectId entityId, double newLength)
        {
            using (Transaction tr = _db.TransactionManager.StartTransaction())
            {
                // Get component data
                Entity ent = tr.GetObject(entityId, OpenMode.ForRead) as Entity;
                if (ent == null)
                {
                    tr.Commit();
                    return;
                }

                string componentType = GetComponentType(ent);
                if (string.IsNullOrEmpty(componentType))
                {
                    tr.Commit();
                    return;
                }

                // Get child parts
                List<ChildPart> childParts = GetChildParts(ent);
                if (childParts.Count == 0)
                {
                    tr.Commit();
                    return;
                }

                // Check if length change affects any mark numbers
                bool markNumbersChanged = false;

                foreach (var part in childParts)
                {
                    // Skip fixed length parts as they aren't affected by stretching
                    if (part.IsFixedLength)
                        continue;

                    // Calculate the old and new actual lengths
                    double oldActualLength = CalculateActualPartLength(part, GetComponentLength(entityId));
                    double newActualLength = CalculateActualPartLength(part, newLength);

                    // If length changed significantly, update mark number
                    if (Math.Abs(oldActualLength - newActualLength) > 0.001)
                    {
                        // Generate new mark number
                        if (componentType == "Horizontal")
                        {
                            string uniqueKey = GenerateHorizontalPartKey(part, newActualLength);
                            part.MarkNumber = GetOrCreateMarkNumber(uniqueKey, part.PartType, _horizontalPartMarks);
                        }
                        else if (componentType == "Vertical")
                        {
                            // For vertical components, we need attachment information
                            string handle = ent.Handle.ToString();
                            var attachments = GetAttachmentsForComponent(handle, true);

                            string uniqueKey = GenerateVerticalPartKey(part, newActualLength, attachments);
                            part.MarkNumber = GetOrCreateMarkNumber(uniqueKey, part.PartType, _verticalPartMarks);
                        }

                        markNumbersChanged = true;
                    }
                }

                // If any mark numbers changed, save the updated child parts
                if (markNumbersChanged)
                {
                    Entity entForWrite = tr.GetObject(entityId, OpenMode.ForWrite) as Entity;
                    string json = JsonConvert.SerializeObject(childParts);
                    SaveChildPartsToEntity(entForWrite, json);
                }

                // Check if this modification affects any attachments
                if (componentType == "Horizontal")
                {
                    UpdateAttachedComponents(entityId, tr);
                }

                tr.Commit();
            }
        }

        // Called when intersections are recalculated
        public void OnIntersectionsRecalculated(List<ObjectId> verticalIds)
        {
            using (Transaction tr = _db.TransactionManager.StartTransaction())
            {
                foreach (ObjectId verticalId in verticalIds)
                {
                    // Process mark numbers for this vertical component
                    ProcessComponentMarkNumbers(verticalId, tr);
                }

                tr.Commit();
            }
        }

        #endregion

        #region Processing Methods

        // Initialize by loading existing mark numbers from the drawing
        private void InitializeFromDrawing()
        {
            _horizontalPartMarks.Clear();
            _verticalPartMarks.Clear();
            _nextMarkNumbers.Clear();
            _processedEntities.Clear();

            // Start a transaction to read the drawing
            using (Transaction tr = _db.TransactionManager.StartTransaction())
            {
                // Get all polylines in the drawing
                TypedValue[] tvs = new TypedValue[] {
                    new TypedValue((int)DxfCode.Start, "LWPOLYLINE")
                };

                SelectionFilter filter = new SelectionFilter(tvs);
                PromptSelectionResult selRes = Application.DocumentManager.MdiActiveDocument.Editor.SelectAll(filter);

                if (selRes.Status == PromptStatus.OK)
                {
                    SelectionSet ss = selRes.Value;

                    foreach (SelectedObject selObj in ss)
                    {
                        ObjectId id = selObj.ObjectId;
                        Entity ent = tr.GetObject(id, OpenMode.ForRead) as Entity;

                        if (ent != null)
                        {
                            // Check if this entity has metal component data
                            string componentType = GetComponentType(ent);

                            if (!string.IsNullOrEmpty(componentType))
                            {
                                // Get child parts
                                List<ChildPart> childParts = GetChildParts(ent);
                                string handle = ent.Handle.ToString();

                                // Store in cache for each part with a mark number
                                foreach (var part in childParts)
                                {
                                    if (!string.IsNullOrEmpty(part.MarkNumber))
                                    {
                                        // Parse mark number to get the numeric part
                                        string[] markParts = part.MarkNumber.Split('-');
                                        if (markParts.Length == 2 && int.TryParse(markParts[1], out int markNumber))
                                        {
                                            // Track the highest mark number for each part type
                                            string partType = markParts[0];
                                            if (!_nextMarkNumbers.ContainsKey(partType) || _nextMarkNumbers[partType] <= markNumber)
                                            {
                                                _nextMarkNumbers[partType] = markNumber + 1;
                                            }

                                            // Generate the unique key for this part
                                            double length = (ent is Polyline) ? ((Polyline)ent).Length : 0;
                                            double actualLength = CalculateActualPartLength(part, length);

                                            if (componentType == "Horizontal")
                                            {
                                                string uniqueKey = GenerateHorizontalPartKey(part, actualLength);
                                                _horizontalPartMarks[uniqueKey] = part.MarkNumber;
                                            }
                                            else if (componentType == "Vertical")
                                            {
                                                var attachments = GetAttachmentsForComponent(handle, true);
                                                string uniqueKey = GenerateVerticalPartKey(part, actualLength, attachments);
                                                _verticalPartMarks[uniqueKey] = part.MarkNumber;
                                            }
                                        }
                                    }
                                }

                                // Add to processed entities
                                _processedEntities.Add(handle);
                            }
                        }
                    }
                }

                tr.Commit();
            }

            // Log how many mark numbers we loaded
            Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(
                $"\nLoaded {_horizontalPartMarks.Count} horizontal part marks and {_verticalPartMarks.Count} vertical part marks.");
        }

        // Process mark numbers for a component - made internal for use by the commands
        internal void ProcessComponentMarkNumbers(ObjectId entityId, Transaction tr)
        {
            Entity ent = tr.GetObject(entityId, OpenMode.ForRead) as Entity;
            if (ent == null)
                return;

            string componentType = GetComponentType(ent);
            if (string.IsNullOrEmpty(componentType))
                return;

            List<ChildPart> childParts = GetChildParts(ent);
            if (childParts.Count == 0)
                return;

            bool markNumbersChanged = false;
            double length = (ent is Polyline) ? ((Polyline)ent).Length : 0;
            string handle = ent.Handle.ToString();

            foreach (var part in childParts)
            {
                // Skip parts that already have mark numbers
                if (!string.IsNullOrEmpty(part.MarkNumber))
                    continue;

                // Calculate actual length
                double actualLength = CalculateActualPartLength(part, length);

                // Generate mark number based on component type
                if (componentType == "Horizontal")
                {
                    string uniqueKey = GenerateHorizontalPartKey(part, actualLength);
                    part.MarkNumber = GetOrCreateMarkNumber(uniqueKey, part.PartType, _horizontalPartMarks);
                }
                else if (componentType == "Vertical")
                {
                    var attachments = GetAttachmentsForComponent(handle, true);
                    string uniqueKey = GenerateVerticalPartKey(part, actualLength, attachments);
                    part.MarkNumber = GetOrCreateMarkNumber(uniqueKey, part.PartType, _verticalPartMarks);
                }

                markNumbersChanged = true;
            }

            // Save updated child parts if mark numbers changed
            if (markNumbersChanged)
            {
                Entity entForWrite = tr.GetObject(entityId, OpenMode.ForWrite) as Entity;
                string json = JsonConvert.SerializeObject(childParts);
                SaveChildPartsToEntity(entForWrite, json);
            }
        }

        // Update components that are attached to the given component
        private void UpdateAttachedComponents(ObjectId entityId, Transaction tr)
        {
            Entity ent = tr.GetObject(entityId, OpenMode.ForRead) as Entity;
            if (ent == null)
                return;

            string handle = ent.Handle.ToString();

            // Get all attachments related to this component
            var attachments = GetAttachmentsForComponent(handle, false);

            if (attachments.Count > 0)
            {
                // Get the ObjectIds for all affected vertical components
                List<ObjectId> verticalIds = new List<ObjectId>();

                foreach (var attachment in attachments)
                {
                    // Find the vertical component
                    ObjectId verticalId = GetEntityByHandle(attachment.VerticalHandle);

                    if (verticalId != ObjectId.Null && !verticalIds.Contains(verticalId))
                    {
                        verticalIds.Add(verticalId);
                    }
                }

                // Process mark numbers for all affected vertical components
                foreach (ObjectId verticalId in verticalIds)
                {
                    ProcessComponentMarkNumbers(verticalId, tr);
                }
            }
        }

        #endregion

        #region Helper Methods

        // Get or create a mark number for a unique part configuration
        private string GetOrCreateMarkNumber(string uniqueKey, string partType, Dictionary<string, string> markCache)
        {
            // Check if we already have a mark number for this configuration
            if (markCache.ContainsKey(uniqueKey))
            {
                return markCache[uniqueKey];
            }

            // Create a new mark number
            if (!_nextMarkNumbers.ContainsKey(partType))
            {
                _nextMarkNumbers[partType] = 1;
            }

            int markNumber = _nextMarkNumbers[partType]++;
            string newMarkNumber = $"{partType}-{markNumber}";

            // Cache it
            markCache[uniqueKey] = newMarkNumber;

            return newMarkNumber;
        }

        // Generate a unique key for horizontal parts
        private string GenerateHorizontalPartKey(ChildPart part, double actualLength)
        {
            // Round the length to avoid floating point comparison issues
            double roundedLength = Math.Round(actualLength, 3);

            // Create a unique key based on part type, length, finish, and fab
            return $"{part.PartType}|{roundedLength}|{part.Finish}|{part.Fab}";
        }

        // Generate a unique key for vertical parts with attachments
        private string GenerateVerticalPartKey(ChildPart part, double actualLength, List<Attachment> attachments)
        {
            // Round the length to avoid floating point comparison issues
            double roundedLength = Math.Round(actualLength, 3);

            // Start with basic part properties
            StringBuilder keyBuilder = new StringBuilder();
            keyBuilder.Append($"{part.PartType}|{roundedLength}|{part.Finish}|{part.Fab}");

            // Add attachment information
            if (attachments.Count > 0)
            {
                // Sort attachments by position to ensure consistent ordering
                var sortedAttachments = attachments.OrderBy(a => a.Position).ThenBy(a => a.Height).ToList();

                foreach (var attachment in sortedAttachments)
                {
                    keyBuilder.Append($"|{attachment.HorizontalPartType}-{attachment.Side}-{Math.Round(attachment.Position, 3)}-{Math.Round(attachment.Height, 3)}-{attachment.Invert}-{Math.Round(attachment.Adjust, 3)}");
                }
            }

            return keyBuilder.ToString();
        }

        // Get attachments for a component
        private List<Attachment> GetAttachmentsForComponent(string handle, bool isVertical)
        {
            List<Attachment> allAttachments = LoadAttachmentsFromDrawing();

            if (isVertical)
            {
                // Return attachments where this component is the vertical
                return allAttachments.Where(a => a.VerticalHandle == handle).ToList();
            }
            else
            {
                // Return attachments where this component is the horizontal
                return allAttachments.Where(a => a.HorizontalHandle == handle).ToList();
            }
        }

        // Get an entity by its handle
        private ObjectId GetEntityByHandle(string handle)
        {
            ObjectId result = ObjectId.Null;

            try
            {
                Handle h = new Handle(Convert.ToInt64(handle, 16));
                result = _db.GetObjectId(false, h, 0);
            }
            catch
            {
                // Handle not found or invalid
            }

            return result;
        }

        // Get the component type from an entity - made internal for use by commands
        internal string GetComponentType(Entity ent)
        {
            // Get component type from Xdata
            using (ResultBuffer rb = ent.XData)
            {
                if (rb != null)
                {
                    TypedValue[] tvs = rb.AsArray();
                    foreach (TypedValue tv in tvs)
                    {
                        if (tv.TypeCode == (int)DxfCode.ExtendedDataRegAppName)
                        {
                            if (tv.Value.ToString() == "METALCOMP")
                            {
                                // Next value should be the component type
                                int index = Array.IndexOf(tvs, tv) + 1;
                                if (index < tvs.Length && tvs[index].TypeCode == (int)DxfCode.ExtendedDataAsciiString)
                                {
                                    return tvs[index].Value.ToString();
                                }
                            }
                        }
                    }
                }
            }

            return null;
        }

        // Get child parts from an entity
        private List<ChildPart> GetChildParts(Entity ent)
        {
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

            // No parts info found
            if (numChunks == 0)
            {
                return new List<ChildPart>();
            }

            // Build JSON string from chunks
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
                    return JsonConvert.DeserializeObject<List<ChildPart>>(jsonBuilder.ToString());
                }
                catch (System.Exception ex)
                {
                    Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(
                        $"\nError parsing child parts JSON: {ex.Message}");
                }
            }

            return new List<ChildPart>();
        }

        // Calculate actual part length
        private double CalculateActualPartLength(ChildPart part, double componentLength)
        {
            if (part.IsFixedLength)
            {
                return part.FixedLength;
            }
            else
            {
                // Calculate based on component length and adjustments
                return componentLength + part.StartAdjustment + part.EndAdjustment;
            }
        }

        // Get component length
        private double GetComponentLength(ObjectId entityId)
        {
            double length = 0;

            using (Transaction tr = _db.TransactionManager.StartTransaction())
            {
                Entity ent = tr.GetObject(entityId, OpenMode.ForRead) as Entity;
                if (ent is Polyline)
                {
                    Polyline pline = ent as Polyline;
                    length = pline.Length;
                }

                tr.Commit();
            }

            return length;
        }

        // Save child parts to an entity
        private void SaveChildPartsToEntity(Entity ent, string json)
        {
            // Check if we need to chunk the JSON (AutoCAD limits Xdata to 16kb per application)
            const int maxChunkSize = 1000; // Conservative value below AutoCAD's limit

            if (json.Length <= maxChunkSize)
            {
                // Single chunk - simple case
                using (ResultBuffer rb = new ResultBuffer(
                    new TypedValue((int)DxfCode.ExtendedDataRegAppName, "METALPARTS"),
                    new TypedValue((int)DxfCode.ExtendedDataAsciiString, json)))
                {
                    ent.XData = rb;
                }
            }
            else
            {
                // Need to chunk the data
                int chunkCount = (int)Math.Ceiling((double)json.Length / maxChunkSize);

                for (int i = 0; i < chunkCount; i++)
                {
                    int startIndex = i * maxChunkSize;
                    int length = Math.Min(maxChunkSize, json.Length - startIndex);
                    string chunk = json.Substring(startIndex, length);

                    string appName = (i == 0) ? "METALPARTS" : $"METALPARTS_{i}";

                    using (ResultBuffer rb = new ResultBuffer(
                        new TypedValue((int)DxfCode.ExtendedDataRegAppName, appName),
                        new TypedValue((int)DxfCode.ExtendedDataAsciiString, chunk)))
                    {
                        ent.XData = rb;
                    }
                }
            }
        }

        // Load attachments from drawing
        private List<Attachment> LoadAttachmentsFromDrawing()
        {
            List<Attachment> attachments = new List<Attachment>();

            using (Transaction tr = _db.TransactionManager.StartTransaction())
            {
                // Get named objects dictionary
                DBDictionary nod = (DBDictionary)tr.GetObject(_db.NamedObjectsDictionaryId, OpenMode.ForRead);

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
                                attachments = JsonConvert.DeserializeObject<List<Attachment>>(json);
                            }
                        }
                    }
                }

                tr.Commit();
            }

            return attachments;
        }

        #endregion

        #region Public Methods

        // Add these methods for the report command
        public int GetHorizontalPartCount()
        {
            return _horizontalPartMarks.Count;
        }

        public int GetVerticalPartCount()
        {
            return _verticalPartMarks.Count;
        }

        #endregion

        #region Static Methods for Commands

        // Command to regenerate all mark numbers in the drawing
        public static void RegenerateAllMarkNumbers()
        {
            // Reset the manager and initialize from scratch
            _instance = null;
            MarkNumberManager manager = Instance;

            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;

            try
            {
                ed.WriteMessage("\nResetting all mark numbers...");

                using (Transaction tr = Application.DocumentManager.MdiActiveDocument.Database.TransactionManager.StartTransaction())
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

                        foreach (SelectedObject selObj in ss)
                        {
                            ObjectId id = selObj.ObjectId;
                            Entity ent = tr.GetObject(id, OpenMode.ForRead) as Entity;

                            if (ent != null)
                            {
                                // Check if this entity has metal component data
                                string componentType = manager.GetComponentType(ent);

                                if (!string.IsNullOrEmpty(componentType))
                                {
                                    // Reset the mark numbers
                                    List<ChildPart> childParts = manager.GetChildParts(ent);
                                    bool markNumbersChanged = false;

                                    foreach (var part in childParts)
                                    {
                                        if (!string.IsNullOrEmpty(part.MarkNumber))
                                        {
                                            part.MarkNumber = null;
                                            markNumbersChanged = true;
                                        }
                                    }

                                    // Save updated child parts if mark numbers changed
                                    if (markNumbersChanged)
                                    {
                                        Entity entForWrite = tr.GetObject(id, OpenMode.ForWrite) as Entity;
                                        string json = JsonConvert.SerializeObject(childParts);
                                        manager.SaveChildPartsToEntity(entForWrite, json);
                                    }
                                }
                            }
                        }
                    }

                    tr.Commit();
                }

                // Force the manager to reinitialize
                _instance = null;

                ed.WriteMessage("\nMark numbers reset. Use GENERATEMARKNUMBERS to regenerate all mark numbers.");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage("\nError: " + ex.Message);
            }
        }

        // Command to generate all mark numbers in the drawing
        public static void GenerateAllMarkNumbers()
        {
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;

            try
            {
                ed.WriteMessage("\nGenerating mark numbers for all components...");

                MarkNumberManager manager = Instance;

                // Process all metal components in the drawing
                using (Transaction tr = Application.DocumentManager.MdiActiveDocument.Database.TransactionManager.StartTransaction())
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

                        // First process all horizontal components
                        foreach (SelectedObject selObj in ss)
                        {
                            ObjectId id = selObj.ObjectId;
                            Entity ent = tr.GetObject(id, OpenMode.ForRead) as Entity;

                            if (ent != null)
                            {
                                string componentType = manager.GetComponentType(ent);

                                if (componentType == "Horizontal")
                                {
                                    manager.ProcessComponentMarkNumbers(id, tr);
                                }
                            }
                        }

                        // Then process all vertical components
                        foreach (SelectedObject selObj in ss)
                        {
                            ObjectId id = selObj.ObjectId;
                            Entity ent = tr.GetObject(id, OpenMode.ForRead) as Entity;

                            if (ent != null)
                            {
                                string componentType = manager.GetComponentType(ent);

                                if (componentType == "Vertical")
                                {
                                    manager.ProcessComponentMarkNumbers(id, tr);
                                }
                            }
                        }
                    }

                    tr.Commit();
                }

                ed.WriteMessage("\nMark number generation complete.");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage("\nError: " + ex.Message);
            }
        }

        #endregion
    }
}