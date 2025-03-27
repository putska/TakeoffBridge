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
using System.Security.Cryptography;

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

            // Connect to events
            ConnectToEvents();

        }

        // Singleton access
        public static MarkNumberManager Instance
        {
            get
            {
                Document doc = Application.DocumentManager.MdiActiveDocument;
                if (_instance == null || _instance._db != doc.Database)
                {
                    _instance = new MarkNumberManager();
                }
                return _instance;
            }
        }


        private HashSet<string> _currentlyProcessing = new HashSet<string>();

        private Dictionary<string, DateTime> _lastProcessTime = new Dictionary<string, DateTime>();

        private bool _suppressEvents = false;

        private Dictionary<string, double> _lastComponentLengths = new Dictionary<string, double>();

        private readonly TimeSpan _throttleInterval = TimeSpan.FromSeconds(5); // 5 seconds between processing the same object

        // Update the throttling method to be more selective
        private bool ShouldThrottleProcessing(string handle)
        {
            DateTime now = DateTime.Now;

            if (_lastProcessTime.ContainsKey(handle))
            {
                TimeSpan elapsed = now - _lastProcessTime[handle];
                if (elapsed < _throttleInterval)
                {
                    // If we're throttling, only update the time if it's been at least 1 second
                    // This prevents extending the throttle period indefinitely
                    if (elapsed.TotalSeconds > 1)
                    {
                        _lastProcessTime[handle] = now;
                    }
                    return true; // Should throttle
                }
            }

            // Update last process time
            _lastProcessTime[handle] = now;
            return false; // Don't throttle
        }

        // Add to the MarkNumberManager class
        public static void ForceProcess(ObjectId entityId)
        {
            MarkNumberManager manager = Instance;

            using (Transaction tr = Application.DocumentManager.MdiActiveDocument.Database.TransactionManager.StartTransaction())
            {
                // Get the entity
                Entity ent = tr.GetObject(entityId, OpenMode.ForRead) as Entity;
                if (ent == null)
                {
                    tr.Commit();
                    return;
                }

                string handle = ent.Handle.ToString();

                // Clear any processing flags for this object
                manager._currentlyProcessing.Remove(handle);

                // Force immediate processing
                System.Diagnostics.Debug.WriteLine($"Forcing processing of {handle}");

                if (ent is Polyline)
                {
                    Polyline pline = ent as Polyline;
                    double length = pline.Length;
                    manager.OnComponentStretched(entityId, length, true);
                }
                else
                {
                    manager.OnComponentModified(entityId);
                }

                tr.Commit();
            }
        }


        // Track if we're currently connected to events
        private bool _eventsConnected = false;

        // Initialize by connecting to events
        private void ConnectToEvents()
        {
            if (_eventsConnected) return;

            try
            {
                // Connect to document events
                Document doc = Application.DocumentManager.MdiActiveDocument;
                if (doc != null)
                {
                    // Connect to database events
                    Database db = doc.Database;
                    db.ObjectModified += Database_ObjectModified;
                    db.ObjectAppended += Database_ObjectAppended;
                    db.ObjectErased += Database_ObjectErased;

                    // Connect to document events
                    Application.DocumentManager.DocumentActivated += DocumentManager_DocumentActivated;
                    Application.DocumentManager.DocumentCreated += DocumentManager_DocumentCreated;

                    _eventsConnected = true;

                    System.Diagnostics.Debug.WriteLine("MarkNumberManager connected to database events");
                }
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error connecting to events: {ex.Message}");
            }
        }

        // Disconnect from events when necessary
        private void DisconnectFromEvents()
        {
            if (!_eventsConnected) return;

            try
            {
                // Disconnect from database events
                Document doc = Application.DocumentManager.MdiActiveDocument;
                if (doc != null)
                {
                    Database db = doc.Database;
                    db.ObjectModified -= Database_ObjectModified;
                    db.ObjectAppended -= Database_ObjectAppended;
                    db.ObjectErased -= Database_ObjectErased;
                }

                // Disconnect from document events
                Application.DocumentManager.DocumentActivated -= DocumentManager_DocumentActivated;
                Application.DocumentManager.DocumentCreated -= DocumentManager_DocumentCreated;

                _eventsConnected = false;

                System.Diagnostics.Debug.WriteLine("MarkNumberManager disconnected from database events");
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error disconnecting from events: {ex.Message}");
            }
        }

        // Handle document switching
        private void DocumentManager_DocumentActivated(object sender, DocumentCollectionEventArgs e)
        {
            // Reset instance when document changes to ensure we're working with the right database
            _instance = null;
        }

        // Handle new document creation
        private void DocumentManager_DocumentCreated(object sender, DocumentCollectionEventArgs e)
        {
            // Reset instance when a new document is created
            _instance = null;
        }

        // Handle object modification event
        private void Database_ObjectModified(object sender, ObjectEventArgs e)
        {
            try
            {
                // Skip if not from current document
                Document doc = Application.DocumentManager.MdiActiveDocument;
                if (doc == null) return;

                Database eventDb = sender as Database;
                if (eventDb != doc.Database) return;

                // Skip if suppressed
                if (_suppressEvents) return;

                // Get the object ID for later use
                ObjectId objId = e.DBObject.ObjectId;

                // Schedule processing for after the transaction completes
                Application.Idle += delegate
                {
                    Application.Idle -= delegate { };

                    try
                    {
                        _suppressEvents = true;

                        // Process the component directly
                        using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
                        {
                            // Check if the object is still valid
                            if (!objId.IsValid || objId.IsErased)
                            {
                                tr.Commit();
                                _suppressEvents = false;
                                return;
                            }

                            // Process the component
                            ProcessComponentMarkNumbers(objId, tr);
                            tr.Commit();
                        }

                        _suppressEvents = false;
                    }
                    catch (System.Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error in Idle handler: {ex.Message}");
                        _suppressEvents = false;
                    }
                };
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in ObjectModified: {ex.Message}");
            }
        }

        // Handle object appended event
        private void Database_ObjectAppended(object sender, ObjectEventArgs e)
        {
            try
            {
                // Skip if not from current document
                Document doc = Application.DocumentManager.MdiActiveDocument;
                if (doc == null) return;

                Database eventDb = sender as Database;
                if (eventDb != doc.Database) return;

                // Skip if suppressed
                if (_suppressEvents) return;

                // Get the object ID for later use
                ObjectId objId = e.DBObject.ObjectId;

                // Schedule processing for after the transaction completes
                Application.Idle += delegate
                {
                    Application.Idle -= delegate { };

                    try
                    {
                        _suppressEvents = true;

                        // Process the component directly
                        using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
                        {
                            // Check if the object is still valid
                            if (!objId.IsValid || objId.IsErased)
                            {
                                tr.Commit();
                                _suppressEvents = false;
                                return;
                            }

                            // Process the component
                            ProcessComponentMarkNumbers(objId, tr);
                            tr.Commit();
                        }

                        _suppressEvents = false;
                    }
                    catch (System.Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error in Idle handler: {ex.Message}");
                        _suppressEvents = false;
                    }
                };
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in ObjectAppended: {ex.Message}");
            }
        }

        // Handle object erased event
        private void Database_ObjectErased(object sender, ObjectErasedEventArgs e)
        {
            // We don't need to update mark numbers when an object is erased,
            // but we might want to clean up our cache of processed entities
            if (e.Erased)
            {
                try
                {
                    // Get the handle of the erased object
                    string handle = e.DBObject.Handle.ToString();

                    // Remove from our processed set if it exists
                    _processedEntities.Remove(handle);
                }
                catch (System.Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error in ObjectErased: {ex.Message}");
                }
            }
        }

        #region Event Handlers

        // Called when a new component is created
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

                    // Process the new component - pass the transaction
                    ProcessComponentMarkNumbers(entityId, tr);

                    // Add to processed set
                    _processedEntities.Add(handle);
                }
                tr.Commit();
            }
        }

        // Called when a component is modified
        // Called when a component is modified
        public void OnComponentModified(ObjectId entityId)
        {
            using (Transaction tr = _db.TransactionManager.StartTransaction())
            {
                // Update this component and any connected components - pass the transaction
                ProcessComponentMarkNumbers(entityId, tr);

                // Check if this modification affects any attachments - pass the transaction
                UpdateAttachedComponents(entityId, tr);

                tr.Commit();
            }
        }

        // Called when a component is stretched (changing its length)

        // Update the method signature to include an isFromStretch parameter with a default value
        public void OnComponentStretched(ObjectId entityId, double newLength, bool isFromStretch = false)
        {
            using (Transaction tr = _db.TransactionManager.StartTransaction())
            {
                try
                {
                    // Get component data
                    Entity ent = tr.GetObject(entityId, OpenMode.ForRead) as Entity;
                    if (ent == null)
                    {
                        tr.Commit();
                        return;
                    }

                    string handle = ent.Handle.ToString();
                    System.Diagnostics.Debug.WriteLine($"OnComponentStretched called for {handle}, isFromStretch={isFromStretch}");

                    string componentType = GetComponentType(ent, tr);
                    if (string.IsNullOrEmpty(componentType))
                    {
                        tr.Commit();
                        return;
                    }

                    // Get child parts
                    List<ChildPart> childParts = GetChildParts(ent, tr);
                    if (childParts.Count == 0)
                    {
                        tr.Commit();
                        return;
                    }

                    // Check if length change affects any mark numbers
                    bool markNumbersChanged = false;
                    double oldComponentLength = GetComponentLength(entityId, tr);

                    // Only process if the length actually changed
                    if (isFromStretch || Math.Abs(oldComponentLength - newLength) > 0.001)
                    {
                        System.Diagnostics.Debug.WriteLine($"Component length changed: {oldComponentLength} -> {newLength}");

                        foreach (var part in childParts)
                        {
                            // Skip fixed length parts as they aren't affected by stretching
                            if (part.IsFixedLength)
                            {
                                System.Diagnostics.Debug.WriteLine($"  Skipping fixed length part: {part.Name}");
                                continue;
                            }

                            // Calculate the old and new actual lengths
                            double oldActualLength = CalculateActualPartLength(part, oldComponentLength);
                            double newActualLength = CalculateActualPartLength(part, newLength);

                            System.Diagnostics.Debug.WriteLine($"  Part {part.Name}: Length {oldActualLength} -> {newActualLength}");

                            // If length changed significantly, update mark number
                            if (Math.Abs(oldActualLength - newActualLength) > 0.001)
                            {
                                string oldMarkNumber = part.MarkNumber;

                                // Generate new mark number
                                if (componentType == "Horizontal")
                                {
                                    string uniqueKey = GenerateHorizontalPartKey(part, newActualLength);
                                    part.MarkNumber = GetOrCreateMarkNumber(uniqueKey, part.PartType, _horizontalPartMarks);
                                }
                                else if (componentType == "Vertical")
                                {
                                    // For vertical components, we need attachment information
                                    var attachments = GetAttachmentsForComponent(handle, true, tr);

                                    string uniqueKey = GenerateVerticalPartKey(part, newActualLength, attachments);
                                    part.MarkNumber = GetOrCreateMarkNumber(uniqueKey, part.PartType, _verticalPartMarks);
                                }

                                System.Diagnostics.Debug.WriteLine($"  Mark number changed: {oldMarkNumber} -> {part.MarkNumber}");
                                markNumbersChanged = true;
                            }
                        }
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine($"Component length unchanged: {oldComponentLength}");
                    }

                    // If any mark numbers changed, save the updated child parts
                    if (markNumbersChanged)
                    {
                        ent.Dispose(); // Close the read version

                        Entity entForWrite = tr.GetObject(entityId, OpenMode.ForWrite) as Entity;
                        string json = JsonConvert.SerializeObject(childParts);
                        SaveChildPartsToEntity(entForWrite, json, tr);

                        System.Diagnostics.Debug.WriteLine($"Saved updated mark numbers after stretch for {handle}");
                    }

                    // Check if this modification affects any attachments
                    if (componentType == "Horizontal")
                    {
                        UpdateAttachedComponents(entityId, tr);
                    }

                    tr.Commit();
                }
                catch (System.Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error in OnComponentStretched: {ex.Message}");
                    System.Diagnostics.Debug.WriteLine($"Stack trace: {ex.StackTrace}");
                    tr.Abort();
                }
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
                                                var attachments = GetAttachmentsForComponent(handle, true, tr);
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

        //private void RegisterXdataApps()
        //{
        //    using (Transaction tr = _db.TransactionManager.StartTransaction())
        //    {
        //        // Get the RegAppTable
        //        RegAppTable regTable = (RegAppTable)tr.GetObject(_db.RegAppTableId, OpenMode.ForWrite);

        //        // Register all the application names we might use
        //        RegisterApp(regTable, "METALCOMP", tr);
        //        RegisterApp(regTable, "METALPARTSINFO", tr);
        //        RegisterApp(regTable, "METALPARTS", tr);

        //        // Register chunk apps preemptively
        //        for (int i = 0; i < 10; i++) // Assume we won't need more than 10 chunks
        //        {
        //            RegisterApp(regTable, $"METALPARTS{i}", tr);
        //        }

        //        tr.Commit();
        //    }
        //}

        //private void RegisterApp(RegAppTable regTable, string appName, Transaction tr)
        //{
        //    if (!regTable.Has(appName))
        //    {
        //        RegAppTableRecord record = new RegAppTableRecord();
        //        record.Name = appName;
        //        regTable.Add(record);
        //        tr.AddNewlyCreatedDBObject(record, true);
        //    }
        //}

        // Process mark numbers for a component - made internal for use by the commands
        internal void ProcessComponentMarkNumbers(ObjectId entityId, Transaction tr)
        {
            try
            {
                Entity ent = tr.GetObject(entityId, OpenMode.ForRead) as Entity;
                if (ent == null) return;

                string handle = ent.Handle.ToString();

                string componentType = GetComponentType(ent, tr);
                if (string.IsNullOrEmpty(componentType)) return;

                List<ChildPart> childParts = GetChildParts(ent, tr);
                if (childParts.Count == 0) return;

                // Get the current length
                double currentLength = 0;
                if (ent is Polyline)
                {
                    Polyline pline = ent as Polyline;
                    currentLength = pline.Length;
                }

                // Check if length has changed
                bool lengthChanged = false;
                if (_lastComponentLengths.ContainsKey(handle))
                {
                    double lastLength = _lastComponentLengths[handle];
                    lengthChanged = Math.Abs(lastLength - currentLength) > 0.001;

                    if (lengthChanged)
                    {
                        System.Diagnostics.Debug.WriteLine($"Length changed for {handle}: {lastLength} -> {currentLength}");
                    }
                }

                // Store current length
                _lastComponentLengths[handle] = currentLength;

                // Only process if this is a new component (not in our cache)
                // or if the length has changed
                bool isNewComponent = !_processedEntities.Contains(handle);

                if (isNewComponent || lengthChanged)
                {
                    System.Diagnostics.Debug.WriteLine($"Processing marks for {handle}: {(isNewComponent ? "new component" : "length changed")}");

                    bool markNumbersChanged = false;

                    foreach (var part in childParts)
                    {
                        // Skip shop use parts
                        if (part.IsShopUse) continue;

                        // Skip fixed length parts if we're just handling a length change
                        if (lengthChanged && part.IsFixedLength) continue;

                        // Calculate actual length
                        double actualLength = CalculateActualPartLength(part, currentLength);

                        // Store original mark number for comparison
                        string oldMarkNumber = part.MarkNumber;

                        // Generate mark number based on component type
                        if (componentType == "Horizontal")
                        {
                            string uniqueKey = GenerateHorizontalPartKey(part, actualLength);
                            part.MarkNumber = GetOrCreateMarkNumber(uniqueKey, part.PartType, _horizontalPartMarks);
                        }
                        else if (componentType == "Vertical")
                        {
                            var attachments = GetAttachmentsForComponent(handle, true, tr);
                            string uniqueKey = GenerateVerticalPartKey(part, actualLength, attachments);
                            part.MarkNumber = GetOrCreateMarkNumber(uniqueKey, part.PartType, _verticalPartMarks);
                        }

                        // Check if mark number changed
                        if (oldMarkNumber != part.MarkNumber)
                        {
                            System.Diagnostics.Debug.WriteLine($"  Mark number changed for {part.Name}: {oldMarkNumber} -> {part.MarkNumber}");
                            markNumbersChanged = true;
                        }
                    }

                    // If any mark numbers changed, save the updated child parts
                    if (markNumbersChanged)
                    {
                        // Need to close entity opened for read
                        ent.Dispose();

                        // Re-open for write
                        Entity entForWrite = tr.GetObject(entityId, OpenMode.ForWrite) as Entity;
                        string json = JsonConvert.SerializeObject(childParts);
                        SaveChildPartsToEntity(entForWrite, json, tr);

                        System.Diagnostics.Debug.WriteLine($"Mark numbers saved for {handle}");

                        // Update our cache
                        _processedEntities.Add(handle);
                    }
                }
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in ProcessComponentMarkNumbers: {ex.Message}");
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
            var attachments = GetAttachmentsForComponent(handle, false, tr);  // Pass transaction

            if (attachments.Count > 0)
            {
                // Get the ObjectIds for all affected vertical components
                List<ObjectId> verticalIds = new List<ObjectId>();

                foreach (var attachment in attachments)
                {
                    // Find the vertical component
                    ObjectId verticalId = GetEntityByHandle(attachment.VerticalHandle, ent.Database);  // Pass database

                    if (verticalId != ObjectId.Null && !verticalIds.Contains(verticalId))
                    {
                        verticalIds.Add(verticalId);
                    }
                }

                // Process mark numbers for all affected vertical components
                foreach (ObjectId verticalId in verticalIds)
                {
                    ProcessComponentMarkNumbers(verticalId, tr);  // Pass transaction
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
        private List<Attachment> GetAttachmentsForComponent(string handle, bool isVertical, Transaction tr)
        {
            List<Attachment> allAttachments = LoadAttachmentsFromDrawing(tr);

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
        private ObjectId GetEntityByHandle(string handle, Database db)
        {
            ObjectId result = ObjectId.Null;

            try
            {
                Handle h = new Handle(Convert.ToInt64(handle, 16));
                result = db.GetObjectId(false, h, 0);
            }
            catch
            {
                // Handle not found or invalid
            }

            return result;
        }

        // Get the component type from an entity - made internal for use by commands
        internal string GetComponentType(Entity ent, Transaction tr = null)
        {
            bool ownsTransaction = (tr == null);

            try
            {
                if (ownsTransaction)
                {
                    tr = _db.TransactionManager.StartTransaction();
                }

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

                if (ownsTransaction)
                {
                    tr.Commit();
                }

                return null;
            }
            catch
            {
                if (ownsTransaction)
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

        // Get child parts from an entity
        private List<ChildPart> GetChildParts(Entity ent, Transaction tr = null)
        {
            bool ownsTransaction = (tr == null);
            List<ChildPart> result = new List<ChildPart>();

            try
            {
                if (ownsTransaction)
                {
                    tr = _db.TransactionManager.StartTransaction();
                }

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
                            Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(
                                $"\nError parsing child parts JSON: {ex.Message}");
                        }
                    }
                }

                if (ownsTransaction)
                {
                    tr.Commit();
                }

                return result;
            }
            catch
            {
                if (ownsTransaction)
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

        private double GetComponentLength(ObjectId entityId, Transaction tr = null)
        {
            double length = 0;
            bool ownsTransaction = (tr == null);

            try
            {
                if (ownsTransaction)
                {
                    tr = _db.TransactionManager.StartTransaction();
                }

                Entity ent = tr.GetObject(entityId, OpenMode.ForRead) as Entity;
                if (ent is Polyline)
                {
                    Polyline pline = ent as Polyline;
                    length = pline.Length;
                }

                if (ownsTransaction)
                {
                    tr.Commit();
                }
            }
            catch
            {
                if (ownsTransaction)
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

            return length;
        }

        // Save child parts to an entity
        private void SaveChildPartsToEntity(Entity ent, string json, Transaction tr)
        {
            // Register the app names first
            RegAppTable regTable = (RegAppTable)tr.GetObject(ent.Database.RegAppTableId, OpenMode.ForWrite);

            // Register METALPARTSINFO app
            RegisterApp(regTable, "METALPARTSINFO", tr);

            // Calculate chunks and register chunk apps
            const int maxChunkSize = 250;
            int chunkCount = (int)Math.Ceiling((double)json.Length / maxChunkSize);

            for (int i = 0; i < chunkCount; i++)
            {
                RegisterApp(regTable, $"METALPARTS{i}", tr);
            }

            // Write METALPARTSINFO with chunk count
            using (ResultBuffer rbInfo = new ResultBuffer(
                new TypedValue((int)DxfCode.ExtendedDataRegAppName, "METALPARTSINFO"),
                new TypedValue((int)DxfCode.ExtendedDataInteger32, chunkCount)))
            {
                ent.XData = rbInfo;
            }

            // Write chunks
            for (int i = 0; i < chunkCount; i++)
            {
                int startIndex = i * maxChunkSize;
                int length = Math.Min(maxChunkSize, json.Length - startIndex);
                string chunk = json.Substring(startIndex, length);

                using (ResultBuffer rb = new ResultBuffer(
                    new TypedValue((int)DxfCode.ExtendedDataRegAppName, $"METALPARTS{i}"),
                    new TypedValue((int)DxfCode.ExtendedDataAsciiString, chunk)))
                {
                    ent.XData = rb;
                }
            }
        }

        // Load attachments from drawing
        private List<Attachment> LoadAttachmentsFromDrawing(Transaction tr)
        {
            List<Attachment> attachments = new List<Attachment>();

            // Use the stored database reference
            Database db = _db;

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
                            attachments = JsonConvert.DeserializeObject<List<Attachment>>(json);
                        }
                    }
                }
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
                                        manager.SaveChildPartsToEntity(entForWrite, json, tr);
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
            Database db = Application.DocumentManager.MdiActiveDocument.Database;

            try
            {
                ed.WriteMessage("\nGenerating mark numbers for all components...");

                MarkNumberManager manager = Instance;

                // First register all apps in a separate transaction
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    RegAppTable regTable = (RegAppTable)tr.GetObject(db.RegAppTableId, OpenMode.ForWrite);

                    // Register apps
                    RegisterApp(regTable, "METALCOMP", tr);
                    RegisterApp(regTable, "METALPARTSINFO", tr);

                    for (int i = 0; i < 10; i++)
                    {
                        RegisterApp(regTable, $"METALPARTS{i}", tr);
                    }

                    tr.Commit();
                }

                // First process horizontal components
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

                        foreach (SelectedObject selObj in ss)
                        {
                            ObjectId id = selObj.ObjectId;
                            Entity ent = tr.GetObject(id, OpenMode.ForRead) as Entity;

                            if (ent != null)
                            {
                                string componentType = manager.GetComponentType(ent, tr);

                                if (componentType == "Horizontal")
                                {
                                    manager.ProcessComponentMarkNumbers(id, tr);
                                }
                            }
                        }
                    }

                    tr.Commit();
                }

                // Then process vertical components
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

                        foreach (SelectedObject selObj in ss)
                        {
                            ObjectId id = selObj.ObjectId;
                            Entity ent = tr.GetObject(id, OpenMode.ForRead) as Entity;

                            if (ent != null)
                            {
                                string componentType = manager.GetComponentType(ent, tr);

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
                ed.WriteMessage($"\nError: {ex.Message}");
                ed.WriteMessage($"\nStack trace: {ex.StackTrace}");
            }
        }

        private static void RegisterApp(RegAppTable regTable, string appName, Transaction tr)
        {
            if (!regTable.Has(appName))
            {
                RegAppTableRecord record = new RegAppTableRecord();
                record.Name = appName;
                regTable.Add(record);
                tr.AddNewlyCreatedDBObject(record, true);
            }
        }

        // Add this method at the end of your class
        public void Cleanup()
        {
            DisconnectFromEvents();
        }

        #endregion
    }
}