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
using System.ComponentModel;

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
        private HashSet<string> _currentlyProcessingHandles = new HashSet<string>();
        private HashSet<ObjectId> _pendingIdleProcessing = new HashSet<ObjectId>();
        private bool _idleHandlerAttached = false;
        private HashSet<long> _processedInvalidIds = new HashSet<long>();
        private static HashSet<long> _invalidObjectIds = new HashSet<long>();
        private Action<ObjectId> _markNumbersProcessedCallback;

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

        private bool _suppressEvents = false;

        private Dictionary<string, double> _lastComponentLengths = new Dictionary<string, double>();



        #region Event Handlers

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
                    //db.ObjectModified += Database_ObjectModified;
                    //db.ObjectAppended += Database_ObjectAppended;
                    //db.ObjectErased += Database_ObjectErased;

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
                    //db.ObjectModified -= Database_ObjectModified;
                    //db.ObjectAppended -= Database_ObjectAppended;
                    //db.ObjectErased -= Database_ObjectErased;
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

        //// Handle object modification event
        //private void Database_ObjectModified(object sender, ObjectEventArgs e)
        //{
        //    try
        //    {
        //        // Skip if suppressed
        //        if (_suppressEvents) return;

        //        // Skip if not from current document
        //        Document doc = Application.DocumentManager.MdiActiveDocument;
        //        if (doc == null) return;

        //        Database eventDb = sender as Database;
        //        if (eventDb != doc.Database) return;



        //        // Get the object ID for later use
        //        ObjectId objId = e.DBObject.ObjectId;

        //        // Skip if we already know it's invalid
        //        if (_invalidObjectIds.Contains(objId.Handle.Value)) return;

        //        // Skip if this object is already being processed
        //        if (e.DBObject is Entity ent && _currentlyProcessingHandles.Contains(ent.Handle.ToString()))
        //        {
        //            return;
        //        }

        //        // Add to pending set and attach idle handler if needed
        //        lock (_pendingIdleProcessing)
        //        {
        //            _pendingIdleProcessing.Add(objId);

        //            if (!_idleHandlerAttached)
        //            {
        //                Application.Idle += Application_Idle;
        //                _idleHandlerAttached = true;
        //            }
        //        }
        //    }
        //    catch (System.Exception ex)
        //    {
        //        System.Diagnostics.Debug.WriteLine($"Error in ObjectModified: {ex.Message}");
        //    }
        //}

        // In the MarkNumberManager class, add a method to safely remove invalid ObjectIds
        private void RemoveInvalidObjectId(ObjectId objId)
        {
            lock (_pendingIdleProcessing)
            {
                _pendingIdleProcessing.Remove(objId);

                // Check if we need to clean up garbage collected ObjectIds
                // This is a more aggressive cleanup that might help with persistent issues
                List<ObjectId> toRemove = new List<ObjectId>();
                foreach (var id in _pendingIdleProcessing)
                {
                    if (!id.IsValid || id.IsErased)
                    {
                        toRemove.Add(id);
                    }
                }

                foreach (var id in toRemove)
                {
                    _pendingIdleProcessing.Remove(id);
                }
            }
        }

        // Then update the Application_Idle method to call this
        private void Application_Idle(object sender, EventArgs e)
        {
            try
            {
                // First, detach the handler to prevent re-entry
                Application.Idle -= Application_Idle;
                _idleHandlerAttached = false;

                // Get a copy of the pending IDs and clear the original
                ObjectId[] idsToProcess;
                lock (_pendingIdleProcessing)
                {
                    if (_pendingIdleProcessing.Count == 0)
                    {
                        return;
                    }

                    idsToProcess = _pendingIdleProcessing.ToArray();
                    _pendingIdleProcessing.Clear();
                }

                _suppressEvents = true;

                // Process the components directly
                Document doc = Application.DocumentManager.MdiActiveDocument;
                if (doc != null)
                {
                    using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
                    {
                        foreach (ObjectId objId in idsToProcess)
                        {
                            // Convert ID to a long value for tracking
                            long objIdValue = objId.Handle.Value;

                            // Skip if we've already found this ID to be invalid
                            if (_processedInvalidIds.Contains(objIdValue))
                            {
                                continue;
                            }

                            // Check if the object is still valid
                            if (objId == ObjectId.Null || !objId.IsValid || objId.IsErased)
                            {
                                System.Diagnostics.Debug.WriteLine($"Skipping invalid ObjectId {objId}");
                                _processedInvalidIds.Add(objIdValue);
                                continue;
                            }

                            try
                            {
                                // Check if the entity actually exists
                                Entity ent = tr.GetObject(objId, OpenMode.ForRead, false, true) as Entity;
                                if (ent == null)
                                {
                                    System.Diagnostics.Debug.WriteLine($"Entity is null for {objId}");
                                    _processedInvalidIds.Add(objIdValue);
                                    continue;
                                }

                                // Process the component
                                ProcessComponentMarkNumbers(objId, tr);
                            }
                            catch (Autodesk.AutoCAD.Runtime.Exception acEx)
                            {
                                System.Diagnostics.Debug.WriteLine($"AutoCAD error processing {objId}: {acEx.Message}");
                                _processedInvalidIds.Add(objIdValue);
                            }
                            catch (System.Exception ex)
                            {
                                System.Diagnostics.Debug.WriteLine($"Error processing {objId}: {ex.Message}");
                                _processedInvalidIds.Add(objIdValue);
                            }
                        }

                        tr.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in Application_Idle: {ex.Message}");
            }
            finally
            {
                _suppressEvents = false;
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

            // Get components with parts using the centralized manager
            List<DrawingComponentManager.Component> components =
                DrawingComponentManager.Instance.GetComponentsWithAttachments();

            foreach (var comp in components)
            {
                string handle = comp.Handle;
                string componentType = comp.Type;
                double length = comp.Length;

                // Process each part
                foreach (var part in comp.Parts)
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

                            // Calculate actual length
                            double actualLength = part.IsFixedLength ? part.FixedLength :
                                                  length + part.StartAdjustment + part.EndAdjustment;

                            // Generate and store unique keys
                            if (componentType == "Horizontal")
                            {
                                string uniqueKey = GenerateHorizontalPartKey(part, actualLength);
                                _horizontalPartMarks[uniqueKey] = part.MarkNumber;
                            }
                            else if (componentType == "Vertical")
                            {
                                // Convert attachments to the format expected by GenerateVerticalPartKey
                                List<Attachment> attachments = new List<Attachment>();
                                foreach (var attachment in part.Attachments)
                                {
                                    attachments.Add(new Attachment
                                    {
                                        HorizontalPartType = attachment.AttachedPartType,
                                        Side = attachment.Side,
                                        Position = attachment.Position,
                                        Height = attachment.Height,
                                        Invert = attachment.Invert,
                                        Adjust = attachment.Adjust
                                    });
                                }

                                string uniqueKey = GenerateVerticalPartKey(part, actualLength, attachments);
                                _verticalPartMarks[uniqueKey] = part.MarkNumber;
                            }
                        }
                    }

                    // Add to processed entities
                    _processedEntities.Add(handle);
                }
            }

            // Log how many mark numbers we loaded
            Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage(
                $"\nLoaded {_horizontalPartMarks.Count} horizontal part marks and {_verticalPartMarks.Count} vertical part marks.");
        }

        // Display mark numbers when ProcessComponentMarkNumbers is called
        public void RegisterMarkNumbersProcessedCallback(Action<ObjectId> callback)
        {
            _markNumbersProcessedCallback = callback;
        }

        // Process mark numbers for a component - made internal for use by the commands
        internal void ProcessComponentMarkNumbers(ObjectId entityId, Transaction tr, bool forceProcess = false)
        {
            try
            {
                // Check if this ID is already known to be invalid
                if (_invalidObjectIds.Contains(entityId.Handle.Value))
                {
                    //System.Diagnostics.Debug.WriteLine($"Skipping known invalid ObjectId {entityId}");
                    return;
                }

                // Check if the ID is valid
                if (!entityId.IsValid || entityId.IsErased)
                {
                    System.Diagnostics.Debug.WriteLine($"Adding invalid ObjectId to ignore list: {entityId}");
                    _invalidObjectIds.Add(entityId.Handle.Value);
                    return;
                }

                // Check if the ID is valid before attempting to access it
                if (entityId == ObjectId.Null || !entityId.IsValid || entityId.IsErased)
                {
                    System.Diagnostics.Debug.WriteLine($"ProcessComponentMarkNumbers: Invalid ObjectId {entityId}");
                    return;
                }

                // Now try to get the entity with error handling
                Entity ent;
                try
                {
                    ent = tr.GetObject(entityId, OpenMode.ForRead) as Entity;
                }
                catch
                {
                    System.Diagnostics.Debug.WriteLine($"Error getting entity {entityId}, adding to invalid list");
                    _invalidObjectIds.Add(entityId.Handle.Value);
                    return;
                }

                if (ent == null)
                {
                    System.Diagnostics.Debug.WriteLine($"Entity is null for {entityId}, adding to invalid list");
                    _invalidObjectIds.Add(entityId.Handle.Value);
                    return;
                }

                string handle = ent.Handle.ToString();


                // Check if we're already processing this handle to prevent recursion
                if (_currentlyProcessingHandles.Contains(handle))
                {
                    System.Diagnostics.Debug.WriteLine($"Skipping recursive mark number processing for handle: {handle}");
                    return;
                }

                // Add to processing set
                _currentlyProcessingHandles.Add(handle);

                System.Diagnostics.Debug.WriteLine($"Processing mark numbers for entity with handle: {handle}");
                try
                {

                    string componentType = GetComponentType(ent, tr);
                    if (string.IsNullOrEmpty(componentType))
                    {
                        System.Diagnostics.Debug.WriteLine($"No component type for {handle}");
                        return;
                    }

                    List<ChildPart> childParts = GetChildParts(ent, tr);
                    if (childParts.Count == 0)
                    {
                        System.Diagnostics.Debug.WriteLine($"No child parts for {handle}");
                        return;
                    }
                    System.Diagnostics.Debug.WriteLine($"Found {childParts.Count} parts");

                    // Get the current length
                    double currentLength = 0;
                    if (ent is Polyline)
                    {
                        Polyline pline = ent as Polyline;
                        currentLength = pline.Length;
                    }
                    System.Diagnostics.Debug.WriteLine($"Part is {currentLength} long");

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
                    // or if forceProcess is true
                    bool isNewComponent = !_processedEntities.Contains(handle);

                    // We'll always process marks for all parts, but our decision to modify the
                    // entity will depend on whether any mark numbers changed
                    bool shouldProcessMarks = isNewComponent || lengthChanged || forceProcess;

                    if (shouldProcessMarks)
                    {
                        System.Diagnostics.Debug.WriteLine($"Processing marks for {handle}: {(isNewComponent ? "new component" : forceProcess ? "forced update" : "length changed")}");

                        bool markNumbersChanged = false;

                        foreach (var part in childParts)
                        {
                            // Skip shop use parts
                            //if (part.IsShopUse) continue;

                            // Calculate actual length
                            double actualLength = CalculateActualPartLength(part, currentLength);

                            // Store original mark number for comparison
                            string oldMarkNumber = part.MarkNumber;

                            // Generate mark number based on component type
                            if (componentType == "Horizontal")
                            {
                                string uniqueKey = GenerateHorizontalPartKey(part, actualLength);
                                System.Diagnostics.Debug.WriteLine($"  Horizontal key for {part.Name}: {uniqueKey}");
                                part.MarkNumber = GetOrCreateMarkNumber(uniqueKey, part.PartType, _horizontalPartMarks);
                            }
                            else if (componentType == "Vertical")
                            {
                                var attachments = GetAttachmentsForComponent(handle, true, tr);
                                string uniqueKey = GenerateVerticalPartKey(part, actualLength, attachments);
                                System.Diagnostics.Debug.WriteLine($"  Vertical key for {part.Name}: {uniqueKey}");
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


                            // Then update the display for this component
                            MarkNumberDisplay.Instance.UpdateMarkNumbersForComponent(entityId);

                        }
                        else
                        {
                            System.Diagnostics.Debug.WriteLine($"No mark numbers changed for {handle}");
                        }
                    }
                    else
                    {
                        System.Diagnostics.Debug.WriteLine($"Skipping mark number processing for {handle}");
                    }
                }
                catch (System.Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error processing mark numbers for {handle}: {ex.Message}");
                }
                finally
                {
                    // Remove from processing set
                    _currentlyProcessingHandles.Remove(handle);
                }
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in ProcessComponentMarkNumbers: {ex.Message}");
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
            string key = $"{part.PartType}|{roundedLength}|{part.Finish}|{part.Fab}";

            System.Diagnostics.Debug.WriteLine($"Generated horizontal part key: {key}");

            return key;
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


        // Save child parts to an entity
        private void SaveChildPartsToEntity(Entity ent, string json, Transaction tr)
        {
            // Check if the JSON is different from the existing data
            string existingJson = GetExistingPartsJson(ent);
            if (existingJson == json)
            {
                System.Diagnostics.Debug.WriteLine($"No changes to parts data for handle: {ent.Handle}");
                return; // No changes needed
            }

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

        // Add this helper method
        private string GetExistingPartsJson(Entity ent)
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

                return jsonBuilder.ToString();
            }

            return string.Empty;
        }

        private List<Attachment> LoadAttachmentsFromDrawing(Transaction tr)
        {
            // Use the centralized method with the transaction
            return DrawingComponentManager.LoadAttachmentsFromDrawing(tr);
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

        

        public static void GenerateMarkNumbers()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            try
            {
                ed.WriteMessage("\nGenerating mark numbers using the fixed method...");

                MarkNumberManager manager = MarkNumberManager.Instance;

                // First, load attachments to ensure they're available for mark number generation
                List<Attachment> attachments = new List<Attachment>();

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    // Get named objects dictionary
                    DBDictionary nod = (DBDictionary)tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForRead);

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
                                ed.WriteMessage($"\nLoaded {attachments.Count} attachments from drawing.");
                            }
                        }
                    }

                    tr.Commit();
                }

                // Important: Force refresh of the mark number manager
                // This might be what's missing in your current implementation
                Type managerType = manager.GetType();
                var initializeMethod = managerType.GetMethod("InitializeFromDrawing",
                    System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Instance);

                if (initializeMethod != null)
                {
                    ed.WriteMessage("\nReinitializing mark number manager...");
                    initializeMethod.Invoke(manager, null);
                }

                // Process all components in a single transaction for consistency
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
                        ed.WriteMessage($"\nFound {selRes.Value.Count} polylines to process.");

                        int processedCount = 0;

                        foreach (SelectedObject selObj in selRes.Value)
                        {
                            ObjectId id = selObj.ObjectId;
                            Entity ent = tr.GetObject(id, OpenMode.ForRead) as Entity;

                            if (ent != null)
                            {
                                // Check if it's a metal component
                                ResultBuffer rbComp = ent.GetXDataForApplication("METALCOMP");
                                if (rbComp != null)
                                {
                                    // Force process this component's mark numbers
                                    manager.ProcessComponentMarkNumbers(id, tr, true);
                                    processedCount++;
                                }
                            }
                        }

                        ed.WriteMessage($"\nProcessed mark numbers for {processedCount} components.");
                    }

                    tr.Commit();
                }

                // Update the display
                ed.WriteMessage("\nUpdating mark number displays...");
                MarkNumberDisplay.Instance.UpdateAllMarkNumbers();

                ed.WriteMessage("\nMark number generation completed using the fixed method.");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in fixed mark number generation: {ex.Message}");
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