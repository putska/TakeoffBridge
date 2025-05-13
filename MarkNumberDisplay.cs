using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.EditorInput;
using Newtonsoft.Json;
using System.Text;

namespace TakeoffBridge
{
    /// <summary>
    /// Class to manage mark number text display in AutoCAD
    /// </summary>
    public class MarkNumberDisplay
    {
        // Singleton instance
        private static MarkNumberDisplay _instance;

        // Dictionary to track MText objects for each component
        private Dictionary<string, List<MarkNumberText>> _markNumberTexts = new Dictionary<string, List<MarkNumberText>>();

        // Dictionary to store custom offsets for mark numbers
        private Dictionary<string, Dictionary<string, Point3d>> _customOffsets = new Dictionary<string, Dictionary<string, Point3d>>();

        // Track if we're suppressing events to prevent recursion
        private bool _suppressEvents = false;

        // Default formatting values
        private double _textHeight = 2.25;
        private string _textStyleName = "Arial Narrow";
        private double _defaultHorizontalOffset = 5.0;
        private double _defaultVerticalOffset = 5.0;
        private double _defaultSpacing = 2.5;

        // Constructor is private for singleton pattern
        private MarkNumberDisplay()
        {
            // Register required application names for XData
            RegisterApplicationNames();

            // Load saved custom positions from the drawing
            LoadCustomPositions();

            // Connect to events
            ConnectToEvents();
        }

        // Singleton access
        public static MarkNumberDisplay Instance
        {
            get
            {
                Document doc = Application.DocumentManager.MdiActiveDocument;
                if (_instance == null)
                {
                    _instance = new MarkNumberDisplay();
                }
                return _instance;
            }
        }

        #region Event Handling

        private bool _eventsConnected = false;

        private void ConnectToEvents()
        {
            if (_eventsConnected) return;

            try
            {
                // Connect to document events
                Document doc = Application.DocumentManager.MdiActiveDocument;
                //if (doc != null)
                //{
                //    // Connect to database events for detecting changes to components
                //    Database db = doc.Database;
                //    db.ObjectModified += Database_ObjectModified;
                //    db.ObjectErased += Database_ObjectErased;

                //    // Connect to MText events for detecting when user moves mark texts
                //    db.ObjectModified += Database_MTextModified;

                //    // Connect to document events
                //    Application.DocumentManager.DocumentActivated += DocumentManager_DocumentActivated;
                //    Application.DocumentManager.DocumentCreated += DocumentManager_DocumentCreated;

                //    _eventsConnected = true;

                //    System.Diagnostics.Debug.WriteLine("MarkNumberDisplay connected to events");
                //}
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error connecting to events: {ex.Message}");
            }
        }

        private void DisconnectFromEvents()
        {
            if (!_eventsConnected) return;

            try
            {
                // Disconnect from database events
                Document doc = Application.DocumentManager.MdiActiveDocument;
                //if (doc != null)
                //{
                //    Database db = doc.Database;
                //    db.ObjectModified -= Database_ObjectModified;
                //    db.ObjectErased -= Database_ObjectErased;
                //    db.ObjectModified -= Database_MTextModified;
                //}

                //// Disconnect from document events
                //Application.DocumentManager.DocumentActivated -= DocumentManager_DocumentActivated;
                //Application.DocumentManager.DocumentCreated -= DocumentManager_DocumentCreated;

                //_eventsConnected = false;

                //System.Diagnostics.Debug.WriteLine("MarkNumberDisplay disconnected from events");
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error disconnecting from events: {ex.Message}");
            }
        }

        private void DocumentManager_DocumentActivated(object sender, DocumentCollectionEventArgs e)
        {
            // Reset the display when a document is activated
            _instance = null;
        }

        private void DocumentManager_DocumentCreated(object sender, DocumentCollectionEventArgs e)
        {
            // Reset the display when a new document is created
            _instance = null;
        }

        private void Database_ObjectModified(object sender, ObjectEventArgs e)
        {
            if (_suppressEvents) return;

            try
            {
                Document doc = Application.DocumentManager.MdiActiveDocument;
                //if (doc == null) return;

                //// Check if this is a component we care about
                //if (!(e.DBObject is Polyline)) return;

                //Polyline polyline = e.DBObject as Polyline;

                //// Check if it has METALCOMP Xdata
                //ResultBuffer rbComp = polyline.GetXDataForApplication("METALCOMP");
                //if (rbComp == null) return;

                //// This is a metal component, update its mark numbers
                //string handle = polyline.Handle.ToString();

                //// Schedule an idle-time update
                //Application.Idle += delegate
                //{
                //    // Make sure we only run once
                //    Application.Idle -= delegate { };

                //    try
                //    {
                //        _suppressEvents = true;
                //        UpdateMarkNumbersForComponent(polyline.ObjectId);
                //    }
                //    finally
                //    {
                //        _suppressEvents = false;
                //    }
                //};
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in Database_ObjectModified: {ex.Message}");
            }
        }

       

        #endregion

        #region Mark Number Display Methods

        /// <summary>
        /// Updates or creates mark number text objects for the specified component
        /// </summary>
        public void UpdateMarkNumbersForComponent(ObjectId componentId)
        {
            // Register application names
            RegisterApplicationNames();

            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
            {
                try
                {
                    // Get the component
                    Entity ent = tr.GetObject(componentId, OpenMode.ForRead) as Entity;
                    if (!(ent is Polyline))
                    {
                        tr.Commit();
                        return;
                    }

                    Polyline pline = ent as Polyline;
                    string handle = pline.Handle.ToString();

                    // Get component type
                    string componentType = GetComponentType(pline);
                    if (string.IsNullOrEmpty(componentType))
                    {
                        tr.Commit();
                        return;
                    }

                    // Check component orientation
                    bool isHorizontal = !componentType.ToUpper().Contains("VERTICAL");

                    // Get child parts
                    List<ChildPart> childParts = GetChildParts(pline);

                    // Filter to only non-shop-use parts with mark numbers
                    var partsToDisplay = childParts
                        .Where(p => !p.IsShopUse && !string.IsNullOrEmpty(p.MarkNumber))
                        .ToList();

                    System.Diagnostics.Debug.WriteLine($"Processing {partsToDisplay.Count} parts for component {handle}");

                    // Get existing mark texts for this component
                    List<MarkNumberText> existingTexts = _markNumberTexts.ContainsKey(handle)
                        ? _markNumberTexts[handle]
                        : new List<MarkNumberText>();

                    // Create a new list for updated texts
                    List<MarkNumberText> updatedTexts = new List<MarkNumberText>();

                    // Use a simple global index for positioning across all parts
                    int globalIndex = 0;

                    // Process each part - now using a global index
                    foreach (var part in partsToDisplay)
                    {
                        // Check if we already have a text for this part
                        MarkNumberText existingText = existingTexts.FirstOrDefault(t => t.PartType == part.PartType);

                        if (existingText != null)
                        {
                            // Update the existing text
                            MText mtext = tr.GetObject(existingText.TextObjectId, OpenMode.ForWrite) as MText;
                            if (mtext != null)
                            {
                                // Update the content if needed
                                if (mtext.Contents != part.MarkNumber)
                                {
                                    mtext.Contents = part.MarkNumber;
                                }

                                // Keep the existing text object ID
                                updatedTexts.Add(new MarkNumberText
                                {
                                    TextObjectId = existingText.TextObjectId,
                                    PartType = part.PartType,
                                    MarkNumber = part.MarkNumber
                                });
                            }
                        }
                        else
                        {
                            // Create a new text for this part - use the global index
                            System.Diagnostics.Debug.WriteLine($"Creating new text for {part.PartType} with index {globalIndex}");
                            ObjectId textId = CreateMarkText(tr, part.MarkNumber, pline, part.PartType, isHorizontal, globalIndex);

                            // Add to list of updated texts
                            updatedTexts.Add(new MarkNumberText
                            {
                                TextObjectId = textId,
                                PartType = part.PartType,
                                MarkNumber = part.MarkNumber
                            });

                            // Increment global index after creating a new text
                            globalIndex++;
                        }
                    }

                    // Remove any texts that no longer have parts
                    foreach (var existingText in existingTexts)
                    {
                        if (!partsToDisplay.Any(p => p.PartType == existingText.PartType))
                        {
                            // This part no longer exists or is now shop use - erase the text
                            MText mtext = tr.GetObject(existingText.TextObjectId, OpenMode.ForWrite) as MText;
                            if (mtext != null)
                            {
                                mtext.Erase();
                            }
                        }
                    }

                    // Update our tracking dictionary
                    _markNumberTexts[handle] = updatedTexts;

                    tr.Commit();
                }
                catch (System.Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error in UpdateMarkNumbersForComponent: {ex.Message}");
                    tr.Abort();
                }
            }
        }

        public static void ClearAttachments()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            try
            {
                ed.WriteMessage("\nClearing all attachments...");

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    // Get named objects dictionary
                    DBDictionary nod = (DBDictionary)tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForWrite);

                    // Check if attachment entry exists
                    const string dictName = "METALATTACHMENTS";
                    if (nod.Contains(dictName))
                    {
                        // Remove the entry
                        ObjectId attachmentsId = nod.GetAt(dictName);
                        nod.Remove(attachmentsId);

                        // Also erase the object
                        tr.GetObject(attachmentsId, OpenMode.ForWrite).Erase();

                        ed.WriteMessage("\nAll attachments cleared from drawing.");
                    }
                    else
                    {
                        ed.WriteMessage("\nNo attachments found in drawing.");
                    }

                    tr.Commit();
                }

                // Force the manager to reinitialize
                _instance = null;

                ed.WriteMessage("\nAttachments have been cleared. Run DETECTATTACHMENTS to detect new attachments.");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError clearing attachments: {ex.Message}");
            }
        }

        /// <summary>
        /// Creates a new MText object for a mark number
        /// </summary>
        private ObjectId CreateMarkText(Transaction tr, string markNumber, Polyline component, string partType, bool isHorizontal, int index)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            // Get component points
            Point3d startPoint = component.GetPoint3dAt(0);
            Point3d endPoint = component.GetPoint3dAt(1);

            // Calculate middle point of the component
            Point3d midPoint = new Point3d(
                (startPoint.X + endPoint.X) / 2,
                (startPoint.Y + endPoint.Y) / 2,
                0
            );

            // Calculate text position
            Point3d textPosition;
            double rotationAngle = 0;

            // Check if we have a custom offset for this part
            string handle = component.Handle.ToString();
            Point3d customOffset = GetCustomOffset(handle, partType);

            if (customOffset != Point3d.Origin)
            {
                // Use the custom offset
                textPosition = customOffset;
            }
            else
            {
                // For the default position, simply use the index parameter
                if (isHorizontal)
                {
                    // For horizontal components, stack mark numbers above
                    textPosition = new Point3d(
                        midPoint.X,
                        midPoint.Y + _defaultVerticalOffset + (index * _defaultSpacing),
                        0
                    );

                    // Horizontal text
                    rotationAngle = 0;
                }
                else
                {
                    // For vertical components, stack mark numbers to the left
                    textPosition = new Point3d(
                        midPoint.X - _defaultHorizontalOffset - (index * _defaultSpacing),
                        midPoint.Y,
                        0
                    );

                    // Vertical text (90 degrees in radians)
                    rotationAngle = Math.PI / 2;
                }
            }

            // Create the MText object
            MText mtext = new MText();
            mtext.Contents = markNumber;
            mtext.Location = textPosition;
            mtext.TextHeight = _textHeight;
            mtext.Rotation = rotationAngle;

            // Set text style
            ObjectId textStyleId = GetTextStyleId(tr, _textStyleName);
            if (textStyleId != ObjectId.Null)
            {
                mtext.TextStyleId = textStyleId;
            }

            // Register app name for XData
            RegisterApp(tr, "MARKNUMBERTEXT");

            // Add custom data to link this text to the component and part
            ResultBuffer rb = new ResultBuffer(
                new TypedValue((int)DxfCode.ExtendedDataRegAppName, "MARKNUMBERTEXT"),
                new TypedValue((int)DxfCode.ExtendedDataAsciiString, handle),
                new TypedValue((int)DxfCode.ExtendedDataAsciiString, partType)
            );
            mtext.XData = rb;

            // Add to model space
            BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
            BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

            ObjectId textId = btr.AppendEntity(mtext);
            tr.AddNewlyCreatedDBObject(mtext, true);

            return textId;
        }

        /// <summary>
        /// Gets the ObjectId for a text style
        /// </summary>
        private ObjectId GetTextStyleId(Transaction tr, string styleName)
        {
            Database db = Application.DocumentManager.MdiActiveDocument.Database;

            TextStyleTable textStyles = (TextStyleTable)tr.GetObject(db.TextStyleTableId, OpenMode.ForRead);

            if (textStyles.Has(styleName))
            {
                return textStyles[styleName];
            }

            return ObjectId.Null;
        }

        

        /// <summary>
        /// Updates all mark number displays in the drawing
        /// </summary>
        public void UpdateAllMarkNumbers()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            Editor ed = doc.Editor;
            ed.WriteMessage("\nUpdating all mark number displays...");

            using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
            {
                try
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

                        foreach (SelectedObject selObj in ss)
                        {
                            ObjectId id = selObj.ObjectId;
                            Entity ent = tr.GetObject(id, OpenMode.ForRead) as Entity;

                            if (ent != null)
                            {
                                string componentType = GetComponentType(ent);

                                if (!string.IsNullOrEmpty(componentType))
                                {
                                    // This is a metal component - update its mark numbers
                                    UpdateMarkNumbersForComponent(id);
                                    count++;
                                }
                            }
                        }

                        ed.WriteMessage($"\nUpdated mark number displays for {count} components.");
                    }

                    tr.Commit();
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nError updating mark number displays: {ex.Message}");
                    tr.Abort();
                }
            }
        }

        /// <summary>
        /// Deletes mark texts for a specific component
        /// </summary>
        private void DeleteMarkTextsForComponent(string handle)
        {
            if (!_markNumberTexts.ContainsKey(handle)) return;

            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
            {
                try
                {
                    List<MarkNumberText> texts = _markNumberTexts[handle];

                    foreach (var text in texts)
                    {
                        if (text.TextObjectId.IsValid && !text.TextObjectId.IsErased)
                        {
                            MText mtext = tr.GetObject(text.TextObjectId, OpenMode.ForWrite) as MText;
                            if (mtext != null)
                            {
                                mtext.Erase();
                            }
                        }
                    }

                    _markNumberTexts.Remove(handle);

                    tr.Commit();
                }
                catch (System.Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error in DeleteMarkTextsForComponent: {ex.Message}");
                    tr.Abort();
                }
            }
        }

        /// <summary>
        /// Clears all mark texts from the drawing
        /// </summary>
        public void ClearAllMarkTexts()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            Editor ed = doc.Editor;
            ed.WriteMessage("\nClearing all mark number displays...");

            using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
            {
                try
                {
                    // Get all MText objects with our XData
                    TypedValue[] tvs = new TypedValue[] {
                        new TypedValue((int)DxfCode.Start, "MTEXT"),
                        new TypedValue((int)DxfCode.ExtendedDataRegAppName, "MARKNUMBERTEXT")
                    };

                    SelectionFilter filter = new SelectionFilter(tvs);
                    PromptSelectionResult selRes = ed.SelectAll(filter);

                    if (selRes.Status == PromptStatus.OK)
                    {
                        SelectionSet ss = selRes.Value;

                        foreach (SelectedObject selObj in ss)
                        {
                            ObjectId id = selObj.ObjectId;
                            MText mtext = tr.GetObject(id, OpenMode.ForWrite) as MText;
                            if (mtext != null)
                            {
                                mtext.Erase();
                            }
                        }

                        ed.WriteMessage($"\nRemoved {ss.Count} mark number displays.");
                    }

                    // Clear our tracking dictionary
                    _markNumberTexts.Clear();

                    tr.Commit();
                }
                catch (System.Exception ex)
                {
                    ed.WriteMessage($"\nError clearing mark number displays: {ex.Message}");
                    tr.Abort();
                }
            }
        }



        #endregion

        #region Custom Position Management

        /// <summary>
        /// Updates the custom offset for a part
        /// </summary>
        private void UpdateCustomOffset(string componentHandle, string partType, Point3d position)
        {
            // Add detailed logging
            System.Diagnostics.Debug.WriteLine($"Attempting to update custom offset for {componentHandle}.{partType} to {position.X}, {position.Y}");

            try
            {
                // Validate inputs
                if (string.IsNullOrEmpty(componentHandle) || string.IsNullOrEmpty(partType))
                {
                    System.Diagnostics.Debug.WriteLine("Cannot update custom offset: componentHandle or partType is null or empty");
                    return;
                }

                // Initialize the dictionary for this component if it doesn't exist
                if (!_customOffsets.ContainsKey(componentHandle))
                {
                    _customOffsets[componentHandle] = new Dictionary<string, Point3d>();
                    System.Diagnostics.Debug.WriteLine($"Created new custom offset dictionary for component {componentHandle}");
                }

                // Update or add the position
                _customOffsets[componentHandle][partType] = position;
                System.Diagnostics.Debug.WriteLine($"Updated custom offset for {componentHandle}.{partType} to {position.X}, {position.Y}");

                // Save the updated custom positions - wrap in try/catch
                try
                {
                    SaveCustomPositions();
                }
                catch (System.Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error saving custom positions: {ex.Message}");
                }
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error updating custom offset: {ex.Message}");
                System.Diagnostics.Debug.WriteLine($"Stack trace: {ex.StackTrace}");
            }
        }

        /// <summary>
        /// Gets the custom offset for a part if it exists
        /// </summary>
        private Point3d GetCustomOffset(string componentHandle, string partType)
        {
            if (_customOffsets.ContainsKey(componentHandle) &&
                _customOffsets[componentHandle].ContainsKey(partType))
            {
                return _customOffsets[componentHandle][partType];
            }

            return Point3d.Origin;
        }

        /// <summary>
        /// Saves custom position data to the drawing
        /// </summary>
        private void SaveCustomPositions()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null)
            {
                System.Diagnostics.Debug.WriteLine("Cannot save custom positions: No active document");
                return;
            }

            using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
            {
                try
                {
                    // Get named objects dictionary
                    DBDictionary nod = (DBDictionary)tr.GetObject(
                        doc.Database.NamedObjectsDictionaryId,
                        OpenMode.ForWrite);

                    // Make sure the RegApp is registered
                    RegisterApp(tr, "MARKNUMBERPOSITIONS");

                    // Convert positions to a serializable format with detailed error handling
                    var serializableOffsets = new Dictionary<string, Dictionary<string, double[]>>();

                    foreach (var kvp in _customOffsets)
                    {
                        string componentHandle = kvp.Key;
                        if (string.IsNullOrEmpty(componentHandle)) continue;

                        serializableOffsets[componentHandle] = new Dictionary<string, double[]>();

                        foreach (var partKvp in kvp.Value)
                        {
                            string partType = partKvp.Key;
                            if (string.IsNullOrEmpty(partType)) continue;

                            Point3d position = partKvp.Value;

                            serializableOffsets[componentHandle][partType] =
                                new double[] { position.X, position.Y, position.Z };
                        }
                    }

                    // Serialize to JSON
                    string json;
                    try
                    {
                        json = JsonConvert.SerializeObject(serializableOffsets);
                        System.Diagnostics.Debug.WriteLine($"Serialized {serializableOffsets.Count} component offsets to JSON");
                    }
                    catch (System.Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error serializing offsets: {ex.Message}");
                        tr.Abort();
                        return;
                    }

                    // Create or update the Xrecord
                    Xrecord xrec;
                    const string dictName = "MARKNUMBERPOSITIONS";

                    if (nod.Contains(dictName))
                    {
                        // Update existing
                        try
                        {
                            xrec = (Xrecord)tr.GetObject(nod.GetAt(dictName), OpenMode.ForWrite);
                        }
                        catch (System.Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Error getting existing Xrecord: {ex.Message}");
                            // Create new if we can't get existing
                            xrec = new Xrecord();
                            nod.SetAt(dictName, xrec);
                            tr.AddNewlyCreatedDBObject(xrec, true);
                        }
                    }
                    else
                    {
                        // Create new
                        xrec = new Xrecord();
                        nod.SetAt(dictName, xrec);
                        tr.AddNewlyCreatedDBObject(xrec, true);
                    }

                    // Set the data
                    ResultBuffer rb = new ResultBuffer(
                        new TypedValue((int)DxfCode.Text, json)
                    );
                    xrec.Data = rb;

                    tr.Commit();
                    System.Diagnostics.Debug.WriteLine("Custom positions saved successfully");
                }
                catch (System.Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error saving custom positions: {ex.Message}");
                    System.Diagnostics.Debug.WriteLine($"Stack trace: {ex.StackTrace}");
                    tr.Abort();
                }
            }
        }

        /// <summary>
        /// Loads custom position data from the drawing
        /// </summary>
        private void LoadCustomPositions()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
            {
                try
                {
                    // Get named objects dictionary
                    DBDictionary nod = (DBDictionary)tr.GetObject(
                        doc.Database.NamedObjectsDictionaryId,
                        OpenMode.ForRead);

                    const string dictName = "MARKNUMBERPOSITIONS";

                    if (nod.Contains(dictName))
                    {
                        Xrecord xrec = (Xrecord)tr.GetObject(nod.GetAt(dictName), OpenMode.ForRead);
                        ResultBuffer rb = xrec.Data;

                        if (rb != null)
                        {
                            TypedValue[] values = rb.AsArray();
                            if (values.Length > 0 && values[0].TypeCode == (int)DxfCode.Text)
                            {
                                string json = values[0].Value.ToString();

                                // Deserialize the data
                                var serializableOffsets = JsonConvert.DeserializeObject<
                                    Dictionary<string, Dictionary<string, double[]>>>(json);

                                // Convert back to Point3d format
                                _customOffsets.Clear();

                                foreach (var kvp in serializableOffsets)
                                {
                                    string componentHandle = kvp.Key;
                                    _customOffsets[componentHandle] = new Dictionary<string, Point3d>();

                                    foreach (var partKvp in kvp.Value)
                                    {
                                        string partType = partKvp.Key;
                                        double[] coords = partKvp.Value;

                                        if (coords.Length >= 3)
                                        {
                                            _customOffsets[componentHandle][partType] =
                                                new Point3d(coords[0], coords[1], coords[2]);
                                        }
                                    }
                                }
                            }
                        }
                    }

                    tr.Commit();
                }
                catch (System.Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error loading custom positions: {ex.Message}");
                    tr.Abort();
                }
            }
        }

        /// <summary>
        /// Copies custom positions from one component to another
        /// </summary>
        public void CopyCustomPositions(string sourceHandle, string targetHandle, Matrix3d transform)
        {
            if (!_customOffsets.ContainsKey(sourceHandle)) return;

            // Create a dictionary for the target if it doesn't exist
            if (!_customOffsets.ContainsKey(targetHandle))
            {
                _customOffsets[targetHandle] = new Dictionary<string, Point3d>();
            }

            // Copy each position with the transform applied
            foreach (var kvp in _customOffsets[sourceHandle])
            {
                string partType = kvp.Key;
                Point3d sourcePosition = kvp.Value;

                // Apply the transformation
                Point3d targetPosition = sourcePosition.TransformBy(transform);

                // Store the new position
                _customOffsets[targetHandle][partType] = targetPosition;
            }

            // Save the updated positions
            SaveCustomPositions();
        }

        #endregion

        #region Helper Methods


        /// <summary>
        /// Registers required application names for XData
        /// </summary>
        private void RegisterApplicationNames()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
            {
                try
                {
                    // Get the RegApp table
                    RegAppTable regTable = (RegAppTable)tr.GetObject(
                        doc.Database.RegAppTableId,
                        OpenMode.ForWrite);

                    // Register our application name
                    if (!regTable.Has("MARKNUMBERTEXT"))
                    {
                        RegAppTableRecord record = new RegAppTableRecord();
                        record.Name = "MARKNUMBERTEXT";
                        regTable.Add(record);
                        tr.AddNewlyCreatedDBObject(record, true);
                    }

                    tr.Commit();
                }
                catch (System.Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error registering application names: {ex.Message}");
                    tr.Abort();
                }
            }
        }


        /// <summary>
        /// Gets the component type from an entity
        /// </summary>
        /// <summary>
        /// Gets the component type from an entity (made public for command access)
        /// </summary>
        public string GetComponentType(Entity ent)
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

        /// <summary>
        /// Helper method to register an application name
        /// </summary>
        private void RegisterApp(Transaction tr, string appName)
        {
            Database db = Application.DocumentManager.MdiActiveDocument.Database;
            RegAppTable regTable = (RegAppTable)tr.GetObject(db.RegAppTableId, OpenMode.ForWrite);

            if (!regTable.Has(appName))
            {
                RegAppTableRecord record = new RegAppTableRecord();
                record.Name = appName;
                regTable.Add(record);
                tr.AddNewlyCreatedDBObject(record, true);
            }
        }

        /// <summary>
        /// Gets child parts from an entity
        /// </summary>
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

        /// <summary>
        /// Class to track mark number text objects
        /// </summary>
        private class MarkNumberText
        {
            /// <summary>
            /// ObjectId of the MText object
            /// </summary>
            public ObjectId TextObjectId { get; set; }

            /// <summary>
            /// Part type associated with this text
            /// </summary>
            public string PartType { get; set; }

            /// <summary>
            /// Current mark number text
            /// </summary>
            public string MarkNumber { get; set; }
        }
        #endregion
    }
}
