using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Newtonsoft.Json;
using TakeoffBridge;

// This line is necessary to tell AutoCAD to load your plugin
[assembly: ExtensionApplication(typeof(GlassTakeoffBridge.GlassTakeoffApp))]
[assembly: CommandClass(typeof(GlassTakeoffBridge.GlassTakeoffCommands))]
[assembly: CommandClass(typeof(TakeoffBridge.MetalComponentCommands))]
[assembly: CommandClass(typeof(TakeoffBridge.MarkNumberManagerCommands))]

namespace GlassTakeoffBridge
{
    // Glass data class to capture all necessary information
    public class GlassData
    {
        public string Handle { get; set; }
        public string GlassType { get; set; }
        public string Floor { get; set; }
        public string Elevation { get; set; }
        public List<double> Coordinates { get; set; }
        public double GlassBiteLeft { get; set; }
        public double GlassBiteRight { get; set; }
        public double GlassBiteTop { get; set; }
        public double GlassBiteBottom { get; set; }
    }

    // Response from web app
    public class ProcessedGlassData
    {
        public string Handle { get; set; }
        public string MarkNumber { get; set; }
        // Add other fields as needed
    }

    // Main application class
    public class GlassTakeoffApp : IExtensionApplication
    {
        // Static reference to our mark number manager
        private static TakeoffBridge.MarkNumberManager _markNumberManager;

        // Public property to access the manager from other classes
        public static TakeoffBridge.MarkNumberManager MarkNumberManager
        {
            get
            {
                if (_markNumberManager == null)
                {
                    _markNumberManager = TakeoffBridge.MarkNumberManager.Instance;
                }
                return _markNumberManager;
            }
        }

        // Keep a static reference to your panel for access from commands
        private static EnhancedMetalComponentPanel _panel;

        public static EnhancedMetalComponentPanel Panel
        {
            get { return _panel; }
        }

        public void Initialize()
        {
            try
            {
                // Called when AutoCAD loads the plugin
                Document doc = Application.DocumentManager.MdiActiveDocument;
                if (doc != null)
                {
                    doc.Editor.WriteMessage("\nMetal Component Takeoff plugin loaded successfully.");
                }

                // Initialize the mark number manager
                _markNumberManager = TakeoffBridge.MarkNumberManager.Instance;

                // Subscribe to AutoCAD events
                SetupDocumentEvents();

                // Register reactor for document creation events
                Application.DocumentManager.DocumentCreated += DocumentManager_DocumentCreated;
                Application.DocumentManager.DocumentActivated += DocumentManager_DocumentActivated;

                if (doc != null)
                {
                    doc.Editor.WriteMessage("\nMark Number Manager initialized successfully.");
                }
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in Initialize: {ex.Message}");
            }
        }

        public void Terminate()
        {
            // Called when AutoCAD unloads the plugin
            if (Application.DocumentManager != null)
            {
                Application.DocumentManager.DocumentCreated -= DocumentManager_DocumentCreated;
                Application.DocumentManager.DocumentActivated -= DocumentManager_DocumentActivated;

                // Also unsubscribe from database events
                Document doc = Application.DocumentManager.MdiActiveDocument;
                if (doc != null && doc.Database != null)
                {
                    //doc.Database.ObjectModified -= Database_ObjectModified;
                    //doc.Database.ObjectAppended -= Database_ObjectAppended;
                    doc.Database.ObjectErased -= Database_ObjectErased;
                }
            }
        }

        private void SetupDocumentEvents()
        {
            try
            {
                // Set up the current document's events
                if (Application.DocumentManager.MdiActiveDocument != null)
                {
                    SetupCurrentDocumentEvents();
                }
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in SetupDocumentEvents: {ex.Message}");
            }
        }

        private void DocumentManager_DocumentCreated(object sender, DocumentCollectionEventArgs e)
        {
            // Reset mark number manager when a new document is created
            _markNumberManager = null;

            // Setup events for the new document
            SetupCurrentDocumentEvents();
        }

        private void DocumentManager_DocumentActivated(object sender, DocumentCollectionEventArgs e)
        {
            // Reset mark number manager when switching documents
            _markNumberManager = null;

            // Setup events for the newly activated document
            SetupCurrentDocumentEvents();
        }

        private void SetupCurrentDocumentEvents()
        {
            try
            {
                Document doc = Application.DocumentManager.MdiActiveDocument;
                if (doc != null)
                {
                    // Unsubscribe from any existing events first to avoid duplicates
                    Database db = doc.Database;
                    //db.ObjectModified -= Database_ObjectModified;
                    //db.ObjectAppended -= Database_ObjectAppended;
                    db.ObjectErased -= Database_ObjectErased;

                    // Now hook into the document events
                    //db.ObjectModified += Database_ObjectModified;
                    //db.ObjectAppended += Database_ObjectAppended;
                    db.ObjectErased += Database_ObjectErased;

                    System.Diagnostics.Debug.WriteLine($"Set up events for database: {db.Filename}");
                }
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error in SetupCurrentDocumentEvents: {ex.Message}");
            }
        }

        //private void Database_ObjectAppended(object sender, ObjectEventArgs e)
        //{
        //    // When a new object is added to the drawing
        //    try
        //    {
        //        // Check if this database is the current document's database
        //        Document doc = Application.DocumentManager.MdiActiveDocument;
        //        if (doc == null) return;

        //        Database eventDb = sender as Database;
        //        Database activeDb = doc.Database;

        //        // If the event is coming from a different document/database, ignore it
        //        if (eventDb != activeDb)
        //        {
        //            return;
        //        }

        //        using (Transaction tr = activeDb.TransactionManager.StartTransaction())
        //        {
        //            if (!e.DBObject.ObjectId.IsValid || e.DBObject.ObjectId.IsErased)
        //            {
        //                tr.Commit();
        //                return;
        //            }

        //            DBObject obj;
        //            try
        //            {
        //                obj = tr.GetObject(e.DBObject.ObjectId, OpenMode.ForRead);
        //            }
        //            catch (Autodesk.AutoCAD.Runtime.Exception acEx)
        //            {
        //                if (acEx.ErrorStatus == Autodesk.AutoCAD.Runtime.ErrorStatus.NotFromThisDocument)
        //                {
        //                    System.Diagnostics.Debug.WriteLine("Object not from this document - ignoring");
        //                }
        //                tr.Commit();
        //                return;
        //            }

        //            // Check if this is a metal component
        //            if (obj is Entity)
        //            {
        //                Entity ent = obj as Entity;
        //                // Check if it has METALCOMP Xdata
        //                ResultBuffer rbComp = ent.GetXDataForApplication("METALCOMP");
        //                if (rbComp != null)
        //                {
        //                    // Store the ObjectId to process later
        //                    ObjectId idToProcess = e.DBObject.ObjectId;
        //                    // Use the Idle event to process after transaction completes
        //                    Application.Idle += delegate
        //                    {
        //                        Application.Idle -= delegate { };  // Remove the handler

        //                        // Check if the active document is still the same
        //                        Document currentDoc = Application.DocumentManager.MdiActiveDocument;
        //                        if (currentDoc == null || currentDoc.Database != eventDb)
        //                        {
        //                            return; // Document changed, don't process
        //                        }

        //                        // Process the component
        //                        MarkNumberManager.Instance.OnComponentCreated(idToProcess);

        //                        // Notify the UI
        //                        NotifyPanelsOfComponentChange(idToProcess);
        //                    };
        //                }
        //            }
        //            tr.Commit();
        //        }
        //    }
        //    catch (System.Exception ex)
        //    {
        //        // Log error but don't crash
        //        System.Diagnostics.Debug.WriteLine($"Error in ObjectAppended: {ex.Message}");
        //    }
        //}

        //private void Database_ObjectModified(object sender, ObjectEventArgs e)
        //{
        //    // When an object is modified in the drawing
        //    try
        //    {
        //        // Check if this database is the current document's database
        //        Document doc = Application.DocumentManager.MdiActiveDocument;
        //        if (doc == null) return;

        //        Database eventDb = sender as Database;
        //        Database activeDb = doc.Database;

        //        // If the event is coming from a different document/database, ignore it
        //        if (eventDb != activeDb)
        //        {
        //            return;
        //        }

        //        // Store the ID to process later
        //        ObjectId idToProcess = e.DBObject.ObjectId;

        //        // Use Idle event to ensure we process after the transaction completes
        //        Application.Idle += delegate
        //        {
        //            Application.Idle -= delegate { }; // Remove the handler

        //            try
        //            {
        //                // Check if the active document is still the same
        //                Document currentDoc = Application.DocumentManager.MdiActiveDocument;
        //                if (currentDoc == null || currentDoc.Database != eventDb)
        //                {
        //                    return; // Document changed, don't process
        //                }

        //                using (Transaction tr = eventDb.TransactionManager.StartTransaction())
        //                {
        //                    // Skip if object is no longer valid
        //                    if (!idToProcess.IsValid || idToProcess.IsErased)
        //                    {
        //                        System.Diagnostics.Debug.WriteLine("Object was erased before processing");
        //                        tr.Commit();
        //                        return;
        //                    }

        //                    // Try to open the object
        //                    DBObject obj;
        //                    try
        //                    {
        //                        obj = tr.GetObject(idToProcess, OpenMode.ForRead);
        //                    }
        //                    catch (Autodesk.AutoCAD.Runtime.Exception acEx)
        //                    {
        //                        // Handle specific AutoCAD exceptions
        //                        if (acEx.ErrorStatus == Autodesk.AutoCAD.Runtime.ErrorStatus.NotFromThisDocument)
        //                        {
        //                            System.Diagnostics.Debug.WriteLine("Object not from this document - ignoring");
        //                        }
        //                        else
        //                        {
        //                            System.Diagnostics.Debug.WriteLine($"Failed to open object: {acEx.Message}");
        //                        }
        //                        tr.Commit();
        //                        return;
        //                    }
        //                    catch
        //                    {
        //                        System.Diagnostics.Debug.WriteLine("Failed to open object - may have been erased");
        //                        tr.Commit();
        //                        return;
        //                    }

        //                    // Check if this is a metal component
        //                    if (obj is Entity)
        //                    {
        //                        Entity ent = obj as Entity;

        //                        // Check if it has METALCOMP Xdata
        //                        ResultBuffer rbComp = ent.GetXDataForApplication("METALCOMP");
        //                        if (rbComp != null)
        //                        {
        //                            // Get the manager and process the component
        //                            MarkNumberManager manager = MarkNumberManager;

        //                            if (ent is Polyline)
        //                            {
        //                                Polyline pline = ent as Polyline;
        //                                double newLength = pline.Length;
        //                                manager.OnComponentStretched(idToProcess, newLength);
        //                            }
        //                            else
        //                            {
        //                                manager.OnComponentModified(idToProcess);
        //                            }

        //                            // Notify any open panels of the change
        //                            NotifyPanelsOfComponentChange(idToProcess);
        //                        }
        //                    }

        //                    tr.Commit();
        //                }
        //            }
        //            catch (System.Exception ex)
        //            {
        //                System.Diagnostics.Debug.WriteLine($"Error processing modified object: {ex.Message}");
        //            }
        //        };
        //    }
        //    catch (System.Exception ex)
        //    {
        //        // Log error but don't crash
        //        System.Diagnostics.Debug.WriteLine($"Error in ObjectModified: {ex.Message}");
        //    }
        //}

        // Add a method to notify panels of component changes
        private void NotifyPanelsOfComponentChange(ObjectId componentId)
        {
            try
            {
                // If the panel exists and is visible
                if (_panel != null)
                {
                    // Call a new method on the panel to notify of changes
                    _panel.OnComponentChanged(componentId);
                }
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error notifying panel of component change: {ex.Message}");
            }
        }

        private void Database_ObjectErased(object sender, ObjectErasedEventArgs e)
        {
            // When an object is erased from the drawing
            if (e.Erased)
            {
                try
                {
                    // Check if this database is the current document's database
                    Document doc = Application.DocumentManager.MdiActiveDocument;
                    if (doc == null) return;

                    Database eventDb = sender as Database;
                    Database activeDb = doc.Database;

                    // If the event is coming from a different document/database, ignore it
                    if (eventDb != activeDb)
                    {
                        return;
                    }

                    ObjectId erasedId = e.DBObject.ObjectId;

                    // You could potentially clean up references or perform other actions
                    // when metal components are deleted
                    System.Diagnostics.Debug.WriteLine($"Object {erasedId} erased");

                    // If it's a metal component being erased, notify the UI
                    Application.Idle += delegate
                    {
                        Application.Idle -= delegate { };  // Remove the handler

                        try
                        {
                            // Check if the active document is still the same
                            Document currentDoc = Application.DocumentManager.MdiActiveDocument;
                            if (currentDoc == null || currentDoc.Database != eventDb)
                            {
                                return; // Document changed, don't process
                            }

                            // Check if the panel is showing this component
                            if (_panel != null)
                            {
                                // Refresh the panel if it's showing the erased object
                                _panel.OnComponentErased(erasedId);
                            }
                        }
                        catch (System.Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Error handling erased notification: {ex.Message}");
                        }
                    };
                }
                catch (System.Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error in ObjectErased: {ex.Message}");
                }
            }
        }


    }

    // Commands available in AutoCAD
    public class GlassTakeoffCommands
    {
        private static readonly HttpClient client = new HttpClient();
        private const string API_URL = "http://localhost:3000/api/glass-takeoff";

        [CommandMethod("GLASSTAKEOFF")]
        public async void GlassTakeoff()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            ed.WriteMessage("\nGlass Takeoff Bridge started.");

            try
            {
                // Get drawing name for reference
                string drawingName = doc.Name;
                ed.WriteMessage($"\nProcessing drawing: {drawingName}");

                // Collect glass polylines
                List<GlassData> glassDataList = new List<GlassData>();

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                    foreach (ObjectId objId in btr)
                    {
                        Entity ent = tr.GetObject(objId, OpenMode.ForRead) as Entity;

                        if (ent is Polyline)
                        {
                            Polyline pline = ent as Polyline;

                            // Check if it has GLASS Xdata
                            ResultBuffer rbXdata = ent.GetXDataForApplication("GLASS");
                            if (rbXdata != null)
                            {
                                GlassData glassData = new GlassData
                                {
                                    Handle = ent.Handle.ToString(),
                                    Coordinates = new List<double>(),
                                    // Default values
                                    GlassType = "",
                                    Floor = "",
                                    Elevation = ""
                                };

                                // Extract Xdata values
                                TypedValue[] xdata = rbXdata.AsArray();

                                // Process Xdata - adjust indices based on your actual Xdata structure
                                for (int i = 1; i < xdata.Length; i++) // Skip first item (application name)
                                {
                                    // Log what we're seeing to adjust the indices as needed
                                    ed.WriteMessage($"\nXdata[{i}]: {xdata[i].TypeCode} = {xdata[i].Value}");

                                    // Attempt to read values based on the expected structure
                                    // You'll need to adjust these indices based on your Xdata
                                    try
                                    {
                                        if (i == 1) glassData.Elevation = xdata[i].Value.ToString();
                                        if (i == 2) glassData.Floor = xdata[i].Value.ToString();
                                        if (i == 3) glassData.GlassType = xdata[i].Value.ToString();
                                        if (i == 4) glassData.GlassBiteLeft = Convert.ToDouble(xdata[i].Value);
                                        if (i == 5) glassData.GlassBiteBottom = Convert.ToDouble(xdata[i].Value);
                                        if (i == 6) glassData.GlassBiteRight = Convert.ToDouble(xdata[i].Value);
                                        if (i == 7) glassData.GlassBiteTop = Convert.ToDouble(xdata[i].Value);
                                    }
                                    catch (Autodesk.AutoCAD.Runtime.Exception ex)
                                    {
                                        ed.WriteMessage($"\nError parsing Xdata[{i}]: {ex.Message}");
                                    }
                                    catch (System.Exception ex)
                                    {
                                        // Handle general .NET exceptions
                                        ed.WriteMessage($"\nGeneral Error: {ex.Message}");
                                        ed.WriteMessage($"\nStack trace: {ex.StackTrace}");
                                    }
                                }

                                // Extract coordinates
                                for (int i = 0; i < pline.NumberOfVertices; i++)
                                {
                                    Point2d pt = pline.GetPoint2dAt(i);
                                    glassData.Coordinates.Add(pt.X);
                                    glassData.Coordinates.Add(pt.Y);
                                }

                                glassDataList.Add(glassData);

                                // Log what we captured
                                ed.WriteMessage($"\nCaptured glass: Type={glassData.GlassType}, Floor={glassData.Floor}, Coords={glassData.Coordinates.Count / 2} points");
                            }
                        }
                    }

                    tr.Commit();
                }

                ed.WriteMessage($"\nCollected {glassDataList.Count} glass items.");

                // Send data to web API
                if (glassDataList.Count > 0)
                {
                    ed.WriteMessage("\nSending data to web application...");
                    string response = await SendDataToWebApp(glassDataList, drawingName, ed);

                    if (response != null)
                    {
                        ed.WriteMessage("\nData processed successfully by web application.");

                        // Parse the response to get mark numbers (if available)
                        try
                        {
                            List<dynamic> processedItems = JsonConvert.DeserializeObject<List<dynamic>>(response);
                            ed.WriteMessage($"\nReceived {processedItems.Count} processed items from API.");

                            // Update drawing with mark numbers
                            // ... (code to add mark numbers to the drawing)
                        }
                        catch (System.Exception ex)
                        {
                            ed.WriteMessage($"\nError parsing API response: {ex.Message}");
                        }
                    }
                    else
                    {
                        ed.WriteMessage("\nFailed to process data with web application.");
                    }
                }

                ed.WriteMessage("\nGlass Takeoff completed successfully.");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError: {ex.Message}");
                ed.WriteMessage($"\nStack trace: {ex.StackTrace}");
            }
        }


        private async Task<string> SendDataToWebApp(List<GlassData> glassDataList, string drawingName, Editor ed)
        {
            try
            {
                // Create the payload
                var payload = new
                {
                    drawing = drawingName,
                    glassItems = glassDataList
                };

                string jsonContent = JsonConvert.SerializeObject(payload);
                ed.WriteMessage($"\nSending data to web API: {jsonContent.Length} characters");
                ed.WriteMessage($"\nAPI URL: {API_URL}");

                using (HttpClient client = new HttpClient())
                {
                    // Add timeout and avoid SSL certificate validation for testing
                    client.Timeout = TimeSpan.FromSeconds(60);

                    // For testing, you might need to ignore SSL certificate errors
                    ServicePointManager.ServerCertificateValidationCallback =
                        (sender, certificate, chain, sslPolicyErrors) => true;

                    ed.WriteMessage("\nCreating HTTP request...");
                    StringContent content = new StringContent(jsonContent, Encoding.UTF8, "application/json");

                    ed.WriteMessage("\nSending HTTP request...");
                    try
                    {
                        HttpResponseMessage response = await client.PostAsync(API_URL, content);

                        ed.WriteMessage($"\nResponse received. Status: {response.StatusCode}");

                        if (response.IsSuccessStatusCode)
                        {
                            string responseContent = await response.Content.ReadAsStringAsync();
                            ed.WriteMessage($"\nAPI response: {responseContent}");
                            return responseContent;
                        }
                        else
                        {
                            string errorContent = await response.Content.ReadAsStringAsync();
                            ed.WriteMessage($"\nAPI error: {errorContent}");
                            return null;
                        }
                    }
                    catch (TaskCanceledException ex)
                    {
                        ed.WriteMessage($"\nHTTP request timed out: {ex.Message}");
                        return null;
                    }
                    catch (HttpRequestException ex)
                    {
                        ed.WriteMessage($"\nHTTP request failed: {ex.Message}");
                        return null;
                    }
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError sending data to web API: {ex.Message}");
                ed.WriteMessage($"\nStack trace: {ex.StackTrace}");
                return null;
            }
        }
    }
}