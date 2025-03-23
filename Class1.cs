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

        public void Initialize()
        {
            // Called when AutoCAD loads the plugin
            Document doc = Application.DocumentManager.MdiActiveDocument;
            doc.Editor.WriteMessage("\nGlass Takeoff Bridge loaded successfully.");

            // Initialize the mark number manager
            _markNumberManager = TakeoffBridge.MarkNumberManager.Instance;

            // Subscribe to AutoCAD events
            SetupDocumentEvents();

            doc.Editor.WriteMessage("\nMark Number Manager initialized successfully.");
        }

        public void Terminate()
        {
            // Called when AutoCAD unloads the plugin
        }

        private void SetupDocumentEvents()
        {
            // Subscribe to document events
            Application.DocumentManager.DocumentCreated += DocumentManager_DocumentCreated;
            Application.DocumentManager.DocumentActivated += DocumentManager_DocumentActivated;

            // Set up the current document's events
            if (Application.DocumentManager.MdiActiveDocument != null)
            {
                SetupCurrentDocumentEvents();
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
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc != null)
            {
                // Hook into the document events that we need for mark number management
                Database db = doc.Database;

                // ObjectModified events can be used to detect when entities are modified
                db.ObjectModified += Database_ObjectModified;

                // You might also want to handle these events
                db.ObjectAppended += Database_ObjectAppended;
                db.ObjectErased += Database_ObjectErased;
            }
        }

        private void Database_ObjectAppended(object sender, ObjectEventArgs e)
        {
            // When a new object is added to the drawing
            try
            {
                Document doc = Application.DocumentManager.MdiActiveDocument;
                Database db = doc.Database;

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    DBObject obj = tr.GetObject(e.DBObject.ObjectId, OpenMode.ForRead);

                    // Check if this is a metal component
                    if (obj is Entity)
                    {
                        Entity ent = obj as Entity;

                        // Check if it has METALCOMP Xdata
                        ResultBuffer rbComp = ent.GetXDataForApplication("METALCOMP");
                        if (rbComp != null)
                        {
                            // Store the ObjectId to process later
                            ObjectId idToProcess = e.DBObject.ObjectId;

                            // Use the Idle event to process after transaction completes
                            Application.Idle += delegate
                            {
                                Application.Idle -= delegate { };  // Remove the handler
                                MarkNumberManager.OnComponentCreated(idToProcess);
                            };
                        }
                    }

                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                // Log error but don't crash
                System.Diagnostics.Debug.WriteLine($"Error in ObjectAppended: {ex.Message}");
            }
        }

        private void Database_ObjectModified(object sender, ObjectEventArgs e)
        {
            // When an object is modified in the drawing
            try
            {
                Document doc = Application.DocumentManager.MdiActiveDocument;
                Database db = doc.Database;

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    DBObject obj = tr.GetObject(e.DBObject.ObjectId, OpenMode.ForRead);

                    // Check if this is a metal component
                    if (obj is Entity)
                    {
                        Entity ent = obj as Entity;

                        // Check if it has METALCOMP Xdata
                        ResultBuffer rbComp = ent.GetXDataForApplication("METALCOMP");
                        if (rbComp != null)
                        {
                            // Store the ObjectId to process later
                            ObjectId idToProcess = e.DBObject.ObjectId;

                            // Check if it's a stretch operation (by checking if length changed)
                            if (ent is Polyline)
                            {
                                Polyline pline = ent as Polyline;
                                double newLength = pline.Length;

                                // Use the Idle event to process after transaction completes
                                Application.Idle += delegate
                                {
                                    Application.Idle -= delegate { };  // Remove the handler
                                    MarkNumberManager.OnComponentStretched(idToProcess, newLength);
                                };
                            }
                            else
                            {
                                // General modification
                                Application.Idle += delegate
                                {
                                    Application.Idle -= delegate { };  // Remove the handler
                                    MarkNumberManager.OnComponentModified(idToProcess);
                                };
                            }
                        }
                    }

                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                // Log error but don't crash
                System.Diagnostics.Debug.WriteLine($"Error in ObjectModified: {ex.Message}");
            }
        }

        private void Database_ObjectErased(object sender, ObjectErasedEventArgs e)
        {
            // When an object is erased from the drawing
            // This can be used to clean up references when objects are deleted
            if (!e.Erased)
            {
                // Object is being deleted
                try
                {
                    // You could potentially clean up references or perform other actions
                    // when metal components are deleted
                    System.Diagnostics.Debug.WriteLine($"Object {e.DBObject.ObjectId} erased");
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