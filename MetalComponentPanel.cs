﻿using System;
using System.Drawing;
using System.Windows.Forms;
using Autodesk.AutoCAD.Runtime;
using Application = Autodesk.AutoCAD.ApplicationServices.Application;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System.Collections.Generic;
using System.Net.Mail;
using System.Linq;
using Newtonsoft.Json;
using System.Reflection;


namespace TakeoffBridge
{

    public class EnhancedMetalComponentPanel : System.Windows.Forms.UserControl, IDisposable
    {
        private ListView partsList;
        private Button btnAddPart;
        private Button btnEditPart;
        private Button btnDeletePart;
        private TextBox txtComponentType;
        private TextBox txtFloor;
        private TextBox txtElevation;
        private Button btnSaveParent;
        private Button btnSave;
        private Label lblLength;
        private Panel visualPanel; // New visual panel for part representation
        private ObjectId currentComponentId;
        private List<ChildPart> currentParts = new List<ChildPart>();
        private double currentComponentLength = 0.0;
        private bool isHorizontal = true; // Track if current component is horizontal
        private List<Attachment> currentAttachments = new List<Attachment>();
        private Panel attachmentPanel; // New visualization panel for attachments
        private bool forceRefreshOnNextSelection = false;
        private bool hasUnsavedChanges = false;
        private Document currentDocument;
        private Editor currentEditor;
        private bool _handlingSelectionChange = false;
        private bool isEventHandlerAttached = false;
        private bool isProcessingUpdate = false;
        // If you want to add a status label
        private Label statusLabel;
        private bool isCreateMode = false;
        private Point3d startPoint = new Point3d();
        private Point3d endPoint = new Point3d();

        // Constructor modified to support both modes
        public EnhancedMetalComponentPanel(bool createMode = false)
        {
            isCreateMode = createMode;
            InitializeComponent();

            // If in create mode, prepare default data
            if (isCreateMode)
            {
                PrepareForCreation();
            }
            else
            {
                ConnectToDocumentEvents();
            }
        }

        private void PrepareForCreation()
        {
            // Clear current component data without using Invoke
            currentComponentId = ObjectId.Null;
            currentComponentLength = 0.0;

            if (currentParts != null)
                currentParts.Clear();
            else
                currentParts = new List<ChildPart>();

            if (currentAttachments != null)
                currentAttachments.Clear();

            // Set default values directly
            if (txtComponentType != null) txtComponentType.Text = "Horizontal";
            if (txtFloor != null) txtFloor.Text = "1";
            if (txtElevation != null) txtElevation.Text = "A";
            if (lblLength != null) lblLength.Text = "0.0";

            // Add default parts based on component type
            LoadDefaultParts("Horizontal");

            // Enable controls for editing
            SetControlsEnabled(true);

            // Find the button panel in your UI
            // Scan through all child controls to find the TableLayoutPanel for buttons
            TableLayoutPanel buttonPanel = null;
            foreach (Control control in this.Controls)
            {
                if (control is TableLayoutPanel mainPanel)
                {
                    // Check if this main panel has any TableLayoutPanel children that might be our button panel
                    foreach (Control subControl in mainPanel.Controls)
                    {
                        if (subControl is TableLayoutPanel panel && panel.RowCount == 2)
                        {
                            // This is likely our button panel with 2 rows
                            buttonPanel = panel;
                            break;
                        }
                    }
                    if (buttonPanel != null) break;
                }
            }

            if (buttonPanel != null)
            {
                // Create the Create Components button
                Button btnCreateComponent = new Button
                {
                    Text = "Create Components",
                    Dock = DockStyle.Fill,
                    Margin = new Padding(3)
                };
                btnCreateComponent.Click += BtnCreateComponent_Click;

                // Create the Create Opening button
                Button btnCreateOpening = new Button
                {
                    Text = "Create Opening",
                    Dock = DockStyle.Fill,
                    Margin = new Padding(3)
                };
                btnCreateOpening.Click += BtnCreateOpening_Click;

                // Add Create Components button to the panel
                if (buttonPanel.ColumnCount > 4)
                {
                    Control existingControl = buttonPanel.GetControlFromPosition(4, 0);
                    if (existingControl == null)
                    {
                        buttonPanel.Controls.Add(btnCreateComponent, 4, 0);
                    }
                    else
                    {
                        // If position is occupied, try another position or a new row
                        buttonPanel.Controls.Add(btnCreateComponent, 4, 1);
                    }
                }
                else
                {
                    // If panel doesn't have enough columns, add to whatever position is available
                    buttonPanel.Controls.Add(btnCreateComponent, 0, 1);
                }

                // Add Create Opening button to the panel
                if (buttonPanel.ColumnCount > 5)
                {
                    Control existingControl = buttonPanel.GetControlFromPosition(5, 0);
                    if (existingControl == null)
                    {
                        buttonPanel.Controls.Add(btnCreateOpening, 5, 0);
                    }
                    else
                    {
                        // If position is occupied, try another position or a new row
                        buttonPanel.Controls.Add(btnCreateOpening, 5, 1);
                    }
                }
                else if (buttonPanel.ColumnCount > 0)
                {
                    // Try to add to the next row or column
                    int col = 0;
                    int row = 1;

                    // Find an empty spot
                    bool found = false;
                    for (row = 0; row < buttonPanel.RowCount && !found; row++)
                    {
                        for (col = 0; col < buttonPanel.ColumnCount && !found; col++)
                        {
                            if (buttonPanel.GetControlFromPosition(col, row) == null)
                            {
                                found = true;
                                break;
                            }
                        }
                        if (found) break;
                    }

                    // If no empty spot, add to the first cell in second row
                    if (!found)
                    {
                        col = 1;
                        row = 1;
                    }

                    buttonPanel.Controls.Add(btnCreateOpening, col, row);
                }

                // Change Save Changes to Apply Changes for consistency
                if (btnSave != null)
                    btnSave.Text = "Apply Changes";
            }
            else
            {
                // If button panel not found, try to add directly to the control
                Button btnCreateComponent = new Button
                {
                    Text = "Create Components",
                    Location = new System.Drawing.Point(300, 700),
                    Size = new System.Drawing.Size(150, 30)
                };
                btnCreateComponent.Click += BtnCreateComponent_Click;
                this.Controls.Add(btnCreateComponent);

                Button btnCreateOpening = new Button
                {
                    Text = "Create Opening",
                    Location = new System.Drawing.Point(300, 740),
                    Size = new System.Drawing.Size(150, 30)
                };
                btnCreateOpening.Click += BtnCreateOpening_Click;
                this.Controls.Add(btnCreateOpening);
            }
        }

        private void BtnCreateComponent_Click(object sender, EventArgs e)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            try
            {
                // Create a new instance of ParentComponentData
                MetalComponentCommands.ParentComponentData parentData =
                    new MetalComponentCommands.ParentComponentData();

                // Set the values
                parentData.ComponentType = txtComponentType.Text;
                parentData.Floor = txtFloor.Text;
                parentData.Elevation = txtElevation.Text;

                // Assign to the static field
                MetalComponentCommands.PendingParentData = parentData;

                // Create a copy of the parts
                if (currentParts != null && currentParts.Count > 0)
                {
                    // Create a new list to hold the copies
                    List<ChildPart> partsCopy = new List<ChildPart>();

                    // Copy each part
                    foreach (ChildPart part in currentParts)
                    {
                        partsCopy.Add(part);
                    }

                    // Assign to the static field
                    MetalComponentCommands.PendingParts = partsCopy;
                }
                else
                {
                    // Initialize with a new empty list if no parts exist
                    MetalComponentCommands.PendingParts = new List<ChildPart>();
                }

                // Debug output to verify the data is set
                System.Diagnostics.Debug.WriteLine($"Setting PendingParentData: Type={MetalComponentCommands.PendingParentData.ComponentType}, Floor={MetalComponentCommands.PendingParentData.Floor}");
                System.Diagnostics.Debug.WriteLine($"Setting PendingParts: Count={MetalComponentCommands.PendingParts.Count}");

                // Execute the ADDMETALPART command
                doc.SendStringToExecute("ADDMETALPART ", true, false, false);


            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Error initiating component creation: {ex.Message}",
                               "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void EnsureLayerExists(string layerName)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                // Open the Layer table
                LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);

                if (!lt.Has(layerName))
                {
                    // Create the layer if it doesn't exist
                    lt.UpgradeOpen();
                    LayerTableRecord ltr = new LayerTableRecord();
                    ltr.Name = layerName;
                    ltr.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByLayer, 1); // Red

                    lt.Add(ltr);
                    tr.AddNewlyCreatedDBObject(ltr, true);
                }

                tr.Commit();
            }
        }

        private void CreateComponent(Point3d startPt, Point3d endPt)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                try
                {
                    // Get the TAG layer ID
                    //LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                    //ObjectId tagLayerId = ObjectId.Null;

                    //if (lt.Has("TAG"))
                    //{
                    //    tagLayerId = lt["TAG"];
                    //}

                    // Create our polyline
                    BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                    Polyline pline = new Polyline();
                    pline.SetDatabaseDefaults();
                    pline.AddVertexAt(0, new Point2d(startPt.X, startPt.Y), 0, 0, 0);
                    pline.AddVertexAt(1, new Point2d(endPt.X, endPt.Y), 0, 0, 0);

                    // If TAG layer exists, set the polyline's layer to TAG
                    //if (tagLayerId != ObjectId.Null)
                    //{
                    //    pline.LayerId = tagLayerId;
                    //}

                    // Add polyline to database
                    ObjectId plineId = btr.AppendEntity(pline);
                    tr.AddNewlyCreatedDBObject(pline, true);

                    // Add component data
                    ResultBuffer rbComp = new ResultBuffer(
                        new TypedValue((int)DxfCode.ExtendedDataRegAppName, "METALCOMP"),
                        new TypedValue((int)DxfCode.ExtendedDataAsciiString, txtComponentType.Text),
                        new TypedValue((int)DxfCode.ExtendedDataAsciiString, txtFloor.Text),
                        new TypedValue((int)DxfCode.ExtendedDataAsciiString, txtElevation.Text)
                    );
                    pline.XData = rbComp;

                    // Register all required application names
                    RegAppTable regTable = (RegAppTable)tr.GetObject(db.RegAppTableId, OpenMode.ForRead);
                    RegisterApp(regTable, "METALCOMP", tr);
                    RegisterApp(regTable, "METALPARTSINFO", tr);

                    // Register chunk apps preemptively
                    for (int i = 0; i < 10; i++)
                    {
                        RegisterApp(regTable, $"METALPARTS{i}", tr);
                    }

                    // Store child parts
                    string partsJson = Newtonsoft.Json.JsonConvert.SerializeObject(currentParts);

                    // Add info Xdata
                    const int maxChunkSize = 250;
                    int numChunks = (int)Math.Ceiling((double)partsJson.Length / maxChunkSize);

                    ResultBuffer rbInfo = new ResultBuffer(
                        new TypedValue((int)DxfCode.ExtendedDataRegAppName, "METALPARTSINFO"),
                        new TypedValue((int)DxfCode.ExtendedDataInteger32, numChunks)
                    );
                    pline.XData = rbInfo;

                    // Add chunk Xdata
                    for (int i = 0; i < numChunks; i++)
                    {
                        int startIndex = i * maxChunkSize;
                        int length = Math.Min(maxChunkSize, partsJson.Length - startIndex);
                        string chunk = partsJson.Substring(startIndex, length);

                        ResultBuffer rbChunk = new ResultBuffer(
                            new TypedValue((int)DxfCode.ExtendedDataRegAppName, $"METALPARTS{i}"),
                            new TypedValue((int)DxfCode.ExtendedDataAsciiString, chunk)
                        );
                        pline.XData = rbChunk;
                    }

                    tr.Commit();

                    // Process mark numbers
                    using (Transaction trx = doc.Database.TransactionManager.StartTransaction())
                    {
                        MarkNumberManager.Instance.ProcessComponentMarkNumbers(plineId, trx, forceProcess: true);
                        trx.Commit();
                    }

                    // Highlight the newly created component
                    doc.Editor.WriteMessage($"\nComponent created with handle: {pline.Handle}");
                    using (Transaction trh = doc.Database.TransactionManager.StartTransaction())
                    {
                        Entity ent = trh.GetObject(plineId, OpenMode.ForWrite) as Entity;
                        if (ent != null)
                        {
                            ent.Highlight();
                        }
                        trh.Commit();
                    }
                }
                catch
                {
                    // Rollback in case of error
                    throw;
                }
            }
        }

        private void LoadDefaultParts(string componentType)
        {
            currentParts.Clear();

            if (componentType.ToUpper().Contains("HORIZONTAL"))
            {
                // Add default horizontal parts
                currentParts.Add(new ChildPart("Horizontal Body", "HB", 0.0, 0.0, "Aluminum"));
                currentParts.Add(new ChildPart("Flat Filler", "FF", -0.03125, 0.0, "Aluminum"));
                currentParts.Add(new ChildPart("Face Cap", "FC", 0.0, 0.0, "Aluminum"));

                // Left side attachments
                ChildPart sbLeft = new ChildPart("Shear Block Left", "SBL", -1.25, 0.0, "Aluminum");
                sbLeft.Attach = "L";
                currentParts.Add(sbLeft);

                // Right side attachments
                ChildPart sbRight = new ChildPart("Shear Block Right", "SBR", 0.0, -1.25, "Aluminum");
                sbRight.Attach = "R";
                currentParts.Add(sbRight);
            }
            else if (componentType.ToUpper().Contains("VERTICAL"))
            {
                // Add default vertical parts
                ChildPart vb = new ChildPart("Vertical Body", "VB", 0.0, 0.0, "Aluminum");
                vb.Clips = true;
                currentParts.Add(vb);

                currentParts.Add(new ChildPart("Pressure Plate", "PP", -0.0625, 0.0, "Aluminum"));
                currentParts.Add(new ChildPart("Snap Cover", "SC", 0.0, -0.125, "Aluminum"));
            }

            // Update the UI
            UpdatePartsList();
        }



        private ObjectIdCollection lastSelection = new ObjectIdCollection();

        private void ConnectToDocumentEvents()
        {
            try
            {
                // Get the current document
                Document doc = Application.DocumentManager.MdiActiveDocument;
                if (doc == null) return;

                // Track the current document
                currentDocument = doc;
                currentEditor = doc.Editor;

                // Subscribe to document events for selection changes
                doc.ImpliedSelectionChanged += Document_ImpliedSelectionChanged;

                // Subscribe to document switch events
                Application.DocumentManager.DocumentActivated += DocumentManager_DocumentActivated;

                // Connect to command events
                ConnectToCommandEvents();

                isEventHandlerAttached = true;

                System.Diagnostics.Debug.WriteLine("Connected to AutoCAD document events");
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error connecting to document events: {ex.Message}");
            }
        }

        private void DisconnectFromDocumentEvents()
        {
            try
            {
                if (currentDocument != null)
                {
                    currentDocument.ImpliedSelectionChanged -= Document_ImpliedSelectionChanged;
                }

                Application.DocumentManager.DocumentActivated -= DocumentManager_DocumentActivated;

                // Disconnect from command events
                DisconnectFromCommandEvents();

                isEventHandlerAttached = false;
                System.Diagnostics.Debug.WriteLine("Disconnected from AutoCAD document events");
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error disconnecting from document events: {ex.Message}");
            }
        }

        // Rename the method from Editor_ImpliedSelectionChanged to Document_ImpliedSelectionChanged
        private void Document_ImpliedSelectionChanged(object sender, EventArgs e)
        {
            // Prevent recursive calls
            if (_handlingSelectionChange) return;

            try
            {

                _handlingSelectionChange = true;

                if (currentEditor == null || currentDocument == null) return;

                // Check for unsaved changes before loading a new component
                if (hasUnsavedChanges && currentComponentId != ObjectId.Null)
                {
                    SaveChanges();
                }

                // Get the current selection
                PromptSelectionResult selRes = currentEditor.SelectImplied();

                if (selRes.Status == PromptStatus.OK && selRes.Value.Count == 1)
                {
                    ObjectId objId = selRes.Value[0].ObjectId;
                    LoadComponentData(objId);
                }
                else
                {
                    // Clear the panel
                    ClearComponentData();
                }
            }
            catch (System.Exception ex)
            {
                // Log the error but don't crash
                System.Diagnostics.Debug.WriteLine($"Error handling selection change: {ex.Message}");
            }
            finally
            {
                _handlingSelectionChange = false;
            }
        }

        private void DocumentManager_DocumentActivated(object sender, DocumentCollectionEventArgs e)
        {
            // Disconnect from old document
            DisconnectFromDocumentEvents();

            // Connect to new document
            ConnectToDocumentEvents();

            // Clear component data when switching documents
            ClearComponentData();
        }

        private void Editor_ImpliedSelectionChanged(object sender, EventArgs e)
        {
            try
            {
                // Check for unsaved changes before loading a new component
                if (hasUnsavedChanges && currentComponentId != ObjectId.Null)
                {
                    SaveChanges();
                }

                // Get the current selection
                PromptSelectionResult selRes = currentEditor.SelectImplied();

                if (selRes.Status == PromptStatus.OK && selRes.Value.Count == 1)
                {
                    ObjectId objId = selRes.Value[0].ObjectId;
                    LoadComponentData(objId);
                }
                else
                {
                    // Clear the panel
                    ClearComponentData();
                }
            }
            catch (System.Exception ex)
            {
                // Log the error but don't crash
                System.Diagnostics.Debug.WriteLine($"Error handling selection change: {ex.Message}");
            }
        }

        private void ConnectToCommandEvents()
        {
            try
            {
                if (currentDocument != null)
                {
                    // Subscribe to command events
                    currentDocument.CommandWillStart += Document_CommandWillStart;
                    currentDocument.CommandEnded += Document_CommandEnded;

                    System.Diagnostics.Debug.WriteLine("Connected to command events");
                }
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error connecting to command events: {ex.Message}");
            }
        }

        private void DisconnectFromCommandEvents()
        {
            try
            {
                if (currentDocument != null)
                {
                    // Unsubscribe from command events
                    currentDocument.CommandWillStart -= Document_CommandWillStart;
                    currentDocument.CommandEnded -= Document_CommandEnded;
                }
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error disconnecting from command events: {ex.Message}");
            }
        }

        private void Document_CommandWillStart(object sender, CommandEventArgs e)
        {
            // Handle command start - may want to temporarily disable UI
            System.Diagnostics.Debug.WriteLine($"Command starting: {e.GlobalCommandName}");

            // Check if command might modify our component
            if (e.GlobalCommandName == "ERASE" ||
                e.GlobalCommandName == "MOVE" ||
                e.GlobalCommandName == "STRETCH" ||
                e.GlobalCommandName == "TRIM" ||
                e.GlobalCommandName == "EXTEND")
            {
                // Save any unsaved changes before the command executes
                if (hasUnsavedChanges && currentComponentId != ObjectId.Null)
                {
                    SaveChanges();
                }
            }
        }

        private void Document_CommandEnded(object sender, CommandEventArgs e)
        {
            // Handle command end - refresh the component data if needed
            System.Diagnostics.Debug.WriteLine($"Command ended: {e.GlobalCommandName}");

            // Force refresh after commands that might modify components
            if (e.GlobalCommandName == "UPDATEMETAL" ||
                e.GlobalCommandName == "UPDATEPARENT" ||
                e.GlobalCommandName == "EXECUTECOPY" ||
                e.GlobalCommandName == "DETECTATTACHMENTS")
            {
                // Refresh the component data if we have an active selection
                if (currentComponentId != ObjectId.Null)
                {
                    try
                    {
                        // Verify the object still exists before trying to load it
                        using (Transaction tr = currentDocument.Database.TransactionManager.StartTransaction())
                        {
                            if (currentComponentId.IsValid && !currentComponentId.IsErased)
                            {
                                LoadComponentData(currentComponentId);
                            }
                            else
                            {
                                ClearComponentData();
                            }
                            tr.Commit();
                        }
                    }
                    catch
                    {
                        ClearComponentData();
                    }
                }
            }
        }

        // Add this method to handle external notifications of component changes
        public void OnComponentChanged(ObjectId componentId)
        {
            // Only process if this is the component we're currently showing
            if (componentId == currentComponentId)
            {
                // Reload the component data
                LoadComponentData(componentId);

                // Refresh the visualizations
                visualPanel.Invalidate();
                attachmentPanel.Invalidate();
            }
        }

        // Add a method to show processing status
        private void SetProcessingStatus(bool isProcessing, string message = "")
        {
            isProcessingUpdate = isProcessing;

            // You might want to add a status label to your UI
            if (statusLabel != null)
            {
                statusLabel.Text = message;
                statusLabel.Visible = !string.IsNullOrEmpty(message);
            }

            // Optionally change cursor
            this.Cursor = isProcessing ? Cursors.WaitCursor : Cursors.Default;

            // Disable controls during processing
            SetControlsEnabled(!isProcessing && currentComponentId != ObjectId.Null);

            // Force UI to update
            System.Windows.Forms.Application.DoEvents();
        }



        private void InitializeComponent()
        {
            this.Text = "Enhanced Metal Component Editor";
            this.Width = 600;
            this.Height = 800;
            this.AutoScaleMode = AutoScaleMode.Dpi; // Automatically scale based on DPI settings

            // Suspend layout to improve performance during initialization
            this.SuspendLayout();

            // Create the main TableLayoutPanel that will contain all other controls
            TableLayoutPanel mainLayout = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 1,
                RowCount = 5, // Header, Component Props, Visualization, Attachments, Parts List & Buttons
                Padding = new Padding(10),
            };

            // Add row styles to control proportional sizing
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 160)); // Component properties section
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 30));   // Visual panel
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 40));   // Attachment panel
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Percent, 30));   // Parts list
            mainLayout.RowStyles.Add(new RowStyle(SizeType.Absolute, 100)); // Buttons

            this.Controls.Add(mainLayout);

            // ========== COMPONENT PROPERTIES SECTION ==========
            TableLayoutPanel propertiesPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 2,
                RowCount = 6, // Type, Floor, Elevation, Length, Save button
                AutoSize = true
            };

            // Set column styles
            propertiesPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Absolute, 110)); // Label column
            propertiesPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 100));  // Content column

            // Title spanning both columns
            Label lblParentSection = new Label
            {
                Text = "Component Properties:",
                Dock = DockStyle.Fill,
                Font = new System.Drawing.Font(this.Font, System.Drawing.FontStyle.Bold)
            };
            propertiesPanel.Controls.Add(lblParentSection, 0, 0);
            propertiesPanel.SetColumnSpan(lblParentSection, 2);

            // Type row
            Label lblType = new Label { Text = "Type:", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft };
            propertiesPanel.Controls.Add(lblType, 0, 1);

            txtComponentType = new TextBox { Dock = DockStyle.Fill, Margin = new Padding(0, 3, 3, 3) };
            propertiesPanel.Controls.Add(txtComponentType, 1, 1);

            // Floor row
            Label lblFloor = new Label { Text = "Floor:", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft };
            propertiesPanel.Controls.Add(lblFloor, 0, 2);

            txtFloor = new TextBox { Dock = DockStyle.Fill, Margin = new Padding(0, 3, 3, 3) };
            propertiesPanel.Controls.Add(txtFloor, 1, 2);

            // Elevation row
            Label lblElevation = new Label { Text = "Elevation:", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft };
            propertiesPanel.Controls.Add(lblElevation, 0, 3);

            txtElevation = new TextBox { Dock = DockStyle.Fill, Margin = new Padding(0, 3, 3, 3) };
            propertiesPanel.Controls.Add(txtElevation, 1, 3);

            // Length row
            Label lblLengthCaption = new Label { Text = "Length:", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft };
            propertiesPanel.Controls.Add(lblLengthCaption, 0, 4);

            lblLength = new Label { Text = "0.0", Dock = DockStyle.Fill, TextAlign = ContentAlignment.MiddleLeft };
            propertiesPanel.Controls.Add(lblLength, 1, 4);

            // Save button - explicitly in its own row
            btnSaveParent = new Button
            {
                Text = "Save Component",
                Dock = DockStyle.Left, 
                Width = 150,
                Height = 30,
                Margin = new Padding(0, 10, 0, 0)
            };
            btnSaveParent.Click += BtnSaveParent_Click;

            // Add the button to the grid with some margin and alignment
            propertiesPanel.Controls.Add(btnSaveParent, 1, 5);

            mainLayout.Controls.Add(propertiesPanel, 0, 0);

            // ========== VISUALIZATION PANEL SECTION ==========
            Panel visualContainer = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(0, 10, 0, 0)
            };

            Label lblVisualization = new Label
            {
                Text = "",
                Dock = DockStyle.Bottom,
                Font = new System.Drawing.Font(this.Font, System.Drawing.FontStyle.Bold)
            };
            visualContainer.Controls.Add(lblVisualization);

            visualPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.White,
                Margin = new Padding(0, 5, 0, 0)
            };
            visualPanel.Paint += VisualPanel_Paint;
            visualContainer.Controls.Add(visualPanel);

            mainLayout.Controls.Add(visualContainer, 0, 1);

            // ========== ATTACHMENTS PANEL SECTION ==========
            Panel attachmentContainer = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(0, 10, 0, 0)
            };

            Label lblAttachments = new Label
            {
                Text = "Attachments",
                Dock = DockStyle.Bottom,
                Font = new System.Drawing.Font(this.Font, System.Drawing.FontStyle.Bold)
            };
            attachmentContainer.Controls.Add(lblAttachments);

            attachmentPanel = new Panel
            {
                Dock = DockStyle.Fill,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.White,
                Margin = new Padding(0, 5, 0, 0)
            };
            attachmentPanel.Paint += AttachmentPanel_Paint;
            attachmentContainer.Controls.Add(attachmentPanel);

            mainLayout.Controls.Add(attachmentContainer, 0, 2);

            // ========== PARTS LIST SECTION ==========
            Panel partsContainer = new Panel
            {
                Dock = DockStyle.Fill,
                Padding = new Padding(0, 25, 0, 0)
            };

            Label lblParts = new Label
            {
                Text = "Parts List",
                Dock = DockStyle.Bottom,
                Font = new System.Drawing.Font(this.Font, System.Drawing.FontStyle.Bold)
            };
            partsContainer.Controls.Add(lblParts);

            partsList = new ListView
            {
                Dock = DockStyle.Fill,
                View = View.Details,
                FullRowSelect = true,
                Margin = new Padding(0, 0, 0, 0)
            };
            partsList.Columns.Add("Name", 100);
            partsList.Columns.Add("Type", 60);
            partsList.Columns.Add("Length", 70);
            partsList.Columns.Add("Fixed", 50);
            partsList.Columns.Add("ShopUse", 50);
            partsList.Columns.Add("Adj Left", 70);  // Column 5 (0-based index)
            partsList.Columns.Add("Adj Right", 70); // Column 6 (0-based index)
            partsList.Columns.Add("Attach", 50);
            partsList.Columns.Add("Finish", 80);
            partsList.Columns.Add("Fab", 80);

            // Add the double-click handler
            partsList.DoubleClick += PartsList_DoubleClick;

            // Add selection changed handler to update the visualization
            partsList.SelectedIndexChanged += (s, e) => visualPanel.Invalidate();

            partsContainer.Controls.Add(partsList);
            mainLayout.Controls.Add(partsContainer, 0, 3);

            // ========== BUTTONS SECTION ==========
            TableLayoutPanel buttonPanel = new TableLayoutPanel
            {
                Dock = DockStyle.Fill,
                ColumnCount = 8,
                RowCount = 2,
                AutoSize = true
            };

            // Set equal column widths
            for (int i = 0; i < 8; i++)
            {
                buttonPanel.ColumnStyles.Add(new ColumnStyle(SizeType.Percent, 12.5F));
            }

            // First row of buttons
            btnAddPart = new Button { Text = "Add Part", Dock = DockStyle.Fill, Margin = new Padding(3) };
            btnAddPart.Click += BtnAddPart_Click;
            buttonPanel.Controls.Add(btnAddPart, 0, 0);

            btnEditPart = new Button { Text = "Edit Part", Dock = DockStyle.Fill, Margin = new Padding(3) };
            btnEditPart.Click += BtnEditPart_Click;
            buttonPanel.Controls.Add(btnEditPart, 1, 0);

            btnDeletePart = new Button { Text = "Delete Part", Dock = DockStyle.Fill, Margin = new Padding(3) };
            btnDeletePart.Click += BtnDeletePart_Click;
            buttonPanel.Controls.Add(btnDeletePart, 2, 0);

            btnSave = new Button { Text = "Save Changes", Dock = DockStyle.Fill, Margin = new Padding(3) };
            btnSave.Click += BtnSave_Click;
            buttonPanel.Controls.Add(btnSave, 3, 0);

            // Second row of buttons
            Button btnCopyProperty = new Button { Text = "Copy Property", Dock = DockStyle.Fill, Margin = new Padding(3) };
            btnCopyProperty.Click += BtnCopyProperty_Click;
            buttonPanel.Controls.Add(btnCopyProperty, 0, 1);

            Button btnDetectAttachments = new Button { Text = "Detect Attachments", Dock = DockStyle.Fill, Margin = new Padding(3) };
            btnDetectAttachments.Click += BtnDetectAttachments_Click;
            buttonPanel.Controls.Add(btnDetectAttachments, 1, 1);

            Button btnAddMetalPart = new Button { Text = "Add Metal Part", Dock = DockStyle.Fill, Margin = new Padding(3) };
            btnAddMetalPart.Click += BtnAddMetalPart_Click;
            buttonPanel.Controls.Add(btnAddMetalPart, 2, 1);

            mainLayout.Controls.Add(buttonPanel, 0, 4);

            // Status label
            statusLabel = new Label
            {
                Text = "",
                Dock = DockStyle.Bottom,
                Height = 20,
                TextAlign = ContentAlignment.MiddleLeft,
                BorderStyle = BorderStyle.None,
                Visible = false
            };
            this.Controls.Add(statusLabel);

            // Initially disable buttons until something is selected
            SetControlsEnabled(false);

            // Resume layout
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        private void VisualPanel_Paint(object sender, PaintEventArgs e)
        {
            if (currentComponentLength <= 0 || currentParts.Count == 0)
            {
                e.Graphics.DrawString("No component selected", this.Font, Brushes.Gray, 10, 10);
                return;
            }

            Graphics g = e.Graphics;
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

            // Add padding around the edges
            int borderPadding = 25;

            // for fixed length parts
            StringFormat near = new StringFormat { Alignment = StringAlignment.Near }; // left
            StringFormat far = new StringFormat { Alignment = StringAlignment.Far }; // right
            StringFormat none = new StringFormat { Alignment = StringAlignment.Center }; // center

            // Set scaling factor to fit component in panel with padding
            double scaleFactor = (visualPanel.Width - (borderPadding * 2)) / currentComponentLength;

            // Draw the baseline representing the component
            int yBase = (int)(visualPanel.Height * 0.8);
            int xStart = borderPadding;
            int xEnd = xStart + (int)(currentComponentLength * scaleFactor);

            // Draw baseline
            g.DrawLine(Pens.Black, xStart, yBase, xEnd, yBase);

            // Draw start and end markers
            g.DrawLine(Pens.Black, xStart, yBase - 5, xStart, yBase + 5);
            g.DrawLine(Pens.Black, xEnd, yBase - 5, xEnd, yBase + 5);

            // Draw length text
            g.DrawString($"Length: {currentComponentLength:F4}", this.Font, Brushes.Black, (xStart + xEnd) / 2 - 40, yBase + 10);

            // Determine if we're showing a horizontal or vertical component for proper visualization
            string orientation = txtComponentType.Text.ToUpper();
            isHorizontal = !orientation.Contains("VERTICAL");

            // Draw each part with improved spacing
            int partIndex = 0;
            int colorIndex = 0;
            Brush[] partBrushes = new Brush[] { Brushes.LightBlue, Brushes.LightGreen, Brushes.LightPink, Brushes.LightYellow, Brushes.LightCoral };

            // Determine spacing for parts
            int partHeight = 20;  // Height of each part

            // Calculate available space and adjust spacing
            int totalAvailableVerticalSpace = yBase - (borderPadding * 2);

            // Make sure we can fit all the parts with some spacing
            int partSpacing = Math.Max(25, Math.Min(35, (totalAvailableVerticalSpace - (currentParts.Count * partHeight)) / (currentParts.Count + 1)));

            // Debug
            System.Diagnostics.Debug.WriteLine($"Available space: {totalAvailableVerticalSpace}, Parts: {currentParts.Count}, Spacing: {partSpacing}");

            foreach (ChildPart part in currentParts)
            {
                StringFormat fmt = none;      // default
                if (part.Attach == "L") fmt = near;
                else if (part.Attach == "R") fmt = far;

                // Debug
                System.Diagnostics.Debug.WriteLine($"Drawing part {partIndex + 1}: {part.Name}");

                // Calculate part positions based on fixed length or adjustments
                int partStart, partEnd;

                if (part.IsFixedLength)
                {

                    // For fixed length parts, center on the component unless they have an attachment side
                    double fixedLengthInPixels = part.FixedLength * scaleFactor;

                    if (part.Attach == "L")
                    {
                        // Attach at left
                        partStart = xStart;
                        partEnd = xStart + (int)fixedLengthInPixels;
                    }
                    else if (part.Attach == "R")
                    {
                        // Attach at right
                        partEnd = xEnd;
                        partStart = xEnd - (int)fixedLengthInPixels;
                    }
                    else
                    {
                        // Center on the component
                        int middle = (xStart + xEnd) / 2;
                        partStart = middle - (int)(fixedLengthInPixels / 2);
                        partEnd = middle + (int)(fixedLengthInPixels / 2);
                    }
                }
                else
                {
                    // FIXED: Reversed adjustment logic to match expected behavior
                    // Negative adjustments should make parts shorter, positive should make them longer
                    partStart = xStart - (int)(part.StartAdjustment * scaleFactor);
                    partEnd = xEnd + (int)(part.EndAdjustment * scaleFactor);
                }

                // If partStart > partEnd due to extreme adjustments, skip this part
                if (partStart >= partEnd)
                {
                    System.Diagnostics.Debug.WriteLine($"Skipping part {part.Name} because partStart ({partStart}) >= partEnd ({partEnd})");
                    partIndex++;
                    colorIndex++;
                    continue;
                }

                // Calculate vertical position for this part - from top of panel downward
                int partY = borderPadding * 2 + (partIndex * (partHeight + partSpacing));

                // Debug
                System.Diagnostics.Debug.WriteLine($"Part {part.Name} at Y = {partY}, partStart = {partStart}, partEnd = {partEnd}");

                Brush partBrush = partBrushes[colorIndex % partBrushes.Length];

                // Use different border style for fixed length parts
                Pen borderPen = part.IsFixedLength ? new Pen(Color.DarkBlue, 2) : Pens.Black;

                // Draw part with subtle rounded corners for better appearance
                System.Drawing.Drawing2D.GraphicsPath path = new System.Drawing.Drawing2D.GraphicsPath();
                int cornerRadius = 3;
                path.AddArc(partStart, partY, cornerRadius * 2, cornerRadius * 2, 180, 90);
                path.AddArc(partEnd - cornerRadius * 2, partY, cornerRadius * 2, cornerRadius * 2, 270, 90);
                path.AddArc(partEnd - cornerRadius * 2, partY + partHeight - cornerRadius * 2, cornerRadius * 2, cornerRadius * 2, 0, 90);
                path.AddArc(partStart, partY + partHeight - cornerRadius * 2, cornerRadius * 2, cornerRadius * 2, 90, 90);
                path.CloseAllFigures();

                g.FillPath(partBrush, path);
                g.DrawPath(borderPen, path);

                // Highlight selected part
                if (partsList.SelectedItems.Count > 0 &&
                    partsList.SelectedItems[0].Tag == part)
                {
                    g.DrawPath(new Pen(Color.Red, 2), path);
                }

                // Draw part name and length info
                string displayName;
                System.Drawing.Font nameFont = new System.Drawing.Font(this.Font.FontFamily, 8, FontStyle.Regular);

                if (part.IsFixedLength)
                {
                    displayName = $"{part.Name} (Fixed: {part.FixedLength:F2})";

                    // Add "FIXED" indicator with better positioning
                    g.DrawString("FIXED", new System.Drawing.Font(nameFont, FontStyle.Bold), Brushes.Navy,
                                 partStart + (partEnd - partStart) / 2 - 15, partY - nameFont.Height - 5);
                }
                else
                {
                    displayName = $"{part.Name} ({part.StartAdjustment:F2}, {part.EndAdjustment:F2})";
                }

                // --- centred text rectangle ---
                RectangleF nameBox = new RectangleF(
                        partStart,                            // left edge of the bar
                        partY - nameFont.Height - 2,          // just above the bar
                        partEnd - partStart,                  // same width as the bar
                        nameFont.Height + 4);

                StringFormat centred = new StringFormat { Alignment = StringAlignment.Center };
                g.DrawString(displayName, nameFont, Brushes.Black, nameBox, fmt);

                // Add Finish and Fab info if they exist
                string extraInfo = "";
                if (!string.IsNullOrEmpty(part.Finish))
                    extraInfo += $"Finish: {part.Finish}";

                if (!string.IsNullOrEmpty(part.Fab))
                {
                    if (!string.IsNullOrEmpty(extraInfo))
                        extraInfo += ", ";
                    extraInfo += $"Fab: {part.Fab}";
                }

                // Display the extra info if it exists
                if (!string.IsNullOrEmpty(extraInfo))
                {
                    System.Drawing.Font infoFont = new System.Drawing.Font(this.Font.FontFamily, 7, FontStyle.Italic);
                    float extraInfoWidth = g.MeasureString(extraInfo, infoFont).Width;
                    int extraTextX = partStart + (partEnd - partStart) / 2 - (int)(extraInfoWidth / 2);
                    if (string.IsNullOrEmpty(part.Attach))
                    {
                                            // Position it inside the part
                    g.DrawString(extraInfo, infoFont, Brushes.DarkSlateBlue, extraTextX, partY - 1);
                    }

                }

                // Draw attachment indicators if applicable with better positioning
                if (!string.IsNullOrEmpty(part.Attach))
                {
                    string attachText = $"Attach: {part.Attach}";
                    System.Drawing.Font attachFont = new System.Drawing.Font(this.Font.FontFamily, 7, FontStyle.Regular);

                    if (part.Attach == "L")
                    {
                        g.DrawString(attachText, attachFont, Brushes.Blue, partStart, partY + partHeight + 2);
                        // Draw attachment arrow
                        g.DrawLine(new Pen(Color.Blue, 2), partStart, partY + partHeight + 15, partStart, partY + partHeight + 20);
                        g.FillPolygon(Brushes.Blue, new Point[] {
                    new Point(partStart, partY + partHeight + 20),
                    new Point(partStart - 5, partY + partHeight + 15),
                    new Point(partStart + 5, partY + partHeight + 15)
                });
                    }
                    else if (part.Attach == "R")
                    {
                        // Calculate width of attach text to position it properly
                        float attachTextWidth = g.MeasureString(attachText, attachFont).Width;
                        g.DrawString(attachText, attachFont, Brushes.Blue, partEnd - attachTextWidth, partY + partHeight + 2);

                        // Draw attachment arrow
                        g.DrawLine(new Pen(Color.Blue, 2), partEnd, partY + partHeight + 15, partEnd, partY + partHeight + 20);
                        g.FillPolygon(Brushes.Blue, new Point[] {
                    new Point(partEnd, partY + partHeight + 20),
                    new Point(partEnd - 5, partY + partHeight + 15),
                    new Point(partEnd + 5, partY + partHeight + 15)
                });
                    }
                }

                // If part has the Clips property set, indicate it
                if (part.Clips)
                {
                    System.Drawing.Font clipsFont = new System.Drawing.Font(this.Font.FontFamily, 7, FontStyle.Regular);
                    g.DrawString("Clips", clipsFont, Brushes.DarkGreen, partEnd + 5, partY);
                }

                partIndex++;
                colorIndex++;
            }


        }

        private void LoadComponentData(ObjectId objId)
        {
            SetProcessingStatus(true, "Loading component data...");

            try
            {
                Document doc = Application.DocumentManager.MdiActiveDocument;
                if (doc == null)
                {
                    ClearComponentData();
                    return;
                }

                Database db = doc.Database;

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    Entity ent = tr.GetObject(objId, OpenMode.ForRead) as Entity;
                if (ent is Polyline)
                {
                    Polyline pline = ent as Polyline;

                    // Check if it has METALCOMP Xdata
                    ResultBuffer rbComp = ent.GetXDataForApplication("METALCOMP");
                    if (rbComp != null)
                    {
                        // Extract component data from Xdata
                        string componentType = "";
                        string floor = "";
                        string elevation = "";

                        TypedValue[] xdataComp = rbComp.AsArray();
                        for (int i = 1; i < xdataComp.Length; i++) // Skip app name
                        {
                            if (i == 1) componentType = xdataComp[i].Value.ToString();
                            if (i == 2) floor = xdataComp[i].Value.ToString();
                            if (i == 3) elevation = xdataComp[i].Value.ToString();
                        }

                        // If this is a vertical component, load associated attachments
                        if (componentType.ToUpper().Contains("VERTICAL"))
                        {
                            try
                            {
                                // Load attachments from drawing data
                                List<Attachment> allAttachments = LoadAttachmentsFromDrawing();

                                // Filter to just the attachments for this vertical
                                string handleString = ent.Handle.ToString();
                                currentAttachments = allAttachments.Where(a => a.VerticalHandle == handleString).ToList();

                                System.Diagnostics.Debug.WriteLine($"Loaded {currentAttachments.Count} attachments for vertical {handleString}");

                                    // Explicitly force a repaint of the attachment panel
                                    if (attachmentPanel != null)
                                    {
                                        // Use BeginInvoke to ensure this happens after the current method completes
                                        this.BeginInvoke(new Action(() => {
                                            attachmentPanel.Invalidate();
                                            System.Diagnostics.Debug.WriteLine("Explicitly invalidated attachment panel");
                                        }));
                                    }
                                }
                            catch (System.Exception ex)
                            {
                                System.Diagnostics.Debug.WriteLine($"Error loading attachments: {ex.Message}");
                                currentAttachments.Clear();
                            }
                        }
                        else
                        {
                            // Clear attachments if not a vertical
                            currentAttachments.Clear();
                        }

                        // Get component length for part calculations
                        double length = 0;
                        if (pline.NumberOfVertices >= 2)
                        {
                            Point3d startPt = pline.GetPoint3dAt(0);
                            Point3d endPt = pline.GetPoint3dAt(1);
                            length = startPt.DistanceTo(endPt);
                        }

                        // Store the component length for visualization
                        currentComponentLength = length;

                        // Get the parts data
                        string partsJson = MetalComponentCommands.GetPartsJsonFromEntity(pline); // Use public method
                        currentParts.Clear();

                        try
                        {
                            string partsJsonx = MetalComponentCommands.GetPartsJsonFromEntity(pline);
                            System.Diagnostics.Debug.WriteLine($"Parts JSON length: {partsJsonx.Length}");

                            currentParts.Clear();

                            if (!string.IsNullOrEmpty(partsJsonx))
                            {
                                try
                                {
                                    currentParts = Newtonsoft.Json.JsonConvert.DeserializeObject<List<ChildPart>>(partsJson);
                                    System.Diagnostics.Debug.WriteLine($"Deserialized {currentParts.Count} parts");

                                    // Debug info about each part
                                    foreach (var part in currentParts)
                                    {
                                        System.Diagnostics.Debug.WriteLine($"Part: {part.Name}, Type: {part.PartType}, " +
                                            $"StartAdj: {part.StartAdjustment}, EndAdj: {part.EndAdjustment}");
                                    }
                                }
                                catch (System.Exception ex)
                                {
                                    System.Diagnostics.Debug.WriteLine($"Error deserializing parts: {ex.Message}");
                                    MessageBox.Show($"Error deserializing parts: {ex.Message}", "Parts Loading Error");
                                }
                            }
                            else
                            {
                                System.Diagnostics.Debug.WriteLine("No parts JSON found");
                                MessageBox.Show("No parts data found for this component", "No Parts");
                            }
                        }
                        catch (System.Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Exception reading parts: {ex.Message}");
                            MessageBox.Show($"Exception reading parts: {ex.Message}", "Parts Reading Error");
                        }

                        // Update the UI on the UI thread
                        Invoke(new Action(() => {
                            // Set parent data values
                            txtComponentType.Text = componentType;
                            txtFloor.Text = floor;
                            txtElevation.Text = elevation;
                            lblLength.Text = length.ToString("F4");

                            // Determine orientation for visualization
                            isHorizontal = !componentType.ToUpper().Contains("VERTICAL");

                            // Update the column headers based on orientation
                            if (partsList.Columns.Count >= 6) // Make sure we have enough columns
                            {
                                if (isHorizontal)
                                {
                                    // For horizontal components
                                    partsList.Columns[5].Text = "Adj Left";
                                    partsList.Columns[6].Text = "Adj Right";
                                }
                                else
                                {
                                    // For vertical components
                                    partsList.Columns[5].Text = "Adj Bott";
                                    partsList.Columns[6].Text = "Adj Top";
                                }
                            }

                            // Update parts list
                            UpdatePartsList();

                            if (forceRefreshOnNextSelection)
                            {
                                visualPanel.Invalidate();
                                attachmentPanel.Invalidate();
                                forceRefreshOnNextSelection = false;
                            }

                            SetControlsEnabled(true);
                        }));

                        currentComponentId = objId;
                        return;
                    }
                }

                tr.Commit();
            }

            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error loading component data: {ex.Message}");
                ClearComponentData();
            }
            finally
            {
                SetProcessingStatus(false);
            }
        }

        private List<Attachment> LoadAttachmentsFromDrawing()
        {
            // Use the centralized method instead
            return DrawingComponentManager.LoadAttachmentsFromDrawing();
        }

        // Add the Attachment paint method
        private void AttachmentPanel_Paint(object sender, PaintEventArgs e)
        {
            System.Diagnostics.Debug.WriteLine($"AttachmentPanel_Paint called with {currentAttachments?.Count ?? 0} attachments");
            if (currentAttachments == null || currentAttachments.Count == 0)
            {
                e.Graphics.DrawString("No attachments for this component", this.Font, Brushes.Gray, 10, 10);
                return;
            }


            Graphics g = e.Graphics;
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

            // Add padding around the edges
            int borderPadding = 25;

            // Get the vertical component's actual length
            double verticalLength = currentComponentLength;

            // Find the main vertical body part to get adjustments
            double startAdj = 0;
            double endAdj = 0;
            if (currentParts != null && currentParts.Count > 0)
            {
                var verticalBody = currentParts.FirstOrDefault(p => p.Clips);
                if (verticalBody != null)
                {
                    startAdj = verticalBody.StartAdjustment;
                    endAdj = verticalBody.EndAdjustment;

                    // Calculate adjusted length considering both start and end adjustments
                    verticalLength = currentComponentLength + startAdj + endAdj;
                }
            }

            System.Diagnostics.Debug.WriteLine($"Original length: {currentComponentLength}, Adjusted length: {verticalLength}");

            // Set scaling factor based on adjusted vertical length
            double scaleFactor = (attachmentPanel.Height - (borderPadding * 2)) / verticalLength;
            System.Diagnostics.Debug.WriteLine($"Scale factor: {scaleFactor}");

            // Draw the vertical component line
            int xCenter = attachmentPanel.Width / 2;  // Center of panel
            int yTop = borderPadding * 3;
            int yBottom = attachmentPanel.Height - borderPadding * 3; // Leave space at bottom

            // Draw vertical line
            g.DrawLine(new Pen(Color.Black, 3), xCenter, yTop, xCenter, yBottom);

            // Draw top and bottom markers
            g.DrawLine(Pens.Black, xCenter - 5, yTop, xCenter + 5, yTop);
            g.DrawLine(Pens.Black, xCenter - 5, yBottom, xCenter + 5, yBottom);

            // Draw length text
            g.DrawString($"Length: {verticalLength:F4}", this.Font, Brushes.Black, xCenter - 40, yBottom + 5);

            // Get start adjustment of the vertical component (if any)
            double verticalStartAdj = 0;
            if (currentParts != null && currentParts.Count > 0)
            {
                // Find a part with Clips=true, which should be the main vertical body
                var verticalBody = currentParts.FirstOrDefault(p => p.Clips);
                if (verticalBody != null)
                {
                    verticalStartAdj = verticalBody.StartAdjustment;
                }
            }

            // Process attachments by side first
            var leftAttachments = currentAttachments.Where(a => a.Side == "L").ToList();
            var rightAttachments = currentAttachments.Where(a => a.Side == "R").ToList();

            DrawSideAttachments(g, leftAttachments, xCenter, yTop, yBottom, verticalLength, scaleFactor, -1, startAdj);
            DrawSideAttachments(g, rightAttachments, xCenter, yTop, yBottom, verticalLength, scaleFactor, 1, startAdj);

            // Draw labels
            g.DrawString("Left Side", new System.Drawing.Font(this.Font, FontStyle.Bold), Brushes.Blue, 10, 10);
            g.DrawString("Right Side", new System.Drawing.Font(this.Font, FontStyle.Bold), Brushes.Red,
                         attachmentPanel.Width - 80, 10);


        }

        private void DrawSideAttachments(Graphics g, List<Attachment> attachments, int xCenter, int yTop, int yBottom,
                                       double verticalLength, double scaleFactor, int sideDirection, double verticalStartAdj)
        {
            // sideDirection: -1 for left, 1 for right
            int sideOffset = 70 * sideDirection;  // Distance from center
            Color sideColor = sideDirection < 0 ? Color.Blue : Color.Red;
            Brush textBrush = sideDirection < 0 ? Brushes.Blue : Brushes.Red;
            Pen linePen = new Pen(sideColor, 2);

            foreach (var attach in attachments)
            {
                try
                {
                    // send the attach.Height to the debug window
                    System.Diagnostics.Debug.WriteLine($"Attachment height: {attach.Height}");

                    double heightPercentage = attach.Height / verticalLength;
                    int yPos = yBottom - (int)((yBottom - yTop) * heightPercentage);

                    // Calculate adjusted height (accounting for vertical start adjustment)
                    double adjustedHeight = attach.Height + verticalStartAdj;
                    // send the adjustedHeight to the debug window
                    //System.Diagnostics.Debug.WriteLine($"Adjusted height: {adjustedHeight}");

                    // Calculate position on the Y axis, based on adjusted height
                    //int yPos = yBottom - (int)(adjustedHeight * scaleFactor);
                    // send the yPos to the debug window
                    System.Diagnostics.Debug.WriteLine($"Y position: {yPos}");

                    // Ensure it's within the vertical bounds
                    yPos = Math.Max(yTop, Math.Min(yBottom, yPos));

                    // Calculate horizontal position
                    int xPos = xCenter + sideOffset;

                    // Draw horizontal line
                    g.DrawLine(linePen, xCenter, yPos, xPos, yPos);
                    // output the points to the debug console
                    System.Diagnostics.Debug.WriteLine($"Drawing line from ({xCenter}, {yPos}) to ({xPos}, {yPos})");

                    // Add arrow at the end
                    g.FillPolygon(new SolidBrush(sideColor), new Point[] {
                new Point(xPos, yPos),
                new Point(xPos - (5 * sideDirection), yPos - 5),
                new Point(xPos - (5 * sideDirection), yPos + 5)
            });

                    // Draw part info
                    string partInfo = $"{attach.HorizontalPartType} to {attach.VerticalPartType}";
                    string heightInfo = $"Height: {adjustedHeight:F2}" + (attach.Adjust != 0 ? $", Adjust: {attach.Adjust:F2}" : "");
                    string invertInfo = attach.Invert ? "Inverted" : "";

                    // Position text based on side
                    float textX = sideDirection < 0 ? xPos - 200 : xPos + 10;

                    // Draw part type
                    g.DrawString(partInfo, this.Font, textBrush, textX, yPos - 20);

                    // Draw height and adjustment info
                    g.DrawString(heightInfo, new System.Drawing.Font(this.Font.FontFamily, 8), textBrush, textX, yPos - 5);

                    // Draw invert info if needed
                    if (!string.IsNullOrEmpty(invertInfo))
                    {
                        g.DrawString(invertInfo, new System.Drawing.Font(this.Font.FontFamily, 8, FontStyle.Italic), textBrush, textX, yPos + 10);
                    }
                }
                catch (System.Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error drawing attachment: {ex.Message}");
                }
            }
        }


        private void ClearComponentData()
        {
            // Clear the underlying data with null checks
            currentComponentId = ObjectId.Null;
            currentComponentLength = 0.0;

            if (currentParts != null)
            {
                currentParts.Clear();
            }
            else
            {
                // Initialize if it's null
                currentParts = new List<ChildPart>();
            }

            if (attachmentPanel != null)
            {
                currentAttachments.Clear();
                attachmentPanel.Invalidate();
            }

            // Update the UI on the UI thread with null checks
            if (this.IsHandleCreated)
            {
                Invoke(new Action(() =>
                {
                    // Clear parent data fields with null checks
                    if (txtComponentType != null) txtComponentType.Text = "";
                    if (txtFloor != null) txtFloor.Text = "";
                    if (txtElevation != null) txtElevation.Text = "";
                    if (lblLength != null) lblLength.Text = "0.0";

                    // Explicitly clear and refresh the parts list with null check
                    if (partsList != null)
                    {
                        partsList.Items.Clear();
                        partsList.Refresh();
                    }

                    // Update visualization panel
                    if (visualPanel != null)
                    {
                        visualPanel.Invalidate();
                    }

                    // Call SetControlsEnabled with parameter false
                    SetControlsEnabled(false);
                }));
            }
            else
            {
                // Handle is not created yet, just set values directly
                if (txtComponentType != null) txtComponentType.Text = "";
                if (txtFloor != null) txtFloor.Text = "";
                if (txtElevation != null) txtElevation.Text = "";
                if (lblLength != null) lblLength.Text = "0.0";

                // Explicitly clear and refresh the parts list with null check
                if (partsList != null)
                {
                    partsList.Items.Clear();
                    partsList.Refresh();
                }

                // Update visualization panel
                if (visualPanel != null)
                {
                    visualPanel.Invalidate();
                }

                // Call SetControlsEnabled with parameter false
                SetControlsEnabled(false);
            }
        }

        private void SetControlsEnabled(bool enabled)
        {
            // Check and set buttons if they're not null
            if (btnAddPart != null) btnAddPart.Enabled = enabled;
            if (btnEditPart != null) btnEditPart.Enabled = enabled;
            if (btnDeletePart != null) btnDeletePart.Enabled = enabled;
            if (btnSave != null) btnSave.Enabled = enabled;
            if (btnSaveParent != null) btnSaveParent.Enabled = enabled;

            // Enable/disable textboxes
            if (txtComponentType != null) txtComponentType.Enabled = enabled;
            if (txtFloor != null) txtFloor.Enabled = enabled;
            if (txtElevation != null) txtElevation.Enabled = enabled;

            // Enable/disable listview
            if (partsList != null) partsList.Enabled = enabled;
        }

        private void UpdatePartsList()
        {
            try
            {
                System.Diagnostics.Debug.WriteLine($"UpdatePartsList called with {currentParts?.Count ?? 0} parts");

                // Clear existing items
                partsList.Items.Clear();

                // Make sure we have the correct column structure
                if (partsList.Columns.Count < 10)  // includes finish and fabs
                {
                    // Remove existing columns first (if any)
                    partsList.Columns.Clear();

                    // Re-add the columns with updated layout
                    partsList.Columns.Add("Name", 100);
                    partsList.Columns.Add("Type", 60);
                    partsList.Columns.Add("Length", 70);
                    partsList.Columns.Add("Shop?", 40);
                    partsList.Columns.Add("Fixed", 50);   
                    partsList.Columns.Add("Adj Left", 70);
                    partsList.Columns.Add("Adj Bott", 70);
                    partsList.Columns.Add("Attach", 50);
                    partsList.Columns.Add("Finish", 80); 
                    partsList.Columns.Add("Fab", 80);    
                }

                if (currentParts == null)
                {
                    System.Diagnostics.Debug.WriteLine("currentParts is null");
                    return;
                }

                // Add each part to the list
                foreach (ChildPart part in currentParts)
                {
                    try
                    {
                        double actualLength = part.GetActualLength(currentComponentLength);

                        ListViewItem item = new ListViewItem(part.Name);
                        item.SubItems.Add(part.PartType);
                        item.SubItems.Add(actualLength.ToString("F4"));
                        item.SubItems.Add(part.IsShopUse ? "Yes" : "No"); // Show shop use status
                        item.SubItems.Add(part.IsFixedLength ? "Yes" : "No"); // Show fixed length status

                        // Show start/end adjustments (or N/A for fixed length parts)
                        if (part.IsFixedLength)
                        {
                            item.SubItems.Add("N/A");
                            item.SubItems.Add("N/A");
                        }
                        else
                        {
                            item.SubItems.Add(part.StartAdjustment.ToString("F4"));
                            item.SubItems.Add(part.EndAdjustment.ToString("F4"));
                        }

                        item.SubItems.Add(part.Attach ?? ""); // Add attachment property
                        item.SubItems.Add(part.Finish ?? "");
                        item.SubItems.Add(part.Fab ?? "");
                        item.Tag = part; // Store reference to the part object

                        partsList.Items.Add(item);
                        System.Diagnostics.Debug.WriteLine($"Added part to ListView: {part.Name}");
                    }
                    catch (System.Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error adding part {part?.Name}: {ex.Message}");
                    }
                }

                // Ensure the ListView is refreshed
                partsList.Refresh();

                // Update visualization after updating parts list
                visualPanel.Invalidate();
            }
            catch (System.Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Exception in UpdatePartsList: {ex.Message}");
                MessageBox.Show($"Error updating parts list: {ex.Message}", "Parts List Error");
            }
        }

        private void PartsList_DoubleClick(object sender, EventArgs e)
        {
            if (partsList.SelectedItems.Count == 0) return;

            // Get selected part
            ListViewItem item = partsList.SelectedItems[0];
            ChildPart part = item.Tag as ChildPart;
            if (part == null) return;

            // Show edit dialog
            using (PartEditForm form = new PartEditForm(part))
            {
                if (form.ShowDialog() == DialogResult.OK)
                {
                    // Update part in list
                    int index = currentParts.IndexOf(part);
                    if (index >= 0)
                    {
                        currentParts[index] = form.ResultPart;

                        // Refresh list display
                        UpdatePartsList();

                        // Refresh visualization
                        visualPanel.Invalidate();

                        // Mark that we have unsaved changes
                        hasUnsavedChanges = true;

                        // Save changes automatically
                        SaveChanges();
                    }
                }
            }
        }

        // Add this method to the EnhancedMetalComponentPanel class
        public void OnComponentErased(ObjectId erasedId)
        {
            if (currentComponentId == erasedId)
            {
                // Clear the panel since the displayed component was erased
                ClearComponentData();
            }
        }

        private void SaveChanges()
        {
            if (!hasUnsavedChanges || currentComponentId == ObjectId.Null) return;

            SetProcessingStatus(true, "Saving changes...");

            try
            {
                // Store the current component ID for re-selection later
                ObjectId idToReselect = currentComponentId;

                // Create a new list to avoid reference issues
                List<ChildPart> partsCopy = new List<ChildPart>();
                foreach (var part in currentParts)
                {
                    partsCopy.Add(part);
                }

                // Store in static fields for the command to use
                MetalComponentCommands.PendingHandle = currentComponentId.Handle.ToString();
                MetalComponentCommands.PendingParts = partsCopy;

                // Execute the command
                Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                if (doc == null)
                {
                    MessageBox.Show("Could not access active document.", "Error");
                    return;
                }

                doc.SendStringToExecute("UPDATEMETAL ", true, false, false);

                // Reset the flag after saving
                hasUnsavedChanges = false;

                // Set flag to force refresh on next selection
                forceRefreshOnNextSelection = true;

                // Re-select the entity and refresh the data after the command completes
                System.Threading.Tasks.Task.Delay(200).ContinueWith(t =>
                {
                    try
                    {
                        // This must be executed on the main thread
                        this.BeginInvoke(new Action(() =>
                        {
                            Editor ed = doc.Editor;

                            // Clear any current selection
                            ed.SetImpliedSelection(new ObjectId[0]);

                            // Create a new selection set with just our object
                            ed.SetImpliedSelection(new ObjectId[] { idToReselect });

                            // Explicitly reload the component data to refresh the display
                            LoadComponentData(idToReselect);

                            System.Diagnostics.Debug.WriteLine("Re-selected component and refreshed data after save");
                        }));
                    }
                    catch (System.Exception ex)
                    {
                        System.Diagnostics.Debug.WriteLine($"Error re-selecting: {ex.Message}");
                    }
                });

                // Log successful save
                System.Diagnostics.Debug.WriteLine("Changes automatically saved.");
                hasUnsavedChanges = false;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Error auto-saving changes: {ex.Message}", "Save Error");
            }
            finally
            {
                SetProcessingStatus(false);
            }
        }

        // Add this at the end of the class
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                // Disconnect from all events
                DisconnectFromDocumentEvents();
            }

            base.Dispose(disposing);
        }

        private void BtnSaveParent_Click(object sender, EventArgs e)
        {
            if (currentComponentId == ObjectId.Null) return;

            try
            {
                // Store in static fields for the command to use
                MetalComponentCommands.PendingHandle = currentComponentId.Handle.ToString();
                MetalComponentCommands.PendingParentData = new MetalComponentCommands.ParentComponentData
                {
                    ComponentType = txtComponentType.Text,
                    Floor = txtFloor.Text,
                    Elevation = txtElevation.Text
                };

                // Execute the command
                Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                doc.SendStringToExecute("UPDATEPARENT ", true, false, false);

                MessageBox.Show("Component data saved successfully.", "Save Complete");

                // Update visualization based on new component type
                isHorizontal = !txtComponentType.Text.ToUpper().Contains("VERTICAL");
                visualPanel.Invalidate();
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Error saving component data: {ex.Message}", "Save Error");
            }
        }

        private void BtnAddPart_Click(object sender, EventArgs e)
        {
            // Show dialog to add new part
            using (PartEditForm form = new PartEditForm())
            {
                if (form.ShowDialog() == DialogResult.OK)
                {
                    // Add the new part
                    currentParts.Add(form.ResultPart);

                    // Update the UI
                    UpdatePartsList();

                    // Update visualization
                    visualPanel.Invalidate();

                    // Mark that we have unsaved changes
                    hasUnsavedChanges = true;

                    // Save changes automatically
                    SaveChanges();
                }
            }
        }

        private void BtnEditPart_Click(object sender, EventArgs e)
        {
            if (partsList.SelectedItems.Count == 0) return;

            ListViewItem item = partsList.SelectedItems[0];
            ChildPart part = item.Tag as ChildPart;

            using (PartEditForm form = new PartEditForm(part, currentComponentId))
            {
                if (form.ShowDialog() == DialogResult.OK)
                {
                    // Update the part in the list
                    int index = currentParts.IndexOf(part);
                    currentParts[index] = form.ResultPart;

                    // Update the UI
                    //double length = double.Parse(lblLength.Text);
                    //double actualLength = form.ResultPart.GetActualLength(length);

                    //item.Text = form.ResultPart.Name;
                    //item.SubItems[1].Text = form.ResultPart.PartType;
                    //item.SubItems[2].Text = actualLength.ToString("F4");
                    //item.Tag = form.ResultPart;

                    // Update the UI
                    UpdatePartsList();

                    // Update visualization
                    visualPanel.Invalidate();

                    // Mark that we have unsaved changes
                    hasUnsavedChanges = true;

                    // Save changes automatically
                    SaveChanges();
                }
            }
        }

        private void BtnDeletePart_Click(object sender, EventArgs e)
        {
            if (partsList.SelectedItems.Count == 0) return;

            if (MessageBox.Show("Are you sure you want to delete this part?",
                                "Confirm Delete", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                ListViewItem item = partsList.SelectedItems[0];
                ChildPart part = item.Tag as ChildPart;

                currentParts.Remove(part);
               partsList.Items.Remove(item);

                // Update visualization after removing part
                visualPanel.Invalidate();

                // Mark that we have unsaved changes
                hasUnsavedChanges = true;

                // Save changes automatically
                SaveChanges();
            }
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            if (currentComponentId == ObjectId.Null)
            {
                MessageBox.Show("No component currently selected.", "No Selection");
                return;
            }

            try
            {
                Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                if (doc == null)
                {
                    MessageBox.Show("Could not access active document.", "Error");
                    return;
                }

                // First, save the component properties (type, floor, elevation)
                // Similar to what BtnSaveParent_Click does
                MetalComponentCommands.PendingHandle = currentComponentId.Handle.ToString();
                MetalComponentCommands.PendingParentData = new MetalComponentCommands.ParentComponentData
                {
                    ComponentType = txtComponentType.Text,
                    Floor = txtFloor.Text,
                    Elevation = txtElevation.Text
                };
                doc.SendStringToExecute("UPDATEPARENT ", true, false, true); // Note: last parameter is true for synchronous execution

                // Then, save the child parts data
                // Create a new list to avoid reference issues
                List<ChildPart> partsCopy = new List<ChildPart>();
                foreach (var part in currentParts)
                {
                    partsCopy.Add(part);
                }

                // Store in static fields for the command to use
                MetalComponentCommands.PendingHandle = currentComponentId.Handle.ToString();
                MetalComponentCommands.PendingParts = partsCopy;

                // Execute the command
                doc.SendStringToExecute("UPDATEMETAL ", true, false, true); // Use synchronous execution

                MessageBox.Show("Changes saved successfully.", "Save Complete");
                forceRefreshOnNextSelection = true;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Error initiating save: {ex.Message}\n\nStack Trace: {ex.StackTrace}", "Save Error");
            }
        }

        private void BtnCopyProperty_Click(object sender, EventArgs e)
        {
            // Execute the COPYPROPERTY command
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (doc != null)
            {
                try
                {
                    // Save any changes first if a component is currently selected
                    if (currentComponentId != ObjectId.Null)
                    {
                        BtnSave_Click(sender, e);
                    }

                    // Execute the command
                    doc.SendStringToExecute("COPYPROPERTY ", true, false, false);
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show($"Error executing COPYPROPERTY command: {ex.Message}", "Command Error");
                }
            }
        }

        private void BtnDetectAttachments_Click(object sender, EventArgs e)
        {
            // Execute the DETECTATTACHMENTS command
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (doc != null)
            {
                try
                {
                    // Save any changes first if a component is currently selected
                    if (currentComponentId != ObjectId.Null)
                    {
                        BtnSave_Click(sender, e);
                    }

                    // Execute the command
                    doc.SendStringToExecute("DETECTATTACHMENTS ", true, false, false);

                    // If a vertical component is currently selected, refresh the attachment panel
                    string componentType = txtComponentType.Text.ToUpper();
                    if (componentType.Contains("VERTICAL"))
                    {
                        // Reload the component to refresh the attachments
                        LoadComponentData(currentComponentId);
                    }
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show($"Error executing DETECTATTACHMENTS command: {ex.Message}", "Command Error");
                }
            }
        }

        private void BtnAddMetalPart_Click(object sender, EventArgs e)
        {
            // Execute the ADDMETALPART command
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (doc != null)
            {
                try
                {
                    // Execute the command
                    doc.SendStringToExecute("ADDMETALPART ", true, false, false);
                }
                catch (System.Exception ex)
                {
                    MessageBox.Show($"Error executing ADDMETALPART command: {ex.Message}", "Command Error");
                }
            }
        }

        private void BtnCreateOpening_Click(object sender, EventArgs e)
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            try
            {
                // Make sure we have a clean state
                if (MetalComponentCommands.PendingOpeningData == null)
                {
                    MetalComponentCommands.PendingOpeningData = new MetalComponentCommands.OpeningComponentData();
                }

                // Set values from existing panel controls
                MetalComponentCommands.PendingOpeningData.Floor = txtFloor.Text;
                MetalComponentCommands.PendingOpeningData.Elevation = txtElevation.Text;

                // Determine component types based on current component type
                bool isHorizontal = txtComponentType.Text.ToUpper().Contains("HORIZONTAL");
                bool isVertical = txtComponentType.Text.ToUpper().Contains("VERTICAL");

                MetalComponentCommands.PendingOpeningData.TopType = isHorizontal ? "Head" : "Horizontal";
                MetalComponentCommands.PendingOpeningData.BottomType = isHorizontal ? "Sill" : "Horizontal";
                MetalComponentCommands.PendingOpeningData.LeftType = isVertical ? "JambL" : "Vertical";
                MetalComponentCommands.PendingOpeningData.RightType = isVertical ? "JambR" : "Vertical";

                // Enable all sides by default
                MetalComponentCommands.PendingOpeningData.CreateTop = true;
                MetalComponentCommands.PendingOpeningData.CreateBottom = true;
                MetalComponentCommands.PendingOpeningData.CreateLeft = true;
                MetalComponentCommands.PendingOpeningData.CreateRight = true;

                // Force the dialog to open
                MetalComponentCommands.PendingOpeningData.ForceDialogToOpen = true;

                // IMPORTANT: Also set the current parts as PendingParts
                if (currentParts != null && currentParts.Count > 0)
                {
                    // Create a deep copy of the parts
                    List<ChildPart> partsCopy = new List<ChildPart>();
                    foreach (var part in currentParts)
                    {
                        partsCopy.Add(part);
                    }

                    // Set the pending parts
                    MetalComponentCommands.PendingParts = partsCopy;

                    System.Diagnostics.Debug.WriteLine($"Setting {partsCopy.Count} pending parts from panel");
                }

                // Debug output
                System.Diagnostics.Debug.WriteLine("---- SETTING OPENING DATA IN PANEL ----");
                System.Diagnostics.Debug.WriteLine($"Floor: {MetalComponentCommands.PendingOpeningData.Floor}");
                System.Diagnostics.Debug.WriteLine($"Elevation: {MetalComponentCommands.PendingOpeningData.Elevation}");
                System.Diagnostics.Debug.WriteLine($"TopType: {MetalComponentCommands.PendingOpeningData.TopType}");
                System.Diagnostics.Debug.WriteLine($"BottomType: {MetalComponentCommands.PendingOpeningData.BottomType}");
                System.Diagnostics.Debug.WriteLine($"LeftType: {MetalComponentCommands.PendingOpeningData.LeftType}");
                System.Diagnostics.Debug.WriteLine($"RightType: {MetalComponentCommands.PendingOpeningData.RightType}");
                System.Diagnostics.Debug.WriteLine("---- END SETTING OPENING DATA ----");

                // Execute the command
                doc.SendStringToExecute("CREATEOPENINGCOMPONENTS ", true, false, false);

                // Show a message to the user
                MessageBox.Show("Opening component dialog will open. After configuring settings, follow prompts to select the center of an opening.",
                               "Create Opening Components", MessageBoxButtons.OK, MessageBoxIcon.Information);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Error initiating opening component creation: {ex.Message}",
                               "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void RegisterApp(RegAppTable regTable, string appName, Transaction tr)
        {
            if (!regTable.Has(appName))
            {
                regTable.UpgradeOpen();
                RegAppTableRecord appRec = new RegAppTableRecord();
                appRec.Name = appName;
                regTable.Add(appRec);
                tr.AddNewlyCreatedDBObject(appRec, true);
            }
        }



    }

    // Form for adding/editing parts
    public class PartEditForm : Form
    {
        private TextBox txtName;
        private TextBox txtType;
        private CheckBox chkShopUse;
        private TextBox txtStartAdj;
        private TextBox txtEndAdj;
        private TextBox txtMaterial;
        private TextBox txtFinish;
        private TextBox txtFab;

        // New controls for fixed length parts
        private CheckBox chkFixedLength;
        private TextBox txtFixedLength;
        private Label lblFixedLength;
        private Label lblStartAdj;
        private Label lblEndAdj;

        private ComboBox cboAttachControl;
        private CheckBox chkInvertControl;
        private TextBox txtAdjustControl;
        private CheckBox chkClipsControl;

        // The part being edited (if any)
        private ChildPart partToEdit;

        private ObjectId _componentId;

        // Result part after editing
        public ChildPart ResultPart { get; private set; }

        // Constructor - accept an optional part to edit
        public PartEditForm(ChildPart part = null, ObjectId componentId = default)
        {
            // Store the part being edited
            this.partToEdit = part;

            // Store the component ID
            this._componentId = componentId;

            // Initialize the form components
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.Text = "Edit Part";
            this.Width = 500;  // Increased width for more controls
            this.Height = 580; // Increased height for more controls
            this.FormBorderStyle = FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.StartPosition = FormStartPosition.CenterParent;

            // Basic part properties
            Label lblName = new Label { Text = "Name:", Left = 10, Top = 10, Width = 80 };
            this.Controls.Add(lblName);

            txtName = new TextBox { Left = 100, Top = 10, Width = 280 };
            this.Controls.Add(txtName);

            Label lblType = new Label { Text = "Type:", Left = 10, Top = 40, Width = 80 };
            this.Controls.Add(lblType);

            txtType = new TextBox { Left = 100, Top = 40, Width = 280 };
            this.Controls.Add(txtType);

            // Fixed length checkbox
            chkShopUse = new CheckBox
            {
                Text = "Shop Use",
                Left = 300,
                Top = 70,
                Width = 100
            };

            this.Controls.Add(chkShopUse);

            // Fixed length checkbox
            chkFixedLength = new CheckBox
            {
                Text = "Fixed Length Part",
                Left = 100,
                Top = 70,
                Width = 280
            };
            chkFixedLength.CheckedChanged += ChkFixedLength_CheckedChanged;
            this.Controls.Add(chkFixedLength);

            // Fixed length control
            lblFixedLength = new Label { Text = "Fixed Length:", Left = 10, Top = 100, Width = 80 };
            this.Controls.Add(lblFixedLength);

            txtFixedLength = new TextBox { Left = 100, Top = 100, Width = 280, Text = "1.25" };
            this.Controls.Add(txtFixedLength);

            // Start adjustment control
            lblStartAdj = new Label { Text = "Start Adj:", Left = 10, Top = 130, Width = 80 };
            this.Controls.Add(lblStartAdj);

            txtStartAdj = new TextBox { Left = 100, Top = 130, Width = 280, Text = "0.0" };
            this.Controls.Add(txtStartAdj);

            // End adjustment control
            lblEndAdj = new Label { Text = "End Adj:", Left = 10, Top = 160, Width = 80 };
            this.Controls.Add(lblEndAdj);

            txtEndAdj = new TextBox { Left = 100, Top = 160, Width = 280, Text = "0.0" };
            this.Controls.Add(txtEndAdj);

            // Material control
            Label lblMaterial = new Label { Text = "Material:", Left = 10, Top = 190, Width = 80 };
            this.Controls.Add(lblMaterial);

            txtMaterial = new TextBox { Left = 100, Top = 190, Width = 280 };
            this.Controls.Add(txtMaterial);

            // Finish field
            Label lblFinish = new Label { Text = "Finish:", Left = 10, Top = 220, Width = 80 };
            this.Controls.Add(lblFinish);

            txtFinish = new TextBox { Left = 100, Top = 220, Width = 280 };
            this.Controls.Add(txtFinish);

            // Fab field
            Label lblFab = new Label { Text = "Fab:", Left = 10, Top = 250, Width = 80 };
            this.Controls.Add(lblFab);

            txtFab = new TextBox { Left = 100, Top = 250, Width = 280 };
            this.Controls.Add(txtFab);

            // New attachment properties section label - moved down
            Label lblAttachmentSection = new Label
            {
                Text = "Attachment Properties:",
                Left = 10,
                Top = 280,
                Width = 380,
                Font = new System.Drawing.Font(this.Font, System.Drawing.FontStyle.Bold)
            };
            this.Controls.Add(lblAttachmentSection);

            // Attach property (L/R)
            Label lblAttach = new Label { Text = "Attach:", Left = 10, Top = 280, Width = 80 };
            this.Controls.Add(lblAttach);

            ComboBox cboAttach = new ComboBox
            {
                Left = 100,
                Top = 280,
                Width = 280,
                DropDownStyle = ComboBoxStyle.DropDownList
            };
            cboAttach.Items.AddRange(new object[] { "", "L", "R" });
            this.Controls.Add(cboAttach);

            // Invert checkbox
            CheckBox chkInvert = new CheckBox
            {
                Text = "Invert",
                Left = 100,
                Top = 310,
                Width = 280
            };
            this.Controls.Add(chkInvert);

            // Adjust value
            Label lblAdjust = new Label { Text = "Adjust:", Left = 10, Top = 340, Width = 80 };
            this.Controls.Add(lblAdjust);

            TextBox txtAdjust = new TextBox
            {
                Left = 100,
                Top = 340,
                Width = 280,
                Text = "0.0"
            };
            this.Controls.Add(txtAdjust);

            // Clips checkbox (for vertical parts)
            CheckBox chkClips = new CheckBox
            {
                Text = "Allows Clips",
                Left = 100,
                Top = 370,
                Width = 280
            };
            this.Controls.Add(chkClips);

            // Add help text to explain the adjustments
            Label lblHelp = new Label
            {
                Text = "For horizontal parts: Start=Left, End=Right\nFor vertical parts: Start=Bottom, End=Top",
                Left = 100,
                Top = 400,
                Width = 280,
                Height = 40
            };
            this.Controls.Add(lblHelp);

            // OK and Cancel buttons - moved down
            Button btnOK = new Button
            {
                Text = "OK",
                Left = 210,
                Top = 450,
                Width = 80,
                DialogResult = DialogResult.OK
            };
            btnOK.Click += BtnOK_Click;
            this.Controls.Add(btnOK);

            Button btnCancel = new Button
            {
                Text = "Cancel",
                Left = 300,
                Top = 450,
                Width = 80,
                DialogResult = DialogResult.Cancel
            };
            this.Controls.Add(btnCancel);

            this.AcceptButton = btnOK;
            this.CancelButton = btnCancel;

            // Store controls as class fields
            this.cboAttachControl = cboAttach;
            this.chkInvertControl = chkInvert;
            this.txtAdjustControl = txtAdjust;
            this.chkClipsControl = chkClips;

            // If part is provided, set values
            if (partToEdit != null)
            {
                txtName.Text = partToEdit.Name;
                txtType.Text = partToEdit.PartType;
                chkShopUse.Checked = partToEdit.IsShopUse;
                // Set fixed length properties
                chkFixedLength.Checked = partToEdit.IsFixedLength;
                txtFixedLength.Text = partToEdit.FixedLength.ToString();

                // Set start/end adjustment values
                txtStartAdj.Text = partToEdit.StartAdjustment.ToString();
                txtEndAdj.Text = partToEdit.EndAdjustment.ToString();

                txtMaterial.Text = partToEdit.Material;
                // Set values for finish and fab
                txtFinish.Text = partToEdit.Finish ?? "";
                txtFab.Text = partToEdit.Fab ?? "";

                // Set attachment properties
                cboAttachControl.SelectedItem = partToEdit.Attach ?? "";
                chkInvertControl.Checked = partToEdit.Invert;
                txtAdjustControl.Text = partToEdit.Adjust.ToString();
                chkClipsControl.Checked = partToEdit.Clips;
            }

            // Update UI controls based on fixed length setting
            UpdateControlsForFixedLength(chkFixedLength.Checked);
        }

        private void ChkFixedLength_CheckedChanged(object sender, EventArgs e)
        {
            // Enable/disable appropriate controls based on fixed length checkbox
            UpdateControlsForFixedLength(chkFixedLength.Checked);
        }

        private void UpdateControlsForFixedLength(bool isFixedLength)
        {
            // Enable/disable and adjust visibility of controls
            txtFixedLength.Enabled = isFixedLength;
            lblFixedLength.Enabled = isFixedLength;

            txtStartAdj.Enabled = !isFixedLength;
            txtEndAdj.Enabled = !isFixedLength;
            lblStartAdj.Enabled = !isFixedLength;
            lblEndAdj.Enabled = !isFixedLength;
        }

        private void BtnOK_Click(object sender, EventArgs e)
        {
            try
            {
                // Validate basic fields
                if (string.IsNullOrEmpty(txtName.Text) || string.IsNullOrEmpty(txtType.Text))
                {
                    MessageBox.Show("Name and Type are required fields.", "Validation Error");
                    this.DialogResult = DialogResult.None;
                    return;
                }

                // Parse numeric values
                double startAdjustment = double.Parse(txtStartAdj.Text);
                double endAdjustment = double.Parse(txtEndAdj.Text);
                double fixedLength = double.Parse(txtFixedLength.Text);
                double adjust = double.Parse(txtAdjustControl.Text);
                bool isFixedLength = chkFixedLength.Checked;
                bool isShopUse = chkShopUse.Checked;

                // Create result part based on fixed length setting
                if (isFixedLength)
                {
                    // Use the fixed length constructor
                    ResultPart = new ChildPart(
                        txtName.Text,
                        txtType.Text,
                        fixedLength,
                        txtMaterial.Text,
                        true // isFixed
                    );
                }
                else
                {
                    // Use the adjustable length constructor
                    ResultPart = new ChildPart(
                        txtName.Text,
                        txtType.Text,
                        startAdjustment,
                        endAdjustment,
                        txtMaterial.Text
                    );
                }

                // Set attachment properties
                ResultPart.Attach = cboAttachControl.SelectedItem?.ToString() ?? "";
                ResultPart.Invert = chkInvertControl.Checked;
                ResultPart.Adjust = adjust;
                ResultPart.Clips = chkClipsControl.Checked;
                ResultPart.IsShopUse = chkShopUse.Checked;
                ResultPart.Finish = txtFinish.Text;
                ResultPart.Fab = txtFab.Text;

                // Process mark numbers if we have a valid component ID
                if (_componentId != ObjectId.Null && _componentId.IsValid)
                {
                    Document doc = Application.DocumentManager.MdiActiveDocument;
                    using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
                    {
                        MarkNumberManager.Instance.ProcessComponentMarkNumbers(_componentId, tr);
                        tr.Commit();
                    }
                }
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Validation Error");
                this.DialogResult = DialogResult.None;
            }
        }
    }

   

}
