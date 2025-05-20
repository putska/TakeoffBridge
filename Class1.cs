using System;
using System.Collections.Generic;
using System.Net;
using System.Net.Http;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.Runtime;
using Newtonsoft.Json;
using TakeoffBridge;
using System.Linq;


// This line is necessary to tell AutoCAD to load your plugin
[assembly: CommandClass(typeof(GlassTakeoffBridge.GlassTakeoffCommands))]
[assembly: CommandClass(typeof(GlassTakeoffBridge.GlassTakeoffCommands))]
[assembly: CommandClass(typeof(TakeoffBridge.MetalComponentCommands))]
[assembly: CommandClass(typeof(TakeoffBridge.MarkNumberManagerCommands))]
[assembly: CommandClass(typeof(TakeoffBridge.MarkNumberDisplayCommands))]
[assembly: CommandClass(typeof(TakeoffBridge.ElevationManagerCommands))]
[assembly: CommandClass(typeof(TakeoffBridge.TemplateGeneratorCommands))]
[assembly: CommandClass(typeof(TakeoffBridge.FabricationTicketCommands))]
[assembly: CommandClass(typeof(TakeoffBridge.WorkPointManager))]
[assembly: CommandClass(typeof(TakeoffBridge.MullionPlacementData))]
[assembly: CommandClass(typeof(TakeoffBridge.FabricationManager))]



namespace GlassTakeoffBridge
{
    // Glass data class to store all necessary information 
    public class GlassData
    {
        public string GlassType { get; set; } = "1";
        public double GlassBiteLeft { get; set; } = 0.5;
        public double GlassBiteRight { get; set; } = 0.5;
        public double GlassBiteTop { get; set; } = 0.5;
        public double GlassBiteBottom { get; set; } = 0.5;
        public string Floor { get; set; } = "01";
        public string Elevation { get; set; } = "A";
        public double GlassWidth { get; set; } // Width including bites
        public double GlassHeight { get; set; } // Height including bites
        public double DloWidth { get; set; } // Daylight opening width
        public double DloHeight { get; set; } // Daylight opening height
        public string MarkNumber { get; set; } // Assigned mark number
    }



    // Glass mark number manager to handle unique glass dimensions
    public class GlassMarkNumberManager
    {
        // Singleton instance
        private static GlassMarkNumberManager _instance;

        // Cache of unique glass dimensions and their assigned mark numbers
        private Dictionary<string, string> _glassMarks = new Dictionary<string, string>();

        // Track the next available number for each glass type
        private Dictionary<string, int> _nextMarkNumbers = new Dictionary<string, int>();

        // Private constructor for singleton pattern
        private GlassMarkNumberManager()
        {
            // Initialize by loading existing mark numbers from the drawing
            InitializeFromDrawing();
        }

        // Singleton access
        public static GlassMarkNumberManager Instance
        {
            get
            {
                if (_instance == null)
                {
                    _instance = new GlassMarkNumberManager();
                }
                return _instance;
            }
        }

        // Initialize by loading existing glass mark numbers
        private void InitializeFromDrawing()
        {
            _glassMarks.Clear();
            _nextMarkNumbers.Clear();

            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
            {
                try
                {
                    // Get all polylines in the drawing
                    TypedValue[] tvs = new TypedValue[] {
                        new TypedValue((int)DxfCode.Start, "LWPOLYLINE"),
                        new TypedValue((int)DxfCode.LayerName, "Tag1")
                    };

                    SelectionFilter filter = new SelectionFilter(tvs);
                    PromptSelectionResult selRes = doc.Editor.SelectAll(filter);

                    if (selRes.Status == PromptStatus.OK)
                    {
                        foreach (SelectedObject selObj in selRes.Value)
                        {
                            Entity ent = tr.GetObject(selObj.ObjectId, OpenMode.ForRead) as Entity;
                            if (ent is Polyline pline)
                            {
                                // Check if it has GLASS Xdata
                                ResultBuffer rbXdata = ent.GetXDataForApplication("GLASS");
                                if (rbXdata != null)
                                {
                                    // Extract glass data
                                    GlassData glassData = GlassTakeoffCommands.ExtractGlassDataFromXData(rbXdata);

                                    if (!string.IsNullOrEmpty(glassData.MarkNumber))
                                    {
                                        // Parse mark number to get the numeric part
                                        string[] markParts = glassData.MarkNumber.Split('-');
                                        if (markParts.Length == 2 && int.TryParse(markParts[1], out int markNumber))
                                        {
                                            string glassType = markParts[0];

                                            // Track the highest mark number for each glass type
                                            if (!_nextMarkNumbers.ContainsKey(glassType) ||
                                                _nextMarkNumbers[glassType] <= markNumber)
                                            {
                                                _nextMarkNumbers[glassType] = markNumber + 1;
                                            }

                                            // Generate unique key based on dimensions
                                            string uniqueKey = GenerateGlassKey(glassData);
                                            _glassMarks[uniqueKey] = glassData.MarkNumber;
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
                    doc.Editor.WriteMessage($"\nError initializing glass mark numbers: {ex.Message}");
                    tr.Abort();
                }
            }
        }

        // Generate a unique key for glass based on dimensions
        private string GenerateGlassKey(GlassData glassData)
        {
            // Round dimensions to avoid floating-point comparison issues
            double roundedWidth = Math.Round(glassData.GlassWidth, 3);
            double roundedHeight = Math.Round(glassData.GlassHeight, 3);

            return $"{glassData.GlassType}|{roundedWidth}|{roundedHeight}";
        }

        // Get or create a mark number for a glass configuration
        public string GetOrCreateMarkNumber(GlassData glassData)
        {
            string uniqueKey = GenerateGlassKey(glassData);

            // Check if we already have a mark number for this configuration
            if (_glassMarks.ContainsKey(uniqueKey))
            {
                return _glassMarks[uniqueKey];
            }

            // Create a new mark number
            string glassType = glassData.GlassType;
            if (!_nextMarkNumbers.ContainsKey(glassType))
            {
                _nextMarkNumbers[glassType] = 1;
            }

            int markNumber = _nextMarkNumbers[glassType]++;
            string newMarkNumber = $"{glassType}-{markNumber}";

            // Cache it
            _glassMarks[uniqueKey] = newMarkNumber;

            return newMarkNumber;
        }
    }

    // UI Panel for Glass Takeoff
    public class GlassPanel : System.Windows.Forms.UserControl
    {
        private System.Windows.Forms.TextBox txtGlassType;
        private System.Windows.Forms.TextBox txtGlassBiteLeft;
        private System.Windows.Forms.TextBox txtGlassBiteRight;
        private System.Windows.Forms.TextBox txtGlassBiteTop;
        private System.Windows.Forms.TextBox txtGlassBiteBottom;
        private System.Windows.Forms.TextBox txtFloor;
        private System.Windows.Forms.TextBox txtElevation;
        private System.Windows.Forms.Button btnCreateGlass;
        private System.Windows.Forms.Button btnCopyProperties;
        private System.Windows.Forms.Button btnSelectGlass;  // New button
        private System.Windows.Forms.Button btnUpdateGlass;  // New button
        private System.Windows.Forms.Label lblDimensions;    // New label to display dimensions
        private System.Windows.Forms.Label lblMarkNumber;    // New label to display mark number

        private GlassData _currentGlassData = new GlassData();
        private ObjectId _selectedGlassId = ObjectId.Null;  // Track the selected glass

        public GlassPanel()
        {
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            // Create all controls
            this.txtGlassType = new System.Windows.Forms.TextBox();
            this.txtGlassBiteLeft = new System.Windows.Forms.TextBox();
            this.txtGlassBiteRight = new System.Windows.Forms.TextBox();
            this.txtGlassBiteTop = new System.Windows.Forms.TextBox();
            this.txtGlassBiteBottom = new System.Windows.Forms.TextBox();
            this.txtFloor = new System.Windows.Forms.TextBox();
            this.txtElevation = new System.Windows.Forms.TextBox();
            this.btnCreateGlass = new System.Windows.Forms.Button();
            this.btnCopyProperties = new System.Windows.Forms.Button();
            this.btnSelectGlass = new System.Windows.Forms.Button();
            this.btnUpdateGlass = new System.Windows.Forms.Button();
            this.lblDimensions = new System.Windows.Forms.Label();
            this.lblMarkNumber = new System.Windows.Forms.Label();

            // Labels for text fields
            System.Windows.Forms.Label lblGlassType = new System.Windows.Forms.Label();
            System.Windows.Forms.Label lblGlassBiteLeft = new System.Windows.Forms.Label();
            System.Windows.Forms.Label lblGlassBiteRight = new System.Windows.Forms.Label();
            System.Windows.Forms.Label lblGlassBiteTop = new System.Windows.Forms.Label();
            System.Windows.Forms.Label lblGlassBiteBottom = new System.Windows.Forms.Label();
            System.Windows.Forms.Label lblFloor = new System.Windows.Forms.Label();
            System.Windows.Forms.Label lblElevation = new System.Windows.Forms.Label();

            // 
            // txtGlassType
            // 
            this.txtGlassType.Location = new System.Drawing.Point(120, 20);
            this.txtGlassType.Name = "txtGlassType";
            this.txtGlassType.Size = new System.Drawing.Size(120, 20);
            this.txtGlassType.TabIndex = 0;
            this.txtGlassType.Text = "1";

            // 
            // txtGlassBiteLeft
            // 
            this.txtGlassBiteLeft.Location = new System.Drawing.Point(120, 50);
            this.txtGlassBiteLeft.Name = "txtGlassBiteLeft";
            this.txtGlassBiteLeft.Size = new System.Drawing.Size(120, 20);
            this.txtGlassBiteLeft.TabIndex = 1;
            this.txtGlassBiteLeft.Text = "0.5";

            // 
            // txtGlassBiteRight
            // 
            this.txtGlassBiteRight.Location = new System.Drawing.Point(120, 80);
            this.txtGlassBiteRight.Name = "txtGlassBiteRight";
            this.txtGlassBiteRight.Size = new System.Drawing.Size(120, 20);
            this.txtGlassBiteRight.TabIndex = 2;
            this.txtGlassBiteRight.Text = "0.5";

            // 
            // txtGlassBiteTop
            // 
            this.txtGlassBiteTop.Location = new System.Drawing.Point(120, 110);
            this.txtGlassBiteTop.Name = "txtGlassBiteTop";
            this.txtGlassBiteTop.Size = new System.Drawing.Size(120, 20);
            this.txtGlassBiteTop.TabIndex = 3;
            this.txtGlassBiteTop.Text = "0.5";

            // 
            // txtGlassBiteBottom
            // 
            this.txtGlassBiteBottom.Location = new System.Drawing.Point(120, 140);
            this.txtGlassBiteBottom.Name = "txtGlassBiteBottom";
            this.txtGlassBiteBottom.Size = new System.Drawing.Size(120, 20);
            this.txtGlassBiteBottom.TabIndex = 4;
            this.txtGlassBiteBottom.Text = "0.5";

            // 
            // txtFloor
            // 
            this.txtFloor.Location = new System.Drawing.Point(120, 170);
            this.txtFloor.Name = "txtFloor";
            this.txtFloor.Size = new System.Drawing.Size(120, 20);
            this.txtFloor.TabIndex = 5;
            this.txtFloor.Text = "01";

            // 
            // txtElevation
            // 
            this.txtElevation.Location = new System.Drawing.Point(120, 200);
            this.txtElevation.Name = "txtElevation";
            this.txtElevation.Size = new System.Drawing.Size(120, 20);
            this.txtElevation.TabIndex = 6;
            this.txtElevation.Text = "A";

            // 
            // btnCreateGlass
            // 
            this.btnCreateGlass.Location = new System.Drawing.Point(40, 240);
            this.btnCreateGlass.Name = "btnCreateGlass";
            this.btnCreateGlass.Size = new System.Drawing.Size(180, 30);
            this.btnCreateGlass.TabIndex = 7;
            this.btnCreateGlass.Text = "Create Glass";
            this.btnCreateGlass.UseVisualStyleBackColor = true;
            this.btnCreateGlass.Click += new System.EventHandler(this.btnCreateGlass_Click);

            // 
            // btnCopyProperties
            // 
            this.btnCopyProperties.Location = new System.Drawing.Point(40, 280);
            this.btnCopyProperties.Name = "btnCopyProperties";
            this.btnCopyProperties.Size = new System.Drawing.Size(180, 30);
            this.btnCopyProperties.TabIndex = 8;
            this.btnCopyProperties.Text = "Copy Glass Properties";
            this.btnCopyProperties.UseVisualStyleBackColor = true;
            this.btnCopyProperties.Click += new System.EventHandler(this.btnCopyProperties_Click);

            // 
            // btnSelectGlass
            // 
            this.btnSelectGlass.Location = new System.Drawing.Point(40, 320);
            this.btnSelectGlass.Name = "btnSelectGlass";
            this.btnSelectGlass.Size = new System.Drawing.Size(180, 30);
            this.btnSelectGlass.TabIndex = 9;
            this.btnSelectGlass.Text = "Select Glass";
            this.btnSelectGlass.UseVisualStyleBackColor = true;
            this.btnSelectGlass.Click += new System.EventHandler(this.btnSelectGlass_Click);

            // 
            // btnUpdateGlass
            // 
            this.btnUpdateGlass.Location = new System.Drawing.Point(40, 360);
            this.btnUpdateGlass.Name = "btnUpdateGlass";
            this.btnUpdateGlass.Size = new System.Drawing.Size(180, 30);
            this.btnUpdateGlass.TabIndex = 10;
            this.btnUpdateGlass.Text = "Update Glass";
            this.btnUpdateGlass.UseVisualStyleBackColor = true;
            this.btnUpdateGlass.Enabled = false; // Disabled until glass is selected
            this.btnUpdateGlass.Click += new System.EventHandler(this.btnUpdateGlass_Click);

            // 
            // lblDimensions
            // 
            this.lblDimensions.AutoSize = true;
            this.lblDimensions.Location = new System.Drawing.Point(20, 400);
            this.lblDimensions.Name = "lblDimensions";
            this.lblDimensions.Size = new System.Drawing.Size(100, 13);
            this.lblDimensions.Text = "Dimensions: ";

            // 
            // lblMarkNumber
            // 
            this.lblMarkNumber.AutoSize = true;
            this.lblMarkNumber.Location = new System.Drawing.Point(20, 420);
            this.lblMarkNumber.Name = "lblMarkNumber";
            this.lblMarkNumber.Size = new System.Drawing.Size(100, 13);
            this.lblMarkNumber.Text = "Mark Number: ";

            // 
            // Labels
            // 
            lblGlassType.AutoSize = true;
            lblGlassType.Location = new System.Drawing.Point(20, 23);
            lblGlassType.Name = "lblGlassType";
            lblGlassType.Size = new System.Drawing.Size(62, 13);
            lblGlassType.Text = "Glass Type:";

            lblGlassBiteLeft.AutoSize = true;
            lblGlassBiteLeft.Location = new System.Drawing.Point(20, 53);
            lblGlassBiteLeft.Name = "lblGlassBiteLeft";
            lblGlassBiteLeft.Size = new System.Drawing.Size(81, 13);
            lblGlassBiteLeft.Text = "Left:";

            lblGlassBiteRight.AutoSize = true;
            lblGlassBiteRight.Location = new System.Drawing.Point(20, 83);
            lblGlassBiteRight.Name = "lblGlassBiteRight";
            lblGlassBiteRight.Size = new System.Drawing.Size(88, 13);
            lblGlassBiteRight.Text = "Right:";

            lblGlassBiteTop.AutoSize = true;
            lblGlassBiteTop.Location = new System.Drawing.Point(20, 113);
            lblGlassBiteTop.Name = "lblGlassBiteTop";
            lblGlassBiteTop.Size = new System.Drawing.Size(80, 13);
            lblGlassBiteTop.Text = "Top:";

            lblGlassBiteBottom.AutoSize = true;
            lblGlassBiteBottom.Location = new System.Drawing.Point(20, 143);
            lblGlassBiteBottom.Name = "lblGlassBiteBottom";
            lblGlassBiteBottom.Size = new System.Drawing.Size(97, 13);
            lblGlassBiteBottom.Text = "Bottom:";

            lblFloor.AutoSize = true;
            lblFloor.Location = new System.Drawing.Point(20, 173);
            lblFloor.Name = "lblFloor";
            lblFloor.Size = new System.Drawing.Size(33, 13);
            lblFloor.Text = "Floor:";

            lblElevation.AutoSize = true;
            lblElevation.Location = new System.Drawing.Point(20, 203);
            lblElevation.Name = "lblElevation";
            lblElevation.Size = new System.Drawing.Size(54, 13);
            lblElevation.Text = "Elevation:";

            // 
            // GlassPanel - Add all controls
            // 
            this.Controls.Add(lblGlassType);
            this.Controls.Add(this.txtGlassType);
            this.Controls.Add(lblGlassBiteLeft);
            this.Controls.Add(this.txtGlassBiteLeft);
            this.Controls.Add(lblGlassBiteRight);
            this.Controls.Add(this.txtGlassBiteRight);
            this.Controls.Add(lblGlassBiteTop);
            this.Controls.Add(this.txtGlassBiteTop);
            this.Controls.Add(lblGlassBiteBottom);
            this.Controls.Add(this.txtGlassBiteBottom);
            this.Controls.Add(lblFloor);
            this.Controls.Add(this.txtFloor);
            this.Controls.Add(lblElevation);
            this.Controls.Add(this.txtElevation);
            this.Controls.Add(this.btnCreateGlass);
            this.Controls.Add(this.btnCopyProperties);
            this.Controls.Add(this.btnSelectGlass);
            this.Controls.Add(this.btnUpdateGlass);
            this.Controls.Add(this.lblDimensions);
            this.Controls.Add(this.lblMarkNumber);

            // Set panel properties
            this.Name = "GlassPanel";
            this.Size = new System.Drawing.Size(260, 450);
        }

        private void btnCreateGlass_Click(object sender, EventArgs e)
        {
            try
            {
                // Get values from UI
                UpdateCurrentGlassData();

                // Call command to create glass
                Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                doc.SendStringToExecute("CREATEGLASS ", true, false, false);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Glass Takeoff", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnSelectGlass_Click(object sender, EventArgs e)
        {
            try
            {
                Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

                // Prompt the user to select a glass polyline
                PromptEntityOptions peo = new PromptEntityOptions("\nSelect glass polyline: ");
                peo.SetRejectMessage("\nOnly glass polylines can be selected.");
                peo.AddAllowedClass(typeof(Polyline), false);

                PromptEntityResult per = doc.Editor.GetEntity(peo);
                if (per.Status != PromptStatus.OK) return;

                // Load the properties
                _selectedGlassId = per.ObjectId;
                LoadGlassProperties(_selectedGlassId);

                // Enable the update button
                btnUpdateGlass.Enabled = true;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Glass Takeoff", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void btnUpdateGlass_Click(object sender, EventArgs e)
        {
            try
            {
                // Ensure we have a selected glass
                if (_selectedGlassId == ObjectId.Null)
                {
                    MessageBox.Show("No glass selected. Please select a glass first.",
                                    "Glass Takeoff", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                    return;
                }

                // Update current glass data from UI
                UpdateCurrentGlassData();

                // Use a command to update the glass instead of directly accessing the database
                // This avoids eLock violations when AutoCAD already has pending operations
                Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;

                // Store the data for the command to use
                GlassTakeoffCommands.PendingGlassData = _currentGlassData;
                GlassTakeoffCommands.PendingGlassId = _selectedGlassId;

                // Call the command to update the glass
                doc.SendStringToExecute("UPDATESELECTGLASS ", true, false, false);

                // The command will update the UI labels when it completes
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Glass Takeoff", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }



        private void LoadGlassProperties(ObjectId glassId)
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                Entity ent = tr.GetObject(glassId, OpenMode.ForRead) as Entity;
                if (ent is Polyline pline)
                {
                    // Check if it has GLASS Xdata
                    ResultBuffer rbXdata = ent.GetXDataForApplication("GLASS");
                    if (rbXdata != null)
                    {
                        // Extract glass data
                        _currentGlassData = GlassTakeoffCommands.ExtractGlassDataFromXData(rbXdata);

                        // Update UI
                        txtGlassType.Text = _currentGlassData.GlassType;
                        txtGlassBiteLeft.Text = _currentGlassData.GlassBiteLeft.ToString();
                        txtGlassBiteRight.Text = _currentGlassData.GlassBiteRight.ToString();
                        txtGlassBiteTop.Text = _currentGlassData.GlassBiteTop.ToString();
                        txtGlassBiteBottom.Text = _currentGlassData.GlassBiteBottom.ToString();
                        txtFloor.Text = _currentGlassData.Floor;
                        txtElevation.Text = _currentGlassData.Elevation;

                        // Update dimension and mark number labels
                        lblDimensions.Text = $"Dimensions: {_currentGlassData.GlassWidth} x {_currentGlassData.GlassHeight}";
                        lblMarkNumber.Text = $"Mark Number: {_currentGlassData.MarkNumber}";
                    }
                    else
                    {
                        MessageBox.Show("Selected polyline is not a glass element or has no glass data.",
                                       "Glass Takeoff", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                        _selectedGlassId = ObjectId.Null;
                        btnUpdateGlass.Enabled = false;
                    }
                }

                tr.Commit();
            }
        }

        public void HandleExternalSelection(ObjectId glassId)
        {
            _selectedGlassId = glassId;
            LoadGlassProperties(glassId);
            btnUpdateGlass.Enabled = true;
        }

        public void UpdateDimensionDisplay(GlassData glassData)
        {
            if (glassData != null)
            {
                lblDimensions.Text = $"Dimensions: {glassData.GlassWidth} x {glassData.GlassHeight}";
                lblMarkNumber.Text = $"Mark Number: {glassData.MarkNumber}";
            }
        }

        private void btnCopyProperties_Click(object sender, EventArgs e)
        {
            try
            {
                // Get values from UI
                UpdateCurrentGlassData();

                // Call command to copy properties
                Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                doc.SendStringToExecute("COPYGLASSPROPERTIES ", true, false, false);
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Glass Takeoff", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
        }

        private void UpdateCurrentGlassData()
        {
            // Parse values from UI
            _currentGlassData.GlassType = txtGlassType.Text;
            _currentGlassData.GlassBiteLeft = double.Parse(txtGlassBiteLeft.Text);
            _currentGlassData.GlassBiteRight = double.Parse(txtGlassBiteRight.Text);
            _currentGlassData.GlassBiteTop = double.Parse(txtGlassBiteTop.Text);
            _currentGlassData.GlassBiteBottom = double.Parse(txtGlassBiteBottom.Text);
            _currentGlassData.Floor = txtFloor.Text;
            _currentGlassData.Elevation = txtElevation.Text;

            // Set as pending data for command
            GlassTakeoffCommands.PendingGlassData = _currentGlassData;
        }

        public void UpdateUIFromGlassData(GlassData glassData)
        {
            if (glassData == null) return;

            txtGlassType.Text = glassData.GlassType;
            txtGlassBiteLeft.Text = glassData.GlassBiteLeft.ToString();
            txtGlassBiteRight.Text = glassData.GlassBiteRight.ToString();
            txtGlassBiteTop.Text = glassData.GlassBiteTop.ToString();
            txtGlassBiteBottom.Text = glassData.GlassBiteBottom.ToString();
            txtFloor.Text = glassData.Floor;
            txtElevation.Text = glassData.Elevation;

            _currentGlassData = glassData;
        }
    }

    // Commands for Glass Takeoff
    public class GlassTakeoffCommands
    {
        // Static reference to glass panel
        private static GlassPanel _glassPanel;

        // Pending data for command execution
        public static GlassData PendingGlassData { get; set; } = new GlassData();

        // Static property to store the selected glass ID
        public static ObjectId PendingGlassId { get; set; } = ObjectId.Null;

        // Class for copying glass properties
        public class GlassCopyData
        {
            public string SourceHandle { get; set; }
            public List<string> TargetHandles { get; set; }
            public GlassData GlassData { get; set; }
        }

        public static GlassCopyData PendingCopyData { get; set; }

        [CommandMethod("ADDGLASSPANEL")]
        public void AddGlassPanel()
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            try
            {
                ed.WriteMessage("\nAdding Glass panel to existing palette set...");

                // Get reference to the existing palette set
                System.Reflection.FieldInfo fieldInfo = typeof(TakeoffBridge.MetalComponentCommands).GetField("_paletteSet",
                    System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);

                if (fieldInfo != null)
                {
                    Autodesk.AutoCAD.Windows.PaletteSet paletteSet = fieldInfo.GetValue(null) as Autodesk.AutoCAD.Windows.PaletteSet;

                    if (paletteSet != null)
                    {
                        // Create glass panel if it doesn't exist
                        if (_glassPanel == null)
                        {
                            _glassPanel = new GlassPanel();
                        }

                        // Try to add the panel - if it's already added, this might throw an exception
                        try
                        {
                            paletteSet.Add("Glass", _glassPanel);
                            ed.WriteMessage("\nGlass panel added successfully.");
                        }
                        catch (System.Exception)
                        {
                            // If we got an exception, the panel might already exist
                            ed.WriteMessage("\nGlass panel already exists or couldn't be added.");
                        }

                        // Show the palette set
                        paletteSet.Visible = true;
                    }
                    else
                    {
                        ed.WriteMessage("\nError: Could not access palette set.");
                    }
                }
                else
                {
                    ed.WriteMessage("\nError: Could not find palette set field.");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError: {ex.Message}");
                ed.WriteMessage($"\nStack trace: {ex.StackTrace}");
            }
        }

        [CommandMethod("CREATEGLASS")]
        public void CreateGlass()
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            Editor ed = doc.Editor;

            try
            {
                ed.WriteMessage("\nCreating glass boundary...");

                // Ensure we have pending data
                if (PendingGlassData == null)
                {
                    PendingGlassData = new GlassData();
                }

                // Get glass information from pending data
                GlassData glassData = PendingGlassData;

                // Ensure Tag1 layer exists
                EnsureLayerExists("Tag1");

                // Ensure Tag2 layer exists for mark numbers
                EnsureLayerExists("Tag2");

                // IMPORTANT: Explicitly clear any running command or selection state
                doc.SendStringToExecute("'_.CANCEL ", true, false, true);

                // Wait a moment to ensure the command is canceled
                System.Threading.Thread.Sleep(100);

                // Get the center point using proper prompt
                PromptPointOptions ppo = new PromptPointOptions("\nSelect center of glass lite: ");
                ppo.AllowNone = false;

                Point3d centerPoint;
                PromptPointResult ppr = ed.GetPoint(ppo);
                if (ppr.Status != PromptStatus.OK)
                {
                    // User canceled - abort the command
                    ed.WriteMessage("\nCommand canceled.");
                    return;
                }

                centerPoint = ppr.Value;
                ed.WriteMessage($"\nSelected center point: ({centerPoint.X}, {centerPoint.Y})");

                // Use boundary detection to find the perimeter
                List<Point3d> boundaryPoints = DetectBoundary(centerPoint);
                if (boundaryPoints.Count < 3)
                {
                    ed.WriteMessage("\nCould not detect glass boundary. Try again.");
                    return;
                }

                // Create glass polygon from boundary points
                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    // Get the Tag1 layer
                    LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                    ObjectId tagLayerId = lt["Tag1"];

                    // Create polyline
                    Polyline pline = new Polyline();
                    pline.SetDatabaseDefaults();
                    pline.Closed = true;
                    pline.ConstantWidth = 0.1;

                    // Add vertices from boundary points
                    for (int i = 0; i < boundaryPoints.Count; i++)
                    {
                        pline.AddVertexAt(i, new Point2d(boundaryPoints[i].X, boundaryPoints[i].Y), 0, 0, 0);
                    }

                    // Set the layer to Tag1
                    pline.LayerId = tagLayerId;

                    // Add to model space
                    BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
                    BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

                    ObjectId plineId = btr.AppendEntity(pline);
                    tr.AddNewlyCreatedDBObject(pline, true);

                    // Calculate glass dimensions
                    CalculateGlassDimensions(pline, glassData);

                    // Assign mark number
                    glassData.MarkNumber = GlassMarkNumberManager.Instance.GetOrCreateMarkNumber(glassData);

                    // Register applications for XData
                    RegisterApp(tr, "GLASS");

                    // Add glass data as XData
                    AddGlassXData(pline, glassData);

                    // Create mark number text
                    CreateGlassMarkText(tr, glassData.MarkNumber, pline);

                    tr.Commit();

                    ed.WriteMessage($"\nGlass created successfully with mark number: {glassData.MarkNumber}");
                    ed.WriteMessage($"\nDimensions: {glassData.GlassWidth} x {glassData.GlassHeight} (with bites)");
                    ed.WriteMessage($"\nDaylight Opening: {glassData.DloWidth} x {glassData.DloHeight}");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError: {ex.Message}");
                ed.WriteMessage($"\nStack trace: {ex.StackTrace}");
            }
        }

        [CommandMethod("COPYGLASSPROPERTIES")]
        public void CopyGlassProperties()
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            try
            {
                // Prompt for source glass
                PromptEntityOptions peoSource = new PromptEntityOptions("\nSelect source glass: ");
                peoSource.SetRejectMessage("\nOnly glass polylines can be selected.");
                peoSource.AddAllowedClass(typeof(Polyline), false);

                PromptEntityResult perSource = ed.GetEntity(peoSource);
                if (perSource.Status != PromptStatus.OK) return;

                ObjectId sourceId = perSource.ObjectId;

                // Get source glass data
                GlassData sourceGlassData = null;
                string sourceHandle = string.Empty;

                using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
                {
                    Entity ent = tr.GetObject(sourceId, OpenMode.ForRead) as Entity;
                    if (ent is Polyline pline)
                    {
                        // Check if it has GLASS Xdata
                        ResultBuffer rbXdata = ent.GetXDataForApplication("GLASS");
                        if (rbXdata != null)
                        {
                            sourceGlassData = ExtractGlassDataFromXData(rbXdata);
                            sourceHandle = ent.Handle.ToString();
                        }
                    }

                    tr.Commit();
                }

                if (sourceGlassData == null)
                {
                    ed.WriteMessage("\nSelected object is not a glass polyline or has no glass data.");
                    return;
                }

                // Update UI if panel exists
                if (_glassPanel != null)
                {
                    _glassPanel.UpdateUIFromGlassData(sourceGlassData);
                }

                // Now prompt for target glass polylines
                PromptSelectionOptions pso = new PromptSelectionOptions();
                pso.MessageForAdding = "\nSelect target glass polylines to receive properties: ";

                PromptSelectionResult psr = ed.GetSelection(pso);
                if (psr.Status != PromptStatus.OK) return;

                SelectionSet ss = psr.Value;

                // Store target handles
                List<string> targetHandles = new List<string>();

                using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
                {
                    foreach (SelectedObject so in ss)
                    {
                        if (so.ObjectId == sourceId) continue; // Skip source

                        Entity ent = tr.GetObject(so.ObjectId, OpenMode.ForRead) as Entity;
                        if (ent is Polyline pline)
                        {
                            targetHandles.Add(ent.Handle.ToString());
                        }
                    }

                    tr.Commit();
                }

                if (targetHandles.Count == 0)
                {
                    ed.WriteMessage("\nNo valid target glass polylines selected.");
                    return;
                }

                // Create the copy data
                PendingCopyData = new GlassCopyData
                {
                    SourceHandle = sourceHandle,
                    TargetHandles = targetHandles,
                    GlassData = sourceGlassData
                };

                // Execute the copy operation
                doc.SendStringToExecute("EXECUTEGLASSCOPY ", true, false, false);
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in COPYGLASSPROPERTIES: {ex.Message}");
            }
        }

        [CommandMethod("EXECUTEGLASSCOPY")]
        public void ExecuteGlassCopy()
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            try
            {
                // Check if we have pending data
                if (PendingCopyData == null)
                {
                    ed.WriteMessage("\nNo pending glass copy operation.");
                    return;
                }

                // Process each target
                foreach (string targetHandle in PendingCopyData.TargetHandles)
                {
                    // Get the object ID from handle
                    long longHandle = Convert.ToInt64(targetHandle, 16);
                    Handle h = new Handle(longHandle);
                    ObjectId objId = doc.Database.GetObjectId(false, h, 0);

                    using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
                    {
                        Entity ent = tr.GetObject(objId, OpenMode.ForWrite) as Entity;
                        if (ent is Polyline pline)
                        {
                            // Calculate dimensions for this specific polyline
                            GlassData glassData = new GlassData(); // Create a new instance
                                                                   // Copy all properties from the source
                            glassData.GlassType = PendingCopyData.GlassData.GlassType;
                            glassData.GlassBiteLeft = PendingCopyData.GlassData.GlassBiteLeft;
                            glassData.GlassBiteRight = PendingCopyData.GlassData.GlassBiteRight;
                            glassData.GlassBiteTop = PendingCopyData.GlassData.GlassBiteTop;
                            glassData.GlassBiteBottom = PendingCopyData.GlassData.GlassBiteBottom;
                            glassData.Floor = PendingCopyData.GlassData.Floor;
                            glassData.Elevation = PendingCopyData.GlassData.Elevation;
                            // We don't copy dimensions or mark number as they will be recalculated

                            CalculateGlassDimensions(pline, glassData);

                            // Assign mark number based on dimensions
                            glassData.MarkNumber = GlassMarkNumberManager.Instance.GetOrCreateMarkNumber(glassData);

                            // Register application for XData
                            RegisterApp(tr, "GLASS");

                            // Add glass data as XData
                            AddGlassXData(pline, glassData);

                            // Update mark text or create if it doesn't exist
                            UpdateGlassMarkText(tr, glassData.MarkNumber, pline);
                        }

                        tr.Commit();
                    }
                }

                ed.WriteMessage($"\nGlass properties copied to {PendingCopyData.TargetHandles.Count} glass polylines.");

                // Clear pending data
                PendingCopyData = null;
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in EXECUTEGLASSCOPY: {ex.Message}");
            }
        }

        [CommandMethod("UPDATEGLASSTAKEOFF")]
        public void UpdateGlassTakeoff()
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            try
            {
                ed.WriteMessage("\nUpdating glass mark numbers...");

                // Get all glass polylines
                TypedValue[] tvs = new TypedValue[] {
                    new TypedValue((int)DxfCode.Start, "LWPOLYLINE"),
                    new TypedValue((int)DxfCode.LayerName, "Tag1")
                };

                SelectionFilter filter = new SelectionFilter(tvs);
                PromptSelectionResult selRes = ed.SelectAll(filter);

                if (selRes.Status != PromptStatus.OK)
                {
                    ed.WriteMessage("\nNo glass polylines found.");
                    return;
                }

                SelectionSet ss = selRes.Value;
                int updateCount = 0;

                // Reset mark number manager to start fresh
                System.Reflection.FieldInfo fieldInfo = typeof(GlassMarkNumberManager).GetField("_instance",
                    System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);

                if (fieldInfo != null)
                {
                    fieldInfo.SetValue(null, null);
                }

                // Process each glass polyline
                using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
                {
                    foreach (SelectedObject selObj in ss)
                    {
                        Entity ent = tr.GetObject(selObj.ObjectId, OpenMode.ForRead) as Entity;
                        if (ent is Polyline pline)
                        {
                            // Check if it has GLASS Xdata
                            ResultBuffer rbXdata = ent.GetXDataForApplication("GLASS");
                            if (rbXdata != null)
                            {
                                // Get entity for write
                                ent.UpgradeOpen();

                                // Extract glass data
                                GlassData glassData = ExtractGlassDataFromXData(rbXdata);

                                // Recalculate dimensions
                                CalculateGlassDimensions(pline, glassData);

                                // Reassign mark number
                                glassData.MarkNumber = GlassMarkNumberManager.Instance.GetOrCreateMarkNumber(glassData);

                                // Update XData
                                AddGlassXData(ent as Polyline, glassData);

                                // Update mark text
                                UpdateGlassMarkText(tr, glassData.MarkNumber, ent as Polyline);

                                updateCount++;
                            }
                        }
                    }

                    tr.Commit();
                }

                ed.WriteMessage($"\nUpdated {updateCount} glass mark numbers.");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in UPDATEGLASSTAKEOFF: {ex.Message}");
            }
        }

        [CommandMethod("EDITGLASS")]
        public void EditGlass()
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            try
            {
                // First, ensure the glass panel is visible
                ShowGlassPanel();

                // Prompt for glass selection
                PromptEntityOptions peo = new PromptEntityOptions("\nSelect glass to edit: ");
                peo.SetRejectMessage("\nOnly glass polylines can be selected.");
                peo.AddAllowedClass(typeof(Polyline), false);

                PromptEntityResult per = ed.GetEntity(peo);
                if (per.Status != PromptStatus.OK) return;

                // Forward the selection to the panel
                if (_glassPanel != null)
                {
                    _glassPanel.HandleExternalSelection(per.ObjectId);
                }
                else
                {
                    ed.WriteMessage("\nGlass panel is not available. Run ADDGLASSPANEL first.");
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError: {ex.Message}");
            }
        }

        [CommandMethod("UPDATESELECTGLASS")]
        public void UpdateSelectedGlass()
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            try
            {
                // Check if we have necessary data
                if (PendingGlassId == ObjectId.Null || PendingGlassData == null)
                {
                    ed.WriteMessage("\nNo pending glass update data.");
                    return;
                }

                using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
                {
                    // Get the entity
                    Entity ent = tr.GetObject(PendingGlassId, OpenMode.ForWrite) as Entity;
                    if (ent is Polyline pline)
                    {
                        // Calculate dimensions
                        CalculateGlassDimensions(pline, PendingGlassData);

                        // Assign mark number
                        PendingGlassData.MarkNumber = GlassMarkNumberManager.Instance.GetOrCreateMarkNumber(PendingGlassData);

                        // Register app name for XData
                        RegisterApp(tr, "GLASS");

                        // Add glass data as XData
                        AddGlassXData(pline, PendingGlassData);

                        // Update mark text
                        UpdateGlassMarkText(tr, PendingGlassData.MarkNumber, pline);

                        ed.WriteMessage($"\nGlass updated with mark number: {PendingGlassData.MarkNumber}");

                        // Update UI if glass panel exists
                        if (_glassPanel != null)
                        {
                            // Need to update UI from main thread
                            System.Windows.Forms.Application.DoEvents();
                            _glassPanel.UpdateDimensionDisplay(PendingGlassData);
                        }
                    }

                    tr.Commit();
                }

                // Reset pending data
                PendingGlassId = ObjectId.Null;
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError updating glass: {ex.Message}");
            }
        }



        // Add this helper method to show the glass panel
        private void ShowGlassPanel()
        {
            // Get reference to the existing palette set
            System.Reflection.FieldInfo fieldInfo = typeof(TakeoffBridge.MetalComponentCommands).GetField("_paletteSet",
                System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static);

            if (fieldInfo != null)
            {
                Autodesk.AutoCAD.Windows.PaletteSet paletteSet = fieldInfo.GetValue(null) as Autodesk.AutoCAD.Windows.PaletteSet;

                if (paletteSet != null)
                {
                    // Make sure the panel exists
                    if (_glassPanel == null)
                    {
                        AddGlassPanel();
                    }

                    // Show the palette set
                    paletteSet.Visible = true;

                    // Activate the Glass tab if possible
                    try
                    {
                        for (int i = 0; i < paletteSet.Count; i++)
                        {
                            try
                            {
                                if (Object.ReferenceEquals(paletteSet[i], _glassPanel))
                                {
                                    paletteSet.Activate(i);
                                    break;
                                }
                            }
                            catch
                            {
                                // Skip this index if there's an error
                                continue;
                            }
                        }
                    }
                    catch
                    {
                        // If we can't activate the tab, at least the palette set is visible
                    }
                }
            }
        }

        // Helper method to extract glass data from XData
        public static GlassData ExtractGlassDataFromXData(ResultBuffer rbXdata)
        {
            GlassData glassData = new GlassData();

            TypedValue[] xdata = rbXdata.AsArray();

            // Skip first item (application name)
            for (int i = 1; i < xdata.Length; i++)
            {
                // Process values based on the expected structure
                try
                {
                    if (i == 1) glassData.GlassType = xdata[i].Value.ToString();
                    if (i == 2) glassData.Floor = xdata[i].Value.ToString();
                    if (i == 3) glassData.Elevation = xdata[i].Value.ToString();
                    if (i == 4) glassData.GlassBiteLeft = Convert.ToDouble(xdata[i].Value);
                    if (i == 5) glassData.GlassBiteBottom = Convert.ToDouble(xdata[i].Value);
                    if (i == 6) glassData.GlassBiteRight = Convert.ToDouble(xdata[i].Value);
                    if (i == 7) glassData.GlassBiteTop = Convert.ToDouble(xdata[i].Value);
                    if (i == 8) glassData.GlassWidth = Convert.ToDouble(xdata[i].Value);
                    if (i == 9) glassData.GlassHeight = Convert.ToDouble(xdata[i].Value);
                    if (i == 10) glassData.DloWidth = Convert.ToDouble(xdata[i].Value);
                    if (i == 11) glassData.DloHeight = Convert.ToDouble(xdata[i].Value);
                    if (i == 12) glassData.MarkNumber = xdata[i].Value.ToString();
                }
                catch (System.Exception ex)
                {
                    // Handle parsing errors silently
                    System.Diagnostics.Debug.WriteLine($"Error parsing Xdata[{i}]: {ex.Message}");
                }
            }

            return glassData;
        }

        // Helper method to add glass data to entity as XData
        public static void AddGlassXData(Polyline pline, GlassData glassData)
        {
            // Create XData with glass information
            ResultBuffer rb = new ResultBuffer(
                new TypedValue((int)DxfCode.ExtendedDataRegAppName, "GLASS"),
                new TypedValue((int)DxfCode.ExtendedDataAsciiString, glassData.GlassType),
                new TypedValue((int)DxfCode.ExtendedDataAsciiString, glassData.Floor),
                new TypedValue((int)DxfCode.ExtendedDataAsciiString, glassData.Elevation),
                new TypedValue((int)DxfCode.ExtendedDataReal, glassData.GlassBiteLeft),
                new TypedValue((int)DxfCode.ExtendedDataReal, glassData.GlassBiteBottom),
                new TypedValue((int)DxfCode.ExtendedDataReal, glassData.GlassBiteRight),
                new TypedValue((int)DxfCode.ExtendedDataReal, glassData.GlassBiteTop),
                new TypedValue((int)DxfCode.ExtendedDataReal, glassData.GlassWidth),
                new TypedValue((int)DxfCode.ExtendedDataReal, glassData.GlassHeight),
                new TypedValue((int)DxfCode.ExtendedDataReal, glassData.DloWidth),
                new TypedValue((int)DxfCode.ExtendedDataReal, glassData.DloHeight),
                new TypedValue((int)DxfCode.ExtendedDataAsciiString, glassData.MarkNumber)
            );

            pline.XData = rb;
        }

        // Helper method to calculate glass dimensions
        public static void CalculateGlassDimensions(Polyline pline, GlassData glassData)
        {
            // Get bounding box of the polyline
            Point3dCollection vertices = new Point3dCollection();
            for (int i = 0; i < pline.NumberOfVertices; i++)
            {
                vertices.Add(pline.GetPoint3dAt(i));
            }

            Point3d minPoint = new Point3d(double.MaxValue, double.MaxValue, 0);
            Point3d maxPoint = new Point3d(double.MinValue, double.MinValue, 0);

            foreach (Point3d pt in vertices)
            {
                minPoint = new Point3d(Math.Min(minPoint.X, pt.X), Math.Min(minPoint.Y, pt.Y), 0);
                maxPoint = new Point3d(Math.Max(maxPoint.X, pt.X), Math.Max(maxPoint.Y, pt.Y), 0);
            }

            // Calculate daylight opening dimensions
            glassData.DloWidth = Math.Round(maxPoint.X - minPoint.X, 3);
            glassData.DloHeight = Math.Round(maxPoint.Y - minPoint.Y, 3);

            // Calculate glass dimensions including bites
            glassData.GlassWidth = Math.Round(glassData.DloWidth + glassData.GlassBiteLeft + glassData.GlassBiteRight, 3);
            glassData.GlassHeight = Math.Round(glassData.DloHeight + glassData.GlassBiteTop + glassData.GlassBiteBottom, 3);
        }

        // Helper method to create mark number text
        private static void CreateGlassMarkText(Transaction tr, string markNumber, Polyline pline)
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            // Get the centroid of the polyline
            Point3d centroid = GetPolygonCentroid(pline);

            // Get the Tag2 layer for mark text
            LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
            ObjectId tagLayerId = lt["Tag2"];

            // Create the MText object
            MText mtext = new MText();
            mtext.Contents = markNumber;
            mtext.Location = centroid;
            mtext.TextHeight = 4.0;
            mtext.Attachment = AttachmentPoint.MiddleCenter;
            mtext.LayerId = tagLayerId;

            // Add to model space
            BlockTable bt = (BlockTable)tr.GetObject(db.BlockTableId, OpenMode.ForRead);
            BlockTableRecord btr = (BlockTableRecord)tr.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

            ObjectId textId = btr.AppendEntity(mtext);
            tr.AddNewlyCreatedDBObject(mtext, true);

            // Register app name for XData
            RegisterApp(tr, "GLASSMARKTEXT");

            // Add custom data to link this text to the glass
            ResultBuffer rb = new ResultBuffer(
                new TypedValue((int)DxfCode.ExtendedDataRegAppName, "GLASSMARKTEXT"),
                new TypedValue((int)DxfCode.ExtendedDataHandle, pline.Handle)
            );
            mtext.XData = rb;
        }

        // Helper method to update glass mark text
        private static void UpdateGlassMarkText(Transaction tr, string markNumber, Polyline pline)
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            // Find existing mark text for this polyline
            TypedValue[] tvs = new TypedValue[] {
        new TypedValue((int)DxfCode.Start, "MTEXT"),
        new TypedValue((int)DxfCode.ExtendedDataRegAppName, "GLASSMARKTEXT")
    };

            SelectionFilter filter = new SelectionFilter(tvs);
            PromptSelectionResult selRes = doc.Editor.SelectAll(filter);

            bool foundExisting = false;

            if (selRes.Status == PromptStatus.OK)
            {
                foreach (SelectedObject selObj in selRes.Value)
                {
                    MText mtext = tr.GetObject(selObj.ObjectId, OpenMode.ForRead) as MText;
                    if (mtext != null)
                    {
                        // Check if this text is linked to our polyline
                        ResultBuffer rb = mtext.GetXDataForApplication("GLASSMARKTEXT");
                        if (rb != null)
                        {
                            TypedValue[] tvArray = rb.AsArray();
                            if (tvArray.Length >= 2 && tvArray[1].TypeCode == (int)DxfCode.ExtendedDataHandle)
                            {
                                try
                                {
                                    // Get the handle value - corrected to handle different types properly
                                    string handleString = tvArray[1].Value.ToString();
                                    Handle storedHandle = new Handle(Convert.ToInt64(handleString, 16));

                                    // Compare the handles
                                    if (storedHandle.Value == pline.Handle.Value)
                                    {
                                        // Found our text, update it
                                        mtext.UpgradeOpen();
                                        mtext.Contents = markNumber;
                                        foundExisting = true;
                                        break;
                                    }
                                }
                                catch (System.Exception ex)
                                {
                                    // Log error but continue processing
                                    doc.Editor.WriteMessage($"\nError converting handle: {ex.Message}");
                                }
                            }
                        }
                    }
                }
            }

            // If no existing text found, create a new one
            if (!foundExisting)
            {
                CreateGlassMarkText(tr, markNumber, pline);
            }
        }

        // Helper method to get polygon centroid
        private static Point3d GetPolygonCentroid(Polyline pline)
        {
            double area = 0.0;
            double cx = 0.0;
            double cy = 0.0;

            for (int i = 0; i < pline.NumberOfVertices; i++)
            {
                Point2d p1 = pline.GetPoint2dAt(i);
                Point2d p2 = pline.GetPoint2dAt((i + 1) % pline.NumberOfVertices);

                double a = p1.X * p2.Y - p2.X * p1.Y;
                area += a;
                cx += (p1.X + p2.X) * a;
                cy += (p1.Y + p2.Y) * a;
            }

            area *= 0.5;
            cx /= (6.0 * area);
            cy /= (6.0 * area);

            return new Point3d(cx, cy, 0);
        }

        // Helper method to detect boundary from center point
        private static List<Point3d> DetectBoundary(Point3d centerPoint)
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            List<Point3d> boundaryPoints = new List<Point3d>();

            try
            {
                // Save current layer
                ObjectId originalLayerId = db.Clayer;
                string tempLayer = "TEMP_BOUNDARY";

                using (Transaction tr = db.TransactionManager.StartTransaction())
                {
                    // Ensure temp layer exists
                    EnsureLayerExists(tempLayer);

                    // Get the layer ID for the temp layer
                    LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);
                    if (lt.Has(tempLayer))
                    {
                        // Set active layer to the temp layer
                        db.Clayer = lt[tempLayer];

                        // Cancel any active command
                        doc.SendStringToExecute("'_.CANCEL ", true, false, true);

                        // Create boundary
                        ed.WriteMessage("\nCreating boundary...");

                        ed.Command(
                            "_boundary",
                            centerPoint,
                            "_a",     // All visible layers
                            "i",      // Island options
                            "n",      // No island deteciton
                            "",      // +X for ray casting
                            "",       // Accept default
                            ""        // End command
                        );

                        // Find the most recently created entity
                        TypedValue[] filterList = new TypedValue[]
                        {
                            new TypedValue((int)DxfCode.LayerName, tempLayer)
                        };

                        SelectionFilter filter = new SelectionFilter(filterList);
                        PromptSelectionResult selRes = ed.SelectAll(filter);

                        if (selRes.Status == PromptStatus.OK && selRes.Value.Count > 0)
                        {
                            // Get the most recently created entity
                            ObjectId boundaryId = selRes.Value[selRes.Value.Count - 1].ObjectId;

                            // Extract vertices from the boundary
                            Polyline boundary = (Polyline)tr.GetObject(boundaryId, OpenMode.ForRead);

                            for (int i = 0; i < boundary.NumberOfVertices; i++)
                            {
                                boundaryPoints.Add(boundary.GetPoint3dAt(i));
                            }

                            // Erase the temporary boundary
                            boundary.UpgradeOpen();
                            boundary.Erase();
                        }
                        else
                        {
                            ed.WriteMessage("\nNo boundary created. Try a different point.");
                        }

                        // Restore original layer
                        db.Clayer = originalLayerId;
                    }

                    tr.Commit();
                }
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError detecting boundary: {ex.Message}");
            }

            return boundaryPoints;
        }

        // Helper method to ensure a layer exists
        private static void EnsureLayerExists(string layerName)
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                LayerTable lt = (LayerTable)tr.GetObject(db.LayerTableId, OpenMode.ForRead);

                if (!lt.Has(layerName))
                {
                    lt.UpgradeOpen();
                    LayerTableRecord ltr = new LayerTableRecord();
                    ltr.Name = layerName;

                    // Set color based on layer name
                    if (layerName == "Tag1")
                    {
                        ltr.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByLayer, 1); // Red
                    }
                    else if (layerName == "Tag2")
                    {
                        ltr.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByLayer, 3); // Green
                    }
                    else
                    {
                        ltr.Color = Autodesk.AutoCAD.Colors.Color.FromColorIndex(Autodesk.AutoCAD.Colors.ColorMethod.ByLayer, 7); // White
                    }

                    lt.Add(ltr);
                    tr.AddNewlyCreatedDBObject(ltr, true);
                }

                tr.Commit();
            }
        }

        // Helper method to register an application name
        public static void RegisterApp(Transaction tr, string appName)
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            RegAppTable regTable = (RegAppTable)tr.GetObject(db.RegAppTableId, OpenMode.ForRead);

            if (!regTable.Has(appName))
            {
                regTable.UpgradeOpen();
                RegAppTableRecord record = new RegAppTableRecord();
                record.Name = appName;
                regTable.Add(record);
                tr.AddNewlyCreatedDBObject(record, true);
            }
        }
    }
}