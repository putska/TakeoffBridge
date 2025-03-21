using System;
using System.Windows.Forms;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Windows;
using Autodesk.AutoCAD.Geometry;
using System.Collections.Generic;
using acApp = Autodesk.AutoCAD.ApplicationServices.Application;
using System.IO;


namespace TakeoffBridge
{
    public class MetalComponentPanel : System.Windows.Forms.UserControl, IDisposable
    {

        private ListView partsList;
        private Button btnAddPart;
        private Button btnEditPart;
        private Button btnDeletePart;
        private TextBox txtComponentType;
        private TextBox txtComponentFloor;
        private Label lblLength;
        private System.Windows.Forms.Timer selectionTimer;
        private ObjectId currentComponentId;
        private List<ChildPart> currentParts = new List<ChildPart>();


        public MetalComponentPanel()
        {
            InitializeComponent();

            // Set up a timer to check selection changes
            selectionTimer = new System.Windows.Forms.Timer();
            selectionTimer.Interval = 500; // Check every half second
            selectionTimer.Tick += SelectionTimer_Tick;
            selectionTimer.Start();
        }

        private ObjectIdCollection lastSelection = new ObjectIdCollection();

        private void SelectionTimer_Tick(object sender, EventArgs e)
        {
            try
            {
                Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
                if (doc == null) return;

                // Get the current selection
                Editor ed = doc.Editor;
                PromptSelectionResult selRes = ed.SelectImplied();

                // Check if selection has changed
                bool selectionChanged = false;

                if (selRes.Status == PromptStatus.OK)
                {
                    // Compare with last selection
                    if (selRes.Value.Count != lastSelection.Count)
                    {
                        selectionChanged = true;
                    }
                    else
                    {
                        for (int i = 0; i < selRes.Value.Count; i++)
                        {
                            if (!lastSelection.Contains(selRes.Value[i].ObjectId))
                            {
                                selectionChanged = true;
                                break;
                            }
                        }
                    }

                    // Update last selection
                    lastSelection.Clear();
                    foreach (SelectedObject selObj in selRes.Value)
                    {
                        lastSelection.Add(selObj.ObjectId);
                    }
                }
                else
                {
                    // No selection
                    if (lastSelection.Count > 0)
                    {
                        selectionChanged = true;
                        lastSelection.Clear();
                    }
                }

                // If selection changed, update the panel
                if (selectionChanged)
                {
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
            }
            catch (System.Exception ex)
            {
                // Log the error but don't crash
                System.Diagnostics.Debug.WriteLine($"Error in selection timer: {ex.Message}");
            }

        }

        private void InitializeComponent()
        {
            this.Text = "Metal Component Editor";
            this.Width = 300;
            this.Height = 600;

            // Component info section
            Label lblType = new Label
            {
                Text = "Component Type:",
                Left = 10,
                Top = 10,
                Width = 100
            };
            this.Controls.Add(lblType);

            txtComponentType = new TextBox
            {
                Left = 120,
                Top = 10,
                Width = 150,
                ReadOnly = true
            };
            this.Controls.Add(txtComponentType);

            Label lblFloor = new Label
            {
                Text = "Floor:",
                Left = 10,
                Top = 40,
                Width = 100
            };
            this.Controls.Add(lblFloor);

            txtComponentFloor = new TextBox
            {
                Left = 120,
                Top = 40,
                Width = 150,
                ReadOnly = true
            };
            this.Controls.Add(txtComponentFloor);

            Label lblLengthLabel = new Label
            {
                Text = "Length:",
                Left = 10,
                Top = 70,
                Width = 100
            };
            this.Controls.Add(lblLengthLabel);

            lblLength = new Label
            {
                Text = "0.0",
                Left = 120,
                Top = 70,
                Width = 150
            };
            this.Controls.Add(lblLength);

            // Parts list
            Label lblParts = new Label
            {
                Text = "Child Parts:",
                Left = 10,
                Top = 100,
                Width = 100
            };
            this.Controls.Add(lblParts);

            partsList = new ListView
            {
                Left = 10,
                Top = 130,
                Width = 270,
                Height = 350,
                View = View.Details,
                FullRowSelect = true
            };
            partsList.Columns.Add("Name", 120);
            partsList.Columns.Add("Type", 60);
            partsList.Columns.Add("Length", 80);
            this.Controls.Add(partsList);

            // Buttons
            btnAddPart = new Button
            {
                Text = "Add Part",
                Left = 10,
                Top = 490,
                Width = 80
            };
            btnAddPart.Click += BtnAddPart_Click;
            this.Controls.Add(btnAddPart);

            btnEditPart = new Button
            {
                Text = "Edit Part",
                Left = 100,
                Top = 490,
                Width = 80
            };
            btnEditPart.Click += BtnEditPart_Click;
            this.Controls.Add(btnEditPart);

            btnDeletePart = new Button
            {
                Text = "Delete Part",
                Left = 190,
                Top = 490,
                Width = 80
            };
            btnDeletePart.Click += BtnDeletePart_Click;
            this.Controls.Add(btnDeletePart);

            Button btnSave = new Button
            {
                Text = "Save Changes",
                Left = 10,
                Top = 530,
                Width = 260
            };
            btnSave.Click += BtnSave_Click;
            this.Controls.Add(btnSave);

            // Initially disable buttons until something is selected
            SetControlsEnabled(false);
        }

        private void Editor_SelectionChanged(object sender, EventArgs e)
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            // Get the current selection
            Editor ed = doc.Editor;
            PromptSelectionResult selRes = ed.SelectImplied();

            // Check if a single entity is selected
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

        private void LoadComponentData(ObjectId objId)
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            // Skip if transactions are active
            if (db.TransactionManager.TopTransaction != null) return;

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
                        // Extract component data
                        string componentType = "", floor = "";
                        TypedValue[] xdataComp = rbComp.AsArray();
                        for (int i = 1; i < xdataComp.Length; i++)
                        {
                            if (i == 1) componentType = xdataComp[i].Value.ToString();
                            if (i == 2) floor = xdataComp[i].Value.ToString();
                        }

                        // Get component length
                        double length = 0;
                        if (pline.NumberOfVertices >= 2)
                        {
                            Point3d startPt = pline.GetPoint3dAt(0);
                            Point3d endPt = pline.GetPoint3dAt(1);
                            length = startPt.DistanceTo(endPt);
                        }

                        // Extract parts data (similar to previous methods)
                        string partsJson = GetPartsJsonFromEntity(pline);
                        currentParts.Clear();

                        if (!string.IsNullOrEmpty(partsJson))
                        {
                            try
                            {
                                //MessageBox.Show($"Deserializing JSON: {partsJson.Substring(0, Math.Min(50, partsJson.Length))}...");
                                currentParts = Newtonsoft.Json.JsonConvert.DeserializeObject<List<ChildPart>>(partsJson);
                                //MessageBox.Show($"Deserialized {currentParts.Count} parts"); ;
                            }
                            catch (System.Exception ex)
                            {
                                MessageBox.Show($"JSON deserialization error: {ex.Message}");
                            }
                        }

                        // Update the UI
                        Invoke(new Action(() =>
                        {
                            txtComponentType.Text = componentType;
                            txtComponentFloor.Text = floor;
                            lblLength.Text = length.ToString("F4");

                            // Update parts list
                            partsList.Items.Clear();
                            //MessageBox.Show($"Updating UI with {currentParts.Count} parts");
                            foreach (ChildPart part in currentParts)
                            {
                                double actualLength = part.GetActualLength(length);
                                ListViewItem item = new ListViewItem(part.Name);
                                item.SubItems.Add(part.PartType);
                                item.SubItems.Add(actualLength.ToString("F4"));
                                item.Tag = part; // Store reference to part object
                                partsList.Items.Add(item);
                            }

                            SetControlsEnabled(true);
                        }));

                        currentComponentId = objId;
                        return;
                    }
                }

                tr.Commit();
            }

            // If we get here, it's not a metal component
            ClearComponentData();
        }

        public static string GetPartsJsonFromEntity(Polyline pline)
        {
            // Add debugging message boxes
            //MessageBox.Show("Starting to read parts data");

            string partsJson = "";

            // First check if we have the chunk count info
            ResultBuffer rbInfo = pline.GetXDataForApplication("METALPARTSINFO");
            if (rbInfo == null)
            {
                //MessageBox.Show("No METALPARTSINFO found");
                return "";
            }

            int chunkCount = 0;
            TypedValue[] xdataInfo = rbInfo.AsArray();
            if (xdataInfo.Length >= 2)
            {
                chunkCount = (int)xdataInfo[1].Value;
                //MessageBox.Show($"Found {chunkCount} chunks");
            }

            // Read each chunk using its specific RegApp name
            for (int i = 0; i < chunkCount; i++)
            {
                string chunkAppName = $"METALPARTS{i}";
                ResultBuffer rbChunk = pline.GetXDataForApplication(chunkAppName);

                if (rbChunk != null)
                {
                    TypedValue[] xdataChunk = rbChunk.AsArray();
                    if (xdataChunk.Length >= 2)
                    {
                        partsJson += xdataChunk[1].Value.ToString();
                        //MessageBox.Show($"Chunk {i}: {xdataChunk[1].Value.ToString().Substring(0, Math.Min(20, xdataChunk[1].Value.ToString().Length))}...");
                    }
                }
                else
                {
                    //MessageBox.Show($"Chunk {i} not found");
                }
            }

            if (!string.IsNullOrEmpty(partsJson))
            {
                //MessageBox.Show($"Final JSON length: {partsJson.Length}");
            }
            else
            {
                //MessageBox.Show("No JSON data found");
            }

            return partsJson;
        }

        private void ClearComponentData()
        {
            Invoke(new Action(() =>
            {
                txtComponentType.Text = "";
                txtComponentFloor.Text = "";
                lblLength.Text = "0.0";
                partsList.Items.Clear();
                SetControlsEnabled(false);
            }));

            currentComponentId = ObjectId.Null;
            currentParts.Clear();
        }

        private void SetControlsEnabled(bool enabled)
        {
            btnAddPart.Enabled = enabled;
            btnEditPart.Enabled = enabled;
            btnDeletePart.Enabled = enabled;
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
                    double length = double.Parse(lblLength.Text);
                    double actualLength = form.ResultPart.GetActualLength(length);

                    ListViewItem item = new ListViewItem(form.ResultPart.Name);
                    item.SubItems.Add(form.ResultPart.PartType);
                    item.SubItems.Add(actualLength.ToString("F4"));
                    item.Tag = form.ResultPart;
                    partsList.Items.Add(item);
                }
            }
        }

        private void BtnEditPart_Click(object sender, EventArgs e)
        {
            if (partsList.SelectedItems.Count == 0) return;

            ListViewItem item = partsList.SelectedItems[0];
            ChildPart part = item.Tag as ChildPart;

            using (PartEditForm form = new PartEditForm(part))
            {
                if (form.ShowDialog() == DialogResult.OK)
                {
                    // Update the part in the list
                    int index = currentParts.IndexOf(part);
                    currentParts[index] = form.ResultPart;

                    // Update the UI
                    double length = double.Parse(lblLength.Text);
                    double actualLength = form.ResultPart.GetActualLength(length);

                    item.Text = form.ResultPart.Name;
                    item.SubItems[1].Text = form.ResultPart.PartType;
                    item.SubItems[2].Text = actualLength.ToString("F4");
                    item.Tag = form.ResultPart;
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

                MessageBox.Show("Changes saved successfully.", "Save Complete");
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Error initiating save: {ex.Message}\n\nStack Trace: {ex.StackTrace}", "Save Error");
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
        private TextBox txtLengthAdj;
        private TextBox txtMaterial;

        private ComboBox cboAttachControl;
        private CheckBox chkInvertControl;
        private TextBox txtAdjustControl;
        private CheckBox chkClipsControl;

        // The part being edited (if any)
        private ChildPart partToEdit;

        // Result part after editing
        public ChildPart ResultPart { get; private set; }

        // Constructor - accept an optional part to edit
        public PartEditForm(ChildPart part = null)
        {
            // Store the part being edited
            this.partToEdit = part;

            // Initialize the form components
            InitializeComponent();
        }

        private void InitializeComponent()
        {
            this.Text = "Edit Part";
            this.Width = 500;  // Increased width for more controls
            this.Height = 450; // Increased height for more controls
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

            Label lblLengthAdj = new Label { Text = "Length Adj:", Left = 10, Top = 70, Width = 80 };
            this.Controls.Add(lblLengthAdj);

            txtLengthAdj = new TextBox { Left = 100, Top = 70, Width = 280, Text = "0.0" };
            this.Controls.Add(txtLengthAdj);

            Label lblMaterial = new Label { Text = "Material:", Left = 10, Top = 100, Width = 80 };
            this.Controls.Add(lblMaterial);

            txtMaterial = new TextBox { Left = 100, Top = 100, Width = 280 };
            this.Controls.Add(txtMaterial);

            // New attachment properties section label
            Label lblAttachmentSection = new Label
            {
                Text = "Attachment Properties:",
                Left = 10,
                Top = 130,
                Width = 380,
                Font = new System.Drawing.Font(this.Font, System.Drawing.FontStyle.Bold)
            };
            this.Controls.Add(lblAttachmentSection);

            // Attach property (L/R)
            Label lblAttach = new Label { Text = "Attach:", Left = 10, Top = 160, Width = 80 };
            this.Controls.Add(lblAttach);

            ComboBox cboAttach = new ComboBox
            {
                Left = 100,
                Top = 160,
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
                Top = 190,
                Width = 280
            };
            this.Controls.Add(chkInvert);

            // Adjust value
            Label lblAdjust = new Label { Text = "Adjust:", Left = 10, Top = 220, Width = 80 };
            this.Controls.Add(lblAdjust);

            TextBox txtAdjust = new TextBox
            {
                Left = 100,
                Top = 220,
                Width = 280,
                Text = "0.0"
            };
            this.Controls.Add(txtAdjust);

            // Clips checkbox (for vertical parts)
            CheckBox chkClips = new CheckBox
            {
                Text = "Allows Clips",
                Left = 100,
                Top = 250,
                Width = 280
            };
            this.Controls.Add(chkClips);

            // OK and Cancel buttons
            Button btnOK = new Button
            {
                Text = "OK",
                Left = 210,
                Top = 280,
                Width = 80,
                DialogResult = DialogResult.OK
            };
            btnOK.Click += BtnOK_Click;
            this.Controls.Add(btnOK);

            Button btnCancel = new Button
            {
                Text = "Cancel",
                Left = 300,
                Top = 280,
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
                txtLengthAdj.Text = partToEdit.LengthAdjustment.ToString();
                txtMaterial.Text = partToEdit.Material;

                // Set attachment properties
                cboAttachControl.SelectedItem = partToEdit.Attach ?? "";
                chkInvertControl.Checked = partToEdit.Invert;
                txtAdjustControl.Text = partToEdit.Adjust.ToString();
                chkClipsControl.Checked = partToEdit.Clips;
            }
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
                double lengthAdj = double.Parse(txtLengthAdj.Text);
                double adjust = double.Parse(txtAdjustControl.Text);

                // Create result part
                ResultPart = new ChildPart(txtName.Text, txtType.Text, lengthAdj, txtMaterial.Text);

                // Set attachment properties
                ResultPart.Attach = cboAttachControl.SelectedItem?.ToString() ?? "";
                ResultPart.Invert = chkInvertControl.Checked;
                ResultPart.Adjust = adjust;
                ResultPart.Clips = chkClipsControl.Checked;
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Validation Error");
                this.DialogResult = DialogResult.None;
            }
        }


        // Command to show the panel
        public class MetalPanelCommands
        {
            private static PaletteSet paletteSet;


        }
    }
}
