using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Drawing;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Newtonsoft.Json;

namespace TakeoffBridge
{
    // Data structures
    public class ElevationInstance
    {
        public string Floor { get; set; }
        public int Quantity { get; set; }
        public string Finish { get; set; }
    }

    public class ElevationDefinition
    {
        public string ElevationCode { get; set; }
        public string Description { get; set; }
        public List<ElevationInstance> Instances { get; set; } = new List<ElevationInstance>();
    }

    // Dialog implementation
    public class ElevationManagerDialog : Form
    {
        private ListBox lstElevations;
        private TextBox txtElevationCode;
        private TextBox txtDescription;
        private Button btnAddElevation;
        private Button btnDeleteElevation;

        private DataGridView gridInstances;
        private Button btnAddInstance;
        private Button btnDeleteInstance;

        private Button btnSave;
        private Button btnCancel;

        // Track the current elevation definition
        private ElevationDefinition currentElevation;
        // List of all elevation definitions
        private List<ElevationDefinition> elevationDefinitions = new List<ElevationDefinition>();

        public ElevationManagerDialog()
        {
            InitializeComponent();

            // Load saved elevation definitions
            LoadElevationDefinitions();

            // Display them in the list
            RefreshElevationList();
        }

        private void InitializeComponent()
        {
            // Form setup
            Text = "Elevation Manager";
            Width = 700;
            Height = 500;
            StartPosition = FormStartPosition.CenterParent;

            // Left side - Elevation list and properties
            Label lblElevations = new Label { Text = "Elevations", Left = 10, Top = 10, Width = 100 };
            Controls.Add(lblElevations);

            lstElevations = new ListBox { Left = 10, Top = 30, Width = 200, Height = 200 };
            lstElevations.SelectedIndexChanged += LstElevations_SelectedIndexChanged;
            Controls.Add(lstElevations);

            Label lblCode = new Label { Text = "Elevation Code:", Left = 10, Top = 240, Width = 100 };
            Controls.Add(lblCode);

            txtElevationCode = new TextBox { Left = 110, Top = 240, Width = 100 };
            Controls.Add(txtElevationCode);

            Label lblDescription = new Label { Text = "Description:", Left = 10, Top = 270, Width = 100 };
            Controls.Add(lblDescription);

            txtDescription = new TextBox { Left = 110, Top = 270, Width = 200 };
            Controls.Add(txtDescription);

            btnAddElevation = new Button { Text = "Add", Left = 10, Top = 300, Width = 80 };
            btnAddElevation.Click += BtnAddElevation_Click;
            Controls.Add(btnAddElevation);

            btnDeleteElevation = new Button { Text = "Delete", Left = 100, Top = 300, Width = 80 };
            btnDeleteElevation.Click += BtnDeleteElevation_Click;
            Controls.Add(btnDeleteElevation);

            // Right side - Instances grid
            Label lblInstances = new Label { Text = "Elevation Instances", Left = 320, Top = 10, Width = 150 };
            Controls.Add(lblInstances);

            gridInstances = new DataGridView
            {
                Left = 320,
                Top = 30,
                Width = 350,
                Height = 250,
                AllowUserToAddRows = false,
                AllowUserToDeleteRows = false,
                SelectionMode = DataGridViewSelectionMode.FullRowSelect
            };

            // Set up grid columns
            gridInstances.Columns.Add("Floor", "Floor");
            gridInstances.Columns.Add("Quantity", "Quantity");
            gridInstances.Columns.Add("Finish", "Finish");
            Controls.Add(gridInstances);

            btnAddInstance = new Button { Text = "Add Instance", Left = 320, Top = 290, Width = 100 };
            btnAddInstance.Click += BtnAddInstance_Click;
            Controls.Add(btnAddInstance);

            btnDeleteInstance = new Button { Text = "Delete Instance", Left = 430, Top = 290, Width = 100 };
            btnDeleteInstance.Click += BtnDeleteInstance_Click;
            Controls.Add(btnDeleteInstance);

            // Bottom - Save/Cancel buttons
            btnSave = new Button { Text = "Save", Left = 500, Top = 420, Width = 80, DialogResult = DialogResult.OK };
            btnSave.Click += BtnSave_Click;
            Controls.Add(btnSave);

            btnCancel = new Button { Text = "Cancel", Left = 590, Top = 420, Width = 80, DialogResult = DialogResult.Cancel };
            Controls.Add(btnCancel);
        }

        private void RefreshElevationList()
        {
            lstElevations.Items.Clear();
            foreach (var elev in elevationDefinitions)
            {
                lstElevations.Items.Add(elev.ElevationCode);
            }
        }

        private void LstElevations_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lstElevations.SelectedIndex >= 0)
            {
                // Load selected elevation
                currentElevation = elevationDefinitions[lstElevations.SelectedIndex];

                // Update form fields
                txtElevationCode.Text = currentElevation.ElevationCode;
                txtDescription.Text = currentElevation.Description;

                // Update grid
                RefreshInstancesGrid();
            }
        }

        private void RefreshInstancesGrid()
        {
            gridInstances.Rows.Clear();

            if (currentElevation != null)
            {
                foreach (var instance in currentElevation.Instances)
                {
                    gridInstances.Rows.Add(instance.Floor, instance.Quantity, instance.Finish);
                }
            }
        }

        private void BtnAddElevation_Click(object sender, EventArgs e)
        {
            // Create a new elevation definition
            ElevationDefinition newElevation = new ElevationDefinition
            {
                ElevationCode = "New",
                Description = "New Elevation"
            };

            elevationDefinitions.Add(newElevation);

            // Refresh list and select the new item
            RefreshElevationList();
            lstElevations.SelectedIndex = elevationDefinitions.Count - 1;
        }

        private void BtnDeleteElevation_Click(object sender, EventArgs e)
        {
            if (lstElevations.SelectedIndex >= 0)
            {
                if (MessageBox.Show("Are you sure you want to delete this elevation?",
                    "Confirm Delete", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    elevationDefinitions.RemoveAt(lstElevations.SelectedIndex);
                    currentElevation = null;

                    RefreshElevationList();
                    txtElevationCode.Text = "";
                    txtDescription.Text = "";
                    gridInstances.Rows.Clear();
                }
            }
        }

        private void BtnAddInstance_Click(object sender, EventArgs e)
        {
            if (currentElevation != null)
            {
                // Add a new instance
                ElevationInstance newInstance = new ElevationInstance
                {
                    Floor = "01",
                    Quantity = 1,
                    Finish = "Paint"
                };

                currentElevation.Instances.Add(newInstance);

                // Refresh grid
                RefreshInstancesGrid();
            }
        }

        private void BtnDeleteInstance_Click(object sender, EventArgs e)
        {
            if (currentElevation != null && gridInstances.SelectedRows.Count > 0)
            {
                int selectedIndex = gridInstances.SelectedRows[0].Index;
                currentElevation.Instances.RemoveAt(selectedIndex);

                // Refresh grid
                RefreshInstancesGrid();
            }
        }

        private void BtnSave_Click(object sender, EventArgs e)
        {
            // Save any pending changes to the current elevation
            if (currentElevation != null)
            {
                currentElevation.ElevationCode = txtElevationCode.Text;
                currentElevation.Description = txtDescription.Text;

                // Update instances from grid
                currentElevation.Instances.Clear();
                foreach (DataGridViewRow row in gridInstances.Rows)
                {
                    ElevationInstance instance = new ElevationInstance
                    {
                        Floor = row.Cells["Floor"].Value?.ToString() ?? "01",
                        Quantity = Convert.ToInt32(row.Cells["Quantity"].Value ?? 1),
                        Finish = row.Cells["Finish"].Value?.ToString() ?? "Paint"
                    };

                    currentElevation.Instances.Add(instance);
                }
            }

            // Save all elevation definitions
            SaveElevationDefinitions();
        }

        private void LoadElevationDefinitions()
        {
            // Load from the drawing database
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
            {
                try
                {
                    // Get named objects dictionary
                    DBDictionary nod = (DBDictionary)tr.GetObject(
                        doc.Database.NamedObjectsDictionaryId,
                        OpenMode.ForRead);

                    const string dictName = "ELEVATIONDEFINITIONS";

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
                                elevationDefinitions = Newtonsoft.Json.JsonConvert.DeserializeObject<List<ElevationDefinition>>(json);
                            }
                        }
                    }

                    tr.Commit();
                }
                catch (System.Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error loading elevation definitions: {ex.Message}");
                    tr.Abort();
                }
            }
        }

        private void SaveElevationDefinitions()
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            using (Transaction tr = doc.Database.TransactionManager.StartTransaction())
            {
                try
                {
                    // Get named objects dictionary
                    DBDictionary nod = (DBDictionary)tr.GetObject(
                        doc.Database.NamedObjectsDictionaryId,
                        OpenMode.ForWrite);

                    // Register app name
                    RegAppTable regTable = (RegAppTable)tr.GetObject(doc.Database.RegAppTableId, OpenMode.ForWrite);
                    if (!regTable.Has("ELEVATIONDEFINITIONS"))
                    {
                        RegAppTableRecord record = new RegAppTableRecord();
                        record.Name = "ELEVATIONDEFINITIONS";
                        regTable.Add(record);
                        tr.AddNewlyCreatedDBObject(record, true);
                    }

                    // Create or update the Xrecord
                    Xrecord xrec;
                    const string dictName = "ELEVATIONDEFINITIONS";

                    if (nod.Contains(dictName))
                    {
                        // Update existing
                        xrec = (Xrecord)tr.GetObject(nod.GetAt(dictName), OpenMode.ForWrite);
                    }
                    else
                    {
                        // Create new
                        xrec = new Xrecord();
                        nod.SetAt(dictName, xrec);
                        tr.AddNewlyCreatedDBObject(xrec, true);
                    }

                    // Serialize the data
                    string json = Newtonsoft.Json.JsonConvert.SerializeObject(elevationDefinitions);

                    // Set the data
                    ResultBuffer rb = new ResultBuffer(
                        new TypedValue((int)DxfCode.Text, json)
                    );
                    xrec.Data = rb;

                    tr.Commit();
                }
                catch (System.Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error saving elevation definitions: {ex.Message}");
                    tr.Abort();
                }
            }
        }
    }

    // Commands
    public class ElevationManagerCommands
    {
        [CommandMethod("ELEVATIONMANAGER")]
        public void ShowElevationManager()
        {
            using (ElevationManagerDialog dlg = new ElevationManagerDialog())
            {
                if (dlg.ShowDialog() == DialogResult.OK)
                {
                    // The dialog already saved the definitions to the drawing
                    Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument.Editor.WriteMessage("\nElevation definitions updated.");
                }
            }
        }

        // Helper methods for loading/saving elevation definitions that might be needed by other classes
        public static List<ElevationDefinition> LoadElevationDefinitions()
        {
            List<ElevationDefinition> elevationDefinitions = new List<ElevationDefinition>();

            // Implementation of loading from drawing...

            return elevationDefinitions;
        }
    }
}