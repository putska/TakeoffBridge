using System.Collections.Generic;
using System.Drawing;
using System.Windows.Forms;
using TakeoffBridge;

public class CopyPropertyDialog : Form
{
    public bool CopyComponentType { get; private set; }
    public bool CopyFloor { get; private set; }
    public bool CopyElevation { get; private set; }
    public bool CopyParts { get; private set; }

    private CheckBox chkComponentType;
    private CheckBox chkFloor;
    private CheckBox chkElevation;
    private CheckBox chkParts;
    private ListView partsList;

    // Add these label field declarations
    private Label lblComponentType;
    private Label lblFloor;
    private Label lblElevation;
    private Label lblPartsCount;

    // Add private fields to store the values
    private string _componentType;
    private string _floor;
    private string _elevation;
    private List<ChildPart> _parts;

    public CopyPropertyDialog(string componentType, string floor, string elevation, List<ChildPart> parts)
    {
        // Store values before initializing components
        _componentType = componentType ?? "";
        _floor = floor ?? "";
        _elevation = elevation ?? "";
        _parts = parts ?? new List<ChildPart>();

        // Initialize components first
        InitializeComponent();

        // Now set label texts after initialization
        UpdateLabelTexts();

        // Add parts to the preview list after initialization
        UpdatePartsList();

        // Default all checkboxes to checked
        chkComponentType.Checked = true;
        chkFloor.Checked = true;
        chkElevation.Checked = true;
        chkParts.Checked = true;
    }

    private void UpdateLabelTexts()
    {
        // Set texts on labels after they've been initialized
        lblComponentType.Text = $"Component Type: {_componentType}";
        lblFloor.Text = $"Floor: {_floor}";
        lblElevation.Text = $"Elevation: {_elevation}";
        lblPartsCount.Text = $"Parts: {_parts.Count} items";
    }

    private void UpdatePartsList()
    {
        // Clear existing items
        partsList.Items.Clear();

        // Add parts to the preview list
        foreach (ChildPart part in _parts)
        {
            ListViewItem item = new ListViewItem(part.Name);
            item.SubItems.Add(part.PartType);
            item.SubItems.Add(part.IsFixedLength ? $"Fixed: {part.FixedLength:F2}" :
                             $"Adj: {part.StartAdjustment:F2}, {part.EndAdjustment:F2}");
            partsList.Items.Add(item);
        }
    }

    private void InitializeComponent()
    {
        // Setup dialog
        this.Text = "Copy Properties";
        this.Width = 400;
        this.Height = 500;
        this.FormBorderStyle = FormBorderStyle.FixedDialog;
        this.MaximizeBox = false;
        this.MinimizeBox = false;
        this.StartPosition = FormStartPosition.CenterParent;

        // Create labels for source info
        Label lblSourceInfo = new Label();
        lblSourceInfo.Text = "Source Component Information:";
        lblSourceInfo.Font = new Font(this.Font, FontStyle.Bold);
        lblSourceInfo.Location = new Point(10, 10);
        lblSourceInfo.Size = new Size(380, 20);
        this.Controls.Add(lblSourceInfo);

        // Initialize the label fields
        lblComponentType = new Label();
        lblComponentType.Location = new Point(20, 35);
        lblComponentType.Size = new Size(350, 20);
        this.Controls.Add(lblComponentType);

        lblFloor = new Label();
        lblFloor.Location = new Point(20, 55);
        lblFloor.Size = new Size(350, 20);
        this.Controls.Add(lblFloor);

        lblElevation = new Label();
        lblElevation.Location = new Point(20, 75);
        lblElevation.Size = new Size(350, 20);
        this.Controls.Add(lblElevation);

        lblPartsCount = new Label();
        lblPartsCount.Location = new Point(20, 95);
        lblPartsCount.Size = new Size(350, 20);
        this.Controls.Add(lblPartsCount);

        // Create checkboxes for what to copy
        Label lblCopyOptions = new Label();
        lblCopyOptions.Text = "What to Copy:";
        lblCopyOptions.Font = new Font(this.Font, FontStyle.Bold);
        lblCopyOptions.Location = new Point(10, 125);
        lblCopyOptions.Size = new Size(380, 20);
        this.Controls.Add(lblCopyOptions);

        chkComponentType = new CheckBox();
        chkComponentType.Text = "Component Type";
        chkComponentType.Location = new Point(20, 150);
        chkComponentType.Size = new Size(150, 20);
        this.Controls.Add(chkComponentType);

        chkFloor = new CheckBox();
        chkFloor.Text = "Floor";
        chkFloor.Location = new Point(200, 150);
        chkFloor.Size = new Size(150, 20);
        this.Controls.Add(chkFloor);

        chkElevation = new CheckBox();
        chkElevation.Text = "Elevation";
        chkElevation.Location = new Point(20, 175);
        chkElevation.Size = new Size(150, 20);
        this.Controls.Add(chkElevation);

        chkParts = new CheckBox();
        chkParts.Text = "Child Parts";
        chkParts.Location = new Point(200, 175);
        chkParts.Size = new Size(150, 20);
        this.Controls.Add(chkParts);

        // Add parts preview
        Label lblPartsPreview = new Label();
        lblPartsPreview.Text = "Parts Preview:";
        lblPartsPreview.Font = new Font(this.Font, FontStyle.Bold);
        lblPartsPreview.Location = new Point(10, 205);
        lblPartsPreview.Size = new Size(380, 20);
        this.Controls.Add(lblPartsPreview);

        partsList = new ListView();
        partsList.View = View.Details;
        partsList.FullRowSelect = true;
        partsList.Location = new Point(20, 230);
        partsList.Size = new Size(350, 180);
        partsList.Columns.Add("Name", 120);
        partsList.Columns.Add("Type", 70);
        partsList.Columns.Add("Dimensions", 160);
        this.Controls.Add(partsList);

        // Add OK and Cancel buttons
        Button btnOK = new Button();
        btnOK.Text = "OK";
        btnOK.Location = new Point(210, 420);
        btnOK.Size = new Size(80, 30);
        btnOK.DialogResult = DialogResult.OK;
        btnOK.Click += (s, e) => {
            CopyComponentType = chkComponentType.Checked;
            CopyFloor = chkFloor.Checked;
            CopyElevation = chkElevation.Checked;
            CopyParts = chkParts.Checked;
        };
        this.Controls.Add(btnOK);

        Button btnCancel = new Button();
        btnCancel.Text = "Cancel";
        btnCancel.Location = new Point(300, 420);
        btnCancel.Size = new Size(80, 30);
        btnCancel.DialogResult = DialogResult.Cancel;
        this.Controls.Add(btnCancel);

        this.AcceptButton = btnOK;
        this.CancelButton = btnCancel;
    }
}