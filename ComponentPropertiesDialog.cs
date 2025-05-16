using System;
using System.Drawing;
using System.Windows.Forms;

public class ComponentPropertiesDialog : Form
{
    public string Floor { get; set; }
    public string Elevation { get; set; }

    public string TopType { get; set; }
    public string BottomType { get; set; }
    public string LeftType { get; set; }
    public string RightType { get; set; }

    public bool CreateTop { get; set; }
    public bool CreateBottom { get; set; }
    public bool CreateLeft { get; set; }
    public bool CreateRight { get; set; }

    private TextBox txtFloor;
    private TextBox txtElevation;

    private CheckBox chkTop;
    private CheckBox chkBottom;
    private CheckBox chkLeft;
    private CheckBox chkRight;

    private ComboBox cboTop;
    private ComboBox cboBottom;
    private ComboBox cboLeft;
    private ComboBox cboRight;

    public ComponentPropertiesDialog()
    {
        InitializeComponent();

        // Set default values
        txtFloor.Text = "1";
        txtElevation.Text = "A";

        chkTop.Checked = true;
        chkBottom.Checked = true;
        chkLeft.Checked = true;
        chkRight.Checked = true;

        cboTop.SelectedIndex = 0;
        cboBottom.SelectedIndex = 0;
        cboLeft.SelectedIndex = 0;
        cboRight.SelectedIndex = 0;
    }

    // Add this method to ComponentPropertiesDialog
    public void InitializeValues(string floor, string elevation,
                               string topType, string bottomType,
                               string leftType, string rightType,
                               bool createTop, bool createBottom,
                               bool createLeft, bool createRight)
    {
        // Set text fields
        txtFloor.Text = floor;
        txtElevation.Text = elevation;

        // Set checkboxes
        chkTop.Checked = createTop;
        chkBottom.Checked = createBottom;
        chkLeft.Checked = createLeft;
        chkRight.Checked = createRight;

        // Set combo boxes - we need to find the right indices
        SetComboBoxSelectedValue(cboTop, topType);
        SetComboBoxSelectedValue(cboBottom, bottomType);
        SetComboBoxSelectedValue(cboLeft, leftType);
        SetComboBoxSelectedValue(cboRight, rightType);
    }

    // Helper method to set a ComboBox selected value
    private void SetComboBoxSelectedValue(ComboBox comboBox, string value)
    {
        if (comboBox == null || string.IsNullOrEmpty(value))
            return;

        // First try to find an exact match
        for (int i = 0; i < comboBox.Items.Count; i++)
        {
            if (comboBox.Items[i].ToString().Equals(value, StringComparison.OrdinalIgnoreCase))
            {
                comboBox.SelectedIndex = i;
                return;
            }
        }

        // If no exact match, try to find a partial match
        for (int i = 0; i < comboBox.Items.Count; i++)
        {
            if (comboBox.Items[i].ToString().IndexOf(value, StringComparison.OrdinalIgnoreCase) >= 0)
            {
                comboBox.SelectedIndex = i;
                return;
            }
        }

        // If no match was found, set to first item
        if (comboBox.Items.Count > 0)
            comboBox.SelectedIndex = 0;
    }

    private void InitializeComponent()
    {
        this.Text = "Opening Component Properties";
        this.Width = 450;
        this.Height = 500;
        this.FormBorderStyle = FormBorderStyle.Sizable; // Changed from FixedDialog to make it resizable
        this.MaximizeBox = false;
        this.MinimizeBox = false;
        this.StartPosition = FormStartPosition.CenterParent;

        // Basic properties
        Label lblBasic = new Label();
        lblBasic.Text = "Basic Properties";
        lblBasic.Font = new Font(this.Font, FontStyle.Bold);
        lblBasic.Location = new Point(10, 10);
        lblBasic.AutoSize = true;
        this.Controls.Add(lblBasic);

        Label lblFloor = new Label();
        lblFloor.Text = "Floor:";
        lblFloor.Location = new Point(20, 40);
        lblFloor.AutoSize = true;
        this.Controls.Add(lblFloor);

        txtFloor = new TextBox();
        txtFloor.Location = new Point(100, 37);
        txtFloor.Size = new Size(80, 20);
        this.Controls.Add(txtFloor);

        Label lblElevation = new Label();
        lblElevation.Text = "Elevation:";
        lblElevation.Location = new Point(200, 40);
        lblElevation.AutoSize = true;
        this.Controls.Add(lblElevation);

        txtElevation = new TextBox();
        txtElevation.Location = new Point(270, 37);
        txtElevation.Size = new Size(80, 20);
        this.Controls.Add(txtElevation);

        // Edge selection
        Label lblEdges = new Label();
        lblEdges.Text = "Edge Components";
        lblEdges.Font = new Font(this.Font, FontStyle.Bold);
        lblEdges.Location = new Point(10, 80);
        lblEdges.AutoSize = true;
        this.Controls.Add(lblEdges);

        // Top edge
        chkTop = new CheckBox();
        chkTop.Text = "Create Top Edge";
        chkTop.Location = new Point(20, 110);
        chkTop.AutoSize = true;
        chkTop.CheckedChanged += (s, e) => cboTop.Enabled = chkTop.Checked;
        this.Controls.Add(chkTop);

        Label lblTopType = new Label();
        lblTopType.Text = "Type:";
        lblTopType.Location = new Point(150, 110);
        lblTopType.AutoSize = true;
        this.Controls.Add(lblTopType);

        cboTop = new ComboBox();
        cboTop.Location = new Point(190, 107);
        cboTop.Size = new Size(100, 21);
        cboTop.DropDownStyle = ComboBoxStyle.DropDownList;
        cboTop.Items.AddRange(new object[] { "Head", "Horizontal" });
        this.Controls.Add(cboTop);

        // Bottom edge
        chkBottom = new CheckBox();
        chkBottom.Text = "Create Bottom Edge";
        chkBottom.Location = new Point(20, 140);
        chkBottom.AutoSize = true;
        chkBottom.CheckedChanged += (s, e) => cboBottom.Enabled = chkBottom.Checked;
        this.Controls.Add(chkBottom);

        Label lblBottomType = new Label();
        lblBottomType.Text = "Type:";
        lblBottomType.Location = new Point(150, 140);
        lblBottomType.AutoSize = true;
        this.Controls.Add(lblBottomType);

        cboBottom = new ComboBox();
        cboBottom.Location = new Point(190, 137);
        cboBottom.Size = new Size(100, 21);
        cboBottom.DropDownStyle = ComboBoxStyle.DropDownList;
        cboBottom.Items.AddRange(new object[] { "Sill", "Horizontal" });
        this.Controls.Add(cboBottom);

        // Left edge
        chkLeft = new CheckBox();
        chkLeft.Text = "Create Left Edge";
        chkLeft.Location = new Point(20, 170);
        chkLeft.AutoSize = true;
        chkLeft.CheckedChanged += (s, e) => cboLeft.Enabled = chkLeft.Checked;
        this.Controls.Add(chkLeft);

        Label lblLeftType = new Label();
        lblLeftType.Text = "Type:";
        lblLeftType.Location = new Point(150, 170);
        lblLeftType.AutoSize = true;
        this.Controls.Add(lblLeftType);

        cboLeft = new ComboBox();
        cboLeft.Location = new Point(190, 167);
        cboLeft.Size = new Size(100, 21);
        cboLeft.DropDownStyle = ComboBoxStyle.DropDownList;
        cboLeft.Items.AddRange(new object[] { "JambL", "Vertical" });
        this.Controls.Add(cboLeft);

        // Right edge
        chkRight = new CheckBox();
        chkRight.Text = "Create Right Edge";
        chkRight.Location = new Point(20, 200);
        chkRight.AutoSize = true;
        chkRight.CheckedChanged += (s, e) => cboRight.Enabled = chkRight.Checked;
        this.Controls.Add(chkRight);

        Label lblRightType = new Label();
        lblRightType.Text = "Type:";
        lblRightType.Location = new Point(150, 200);
        lblRightType.AutoSize = true;
        this.Controls.Add(lblRightType);

        cboRight = new ComboBox();
        cboRight.Location = new Point(190, 197);
        cboRight.Size = new Size(100, 21);
        cboRight.DropDownStyle = ComboBoxStyle.DropDownList;
        cboRight.Items.AddRange(new object[] { "JambR", "Vertical" });
        this.Controls.Add(cboRight);

        // Preview
        Panel preview = new Panel();
        preview.BorderStyle = BorderStyle.FixedSingle;
        preview.Location = new Point(20, 230);
        preview.Size = new Size(350, 80);
        preview.BackColor = Color.White;
        preview.Paint += (s, e) => {
            Graphics g = e.Graphics;
            g.SmoothingMode = System.Drawing.Drawing2D.SmoothingMode.AntiAlias;

            // Draw a rectangle representing the opening
            int margin = 10;
            Rectangle rect = new Rectangle(margin, margin,
                                         preview.Width - (2 * margin),
                                         preview.Height - (2 * margin));

            // Draw each edge with color based on selection
            Pen topPen = chkTop.Checked ? Pens.Red : Pens.LightGray;
            Pen bottomPen = chkBottom.Checked ? Pens.Red : Pens.LightGray;
            Pen leftPen = chkLeft.Checked ? Pens.Red : Pens.LightGray;
            Pen rightPen = chkRight.Checked ? Pens.Red : Pens.LightGray;

            g.DrawLine(topPen, rect.Left, rect.Top, rect.Right, rect.Top);
            g.DrawLine(bottomPen, rect.Left, rect.Bottom, rect.Right, rect.Bottom);
            g.DrawLine(leftPen, rect.Left, rect.Top, rect.Left, rect.Bottom);
            g.DrawLine(rightPen, rect.Right, rect.Top, rect.Right, rect.Bottom);

            // Draw center point
            g.FillEllipse(Brushes.Blue, rect.Width / 2 + margin - 3,
                         rect.Height / 2 + margin - 3, 6, 6);

            // Draw text
            g.DrawString("Selection Point", new Font(this.Font.FontFamily, 8),
                        Brushes.Black, rect.Width / 2 - 30, rect.Height / 2 + 10);
        };
        this.Controls.Add(preview);

        // Buttons
        Button btnOK = new Button();
        btnOK.Text = "OK";
        btnOK.DialogResult = DialogResult.OK;
        btnOK.Location = new Point(200, 360);
        btnOK.Size = new Size(100, 30);
        btnOK.Click += (s, e) => {
            // Save values when OK is clicked
            Floor = txtFloor.Text;
            Elevation = txtElevation.Text;

            CreateTop = chkTop.Checked;
            CreateBottom = chkBottom.Checked;
            CreateLeft = chkLeft.Checked;
            CreateRight = chkRight.Checked;

            TopType = cboTop.SelectedItem?.ToString() ?? "Head";
            BottomType = cboBottom.SelectedItem?.ToString() ?? "Sill";
            LeftType = cboLeft.SelectedItem?.ToString() ?? "JambL";
            RightType = cboRight.SelectedItem?.ToString() ?? "JambR";
        };
        this.Controls.Add(btnOK);

        Button btnCancel = new Button();
        btnCancel.Text = "Cancel";
        btnCancel.DialogResult = DialogResult.Cancel;
        btnCancel.Location = new Point(310, 360);
        btnCancel.Size = new Size(100, 30);
        this.Controls.Add(btnCancel);

        this.AcceptButton = btnOK;
        this.CancelButton = btnCancel;
    }
}