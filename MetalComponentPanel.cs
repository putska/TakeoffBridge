using System;
using System.Drawing;
using System.Windows.Forms;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using System.Collections.Generic;
using System.Net.Mail;
using System.Linq;


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
        private System.Windows.Forms.Timer selectionTimer;
        private ObjectId currentComponentId;
        private List<ChildPart> currentParts = new List<ChildPart>();
        private double currentComponentLength = 0.0;
        private bool isHorizontal = true; // Track if current component is horizontal
        private List<Attachment> currentAttachments = new List<Attachment>();
        private Panel attachmentPanel; // New visualization panel for attachments
        private bool forceRefreshOnNextSelection = false;
        private bool hasUnsavedChanges = false;

        public EnhancedMetalComponentPanel()
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

                // Check if we went from having something selected to nothing selected
                bool hadSelectionBefore = lastSelection.Count > 0;
                bool hasSelectionNow = selRes.Status == PromptStatus.OK && selRes.Value.Count > 0;

                // If selection is about to change and we have unsaved changes, save them
                if ((selectionChanged || !hasSelectionNow) && hasUnsavedChanges && currentComponentId != ObjectId.Null)
                {
                    // Save any pending changes before changing selection
                    SaveChanges();
                }

                if (!hasSelectionNow)
                {
                    // Nothing selected - make sure UI is cleared
                    if (currentComponentId != ObjectId.Null)
                    {
                        ClearComponentData();
                        // Add debug message to confirm clearing happened
                        System.Diagnostics.Debug.WriteLine("Selection cleared - UI should be reset");
                    }
                    // Clear the last selection
                    lastSelection.Clear();
                    return;
                }

                if (selRes.Status == PromptStatus.OK)
                {
                    // Force refresh if flag is set or if we previously had nothing selected
                    if (forceRefreshOnNextSelection || !hadSelectionBefore)
                    {
                        selectionChanged = true;
                        forceRefreshOnNextSelection = false; // Reset the flag
                    }
                    else
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
                    }

                    // Update last selection
                    lastSelection.Clear();
                    foreach (SelectedObject selObj in selRes.Value)
                    {
                        lastSelection.Add(selObj.ObjectId);
                    }
                }

                // If selection changed, update the panel
                if (selectionChanged)
                {
                    // Check for unsaved changes before loading a new component
                    if (hasUnsavedChanges && currentComponentId != ObjectId.Null)
                    {
                        SaveChanges();
                    }

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
            this.Text = "Enhanced Metal Component Editor";
            this.Width = 600; // Increased width for visualization panel
            this.Height = 700; // Increased height for visualization panel

            // Parent info section
            Label lblParentSection = new Label
            {
                Text = "Component Properties:",
                Left = 10,
                Top = 10,
                Width = 380,
                Font = new System.Drawing.Font(this.Font, System.Drawing.FontStyle.Bold)
            };
            this.Controls.Add(lblParentSection);

            Label lblType = new Label { Text = "Type:", Left = 10, Top = 40, Width = 100 };
            this.Controls.Add(lblType);

            txtComponentType = new TextBox { Left = 120, Top = 40, Width = 150 };
            this.Controls.Add(txtComponentType);

            Label lblFloor = new Label { Text = "Floor:", Left = 10, Top = 70, Width = 100 };
            this.Controls.Add(lblFloor);

            txtFloor = new TextBox { Left = 120, Top = 70, Width = 150 };
            this.Controls.Add(txtFloor);

            Label lblElevation = new Label { Text = "Elevation:", Left = 10, Top = 100, Width = 100 };
            this.Controls.Add(lblElevation);

            txtElevation = new TextBox { Left = 120, Top = 100, Width = 150 };
            this.Controls.Add(txtElevation);

            // Add length display
            Label lblLengthCaption = new Label { Text = "Length:", Left = 10, Top = 130, Width = 100 };
            this.Controls.Add(lblLengthCaption);

            lblLength = new Label { Text = "0.0", Left = 120, Top = 130, Width = 150 };
            this.Controls.Add(lblLength);

            btnSaveParent = new Button
            {
                Text = "Save Component",
                Left = 120,
                Top = 160,
                Width = 150
            };
            btnSaveParent.Click += BtnSaveParent_Click;
            this.Controls.Add(btnSaveParent);

            // Add visualization panel
            Label lblVisualization = new Label
            {
                Text = "Part Visualization:",
                Left = 10,
                Top = 200,
                Width = 580,
                Font = new System.Drawing.Font(this.Font, System.Drawing.FontStyle.Bold)
            };
            this.Controls.Add(lblVisualization);

            visualPanel = new Panel
            {
                Left = 10,
                Top = 230,
                Width = 580,
                Height = 150,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.White
            };
            visualPanel.Paint += VisualPanel_Paint;
            this.Controls.Add(visualPanel);

            // Create an attachment visualization panel
            Label lblAttachments = new Label
            {
                Text = "Attachments:",
                Left = 10,
                Top = 390,
                Width = 580,
                Font = new System.Drawing.Font(this.Font, System.Drawing.FontStyle.Bold)
            };
            this.Controls.Add(lblAttachments);

            attachmentPanel = new Panel
            {
                Left = 10,
                Top = 420,
                Width = 580,
                Height = 150,
                BorderStyle = BorderStyle.FixedSingle,
                BackColor = Color.White
            };
            attachmentPanel.Paint += AttachmentPanel_Paint;
            this.Controls.Add(attachmentPanel);

            // Child parts section
            Label lblParts = new Label
            {
                Text = "Child Parts:",
                Left = 10,
                Top = 580,
                Width = 100,
                Font = new System.Drawing.Font(this.Font, System.Drawing.FontStyle.Bold)
            };
            this.Controls.Add(lblParts);

            partsList = new ListView
            {
                Left = 10,
                Top = 610,
                Width = 580,
                Height = 200,
                View = View.Details,
                FullRowSelect = true
            };
            partsList.Columns.Add("Name", 120);
            partsList.Columns.Add("Type", 80);
            partsList.Columns.Add("Length", 80);
            partsList.Columns.Add("Start Adj", 80);
            partsList.Columns.Add("End Adj", 80);
            partsList.Columns.Add("Attach", 80);
            this.Controls.Add(partsList);

            // Add the double-click handler
            partsList.DoubleClick += PartsList_DoubleClick;

            // Add selection changed handler to update the visualization
            partsList.SelectedIndexChanged += (s, e) => visualPanel.Invalidate();

            // Part management buttons
            btnAddPart = new Button { Text = "Add Part", Left = 10, Top = 850, Width = 80 };
            btnAddPart.Click += BtnAddPart_Click;
            this.Controls.Add(btnAddPart);

            btnEditPart = new Button { Text = "Edit Part", Left = 100, Top = 850, Width = 80 };
            btnEditPart.Click += BtnEditPart_Click;
            this.Controls.Add(btnEditPart);

            btnDeletePart = new Button { Text = "Delete Part", Left = 190, Top = 850, Width = 80 };
            btnDeletePart.Click += BtnDeletePart_Click;
            this.Controls.Add(btnDeletePart);

            btnSave = new Button { Text = "Save Changes", Left = 280, Top = 850, Width = 100 };
            btnSave.Click += BtnSave_Click;
            this.Controls.Add(btnSave);

            // Add a separator line before the command buttons
            Panel separatorLine = new Panel
            {
                Left = 10,
                Top = 870, // Just above the buttons
                Width = 580,
                Height = 2,
                BackColor = Color.Gray
            };
            this.Controls.Add(separatorLine);

            Button btnCopyProperty = new Button
            {
                Text = "Copy Property",
                Left = 10,
                Top = 900, // Adjust this based on your current layout
                Width = 120
            };
            btnCopyProperty.Click += BtnCopyProperty_Click;
            this.Controls.Add(btnCopyProperty);

            Button btnDetectAttachments = new Button
            {
                Text = "Detect Attachments",
                Left = 140,
                Top = 900, // Same top position as the previous button
                Width = 120
            };
            btnDetectAttachments.Click += BtnDetectAttachments_Click;
            this.Controls.Add(btnDetectAttachments);

            Button btnAddMetalPart = new Button
            {
                Text = "Add Metal Part",
                Left = 270,
                Top = 900, // Same top position
                Width = 120
            };
            btnAddMetalPart.Click += BtnAddMetalPart_Click;
            this.Controls.Add(btnAddMetalPart);

            // Initially disable buttons until something is selected
            SetControlsEnabled(false);
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

            // Set scaling factor to fit component in panel with padding
            double scaleFactor = (visualPanel.Width - (borderPadding * 2)) / currentComponentLength;

            // Draw the baseline representing the component
            int yBase = visualPanel.Height / 2;
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
            int partHeight = 15;  // Height of each part

            // Calculate available space and adjust spacing
            int totalAvailableVerticalSpace = yBase - borderPadding;

            // Make sure we can fit all the parts with some spacing
            int partSpacing = Math.Max(5, Math.Min(15, (totalAvailableVerticalSpace - (currentParts.Count * partHeight)) / (currentParts.Count + 1)));

            // Debug
            System.Diagnostics.Debug.WriteLine($"Available space: {totalAvailableVerticalSpace}, Parts: {currentParts.Count}, Spacing: {partSpacing}");

            foreach (ChildPart part in currentParts)
            {
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
                int partY = borderPadding + (partIndex * (partHeight + partSpacing));

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

                // Center the text and adjust position for readability
                int textX = partStart + (partEnd - partStart) / 2 - (int)(g.MeasureString(displayName, nameFont).Width / 2);
                g.DrawString(displayName, nameFont, Brushes.Black, textX, partY - nameFont.Height - 2);

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

            // Draw a subtle border around the entire visualization area
            g.DrawRectangle(new Pen(Color.Gray, 1),
                            borderPadding / 2,
                            borderPadding / 2,
                            visualPanel.Width - borderPadding,
                            visualPanel.Height - borderPadding);
        }

        private void LoadComponentData(ObjectId objId)
        {
            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
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

                            // Update parts list
                            UpdatePartsList();

                            // Update visualization
                            visualPanel.Invalidate();

                            // also update attachment visualization
                            attachmentPanel.Invalidate();

                            SetControlsEnabled(true);
                        }));

                        currentComponentId = objId;
                        return;
                    }
                }

                tr.Commit();
            }

            // Clear data if no valid component found
            ClearComponentData();
        }

        private List<Attachment> LoadAttachmentsFromDrawing()
        {
            List<Attachment> attachments = new List<Attachment>();

            Document doc = Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                // Get named objects dictionary
                DBDictionary nod = (DBDictionary)tr.GetObject(db.NamedObjectsDictionaryId, OpenMode.ForRead);

                // Check if entry exists
                const string dictName = "METALATTACHMENTS";

                if (nod.Contains(dictName))
                {
                    DBObject obj = tr.GetObject(nod.GetAt(dictName), OpenMode.ForRead);
                    if (obj is Xrecord)
                    {
                        Xrecord xrec = obj as Xrecord;
                        ResultBuffer rb = xrec.Data;

                        if (rb != null)
                        {
                            TypedValue[] values = rb.AsArray();
                            if (values.Length > 0 && values[0].TypeCode == (int)DxfCode.Text)
                            {
                                string json = values[0].Value.ToString();
                                attachments = Newtonsoft.Json.JsonConvert.DeserializeObject<List<Attachment>>(json);
                            }
                        }
                    }
                }

                tr.Commit();
            }

            return attachments;
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

            // Set scaling factor to fit vertical component in panel height with padding
            double scaleFactor = (attachmentPanel.Height - (borderPadding * 2)) / verticalLength;

            // Draw the vertical component line
            int xCenter = attachmentPanel.Width / 2;  // Center of panel
            int yTop = borderPadding;
            int yBottom = yTop + (int)(verticalLength * scaleFactor);

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

            // Draw the attachments on each side
            DrawSideAttachments(g, leftAttachments, xCenter, yTop, yBottom, verticalLength, scaleFactor, -1, verticalStartAdj);
            DrawSideAttachments(g, rightAttachments, xCenter, yTop, yBottom, verticalLength, scaleFactor, 1, verticalStartAdj);

            // Draw labels
            g.DrawString("Left Side", new System.Drawing.Font(this.Font, FontStyle.Bold), Brushes.Blue, 10, 10);
            g.DrawString("Right Side", new System.Drawing.Font(this.Font, FontStyle.Bold), Brushes.Red,
                         attachmentPanel.Width - 80, 10);

            // Draw a subtle border around the entire visualization area
            g.DrawRectangle(new Pen(Color.Gray, 1),
                            borderPadding / 2,
                            borderPadding / 2,
                            attachmentPanel.Width - borderPadding,
                            attachmentPanel.Height - borderPadding);
        }

        private void DrawSideAttachments(Graphics g, List<Attachment> attachments, int xCenter, int yTop, int yBottom,
                                       double verticalLength, double scaleFactor, int sideDirection, double verticalStartAdj)
        {
            // sideDirection: -1 for left, 1 for right
            int sideOffset = 70 * sideDirection;  // Distance from center
            Color sideColor = sideDirection < 0 ? Color.Blue : Color.Red;
            Brush textBrush = sideDirection < 0 ? Brushes.Blue : Brushes.Red;
            Pen linePen = new Pen(sideColor, 2);

            // Horizontal line length
            int horizontalLineLength = 60;

            foreach (var attach in attachments)
            {
                try
                {
                    // Calculate adjusted height (accounting for vertical start adjustment)
                    double adjustedHeight = attach.Height + verticalStartAdj;

                    // Calculate position on the Y axis, based on adjusted height
                    int yPos = yBottom - (int)(adjustedHeight * scaleFactor);

                    // Ensure it's within the vertical bounds
                    yPos = Math.Max(yTop, Math.Min(yBottom, yPos));

                    // Calculate horizontal position
                    int xPos = xCenter + sideOffset;

                    // Draw horizontal line
                    g.DrawLine(linePen, xCenter, yPos, xPos, yPos);

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

            // Then update the UI on the UI thread with null checks
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
                if (partsList.Columns.Count < 7)  // Added column for Fixed Length
                {
                    // Remove existing columns first (if any)
                    partsList.Columns.Clear();

                    // Re-add the columns with updated layout
                    partsList.Columns.Add("Name", 100);
                    partsList.Columns.Add("Type", 60);
                    partsList.Columns.Add("Length", 70);
                    partsList.Columns.Add("Fixed", 50);   // New column for Fixed flag
                    partsList.Columns.Add("Start Adj", 70);
                    partsList.Columns.Add("End Adj", 70);
                    partsList.Columns.Add("Attach", 50);
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

        private void SaveChanges()
        {
            if (!hasUnsavedChanges || currentComponentId == ObjectId.Null) return;

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
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Error auto-saving changes: {ex.Message}", "Save Error");
            }
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

            using (PartEditForm form = new PartEditForm(part))
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
        private TextBox txtStartAdj;
        private TextBox txtEndAdj;
        private TextBox txtMaterial;

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
            this.Height = 520; // Increased height for more controls
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

            // New attachment properties section label - moved down
            Label lblAttachmentSection = new Label
            {
                Text = "Attachment Properties:",
                Left = 10,
                Top = 220,
                Width = 380,
                Font = new System.Drawing.Font(this.Font, System.Drawing.FontStyle.Bold)
            };
            this.Controls.Add(lblAttachmentSection);

            // Attach property (L/R)
            Label lblAttach = new Label { Text = "Attach:", Left = 10, Top = 250, Width = 80 };
            this.Controls.Add(lblAttach);

            ComboBox cboAttach = new ComboBox
            {
                Left = 100,
                Top = 250,
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
                Top = 280,
                Width = 280
            };
            this.Controls.Add(chkInvert);

            // Adjust value
            Label lblAdjust = new Label { Text = "Adjust:", Left = 10, Top = 310, Width = 80 };
            this.Controls.Add(lblAdjust);

            TextBox txtAdjust = new TextBox
            {
                Left = 100,
                Top = 310,
                Width = 280,
                Text = "0.0"
            };
            this.Controls.Add(txtAdjust);

            // Clips checkbox (for vertical parts)
            CheckBox chkClips = new CheckBox
            {
                Text = "Allows Clips",
                Left = 100,
                Top = 340,
                Width = 280
            };
            this.Controls.Add(chkClips);

            // Add help text to explain the adjustments
            Label lblHelp = new Label
            {
                Text = "For horizontal parts: Start=Left, End=Right\nFor vertical parts: Start=Bottom, End=Top",
                Left = 100,
                Top = 370,
                Width = 280,
                Height = 40
            };
            this.Controls.Add(lblHelp);

            // OK and Cancel buttons - moved down
            Button btnOK = new Button
            {
                Text = "OK",
                Left = 210,
                Top = 420,
                Width = 80,
                DialogResult = DialogResult.OK
            };
            btnOK.Click += BtnOK_Click;
            this.Controls.Add(btnOK);

            Button btnCancel = new Button
            {
                Text = "Cancel",
                Left = 300,
                Top = 420,
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

                // Set fixed length properties
                chkFixedLength.Checked = partToEdit.IsFixedLength;
                txtFixedLength.Text = partToEdit.FixedLength.ToString();

                // Set start/end adjustment values
                txtStartAdj.Text = partToEdit.StartAdjustment.ToString();
                txtEndAdj.Text = partToEdit.EndAdjustment.ToString();

                txtMaterial.Text = partToEdit.Material;

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
            }
            catch (System.Exception ex)
            {
                MessageBox.Show($"Error: {ex.Message}", "Validation Error");
                this.DialogResult = DialogResult.None;
            }
        }
    }

    public class Attachment
    {
        public string HorizontalHandle { get; set; }
        public string VerticalHandle { get; set; }
        public string HorizontalPartType { get; set; }
        public string VerticalPartType { get; set; }
        public string Side { get; set; }
        public double Position { get; set; }
        public double Height { get; set; }
        public bool Invert { get; set; }
        public double Adjust { get; set; }
    }

}
