using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace TakeoffBridge
{
    public partial class FabricationManager
    {
        // Add this to the existing FabricationManager class

        #region Excel Export Methods

        /// <summary>
        /// Command method to export all parts data to Excel
        /// </summary>
        [CommandMethod("ExcelExport")]
        public static void ExcelExportCommand()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            if (doc == null) return;

            Editor ed = doc.Editor;
            ed.WriteMessage("\n--- Starting Excel Export Process ---");

            try
            {
                // Create a new instance of FabricationManager
                FabricationManager manager = new FabricationManager();

                // Execute the export
                manager.ExportPartsToExcel();

                ed.WriteMessage("\n--- Excel Export Process Completed Successfully ---");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError in Excel export process: {ex.Message}");
                if (ex.InnerException != null)
                {
                    ed.WriteMessage($"\nInner exception: {ex.InnerException.Message}");
                }
            }
        }

        /// <summary>
        /// Exports all parts data to an Excel file
        /// </summary>
        public void ExportPartsToExcel()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            ed.WriteMessage("\nCollecting parts data from drawing...");

            // Get all parts from the drawing
            List<PartInfo> allParts = CollectPartsDataFromDrawing();

            if (allParts.Count == 0)
            {
                ed.WriteMessage("\nNo parts found in the drawing to export.");
                return;
            }

            ed.WriteMessage($"\nFound {allParts.Count} parts in the drawing.");

            // Group similar parts and calculate quantities
            var groupedParts = allParts
                .GroupBy(p => new {
                    p.PartNumber,
                    p.Length,
                    p.MarkNumber,
                    p.Finish,
                    p.Fab
                })
                .Select(group => new {
                    PartNumber = group.Key.PartNumber,
                    Length = group.Key.Length,
                    MarkNumber = group.Key.MarkNumber,
                    Finish = group.Key.Finish,
                    Fab = group.Key.Fab,
                    Quantity = group.Count()
                })
                .OrderBy(p => p.PartNumber)
                .ThenBy(p => p.Length)
                .ToList();

            ed.WriteMessage($"\nGrouped into {groupedParts.Count} unique parts.");

            // Collect glass data
            ed.WriteMessage("\nCollecting glass data from drawing...");
            List<GlassItemInfo> allGlass = CollectGlassDataFromDrawing();

            if (allGlass.Count == 0)
            {
                ed.WriteMessage("\nNo glass items found in the drawing.");
            }
            else
            {
                ed.WriteMessage($"\nFound {allGlass.Count} glass items in the drawing.");
            }

            // Group similar glass items and calculate quantities
            var groupedGlass = allGlass
                .GroupBy(g => new {
                    g.Width,
                    g.Height,
                    g.MarkNumber,
                    g.GlassType
                })
                .Select(group => new {
                    Width = group.Key.Width,
                    Height = group.Key.Height,
                    MarkNumber = group.Key.MarkNumber,
                    GlassType = group.Key.GlassType,
                    Quantity = group.Count()
                })
                .OrderBy(g => g.MarkNumber)
                .ThenBy(g => g.Width)
                .ThenBy(g => g.Height)
                .ToList();

            if (groupedGlass.Count > 0)
            {
                ed.WriteMessage($"\nGrouped into {groupedGlass.Count} unique glass items.");
            }

            // Generate file path
            string filePath = Path.Combine(RootPath, "TakeoffExport.xlsx");

            try
            {
                // Create Excel file with both metal and glass data
                CreateExcelFile(filePath, groupedParts, groupedGlass);
                ed.WriteMessage($"\nExcel file created successfully: {filePath}");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError creating Excel file: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Creates an Excel file with both metal parts and glass data
        /// </summary>
        private void CreateExcelFile(string filePath, IEnumerable<dynamic> parts, IEnumerable<dynamic> glass)
        {
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet metalSheet = null;
            Excel.Worksheet glassSheet = null;

            try
            {
                // Create a new Excel application
                excelApp = new Excel.Application();
                excelApp.Visible = false;
                excelApp.DisplayAlerts = false;

                // Create a new workbook
                workbook = excelApp.Workbooks.Add(Type.Missing);

                // ---------------------- METAL PARTS SHEET ----------------------
                // Get the first worksheet for metal parts
                metalSheet = (Excel.Worksheet)workbook.Sheets[1];
                metalSheet.Name = "Metal Parts";

                // Add headers to the first row
                metalSheet.Cells[1, 1] = "Qty";
                metalSheet.Cells[1, 2] = "Part No";
                metalSheet.Cells[1, 3] = "Length";
                metalSheet.Cells[1, 4] = "Mark No";
                metalSheet.Cells[1, 5] = "Finish";
                metalSheet.Cells[1, 6] = "Fab";

                // Format the header row
                Excel.Range metalHeaderRange = metalSheet.Range["A1:F1"];
                metalHeaderRange.Font.Bold = true;
                metalHeaderRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                metalHeaderRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                // Add data rows for metal parts
                int row = 2;
                foreach (var part in parts)
                {
                    metalSheet.Cells[row, 1] = part.Quantity;
                    metalSheet.Cells[row, 2] = part.PartNumber;

                    // Format the length as needed
                    string formattedLength = FormatInches(part.Length);
                    metalSheet.Cells[row, 3] = formattedLength;

                    metalSheet.Cells[row, 4] = part.MarkNumber;
                    metalSheet.Cells[row, 5] = part.Finish;
                    metalSheet.Cells[row, 6] = part.Fab;

                    row++;
                }

                // Auto-fit columns for better readability
                metalSheet.Columns.AutoFit();

                // Apply borders to all data
                Excel.Range metalDataRange = metalSheet.Range[$"A1:F{row - 1}"];
                metalDataRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                // ---------------------- GLASS SHEET ----------------------
                // Add a new worksheet for glass
                glassSheet = (Excel.Worksheet)workbook.Sheets.Add(After: metalSheet);
                glassSheet.Name = "Glass";

                // Add headers to the first row of glass sheet
                glassSheet.Cells[1, 1] = "Qty";
                glassSheet.Cells[1, 2] = "Width";
                glassSheet.Cells[1, 3] = "Height";
                glassSheet.Cells[1, 4] = "Mark No";
                glassSheet.Cells[1, 5] = "Type";

                // Format the header row
                Excel.Range glassHeaderRange = glassSheet.Range["A1:E1"];
                glassHeaderRange.Font.Bold = true;
                glassHeaderRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                glassHeaderRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                // Add data rows for glass
                row = 2;
                foreach (var item in glass)
                {
                    glassSheet.Cells[row, 1] = item.Quantity;

                    // Format the width and height as needed
                    string formattedWidth = FormatInches(item.Width);
                    string formattedHeight = FormatInches(item.Height);

                    glassSheet.Cells[row, 2] = formattedWidth;
                    glassSheet.Cells[row, 3] = formattedHeight;
                    glassSheet.Cells[row, 4] = item.MarkNumber;
                    glassSheet.Cells[row, 5] = item.GlassType;

                    row++;
                }

                // Auto-fit columns for better readability
                glassSheet.Columns.AutoFit();

                // Apply borders to all data
                Excel.Range glassDataRange = glassSheet.Range[$"A1:E{row - 1}"];
                glassDataRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                // Save and close
                workbook.SaveAs(filePath);
                Document doc = Application.DocumentManager.MdiActiveDocument;
                doc.Editor.WriteMessage($"\nExcel file saved to: {filePath}");
            }
            finally
            {
                // Clean up COM objects to prevent memory leaks
                if (workbook != null)
                {
                    workbook.Close(true);
                    Marshal.ReleaseComObject(workbook);
                }

                if (excelApp != null)
                {
                    excelApp.Quit();
                    Marshal.ReleaseComObject(excelApp);
                }

                if (metalSheet != null)
                {
                    Marshal.ReleaseComObject(metalSheet);
                }

                if (glassSheet != null)
                {
                    Marshal.ReleaseComObject(glassSheet);
                }

                // Force garbage collection
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        /// <summary>
        /// Class to store glass item information
        /// </summary>
        private class GlassItemInfo
        {
            public string GlassType { get; set; }
            public double Width { get; set; }
            public double Height { get; set; }
            public string MarkNumber { get; set; }
            public string Floor { get; set; }
            public string Elevation { get; set; }
        }

        /// <summary>
        /// Collects glass data from the drawing
        /// </summary>
        private List<GlassItemInfo> CollectGlassDataFromDrawing()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Database db = doc.Database;
            List<GlassItemInfo> glassItems = new List<GlassItemInfo>();

            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                try
                {
                    // Use a selection filter to find glass polylines (on Tag1 layer)
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
                                    GlassItemInfo glassInfo = ExtractGlassDataFromXData(rbXdata);
                                    if (glassInfo != null)
                                    {
                                        glassItems.Add(glassInfo);
                                    }
                                }
                            }
                        }
                    }

                    tr.Commit();
                }
                catch (System.Exception ex)
                {
                    doc.Editor.WriteMessage($"\nError collecting glass data: {ex.Message}");
                    tr.Abort();
                }
            }

            return glassItems;
        }

        /// <summary>
        /// Extracts glass data from XData
        /// </summary>
        private GlassItemInfo ExtractGlassDataFromXData(ResultBuffer rbXdata)
        {
            GlassItemInfo glassInfo = new GlassItemInfo();

            TypedValue[] xdata = rbXdata.AsArray();

            // Skip first item (application name)
            for (int i = 1; i < xdata.Length; i++)
            {
                // Process values based on the expected structure
                try
                {
                    if (i == 1) glassInfo.GlassType = xdata[i].Value.ToString();
                    if (i == 2) glassInfo.Floor = xdata[i].Value.ToString();
                    if (i == 3) glassInfo.Elevation = xdata[i].Value.ToString();
                    // Skip bite values at indices 4-7
                    if (i == 8) glassInfo.Width = Convert.ToDouble(xdata[i].Value);
                    if (i == 9) glassInfo.Height = Convert.ToDouble(xdata[i].Value);
                    // Skip DLO dimensions at indices 10-11
                    if (i == 12) glassInfo.MarkNumber = xdata[i].Value.ToString();
                }
                catch (System.Exception)
                {
                    // Handle parsing errors silently
                }
            }

            return glassInfo;
        }

        #endregion
    }
}