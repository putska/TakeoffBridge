using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.AutoCAD.ApplicationServices;
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

            // Generate file path
            string filePath = Path.Combine(RootPath, "MetalExport.xlsx");

            try
            {
                // Create Excel file
                CreateExcelFile(filePath, groupedParts);
                ed.WriteMessage($"\nExcel file created successfully: {filePath}");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError creating Excel file: {ex.Message}");
                throw;
            }
        }

        /// <summary>
        /// Creates an Excel file with the parts data
        /// </summary>
        private void CreateExcelFile(string filePath, IEnumerable<dynamic> parts)
        {
            Excel.Application excelApp = null;
            Excel.Workbook workbook = null;
            Excel.Worksheet worksheet = null;

            try
            {
                // Create a new Excel application
                excelApp = new Excel.Application();
                excelApp.Visible = false;
                excelApp.DisplayAlerts = false;

                // Create a new workbook
                workbook = excelApp.Workbooks.Add(Type.Missing);

                // Get the first worksheet
                worksheet = (Excel.Worksheet)workbook.Sheets[1];
                worksheet.Name = "Sheet1";

                // Add headers to the first row
                worksheet.Cells[1, 1] = "Qty";
                worksheet.Cells[1, 2] = "Part No";
                worksheet.Cells[1, 3] = "Length";
                worksheet.Cells[1, 4] = "Mark No";
                worksheet.Cells[1, 5] = "Finish";
                worksheet.Cells[1, 6] = "Fab";

                // Format the header row
                Excel.Range headerRange = worksheet.Range["A1:F1"];
                headerRange.Font.Bold = true;
                headerRange.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                headerRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

                // Add data rows
                int row = 2;
                foreach (var part in parts)
                {
                    worksheet.Cells[row, 1] = part.Quantity;
                    worksheet.Cells[row, 2] = part.PartNumber;

                    // Format the length as needed
                    string formattedLength = FormatInches(part.Length);
                    worksheet.Cells[row, 3] = formattedLength;

                    worksheet.Cells[row, 4] = part.MarkNumber;
                    worksheet.Cells[row, 5] = part.Finish;
                    worksheet.Cells[row, 6] = part.Fab;

                    row++;
                }

                // Auto-fit columns for better readability
                worksheet.Columns.AutoFit();

                // Apply borders to all data
                Excel.Range dataRange = worksheet.Range[$"A1:F{row - 1}"];
                dataRange.Borders.LineStyle = Excel.XlLineStyle.xlContinuous;

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

                if (worksheet != null)
                {
                    Marshal.ReleaseComObject(worksheet);
                }

                // Force garbage collection
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }
        }

        #endregion
    }
}