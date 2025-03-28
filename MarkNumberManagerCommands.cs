using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using System.Text;
using Newtonsoft.Json;

namespace TakeoffBridge
{
    // This should be a separate class file
    public class MarkNumberManagerCommands
    {
        [CommandMethod("GENERATEMARKNUMBERS")]
        public void GenerateMarkNumbers()
        {
            // This calls the static method in MarkNumberManager
            MarkNumberManager.GenerateAllMarkNumbers();
        }

        [CommandMethod("REGENERATEALLMARKNUMBERS")]
        public void RegenerateAllMarkNumbers()
        {
            // This calls the static method in MarkNumberManager
            MarkNumberManager.RegenerateAllMarkNumbers();
        }


        [CommandMethod("MARKNUMBERREPORT")]
        public void GenerateMarkNumberReport()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            // Get the mark number manager
            var manager = GlassTakeoffBridge.GlassTakeoffApp.MarkNumberManager;

            // Generate a simple report
            ed.WriteMessage("\n--- Mark Number Report ---");

            // Get counts from the manager (you'll need to add these methods)
            // int horizontalCount = manager.GetHorizontalPartCount();
            // int verticalCount = manager.GetVerticalPartCount();

            ed.WriteMessage($"\nHorizontal Mark Count: {manager.GetHorizontalPartCount()}");
            ed.WriteMessage($"\nVertical Mark Count: {manager.GetVerticalPartCount()}");

            ed.WriteMessage("\n--- End of Report ---");
        }

        [CommandMethod("MARKSELECTED")]
        public void MarkSelected()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;
            Database db = doc.Database;

            // Prompt for selection
            PromptEntityResult result = ed.GetEntity("\nSelect component to mark: ");
            if (result.Status != PromptStatus.OK)
                return;

            // Process the selected component
            using (Transaction tr = db.TransactionManager.StartTransaction())
            {
                MarkNumberManager.Instance.ProcessComponentMarkNumbers(result.ObjectId, tr);
                tr.Commit();
            }

            ed.WriteMessage("\nComponent marked.");
        }

    }
}