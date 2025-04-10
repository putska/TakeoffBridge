// TemplateGeneratorCommands.cs
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;

namespace TakeoffBridge
{
    public class TemplateGeneratorCommands
    {
        [CommandMethod("GENERATETEMPLATES")]
        public void GenerateTemplates()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor editor = doc.Editor;

            try
            {
                TemplateGenerator generator = new TemplateGenerator();
                generator.GenerateTemplatesFromDrawing();

                editor.WriteMessage("\nTemplate generation complete.");
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage($"\nError generating templates: {ex.Message}");
            }
        }
    }
}