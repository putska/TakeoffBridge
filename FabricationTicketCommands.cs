using Autodesk.AutoCAD.Runtime;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.EditorInput;

namespace TakeoffBridge
{
    public class FabricationTicketCommands
    {
        [CommandMethod("DRAWFABNEW")]
        public void GenerateFabricationTickets()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor ed = doc.Editor;

            try
            {
                ed.WriteMessage("\nGenerating fabrication tickets...");

                // Create a new instance of the fabrication manager
                FabricationManager manager = new FabricationManager();

                // Process the drawing and generate tickets
                manager.ProcessDrawingAndGenerateTickets();

                ed.WriteMessage("\nFabrication ticket generation complete.");
            }
            catch (System.Exception ex)
            {
                ed.WriteMessage($"\nError: {ex.Message}");
                if (ex.InnerException != null)
                {
                    ed.WriteMessage($"\nInner exception: {ex.InnerException.Message}");
                }
            }
        }
    }
}