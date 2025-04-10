using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Runtime;
using System;
using System.Collections.Generic;
using System.IO;

namespace TakeoffBridge
{
    public class FabricationTicketCommands
    {
        /// <summary>
        /// Command to generate both templates and fabrication tickets in one operation
        /// </summary>
        [CommandMethod("GENERATEALLFABRICATION")]
        public void GenerateAllFabrication()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor editor = doc.Editor;

            try
            {
                // First step - generate templates
                editor.WriteMessage("\n=== Generating Templates ===");
                TemplateGenerator templateGen = new TemplateGenerator();
                templateGen.GenerateTemplatesFromDrawing();

                // Second step - generate fabrication tickets
                editor.WriteMessage("\n=== Generating Fabrication Tickets ===");
                FabTicketGenerator ticketGen = new FabTicketGenerator();
                ticketGen.GenerateTicketsFromDrawing();

                editor.WriteMessage("\n=== All Fabrication Generation Complete ===");
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage($"\nError in GenerateAllFabrication: {ex.Message}");
            }
        }

        /// <summary>
        /// Command to set up the Takeoff directory structure if it doesn't exist
        /// </summary>
        [CommandMethod("SETUPTAKEOFFDIRS")]
        public void SetupTakeoffDirectories()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor editor = doc.Editor;

            string baseDir = @"C:\CSE\Takeoff";
            string[] subDirs = new string[] {
                "fabs",
                "fabs\\dies",
                "fabs\\tickets",
                "parts"
            };

            try
            {
                // Create base directory if it doesn't exist
                if (!Directory.Exists(baseDir))
                {
                    Directory.CreateDirectory(baseDir);
                    editor.WriteMessage($"\nCreated base directory: {baseDir}");
                }

                // Create subdirectories
                foreach (string subDir in subDirs)
                {
                    string fullPath = Path.Combine(baseDir, subDir);
                    if (!Directory.Exists(fullPath))
                    {
                        Directory.CreateDirectory(fullPath);
                        editor.WriteMessage($"\nCreated directory: {fullPath}");
                    }
                }

                // Check for template file
                string templatePath = Path.Combine(baseDir, "fabs\\dies\\fabtemplate.dwg");
                if (!File.Exists(templatePath))
                {
                    editor.WriteMessage($"\nWARNING: Base template file not found at {templatePath}");
                    editor.WriteMessage("\nPlease create a base template file before generating templates.");
                }

                editor.WriteMessage("\nTakeoff directory setup complete.");
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage($"\nError in SetupTakeoffDirectories: {ex.Message}");
            }
        }

        [CommandMethod("GENERATEFABTICKETS")]
        public void GenerateFabTickets()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor editor = doc.Editor;

            try
            {
                FabTicketGenerator generator = new FabTicketGenerator();
                generator.GenerateTicketsFromDrawing();

                editor.WriteMessage("\nFabrication ticket generation complete.");
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage($"\nError generating fabrication tickets: {ex.Message}");
            }
        }

        /// <summary>
        /// Command to test template generation with a simple part
        /// </summary>
        [CommandMethod("TESTTEMPLATEGENERATOR")]
        public void TestTemplateGenerator()
        {
            Document doc = Application.DocumentManager.MdiActiveDocument;
            Editor editor = doc.Editor;

            try
            {
                // Create a manual test template
                string partNumber = "601TVG005";
                string fab = "1";

                TemplateGenerator.PartTemplateInfo testTemplate = new TemplateGenerator.PartTemplateInfo
                {
                    PartNumber = partNumber,
                    PartType = "Vertical",
                    Fab = fab,
                    Finish = "Mill",
                    IsVertical = true
                };

                // Generate a template
                PartDrawingGenerator generator = new PartDrawingGenerator();
                bool success = generator.DrawPart(
                    partNumber,
                    Path.Combine(@"C:\CSE\Takeoff\fabs\dies", $"{partNumber}-{fab}.dwg"),
                    30.0,
                    45,
                    90,
                    45,
                    90,
                    PartDrawingGenerator.MirrorType.VerticalMullion,
                    false,
                    null,
                    0
                );

                if (success)
                {
                    editor.WriteMessage($"\nSuccessfully created test template for {partNumber}-{fab}");
                }
                else
                {
                    editor.WriteMessage($"\nFailed to create test template");
                }
            }
            catch (System.Exception ex)
            {
                editor.WriteMessage($"\nError in TestTemplateGenerator: {ex.Message}");
            }
        }
    }
}