using Autodesk.AutoCAD.DatabaseServices;
using System;
using System.Collections.Generic;

namespace TakeoffBridge
{
    /// <summary>
    /// Represents a child part of a metal component
    /// </summary>
    public class ChildPart
    {
        public string Name { get; set; }
        public string PartType { get; set; }

        // Original length adjustment (keeping for backward compatibility)
        public double LengthAdjustment { get; set; }
        public bool IsShopUse { get; set; }

        // End-specific adjustments
        public double StartAdjustment { get; set; } // Left for horizontal, Bottom for vertical
        public double EndAdjustment { get; set; }   // Right for horizontal, Top for vertical

        // Fixed length properties
        public bool IsFixedLength { get; set; }
        public double FixedLength { get; set; }

        public string MarkNumber { get; set; }
        public string Material { get; set; }

        // Attachment properties
        public string Attach { get; set; } // "L", "R", or null/empty
        public bool Invert { get; set; }   // Whether the part should be inverted
        public double Adjust { get; set; } // Vertical adjustment (inches)
        public bool Clips { get; set; }    // For vertical parts - whether clips attach to this part

        public string Finish { get; set; } // The finish of the part
        public string Fab { get; set; }    // The fabrication of the part

        // NEW PROPERTY: List of attachments
        public List<PartAttachment> Attachments { get; set; } = new List<PartAttachment>();

        // Default constructor for JSON deserialization
        public ChildPart()
        {
            // Initialize default values
            Name = "";
            PartType = "";
            LengthAdjustment = 0.0;
            StartAdjustment = 0.0;
            EndAdjustment = 0.0;
            IsFixedLength = false;
            FixedLength = 0.0;
            Material = "";
            MarkNumber = "";
            Attach = "";
            Invert = false;
            Adjust = 0.0;
            Clips = false;
            Finish = "Paint";
            Fab = "1";
            // Attachments already initialized by property initializer
        }

        public ChildPart(string name, string partType, double lengthAdjustment, string material)
        {
            Name = name;
            PartType = partType;
            LengthAdjustment = lengthAdjustment;
            Material = material;
            IsShopUse = false;
            MarkNumber = ""; // Will be assigned later

            // Initialize new end-specific adjustments
            // Initially set them to half the total adjustment on each end
            StartAdjustment = lengthAdjustment / 2;
            EndAdjustment = lengthAdjustment / 2;

            // Default to non-fixed length
            IsFixedLength = false;
            FixedLength = 0.0;

            // Initialize attachment properties with defaults
            Attach = "";
            Invert = false;
            Adjust = 0.0;
            Clips = false;

            Finish = "Paint";
            Fab = "1";
            // Attachments already initialized by property initializer
        }

        // Additional constructor for direct end adjustments
        public ChildPart(string name, string partType, double startAdjustment, double endAdjustment, string material)
        {
            Name = name;
            PartType = partType;
            IsShopUse = false;
            StartAdjustment = startAdjustment;
            EndAdjustment = endAdjustment;
            LengthAdjustment = startAdjustment + endAdjustment; // For backward compatibility
            Material = material;
            MarkNumber = "";

            // Default to non-fixed length
            IsFixedLength = false;
            FixedLength = 0.0;

            // Initialize attachment properties with defaults
            Attach = "";
            Invert = false;
            Adjust = 0.0;
            Clips = false;

            Finish = "Paint";
            Fab = "1";
            // Attachments already initialized by property initializer
        }

        // Constructor for fixed length parts
        public ChildPart(string name, string partType, double fixedLength, string material, bool isFixed)
        {
            if (!isFixed) throw new ArgumentException("This constructor should only be used for fixed length parts");

            Name = name;
            PartType = partType;
            IsShopUse = false;
            FixedLength = fixedLength;
            IsFixedLength = true;
            Material = material;
            MarkNumber = "";

            // Set adjustments to 0 for fixed length parts
            StartAdjustment = 0;
            EndAdjustment = 0;
            LengthAdjustment = 0;

            // Initialize attachment properties with defaults
            Attach = "";
            Invert = false;
            Adjust = 0.0;
            Clips = false;

            Finish = "Paint";
            Fab = "1";
            // Attachments already initialized by property initializer
        }

        // Calculate the actual length based on parent component length
        public double GetActualLength(double parentLength)
        {
            // If fixed length, just return that value
            if (IsFixedLength)
            {
                return FixedLength;
            }

            // Otherwise use the sum of both adjustments to get the total length adjustment
            return parentLength + StartAdjustment + EndAdjustment;
        }

        // Calculate the actual start point offset
        public double GetStartOffset()
        {
            return StartAdjustment;
        }

        // Calculate the actual end point offset
        public double GetEndOffset()
        {
            return EndAdjustment;
        }
    }

    /// <summary>
    /// Represents an attachment between components
    /// </summary>
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

    /// <summary>
    /// Represents a part-level attachment information
    /// </summary>
    public class PartAttachment
    {
        public string Side { get; set; }
        public double Position { get; set; }
        public double Height { get; set; }
        public bool Invert { get; set; }
        public double Adjust { get; set; }
        public string AttachedPartNumber { get; set; }
        public string AttachedPartType { get; set; }
        public string AttachedFab { get; set; }
    }

    
}