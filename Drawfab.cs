﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.EditorInput;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.GraphicsInterface;

namespace TakeoffBridge
{
    public class Drawfab
    {
        private const string WorkPointAppName = "WORKPOINT";
        private double _width;
        private double _height;
        private string _partNumber;

        public Drawfab()
        {
            _width = 0;
            _height = 0;
            _partNumber = string.Empty;
        }

        public Drawfab(string partNumber)
        {
            _partNumber = partNumber;
        }

        public double Width => _width;
        public double Height => _height;
        public string PartNumber => _partNumber;

        /// <summary>
        /// Imports a drawing file as a block in the current database
        /// </summary>
        private BlockTableRecord ImportDrawingAsBlock(Database database, string partDrawingPath, string blockName, Transaction trans)
        {
            // Check if block already exists
            var bt = database.GetBlockTable(trans);

            if (!bt.Has(blockName))
            {
                // Import the drawing file
                bt.UpgradeOpen();

                using (var tempDb = new Database(false, true))
                {
                    tempDb.ReadDwgFile(partDrawingPath, FileOpenMode.OpenForReadAndAllShare, true, null);

                    // Insert as block
                    var blockId = database.Insert(blockName, tempDb, false);
                    bt.DowngradeOpen();

                    return trans.GetObject(blockId, OpenMode.ForRead) as BlockTableRecord;
                }
            }

            return trans.GetObject(bt[blockName], OpenMode.ForRead) as BlockTableRecord;
        }

        /// <summary>
        /// Gets the workpoint from an entity
        /// </summary>
        private Point3d? GetWorkPoint(Entity entity)
        {
            using (ResultBuffer rb = entity.GetXDataForApplication(WorkPointAppName))
            {
                if (rb != null)
                {
                    TypedValue[] values = rb.AsArray();

                    if (values.Length > 1 && values[1].TypeCode == (int)DxfCode.ExtendedDataWorldXCoordinate)
                    {
                        return (Point3d)values[1].Value;
                    }
                }
            }

            return null;
        }

        /// <summary>
        /// Sets a workpoint on an entity
        /// </summary>
        private void SetWorkPoint(Point3d workPoint, Transaction trans, Entity entity)
        {
            // Register application name if needed
            RegisterAppName(trans, WorkPointAppName);

            // Set XData with workpoint
            entity.XData = new ResultBuffer(
                new TypedValue((int)DxfCode.ExtendedDataRegAppName, WorkPointAppName),
                new TypedValue((int)DxfCode.ExtendedDataWorldXCoordinate, workPoint)
            );
        }

        /// <summary>
        /// Registers an application name for XData if not already registered
        /// </summary>
        private void RegisterAppName(Transaction trans, string appName)
        {
            Database db = Application.DocumentManager.MdiActiveDocument.Database;
            using (RegAppTable regTable = (RegAppTable)trans.GetObject(db.RegAppTableId, OpenMode.ForRead))
            {
                if (!regTable.Has(appName))
                {
                    regTable.UpgradeOpen();
                    using (RegAppTableRecord regApp = new RegAppTableRecord())
                    {
                        regApp.Name = appName;
                        regTable.Add(regApp);
                        trans.AddNewlyCreatedDBObject(regApp, true);
                    }
                }
            }
        }

        /// <summary>
        /// Extracts regions from a block and handles polyline conversion
        /// </summary>
        public List<Region> GetExtrusionRegionsFromBlock(Database database, ObjectId blockId, Transaction trans)
        {
            List<Region> finalRegions = new List<Region>();
            Point3d? workPoint = null;

            try
            {
                // Get block table record
                BlockTableRecord btr = trans.GetObject(blockId, OpenMode.ForRead) as BlockTableRecord;
                if (btr == null) return null;

                // First, collect all region IDs
                List<ObjectId> regionIds = new List<ObjectId>();
                foreach (ObjectId objId in btr)
                {
                    if (objId.ObjectClass.DxfName == "REGION")
                    {
                        regionIds.Add(objId);
                    }
                }

                // Process existing regions
                if (regionIds.Count > 0)
                {
                    foreach (ObjectId regionId in regionIds)
                    {
                        Region region = trans.GetObject(regionId, OpenMode.ForRead) as Region;
                        if (region != null)
                        {
                            // Check for workpoint
                            using (ResultBuffer rb = region.GetXDataForApplication(WorkPointAppName))
                            {
                                if (rb != null)
                                {
                                    TypedValue[] values = rb.AsArray();
                                    if (values.Length > 1 && values[1].TypeCode == (int)DxfCode.ExtendedDataWorldXCoordinate)
                                    {
                                        workPoint = (Point3d)values[1].Value;
                                    }
                                }
                            }

                            // Create a new region instead of cloning
                            Region newRegion = new Region();
                            newRegion.CopyFrom(region);
                            finalRegions.Add(newRegion);
                        }
                    }
                }
                else
                {
                    // Collect all curve entities
                    DBObjectCollection curves = new DBObjectCollection();

                    foreach (ObjectId objId in btr)
                    {
                        if (objId.ObjectClass.DxfName == "LWPOLYLINE" ||
                            objId.ObjectClass.DxfName == "POLYLINE" ||
                            objId.ObjectClass.DxfName == "LINE" ||
                            objId.ObjectClass.DxfName == "ARC" ||
                            objId.ObjectClass.DxfName == "CIRCLE" ||
                            objId.ObjectClass.DxfName == "ELLIPSE" ||
                            objId.ObjectClass.DxfName == "SPLINE")
                        {
                            Curve curve = trans.GetObject(objId, OpenMode.ForRead) as Curve;
                            if (curve != null && !curve.IsErased)
                            {
                                // Check for workpoint
                                using (ResultBuffer rb = curve.GetXDataForApplication(WorkPointAppName))
                                {
                                    if (rb != null)
                                    {
                                        TypedValue[] values = rb.AsArray();
                                        if (values.Length > 1 && values[1].TypeCode == (int)DxfCode.ExtendedDataWorldXCoordinate)
                                        {
                                            workPoint = (Point3d)values[1].Value;
                                        }
                                    }
                                }

                                // Create a new curve instead of cloning
                                Curve newCurve = curve.GetType().GetConstructor(Type.EmptyTypes).Invoke(null) as Curve;
                                newCurve.CopyFrom(curve);
                                curves.Add(newCurve);
                            }
                        }
                    }

                    // Create regions from curves
                    if (curves.Count > 0)
                    {
                        try
                        {
                            DBObjectCollection newRegions = Region.CreateFromCurves(curves);
                            foreach (DBObject obj in newRegions)
                            {
                                Region region = obj as Region;
                                if (region != null)
                                {
                                    finalRegions.Add(region);
                                }
                            }
                        }
                        catch (Exception)
                        {
                            throw new Exception("Failed to create regions from curves. Make sure all polylines are closed.");
                        }
                        finally
                        {
                            // Clean up curves
                            foreach (DBObject obj in curves)
                            {
                                obj.Dispose();
                            }
                        }
                    }
                }

                // Process regions for Boolean operations if needed
                if (finalRegions.Count > 1)
                {
                    // Find which regions are contained within other regions
                    List<Region> processedRegions = new List<Region>();

                    // Sort regions by area (largest first)
                    finalRegions.Sort((a, b) => b.Area.CompareTo(a.Area));

                    for (int i = 0; i < finalRegions.Count; i++)
                    {
                        Region currentRegion = finalRegions[i];
                        Extents3d currentExtents = currentRegion.GeometricExtents;
                        bool isContained = false;
                        Region containerRegion = null;

                        // Check if this region is contained within any processed region
                        foreach (Region processedRegion in processedRegions)
                        {
                            Extents3d processedExtents = processedRegion.GeometricExtents;

                            // If the current region's extents are fully contained within a processed region's extents
                            if (currentExtents.MinPoint.X >= processedExtents.MinPoint.X &&
                                currentExtents.MinPoint.Y >= processedExtents.MinPoint.Y &&
                                currentExtents.MaxPoint.X <= processedExtents.MaxPoint.X &&
                                currentExtents.MaxPoint.Y <= processedExtents.MaxPoint.Y)
                            {
                                // Perform more precise containment check using region relationship
                                try
                                {
                                    // If one region is fully inside another, subtract it
                                    containerRegion = processedRegion;
                                    isContained = true;
                                    break;
                                }
                                catch
                                {
                                    // If Boolean check fails, assume they're separate regions
                                    isContained = false;
                                }
                            }
                        }

                        if (isContained && containerRegion != null)
                        {
                            // This region is contained within another region, subtract it
                            containerRegion.BooleanOperation(BooleanOperationType.BoolSubtract, currentRegion);
                            currentRegion.Dispose();
                        }
                        else
                        {
                            // This is a separate region, add it to the processed list
                            processedRegions.Add(currentRegion);
                        }
                    }

                    // Update the final regions list
                    finalRegions = processedRegions;
                }

                // Set workpoint on regions if available
                if (workPoint.HasValue)
                {
                    foreach (Region region in finalRegions)
                    {
                        // Register application name if needed
                        RegisterAppName(trans, WorkPointAppName);

                        // Set XData with workpoint
                        region.XData = new ResultBuffer(
                            new TypedValue((int)DxfCode.ExtendedDataRegAppName, WorkPointAppName),
                            new TypedValue((int)DxfCode.ExtendedDataWorldXCoordinate, workPoint.Value)
                        );
                    }
                }

                return finalRegions;
            }
            catch (Exception ex)
            {
                // Clean up resources in case of error
                foreach (Region region in finalRegions)
                {
                    region.Dispose();
                }
                throw new Exception("Error creating regions: " + ex.Message, ex);
            }
        }

        /// <summary>
        /// Transform regions to standard orientation
        /// </summary>
        private void TransformRegionsToStandardOrientation(List<Region> regions, Matrix3d handedTransformation, bool preserveOriginalOrientation = false)
        {
            // Calculate extents for dimensions
            var extents = new Extents3d();
            foreach (var region in regions)
            {
                extents.AddExtents(region.GeometricExtents);
            }

            // Set dimensions
            _width = extents.MaxPoint.X - extents.MinPoint.X;
            _height = extents.MaxPoint.Y - extents.MinPoint.Y;

            // Move to origin
            var transformation = Matrix3d.Displacement(Point3d.Origin -
                new Point3d(extents.MinPoint.X, extents.MinPoint.Y, extents.MinPoint.Z));

            if (!preserveOriginalOrientation)
            {
                // Standard mode: orient for extrusion along X axis
                // In standard orientation, the part is oriented for extrusion along X
                // with its face in the YZ plane at the origin
                transformation = Matrix3d.Rotation(Math.PI / 2, Vector3d.YAxis, Point3d.Origin) * transformation;
                // Compensate for the -90 degree rotation by rotating around X axis
                transformation = Matrix3d.Rotation(Math.PI / 2, Vector3d.XAxis, Point3d.Origin) * transformation;
            }
            else
            {
                // Preservation mode: keep original orientation but prepare for extrusion
                // When preserving orientation, part should be extruded along Z axis
                // with its face in the XY plane at the origin
            }

            // Apply handed transformation 
            transformation = handedTransformation * transformation;

            // Apply transformation to regions
            foreach (var region in regions)
            {
                region.TransformBy(transformation);
            }
        }

        /// <summary>
        /// Extrudes regions to create solids
        /// </summary>
        private List<Solid3d> ExtrudeRegions(List<Region> regions, double extrusionLength, Database database, Transaction trans, bool preserveOriginalOrientation = false)
        {
            var solids = new List<Solid3d>();
            var sweepOpts = new SweepOptions();
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;

            try
            {
                // Add a small extension to ensure clean cuts
                double extension = 0.1;
                double adjustedLength = extrusionLength + extension;

                // Choose direction vector based on orientation
                Vector3d extrusionDirection = preserveOriginalOrientation ?
                    Vector3d.ZAxis : Vector3d.XAxis;

                foreach (var region in regions)
                {
                    try
                    {
                        Solid3d solid = new Solid3d();
                        solid.SetDatabaseDefaults(database);
                        solid.RecordHistory = false;

                        ed.WriteMessage($"\nExtruding region with area: {region.Area} to length: {adjustedLength}");
                        solid.CreateExtrudedSolid(region, adjustedLength * extrusionDirection, sweepOpts);

                        // Transfer workpoint data if available
                        Point3d? workPoint = GetWorkPoint(region);
                        if (workPoint.HasValue)
                        {
                            SetWorkPoint(workPoint.Value, trans, solid);
                        }

                        solids.Add(solid);
                    }
                    catch (Exception ex)
                    {
                        ed.WriteMessage($"\nError extruding region: {ex.Message}");
                        // Continue with other regions
                    }
                }

                if (solids.Count == 0)
                {
                    throw new Exception("Failed to create any solids from regions");
                }

                return solids;
            }
            catch (Exception ex)
            {
                // Clean up any created solids
                foreach (var solid in solids)
                {
                    solid.Dispose();
                }

                throw new Exception($"Error in ExtrudeRegions: {ex.Message}", ex);
            }
        }

        /// <summary>
        /// Applies miter and tilt cuts to solids
        /// </summary>
        private void ApplyMiterAndTiltCuts(List<Solid3d> solids, double miterLeft, double tiltLeft,
                                   double miterRight, double tiltRight, double length,
                                   bool preserveOriginalOrientation = false, bool visualizeOnly = false)
        {
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;
            ed.WriteMessage($"\nApplying cuts - Left: {miterLeft}°/{tiltLeft}°, Right: {miterRight}°/{tiltRight}°");

            try
            {
                // Calculate part dimensions
                Extents3d extents = new Extents3d();
                foreach (var solid in solids)
                {
                    extents.AddExtents(solid.GeometricExtents);
                }

                double partWidth = extents.MaxPoint.X - extents.MinPoint.X;   // Length in X
                double partDepth = extents.MaxPoint.Y - extents.MinPoint.Y;   // Depth in Y
                double partHeight = extents.MaxPoint.Z - extents.MinPoint.Z;  // Height in Z

                ed.WriteMessage($"\nPart dimensions: Width={partWidth}, Depth={partDepth}, Height={partHeight}");

                // LEFT END CUTTING - Unchanged from previous version
                if (miterLeft != 90 || tiltLeft != 90)
                {
                    ed.WriteMessage("\nProcessing left cut...");

                    // Calculate cutting plane normal
                    double mRad = (miterLeft - 90).ToRadians();
                    double tRad = (tiltLeft - 90).ToRadians();

                    // Normal calculation
                    double nx = Math.Cos(tRad) * Math.Cos(mRad);
                    double ny = Math.Cos(tRad) * Math.Sin(mRad);
                    double nz = Math.Sin(tRad);

                    Vector3d normal = new Vector3d(nx, ny, nz);
                    double len = Math.Sqrt(nx * nx + ny * ny + nz * nz);
                    normal = new Vector3d(nx / len, ny / len, nz / len);

                    // Determine origin point based on angles
                    Point3d planeOrigin = Point3d.Origin;

                    // 1. For miter > 90 (left side), move right
                    if (miterLeft > 90)
                    {
                        // Calculate offset based on miter angle and part depth
                        double miterAngleFromVertical = Math.PI / 2 - mRad; // Angle from Y axis
                        double offsetX = partDepth / Math.Tan(miterAngleFromVertical);

                        // Ensure offset is positive for right movement
                        offsetX = Math.Abs(offsetX);

                        planeOrigin = new Point3d(offsetX, 0, 0);
                        ed.WriteMessage($"\nMoving left cut plane right to X={offsetX} for miter > 90");
                    }

                    // 2. For tilt > 90 (left side), move right
                    if (tiltLeft > 90)
                    {
                        // Calculate offset based on tilt angle and part height
                        double tiltAngleFromVertical = Math.PI / 2 - tRad; // Angle from Z axis
                        double offsetX = partHeight / Math.Tan(tiltAngleFromVertical);

                        // Ensure offset is positive for right movement
                        offsetX = Math.Abs(offsetX);

                        // Use this offset if it's larger than the miter offset
                        if (miterLeft <= 90 || offsetX > planeOrigin.X)
                        {
                            planeOrigin = new Point3d(offsetX, 0, 0);
                            ed.WriteMessage($"\nMoving left cut plane right to X={offsetX} for tilt > 90");
                        }
                    }

                    // Create the cutting plane
                    Plane leftCutPlane = new Plane(planeOrigin, normal);

                    ed.WriteMessage($"\nLeft cut normal: {normal.X}, {normal.Y}, {normal.Z}");
                    ed.WriteMessage($"\nLeft cut plane origin: {planeOrigin.X}, {planeOrigin.Y}, {planeOrigin.Z}");

                    // Apply cutting operation
                    foreach (var solid in solids)
                    {
                        try
                        {
                            solid.Slice(leftCutPlane);
                            ed.WriteMessage("\nLeft cut successful");
                        }
                        catch (Exception ex)
                        {
                            ed.WriteMessage($"\nError on left cut: {ex.Message}");
                        }
                    }
                }

                // RIGHT END CUTTING - MODIFIED for mirror behavior
                if (miterRight != 90 || tiltRight != 90)
                {
                    ed.WriteMessage("\nProcessing right cut...");

                    // MIRROR LOGIC: Convert right angle to equivalent left angle
                    // For a mirror image, 45° right should behave like 135° left (180 - 45)
                    double mirroredMiter = 180 - miterRight;

                    // Calculate using the mirrored angle
                    double mRad = (mirroredMiter - 90).ToRadians();
                    double tRad = (tiltRight - 90).ToRadians();

                    // Normal calculation - now using mirroredMiter
                    double nx = -Math.Cos(tRad) * Math.Cos(mRad);  // Negated for right side
                    double ny = -Math.Cos(tRad) * Math.Sin(mRad);  // Negated for right side with mirrored angle
                    double nz = Math.Sin(tRad);

                    Vector3d normal = new Vector3d(nx, ny, nz);
                    double len = Math.Sqrt(nx * nx + ny * ny + nz * nz);
                    normal = new Vector3d(nx / len, ny / len, nz / len);

                    // Start at right end of part
                    Point3d planeOrigin = new Point3d(length, 0, 0);

                    // RIGHT SIDE LOGIC with mirrored angles

                    // SIMPLIFIED RIGHT MITER LOGIC - direct approach
                    // For 45 degree miter (original), keep plane at end of part
                    if (miterRight == 45)
                    {
                        // No offset needed - stay at end of part
                        planeOrigin = new Point3d(length, 0, 0);
                        ed.WriteMessage($"\nKeeping right cut plane at X={length} for 45° miter");
                    }
                    // For 135 degree miter (original), move left by depth
                    else if (miterRight == 135)
                    {
                        // Move left by depth
                        planeOrigin = new Point3d(length - partDepth, 0, 0);
                        ed.WriteMessage($"\nMoving right cut plane left to X={length - partDepth} for 135° miter");
                    }
                    // Handle any other miter angles
                    else if (miterRight != 90)
                    {
                        // For larger than 90, move left
                        if (miterRight > 90)
                        {
                            double offsetX = partDepth * Math.Tan((miterRight - 90).ToRadians());
                            offsetX = Math.Abs(offsetX);
                            planeOrigin = new Point3d(length - offsetX, 0, 0);
                            ed.WriteMessage($"\nMoving right cut plane left to X={length - offsetX} for miter > 90");
                        }
                        // For less than 90, no movement needed
                        else
                        {
                            planeOrigin = new Point3d(length, 0, 0);
                            ed.WriteMessage($"\nKeeping right cut plane at X={length} for miter < 90");
                        }
                    }

                    // For tilt > 90 (right side), move left
                    if (tiltRight > 90)
                    {
                        // Calculate offset based on tilt angle and part height
                        double tiltAngleFromVertical = Math.PI / 2 - tRad; // Angle from Z axis
                        double offsetX = partHeight / Math.Tan(tiltAngleFromVertical);

                        // Ensure offset is positive
                        offsetX = Math.Abs(offsetX);

                        // Use this offset if it requires more movement than the miter
                        if ((mirroredMiter <= 90 && length - offsetX < planeOrigin.X) ||
                            (mirroredMiter > 90 && length - offsetX < planeOrigin.X))
                        {
                            planeOrigin = new Point3d(length - offsetX, 0, 0);
                            ed.WriteMessage($"\nMoving right cut plane LEFT to X={length - offsetX} for tilt > 90");
                        }
                    }

                    // Create the cutting plane
                    Plane rightCutPlane = new Plane(planeOrigin, normal);

                    ed.WriteMessage($"\nRight cut normal: {normal.X}, {normal.Y}, {normal.Z}");
                    ed.WriteMessage($"\nRight cut plane origin: {planeOrigin.X}, {planeOrigin.Y}, {planeOrigin.Z}");

                    // Apply cutting operation
                    foreach (var solid in solids)
                    {
                        try
                        {
                            solid.Slice(rightCutPlane);
                            ed.WriteMessage("\nRight cut successful");
                        }
                        catch (Exception ex)
                        {
                            ed.WriteMessage($"\nError on right cut: {ex.Message}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                ed.WriteMessage($"\nGeneral error in ApplyMiterAndTiltCuts: {ex.Message}");
            }
        }



        /// <summary>
        /// Creates a handed transformation matrix based on part characteristics
        /// </summary>
        private Matrix3d GetHandedTransformation(bool handed, string handedSide)
        {
            if (handed && !string.IsNullOrEmpty(handedSide) && handedSide.ToUpper() != "L")
            {
                // Mirror for right-handed parts
                return Matrix3d.Mirroring(new Plane(Point3d.Origin, Vector3d.YAxis));
            }

            return Matrix3d.Identity;
        }

        /// <summary>
        /// Main method to create an extruded part
        /// </summary>
        public List<Solid3d> CreateExtrudedPart(
            Database database,
            string partDrawingPath,
            string blockName,
            double length,
            double miterLeft = 90,
            double tiltLeft = 90,
            double miterRight = 90,
            double tiltRight = 90,
            bool handed = false,
            string handedSide = "")
        {
            return CreateExtrudedPart(database, partDrawingPath, blockName, length,
                miterLeft, tiltLeft, miterRight, tiltRight,
                handed, handedSide, Matrix3d.Identity);
        }

        public List<Solid3d> CreateExtrudedPart(
            Database database,
            string partDrawingPath,
            string blockName,
            double length,
            double miterLeft = 90,
            double tiltLeft = 90,
            double miterRight = 90,
            double tiltRight = 90,
            bool handed = false,
            string handedSide = "",
            bool preserveOriginalOrientation = false,
            bool visualizeOnly = false)
        {
            return CreateExtrudedPart(database, partDrawingPath, blockName, length,
                miterLeft, tiltLeft, miterRight, tiltRight,
                handed, handedSide, Matrix3d.Identity, preserveOriginalOrientation, visualizeOnly);
        }

        public List<Solid3d> CreateExtrudedPart(
            Database database,
            string partDrawingPath,
            string blockName,
            double length,
            double miterLeft,
            double tiltLeft,
            double miterRight,
            double tiltRight,
            bool handed,
            string handedSide,
            Matrix3d finalTransform,
            bool preserveOriginalOrientation = false,
            bool visualizeOnly = false)
        {
            List<Solid3d> resultSolids = new List<Solid3d>();
            Editor ed = Application.DocumentManager.MdiActiveDocument.Editor;

            using (var trans = database.TransactionManager.StartTransaction())
            {
                try
                {
                    // Import drawing as block
                    ed.WriteMessage("\nImporting drawing as block...");
                    var btr = ImportDrawingAsBlock(database, partDrawingPath, blockName, trans);
                    if (btr == null)
                    {
                        throw new Exception($"Failed to import drawing from {partDrawingPath}");
                    }

                    // Extract regions from block
                    ed.WriteMessage("\nExtracting regions from block...");
                    var regions = GetExtrusionRegionsFromBlock(database, btr.ObjectId, trans);
                    if (regions == null || regions.Count == 0)
                    {
                        throw new Exception("No valid regions found in the drawing");
                    }
                    ed.WriteMessage($"\nFound {regions.Count} region(s)");

                    // Transform regions to orientation
                    ed.WriteMessage("\nTransforming regions...");
                    var handedTransformation = GetHandedTransformation(handed, handedSide);
                    TransformRegionsToStandardOrientation(regions, handedTransformation, preserveOriginalOrientation);

                    // Extrude regions to create solids
                    ed.WriteMessage("\nExtruding regions...");
                    resultSolids = ExtrudeRegions(regions, length, database, trans, preserveOriginalOrientation);
                    ed.WriteMessage($"\nCreated {resultSolids.Count} solid(s)");

                    /// Apply miter and tilt cuts if needed
                    if (miterLeft != 90 || tiltLeft != 90 || miterRight != 90 || tiltRight != 90)
                    {
                        ed.WriteMessage($"\nApplying miter cuts: Left={miterLeft},{tiltLeft} Right={miterRight},{tiltRight}");
                        ApplyMiterAndTiltCuts(resultSolids, miterLeft, tiltLeft, miterRight, tiltRight, length,
                                            preserveOriginalOrientation, visualizeOnly);
                    }

                    // Apply final transformation
                    if (finalTransform != Matrix3d.Identity)
                    {
                        ed.WriteMessage("\nApplying final transformation...");
                        foreach (var solid in resultSolids)
                        {
                            solid.TransformBy(finalTransform);
                        }
                    }

                    // Add solids to modelspace
                    ed.WriteMessage("\nAdding solids to modelspace...");
                    var ms = database.GetModelspace(trans);
                    ms.UpgradeOpen();

                    foreach (var solid in resultSolids)
                    {
                        solid.Layer = "0"; // Set default layer, change as needed
                        ms.AppendEntity(solid);
                        trans.AddNewlyCreatedDBObject(solid, true);
                    }

                    ms.DowngradeOpen();

                    // Dispose of regions as they're no longer needed
                    foreach (var region in regions)
                    {
                        region.Dispose();
                    }

                    ed.WriteMessage("\nCommitting transaction...");
                    trans.Commit();
                }
                catch (Exception ex)
                {
                    ed.WriteMessage($"\nError in CreateExtrudedPart: {ex.Message}");
                    if (ex.InnerException != null)
                    {
                        ed.WriteMessage($"\nInner exception: {ex.InnerException.Message}");
                    }
                    trans.Abort();
                    throw new Exception($"Failed to create extruded part: {ex.Message}", ex);
                }
            }

            return resultSolids;
        }
    }

    // Extension methods for convenience
    public static class DrawfabExtensions
    {
        public static double ToRadians(this double degrees)
        {
            return degrees * (Math.PI / 180.0);
        }

        public static double ToDegrees(this double radians)
        {
            return radians * (180.0 / Math.PI);
        }

        public static BlockTable GetBlockTable(this Database database, Transaction transaction)
        {
            return (BlockTable)transaction.GetObject(database.BlockTableId, OpenMode.ForRead);
        }

        public static BlockTableRecord GetModelspace(this Database database, Transaction transaction)
        {
            var blockTable = database.GetBlockTable(transaction);
            return (BlockTableRecord)transaction.GetObject(blockTable[BlockTableRecord.ModelSpace], OpenMode.ForRead);
        }
    }
    
}