using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Autodesk.AutoCAD.ApplicationServices;
using Autodesk.AutoCAD.ApplicationServices.Core;
using Autodesk.AutoCAD.Colors;
using Autodesk.AutoCAD.DatabaseServices;
using Autodesk.AutoCAD.Geometry;
using Autodesk.AutoCAD.GraphicsInterface;

namespace TakeoffBridge
{
    public class Drawfab
    {
        private const string WorkPointAppName = "WP";
        private double _width;
        private double _height;
        private string _partNumber;
        private Document _document;
        private bool _ownsDocument;

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

        // New overloaded constructor (accepts explicit document)
        public Drawfab(string partNumber, Document document)
        {
            _partNumber = partNumber;
            _document = document;
            _ownsDocument = false;
        }

        public double Width => _width;
        public double Height => _height;
        public string PartNumber => _partNumber;

        // Then in your methods, check which approach to use:
        private Document GetDocument()
        {
            if (_document != null)
            {
                return _document;
            }
            else
            {
                return Autodesk.AutoCAD.ApplicationServices.Application.DocumentManager.MdiActiveDocument;
            }
        }

        /// <summary>
        /// Imports a drawing file as a block in the current database
        /// </summary>
        private BlockTableRecord ImportDrawingAsBlock(Database database, string partDrawingPath, string blockName, Transaction trans)
        {
            // Check if block already exists
            var bt = database.GetBlockTable(trans);

            if (!bt.Has(blockName))
            {
                try
                {
                    // Create a new block
                    BlockTableRecord newBtr = new BlockTableRecord();
                    newBtr.Name = blockName;

                    // Get the block table for write and add our new block
                    bt.UpgradeOpen();
                    bt.Add(newBtr);
                    trans.AddNewlyCreatedDBObject(newBtr, true);

                    // Now open the newly created block for write so we can add entities
                    newBtr.UpgradeOpen();

                    // Open the source drawing and copy entities directly
                    using (var tempDb = new Database(false, true))
                    {
                        tempDb.ReadDwgFile(partDrawingPath, FileOpenMode.OpenForReadAndAllShare, true, "");

                        using (Transaction tempTr = tempDb.TransactionManager.StartTransaction())
                        {
                            BlockTable tempBt = (BlockTable)tempTr.GetObject(tempDb.BlockTableId, OpenMode.ForRead);
                            BlockTableRecord tempMs = (BlockTableRecord)tempTr.GetObject(
                                tempBt[BlockTableRecord.ModelSpace], OpenMode.ForRead);

                            // Copy each entity to the new block
                            foreach (ObjectId id in tempMs)
                            {
                                try
                                {
                                    Entity ent = tempTr.GetObject(id, OpenMode.ForRead) as Entity;
                                    if (ent != null)
                                    {
                                        // Clone the entity but handle possible errors
                                        Entity clone = null;
                                        try
                                        {
                                            // Create a deep copy of the entity
                                            clone = ent.Clone() as Entity;

                                            // Explicitly clear any XData to avoid database conflicts
                                            clone.XData = null;

                                            // Add to the block
                                            newBtr.AppendEntity(clone);
                                            trans.AddNewlyCreatedDBObject(clone, true);
                                        }
                                        catch (Exception cloneEx)
                                        {
                                            System.Diagnostics.Debug.WriteLine($"Error cloning entity: {cloneEx.Message}");
                                            if (clone != null && !clone.IsDisposed)
                                            {
                                                clone.Dispose();
                                            }
                                        }
                                    }
                                }
                                catch (Exception entEx)
                                {
                                    System.Diagnostics.Debug.WriteLine($"Error processing entity: {entEx.Message}");
                                }
                            }

                            tempTr.Commit();
                        }
                    }

                    // Downgrade the block table when done
                    bt.DowngradeOpen();

                    return newBtr;
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error in ImportDrawingAsBlock: {ex.Message}");

                    // Create an empty block as a fallback
                    try
                    {
                        BlockTableRecord emptyBlock = new BlockTableRecord();
                        emptyBlock.Name = blockName;
                        bt.Add(emptyBlock);
                        trans.AddNewlyCreatedDBObject(emptyBlock, true);

                        System.Diagnostics.Debug.WriteLine($"Created empty block '{blockName}' as fallback");
                        return emptyBlock;
                    }
                    catch (Exception ex2)
                    {
                        System.Diagnostics.Debug.WriteLine($"Failed to create fallback block: {ex2.Message}");
                        throw;
                    }
                }
            }

            // Return existing block
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
        private void SetWorkPoint(Point3d workPoint, Database database, Transaction trans, Entity entity)
        {
            // Register application name if needed
            RegisterAppName(trans, database, WorkPointAppName);

            // Set XData with workpoint
            entity.XData = new ResultBuffer(
                new TypedValue((int)DxfCode.ExtendedDataRegAppName, WorkPointAppName),
                new TypedValue((int)DxfCode.ExtendedDataWorldXCoordinate, workPoint)
            );
        }

        /// <summary>
        /// Registers an application name for XData if not already registered
        /// </summary>
        private void RegisterAppName(Transaction trans, Database database, string appName)
        {

            Database db = database;

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
                        RegisterAppName(trans, database,  WorkPointAppName);

                        // Set XData with workpoint
                        try
                        {
                            region.XData = new ResultBuffer(
                                new TypedValue((int)DxfCode.ExtendedDataRegAppName, WorkPointAppName),
                                new TypedValue((int)DxfCode.ExtendedDataWorldXCoordinate, workPoint.Value)
                            );
                        }
                        catch (Autodesk.AutoCAD.Runtime.Exception ex)
                        {
                            if (ex.ErrorStatus == Autodesk.AutoCAD.Runtime.ErrorStatus.RegisteredApplicationIdNotFound)
                            {
                                // Just log and continue
                                System.Diagnostics.Debug.WriteLine($"Ignoring RegAppIdNotFound: {ex.Message}");
                            }
                            else throw; // Re-throw any other errors
                        }
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

            // Modified: Move to origin with vertical centering
            // Create the displacement vector correctly
            Vector3d displacementVector = new Vector3d(
                -extents.MinPoint.X,
                -(extents.MinPoint.Y + extents.MaxPoint.Y) / 2, // Center vertically
                -extents.MinPoint.Z);

            // Apply displacement
            var transformation = Matrix3d.Displacement(displacementVector);

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

                        solid.CreateExtrudedSolid(region, adjustedLength * extrusionDirection, sweepOpts);

                        // Transfer workpoint data if available
                        Point3d? workPoint = GetWorkPoint(region);
                        if (workPoint.HasValue)
                        {
                            SetWorkPoint(workPoint.Value, database, trans, solid);
                        }

                        solids.Add(solid);
                    }
                    catch (Exception ex)
                    {
                        // write ext to output window not editor
                        System.Diagnostics.Debug.WriteLine($"Error creating solid from region: {ex.Message}");


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
            Document doc = GetDocument();
            Database db = doc.Database;


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


                // LEFT END CUTTING - Unchanged from previous version
                if (miterLeft != 90 || tiltLeft != 90)
                {

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

                        }
                    }

                    // Create the cutting plane
                    Plane leftCutPlane = new Plane(planeOrigin, normal);


                    // Apply cutting operation
                    foreach (var solid in solids)
                    {
                        try
                        {
                            solid.Slice(leftCutPlane);

                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Error on Left Cut: {ex.Message}");
                        }
                    }
                }

                // RIGHT END CUTTING - MODIFIED for mirror behavior
                if (miterRight != 90 || tiltRight != 90)
                {

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

                    }
                    // For 135 degree miter (original), move left by depth
                    else if (miterRight == 135)
                    {
                        // Move left by depth
                        planeOrigin = new Point3d(length - partDepth, 0, 0);

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

                        }
                        // For less than 90, no movement needed
                        else
                        {
                            planeOrigin = new Point3d(length, 0, 0);

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

                        }
                    }

                    // Create the cutting plane
                    Plane rightCutPlane = new Plane(planeOrigin, normal);


                    // Apply cutting operation
                    foreach (var solid in solids)
                    {
                        try
                        {
                            solid.Slice(rightCutPlane);

                        }
                        catch (Exception ex)
                        {
                            System.Diagnostics.Debug.WriteLine($"Error on Right Cut: {ex.Message}");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error on Applying Miter and Tilt Cuts: {ex.Message}");
            }
        }

        private void AddAngleDimensions(List<Solid3d> solids, double miterLeft, double tiltLeft,
                               double miterRight, double tiltRight, double length,
                               Database db, Transaction trans)
        {
            // Get the block table
            BlockTable bt = (BlockTable)trans.GetObject(db.BlockTableId, OpenMode.ForRead);
            BlockTableRecord ms = (BlockTableRecord)trans.GetObject(bt[BlockTableRecord.ModelSpace], OpenMode.ForWrite);

            // Create a layer for dimensions
            LayerTable lt = (LayerTable)trans.GetObject(db.LayerTableId, OpenMode.ForRead);
            string dimLayerName = "DIMENSIONS";

            if (!lt.Has(dimLayerName))
            {
                lt.UpgradeOpen();
                LayerTableRecord ltr = new LayerTableRecord();
                ltr.Name = dimLayerName;
                ltr.Color = Color.FromColorIndex(ColorMethod.ByAci, 3); // Green
                lt.Add(ltr);
                trans.AddNewlyCreatedDBObject(ltr, true);
            }

            // LEFT END DIMENSIONS
            if (miterLeft != 90 || tiltLeft != 90)
            {
                // For miter - create dimension in XY plane
                if (miterLeft != 90)
                {
                    // Create an angular dimension
                    Point3d center = new Point3d(0, 0, 0);
                    Point3d xAxisPt = new Point3d(5, 0, 0);
                    Point3d anglePt;

                    // Calculate point on angle line
                    double angle = miterLeft.ToRadians();
                    anglePt = new Point3d(5 * Math.Cos(angle), 5 * Math.Sin(angle), 0);

                    // Create angular dimension 
                    RotatedDimension dimObj = new RotatedDimension();
                    dimObj.XLine1Point = center;
                    dimObj.XLine2Point = xAxisPt;
                    dimObj.DimLinePoint = anglePt;
                    dimObj.Rotation = 0;
                    dimObj.DimensionText = $"{miterLeft}°";
                    dimObj.TextPosition = new Point3d(3, 3, 0);
                    dimObj.Layer = dimLayerName;

                    ms.AppendEntity(dimObj);
                    trans.AddNewlyCreatedDBObject(dimObj, true);
                }

                // For tilt - create dimension in XZ plane
                if (tiltLeft != 90)
                {
                    // Create an angular dimension for tilt
                    Point3d center = new Point3d(0, 0, 0);
                    Point3d xAxisPt = new Point3d(5, 0, 0);
                    Point3d anglePt;

                    // Calculate point on angle line
                    double angle = tiltLeft.ToRadians();
                    anglePt = new Point3d(5 * Math.Cos(angle), 0, 5 * Math.Sin(angle));

                    // Create angular dimension
                    RotatedDimension dimObj = new RotatedDimension();
                    dimObj.XLine1Point = center;
                    dimObj.XLine2Point = xAxisPt;
                    dimObj.DimLinePoint = anglePt;
                    dimObj.Rotation = 0;
                    dimObj.DimensionText = $"{tiltLeft}°";
                    dimObj.TextPosition = new Point3d(3, 0, 3);
                    dimObj.Layer = dimLayerName;

                    ms.AppendEntity(dimObj);
                    trans.AddNewlyCreatedDBObject(dimObj, true);
                }
            }

            // RIGHT END DIMENSIONS
            if (miterRight != 90 || tiltRight != 90)
            {
                // For miter - create dimension in XY plane
                if (miterRight != 90)
                {
                    // Create an angular dimension
                    Point3d center = new Point3d(length, 0, 0);
                    Point3d xAxisPt = new Point3d(length - 5, 0, 0);
                    Point3d anglePt;

                    // Calculate point on angle line
                    double angle = (180 - miterRight).ToRadians();
                    anglePt = new Point3d(length - 5 * Math.Cos(angle), 5 * Math.Sin(angle), 0);

                    // Create angular dimension
                    RotatedDimension dimObj = new RotatedDimension();
                    dimObj.XLine1Point = center;
                    dimObj.XLine2Point = xAxisPt;
                    dimObj.DimLinePoint = anglePt;
                    dimObj.Rotation = 0;
                    dimObj.DimensionText = $"{miterRight}°";
                    dimObj.TextPosition = new Point3d(length - 3, 3, 0);
                    dimObj.Layer = dimLayerName;

                    ms.AppendEntity(dimObj);
                    trans.AddNewlyCreatedDBObject(dimObj, true);
                }

                // For tilt - create dimension in XZ plane
                if (tiltRight != 90)
                {
                    // Create an angular dimension for tilt
                    Point3d center = new Point3d(length, 0, 0);
                    Point3d xAxisPt = new Point3d(length - 5, 0, 0);
                    Point3d anglePt;

                    // Calculate point on angle line
                    double angle = tiltRight.ToRadians();
                    anglePt = new Point3d(length - 5 * Math.Cos(angle), 0, 5 * Math.Sin(angle));

                    // Create angular dimension
                    RotatedDimension dimObj = new RotatedDimension();
                    dimObj.XLine1Point = center;
                    dimObj.XLine2Point = xAxisPt;
                    dimObj.DimLinePoint = anglePt;
                    dimObj.Rotation = 0;
                    dimObj.DimensionText = $"{tiltRight}°";
                    dimObj.TextPosition = new Point3d(length - 3, 0, 3);
                    dimObj.Layer = dimLayerName;

                    ms.AppendEntity(dimObj);
                    trans.AddNewlyCreatedDBObject(dimObj, true);
                }
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
            bool visualizeOnly = false,
            bool addDimensions = true)
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
            bool visualizeOnly = false,
            bool addDimensions = true)
        {
            List<Solid3d> resultSolids = new List<Solid3d>();


            using (var trans = database.TransactionManager.StartTransaction())
            {
                try
                {
                    // Import drawing as block
                    var btr = ImportDrawingAsBlock(database, partDrawingPath, blockName, trans);
                    if (btr == null)
                    {
                        throw new Exception($"Failed to import drawing from {partDrawingPath}");
                    }
                    try
                    {
                        


                        // Extract regions from block
                    var regions = GetExtrusionRegionsFromBlock(database, btr.ObjectId, trans);
                    if (regions == null || regions.Count == 0)
                    {
                        throw new Exception("No valid regions found in the drawing");
                    }


                    // Transform regions to orientation
                    var handedTransformation = GetHandedTransformation(handed, handedSide);
                    TransformRegionsToStandardOrientation(regions, handedTransformation, preserveOriginalOrientation);

                    // Extrude regions to create solids
                    resultSolids = ExtrudeRegions(regions, length, database, trans, preserveOriginalOrientation);

                    /// Apply miter and tilt cuts if needed
                    if (miterLeft != 90 || tiltLeft != 90 || miterRight != 90 || tiltRight != 90)
                    {
                        ApplyMiterAndTiltCuts(resultSolids, miterLeft, tiltLeft, miterRight, tiltRight, length,
                                            preserveOriginalOrientation, visualizeOnly);
                    }
                    // Add angle dimensions if needed
                    if (addDimensions)
                    {
                        AddAngleDimensions(resultSolids, miterLeft, tiltLeft, miterRight, tiltRight, length, database, trans);
                    }

                    // Apply final transformation
                    if (finalTransform != Matrix3d.Identity)
                    {
                        foreach (var solid in resultSolids)
                        {
                            solid.TransformBy(finalTransform);
                        }
                    }

                    // Add solids to modelspace
                    var ms = database.GetModelspace(trans);
                    ms.UpgradeOpen();

                    foreach (var solid in resultSolids)
                    {
                        solid.Layer = "0"; // Set default layer, change as needed
                        ms.AppendEntity(solid);
                        trans.AddNewlyCreatedDBObject(solid, true);
                    }
                    trans.Commit();
                    //ms.DowngradeOpen();

                    // Dispose of regions as they're no longer needed
                    foreach (var region in regions)
                    {
                        region.Dispose();
                    }

                    }
                    finally
                    {
                        (btr as IDisposable)?.Dispose();
                    }
                }
                catch (Exception ex)
                {
                    System.Diagnostics.Debug.WriteLine($"Error Creating Extruded Part: {ex.Message}");
                    if (ex.InnerException != null)
                    {
                        System.Diagnostics.Debug.WriteLine($"Inner Exception: {ex.InnerException.Message}");
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