using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using JSE_RevitAddin_MEP_OPENINGS.Services;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Transactions;

namespace JSE_RevitAddin_MEP_OPENINGS.Commands
{
    [Transaction(TransactionMode.Manual)]
    public class ClusterMergeCommand : IExternalCommand
    {
        // Static field to store cluster geometry between method calls
        private static ClusterGeometry? _storedClusterGeometry;
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            // Setup logging for this command
            try
            {
                DebugLogger.InitLogFile("clustermerge");
                DebugLogger.Log("=== ClusterMergeCommand START (Multi-Merge Logic) ===");
            }
            catch (Exception ex)
            {
                TaskDialog.Show("Logging Error", $"Failed to initialize logger: {ex.Message}");
                return Result.Failed;
            }

            UIDocument uiDoc = commandData.Application.ActiveUIDocument;
            Document doc = uiDoc.Document;

            try
            {
                // 1. Get pre-selected elements
                ICollection<ElementId> selectedIds = uiDoc.Selection.GetElementIds();
                DebugLogger.Log($"Initial selection count: {selectedIds.Count}");

                if (selectedIds.Count < 2)
                {
                    message = "Please select at least two cluster elements to merge.";
                    TaskDialog.Show("Selection Error", message);
                    DebugLogger.Log($"Error: Not enough elements selected. Aborting. Message: {message}");
                    return Result.Failed;
                }

                // 2. Filter for valid FamilyInstances
                List<FamilyInstance> selectedInstances = selectedIds
                    .Select(id => doc.GetElement(id))
                    .OfType<FamilyInstance>()
                    .ToList();

                DebugLogger.Log($"Found {selectedInstances.Count} valid FamilyInstances in selection.");

                if (selectedInstances.Count < 2)
                {
                    message = "The selection must contain at least two cluster family instances.";
                    TaskDialog.Show("Selection Error", message);
                    DebugLogger.Log($"Error: Not enough FamilyInstances in selection. Aborting. Message: {message}");
                    return Result.Failed;
                }

                // 3. Calculate the combined bounding box of all selected instances
                BoundingBoxXYZ? combinedBbox = null;
                foreach (var instance in selectedInstances)
                {
                    var instanceBbox = instance.get_BoundingBox(null);
                    if (instanceBbox == null)
                    {
                        DebugLogger.Log($"Warning: Instance ID {instance.Id} has no bounding box and will be skipped.");
                        continue;
                    }

                    if (combinedBbox == null)
                    {
                        combinedBbox = instanceBbox;
                    }
                    else
                    {
                        combinedBbox.Min = new XYZ(Math.Min(combinedBbox.Min.X, instanceBbox.Min.X), Math.Min(combinedBbox.Min.Y, instanceBbox.Min.Y), Math.Min(combinedBbox.Min.Z, instanceBbox.Min.Z));
                        combinedBbox.Max = new XYZ(Math.Max(combinedBbox.Max.X, instanceBbox.Max.X), Math.Max(combinedBbox.Max.Y, instanceBbox.Max.Y), Math.Max(combinedBbox.Max.Z, instanceBbox.Max.Z));
                    }
                }

                if (combinedBbox == null)
                {
                    message = "Could not determine the combined boundaries of the selected elements.";
                    TaskDialog.Show("Error", message);
                    DebugLogger.Log($"FATAL: Could not calculate a combined bounding box. Aborting.");
                    return Result.Failed;
                }
                DebugLogger.Log($"Combined BBox Min: {combinedBbox.Min}, Max: {combinedBbox.Max}");

                // 3a. Determine the cluster's orientation from the first selected instance
                var firstInstance = selectedInstances.First();
                var referenceTransform = firstInstance.GetTransform();
                // We only consider rotation in the XY plane (around Z-axis) for wall/floor sleeves
                double rotationAngle = Math.Atan2(referenceTransform.BasisX.Y, referenceTransform.BasisX.X);
                var rotation = Transform.CreateRotation(XYZ.BasisZ, rotationAngle);
                var inverseRotation = rotation.Inverse;
                DebugLogger.Log($"Determined cluster orientation angle: {rotationAngle} radians from instance {firstInstance.Id}.");

                // 3b. Get all corners of all bounding boxes in world coordinates
                var allWorldCorners = new List<XYZ>();
                foreach (var instance in selectedInstances)
                {
                    var instanceBbox = instance.get_BoundingBox(null);
                    if (instanceBbox == null) continue;
                    allWorldCorners.AddRange(GetBoundingBoxCorners(instanceBbox));
                }

                // 3c. Transform corners to the local coordinate system and find the new local bounding box
                var localCorners = allWorldCorners.Select(c => inverseRotation.OfPoint(c)).ToList();

                var localMin = new XYZ(
                    localCorners.Min(c => c.X),
                    localCorners.Min(c => c.Y),
                    localCorners.Min(c => c.Z)
                );
                var localMax = new XYZ(
                    localCorners.Max(c => c.X),
                    localCorners.Max(c => c.Y),
                    localCorners.Max(c => c.Z)
                );
                DebugLogger.Log($"Local BBox (oriented) Min: {localMin}, Max: {localMax}");

                // 4. Find the replacement family symbol
                string replacementFamilyPrefix = "MergeDCluster";
                DebugLogger.Log($"Searching for a replacement family symbol with a family name starting with '{replacementFamilyPrefix}'.");

                FamilySymbol? replacementSymbol = new FilteredElementCollector(doc)
                    .OfClass(typeof(FamilySymbol))
                    .Cast<FamilySymbol>()
                    .FirstOrDefault(s => s.Family != null && s.Family.Name.StartsWith(replacementFamilyPrefix, StringComparison.OrdinalIgnoreCase));

                if (replacementSymbol == null)
                {
                    DebugLogger.Log($"FATAL: No family symbol found with a family name starting with '{replacementFamilyPrefix}'. Please ensure the family is loaded.");
                    message = $"Could not find a replacement family. Please load a family whose name starts with '{replacementFamilyPrefix}'.";
                    TaskDialog.Show("Error", message);
                    return Result.Failed;
                }

                DebugLogger.Log($"Found replacement symbol: Family='{replacementSymbol.Family.Name}', Symbol='{replacementSymbol.Name}' (ID: {replacementSymbol.Id}).");

                // 5. Determine properties for the new merged cluster
                XYZ centerPoint = (combinedBbox.Min + combinedBbox.Max) / 2.0;
                double width = localMax.X - localMin.X;
                double height = localMax.Y - localMin.Y;
                double depth = localMax.Z - localMin.Z;

                Level? level = doc.GetElement(selectedInstances.First().LevelId) as Level;
                if (level == null)
                {
                    // Fallback: find nearest level to the center point
                    level = new FilteredElementCollector(doc)
                        .OfClass(typeof(Level))
                        .Cast<Level>()
                        .OrderBy(l => Math.Abs(l.ProjectElevation - centerPoint.Z))
                        .FirstOrDefault();
                }

                if (level == null)
                {
                    message = "Could not determine a valid level for the new cluster.";
                    TaskDialog.Show("Error", message);
                    DebugLogger.Log($"FATAL: Could not find a level for placement. Aborting.");
                    return Result.Failed;
                }

                DebugLogger.Log($"New cluster properties: Center={centerPoint}, Width={width}, Height={height}, Depth={depth}, Level={level.Name}");

                // 6. Store host information and cluster geometry BEFORE deleting instances
                var originalHost = firstInstance.Host;
                var originalHostId = originalHost?.Id;
                
                // Store cluster geometry for parameter distribution
                var clusterGeometry = new ClusterGeometry
                {
                    MinX = localMin.X,
                    MaxX = localMax.X,
                    MinY = localMin.Y,
                    MaxY = localMax.Y
                };
                StoreClusterGeometry(clusterGeometry);
                
                // 6. Replace the selected instances
                using (Transaction tx = new Transaction(doc, "Merge Clusters"))
                {
                    DebugLogger.Log("Starting transaction to merge clusters.");
                    tx.Start();

                    // Delete the original instances
                    doc.Delete(selectedIds);
                    DebugLogger.Log($"Deleted {selectedIds.Count} original instances.");

                    if (!replacementSymbol.IsActive)
                    {
                        DebugLogger.Log($"Activating replacement symbol '{replacementSymbol.Name}'.");
                        replacementSymbol.Activate();
                    }
                    
                    DebugLogger.Log("Creating new merged family instance.");
                    FamilyInstance newInstance = doc.Create.NewFamilyInstance(centerPoint, replacementSymbol, level, Autodesk.Revit.DB.Structure.StructuralType.NonStructural);

                    if (newInstance == null)
                    {
                        DebugLogger.Log("FATAL: Failed to create new family instance. Aborting transaction.");
                        tx.RollBack();
                        message = "Failed to create the new merged cluster instance.";
                        return Result.Failed;
                    }

                    DebugLogger.Log($"Successfully created new merged instance with ID: {newInstance.Id}");

                    // Log available parameters for debugging
                    LogAvailableParameters(newInstance);

                    // Set size parameters based on family type (Horizontal vs Vertical)
                    SetSizeParametersForFamilyType(newInstance, width, height, depth);

                    // Apply the original cluster's rotation to the new instance
                    double finalRotationAngle = rotationAngle;
                    var host = originalHost;

                    // For X-aligned walls, a 180-degree rotation of the original instance should not be
                    // applied to the new instance, as the new family is assumed to be correctly oriented by default.
                    if (originalHostId != null && Math.Abs(Math.Abs(rotationAngle) - Math.PI) < 0.001) // if angle is ~180 or ~-180 deg
                    {
                        // Get the host from the document since the original reference might be invalid
                        var currentHost = doc.GetElement(originalHostId);
                        if (currentHost is Wall wall)
                        {
                            var loc = wall.Location as LocationCurve;
                            if (loc != null && loc.Curve is Line line)
                            {
                                XYZ wallDirection = line.Direction.Normalize();
                                // If wall direction's X component is much larger than Y, it's an X-aligned wall.
                                if (Math.Abs(wallDirection.X) > Math.Abs(wallDirection.Y))
                                {
                                    DebugLogger.Log($"Original instance was rotated ~180d on an X-aligned wall. Overriding final rotation to 0.");
                                    finalRotationAngle = 0;
                                }
                            }
                        }
                    }

                    if (Math.Abs(finalRotationAngle) > 0.001)
                    {
                        var axis = Line.CreateBound(centerPoint, centerPoint + XYZ.BasisZ);
                        DebugLogger.Log($"Applying rotation of {finalRotationAngle} radians around Z axis.");
                        ElementTransformUtils.RotateElement(doc, newInstance.Id, axis, finalRotationAngle);
                    }

                    tx.Commit();
                    DebugLogger.Log("Transaction committed successfully.");
                }

                TaskDialog.Show("Success", "Cluster has been replaced successfully.");
                DebugLogger.Log("=== ClusterMergeCommand END ===");
                return Result.Succeeded;
            }
            catch (Autodesk.Revit.Exceptions.OperationCanceledException)
            {
                DebugLogger.Log("Operation cancelled by user.");
                return Result.Cancelled;
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"An unhandled exception occurred: {ex.Message}\n{ex.StackTrace}");
                message = $"An error occurred: {ex.Message}";
                return Result.Failed;
            }
        }

        /// <summary>
        /// Gets the 8 corner points of an axis-aligned bounding box.
        /// </summary>
        private IEnumerable<XYZ> GetBoundingBoxCorners(BoundingBoxXYZ bbox)
        {
            return new[]
            {
                bbox.Min,
                new XYZ(bbox.Max.X, bbox.Min.Y, bbox.Min.Z),
                new XYZ(bbox.Max.X, bbox.Max.Y, bbox.Min.Z),
                new XYZ(bbox.Min.X, bbox.Max.Y, bbox.Min.Z),
                new XYZ(bbox.Min.X, bbox.Min.Y, bbox.Max.Z),
                new XYZ(bbox.Max.X, bbox.Min.Y, bbox.Max.Z),
                bbox.Max,
                new XYZ(bbox.Min.X, bbox.Max.Y, bbox.Max.Z)
            };
        }

        /// <summary>
        /// Helper method to set a parameter on an instance and log the outcome.
        /// </summary>
        private bool SetParameter(FamilyInstance instance, string paramName, double value, string dimensionName)
        {
            Parameter? param = instance.LookupParameter(paramName);
            if (param != null && !param.IsReadOnly)
            {
                param.Set(value);
                DebugLogger.Log($"Set parameter '{paramName}' to {value} for new cluster's {dimensionName}.");
                return true;
            }
            else
            {
                DebugLogger.Log($"Warning: Could not find or set parameter '{paramName}' for new cluster.");
                return false;
            }
        }

        /// <summary>
        /// Logs all available parameters on a family instance for debugging purposes.
        /// </summary>
        private void LogAvailableParameters(FamilyInstance instance)
        {
            try
            {
                DebugLogger.Log($"=== Available Parameters for {instance.Symbol.Family.Name} ===");
                foreach (Parameter param in instance.Parameters)
                {
                    if (param != null)
                    {
                        string paramValue = param.StorageType == StorageType.String ? param.AsString() ?? "N/A" :
                                          param.StorageType == StorageType.Double ? param.AsDouble().ToString("F3") :
                                          param.StorageType == StorageType.Integer ? param.AsInteger().ToString() :
                                          param.StorageType == StorageType.ElementId ? param.AsElementId().ToString() :
                                          "N/A";
                        
                        DebugLogger.Log($"Parameter: {param.Definition.Name} (Type: {param.StorageType}, ReadOnly: {param.IsReadOnly}, Value: {paramValue})");
                    }
                }
                DebugLogger.Log("=== End Parameters ===");
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"Error logging parameters: {ex.Message}");
            }
        }

        /// <summary>
        /// Sets size parameters based on family type (Horizontal vs Vertical) and cluster geometry
        /// </summary>
        private void SetSizeParametersForFamilyType(FamilyInstance instance, double width, double height, double depth)
        {
            try
            {
                // Check if this is a horizontal family (has H1, H2, H3 parameters)
                bool isHorizontal = HasParameter(instance, "H1") && HasParameter(instance, "H2") && HasParameter(instance, "H3");
                
                // Check if this is a vertical family (has W1, W2, W3 parameters)
                bool isVertical = HasParameter(instance, "W1") && HasParameter(instance, "W2") && HasParameter(instance, "W3");
                
                DebugLogger.Log($"Family type detection: Horizontal={isHorizontal}, Vertical={isVertical}");
                
                if (isHorizontal)
                {
                    // Horizontal family: W1, W2 for width, H1, H2, H3 for height
                    DebugLogger.Log("Setting parameters for HORIZONTAL family type");
                    
                    // For horizontal families, distribute width and height based on cluster geometry
                    // W1 and W2 represent left and right sections
                    // H1, H2, H3 represent top, middle, and bottom sections
                    DistributeHorizontalParameters(instance, width, height);
                }
                else if (isVertical)
                {
                    // Vertical family: W1, W2, W3 for width, H1, H2 for height
                    DebugLogger.Log("Setting parameters for VERTICAL family type");
                    
                    // For vertical families, distribute width and height based on cluster geometry
                    // W1, W2, W3 represent left, middle, and right sections
                    // H1 and H2 represent top and bottom sections
                    DistributeVerticalParameters(instance, width, height);
                }
                else
                {
                    // Fallback: try standard parameter names
                    DebugLogger.Log("Family type not recognized, trying standard parameter names");
                    SetParameter(instance, "Width", width, "width");
                    SetParameter(instance, "Height", height, "height");
                    SetParameter(instance, "Depth", depth, "depth");
                    
                    // Try alternative parameter names
                    if (!HasParameter(instance, "Width"))
                    {
                        SetParameter(instance, "w", width, "width");
                    }
                    if (!HasParameter(instance, "Height"))
                    {
                        SetParameter(instance, "h", height, "height");
                    }
                    if (!HasParameter(instance, "Depth"))
                    {
                        SetParameter(instance, "d", depth, "depth");
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"Error setting size parameters: {ex.Message}");
            }
        }

        /// <summary>
        /// Checks if a parameter exists on the instance
        /// </summary>
        private bool HasParameter(FamilyInstance instance, string paramName)
        {
            try
            {
                var param = instance.LookupParameter(paramName);
                return param != null && !param.IsReadOnly;
            }
            catch
            {
                return false;
            }
        }

        /// <summary>
        /// Distributes parameters for horizontal family based on cluster geometry
        /// </summary>
        private void DistributeHorizontalParameters(FamilyInstance instance, double width, double height)
        {
            try
            {
                // Use the stored cluster geometry for parameter distribution
                var clusterGeometry = GetStoredClusterGeometry();
                if (clusterGeometry == null)
                {
                    DebugLogger.Log("Warning: Cannot get stored cluster geometry, using default distribution");
                    // Fallback to default distribution
                    if (HasParameter(instance, "W1")) SetParameter(instance, "W1", width / 2.0, "half-width");
                    if (HasParameter(instance, "W2")) SetParameter(instance, "W2", width / 2.0, "half-width");
                    if (HasParameter(instance, "H1")) SetParameter(instance, "H1", height / 3.0, "height-section-1");
                    if (HasParameter(instance, "H2")) SetParameter(instance, "H2", height / 3.0, "height-section-2");
                    if (HasParameter(instance, "H3")) SetParameter(instance, "H3", height / 3.0, "height-section-3");
                    return;
                }

                DebugLogger.Log($"Using stored cluster geometry: MinX={clusterGeometry.MinX}, MaxX={clusterGeometry.MaxX}, MinY={clusterGeometry.MinY}, MaxY={clusterGeometry.MaxY}");
                
                // Distribute width parameters (W1, W2) based on X-axis distribution
                if (HasParameter(instance, "W1") && HasParameter(instance, "W2"))
                {
                    double leftWidth = clusterGeometry.MidX - clusterGeometry.MinX;
                    double rightWidth = clusterGeometry.MaxX - clusterGeometry.MidX;
                    
                    SetParameter(instance, "W1", leftWidth, "left-width");
                    SetParameter(instance, "W2", rightWidth, "right-width");
                    DebugLogger.Log($"Distributed width: W1={leftWidth}, W2={rightWidth}");
                }
                
                // Distribute height parameters (H1, H2, H3) based on Y-axis distribution
                if (HasParameter(instance, "H1") && HasParameter(instance, "H2") && HasParameter(instance, "H3"))
                {
                    double totalHeight = clusterGeometry.MaxY - clusterGeometry.MinY;
                    double sectionHeight = totalHeight / 3.0;
                    
                    // Simple distribution based on cluster position - can be enhanced later
                    // For now, distribute based on which sections have content
                    double topSection = clusterGeometry.MaxY - clusterGeometry.MidY;
                    double middleSection = clusterGeometry.MidY - clusterGeometry.MinY;
                    double bottomSection = totalHeight / 3.0;
                    
                    // Check if sections have significant content (threshold can be adjusted)
                    if (topSection > totalHeight * 0.1) // If top section has significant content
                    {
                        SetParameter(instance, "H1", sectionHeight, "top-height");
                    }
                    else
                    {
                        SetParameter(instance, "H1", 0.0, "top-height-empty");
                    }
                    
                    if (middleSection > totalHeight * 0.1) // If middle section has significant content
                    {
                        SetParameter(instance, "H2", sectionHeight, "middle-height");
                    }
                    else
                    {
                        SetParameter(instance, "H2", 0.0, "middle-height-empty");
                    }
                    
                    if (bottomSection > totalHeight * 0.1) // If bottom section has significant content
                    {
                        SetParameter(instance, "H3", sectionHeight, "bottom-height");
                    }
                    else
                    {
                        SetParameter(instance, "H3", 0.0, "bottom-height-empty");
                    }
                    
                    DebugLogger.Log($"Distributed height: H1={(topSection > totalHeight * 0.1 ? sectionHeight : 0)}, H2={(middleSection > totalHeight * 0.1 ? sectionHeight : 0)}, H3={(bottomSection > totalHeight * 0.1 ? sectionHeight : 0)}");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"Error distributing horizontal parameters: {ex.Message}");
            }
        }

        /// <summary>
        /// Distributes parameters for vertical family based on cluster geometry
        /// </summary>
        private void DistributeVerticalParameters(FamilyInstance instance, double width, double height)
        {
            try
            {
                // Use the stored cluster geometry for parameter distribution
                var clusterGeometry = GetStoredClusterGeometry();
                if (clusterGeometry == null)
                {
                    DebugLogger.Log("Warning: Cannot get stored cluster geometry, using default distribution");
                    // Fallback to default distribution
                    if (HasParameter(instance, "W1")) SetParameter(instance, "W1", width / 3.0, "width-section-1");
                    if (HasParameter(instance, "W2")) SetParameter(instance, "W2", width / 3.0, "width-section-2");
                    if (HasParameter(instance, "W3")) SetParameter(instance, "W3", width / 3.0, "width-section-3");
                    if (HasParameter(instance, "H1")) SetParameter(instance, "H1", height / 2.0, "half-height");
                    if (HasParameter(instance, "H2")) SetParameter(instance, "H2", height / 2.0, "half-height");
                    return;
                }

                DebugLogger.Log($"Using stored cluster geometry: MinX={clusterGeometry.MinX}, MaxX={clusterGeometry.MaxX}, MinY={clusterGeometry.MinY}, MaxY={clusterGeometry.MaxY}");
                
                // Distribute width parameters (W1, W2, W3) based on X-axis distribution
                if (HasParameter(instance, "W1") && HasParameter(instance, "W2") && HasParameter(instance, "W3"))
                {
                    double totalWidth = clusterGeometry.MaxX - clusterGeometry.MinX;
                    double sectionWidth = totalWidth / 3.0;
                    
                    // Simple distribution based on cluster position - can be enhanced later
                    // For now, distribute based on which sections have content
                    double leftSection = clusterGeometry.MidX - clusterGeometry.MinX;
                    double middleSection = totalWidth / 3.0;
                    double rightSection = clusterGeometry.MaxX - clusterGeometry.MidX;
                    
                    // Check if sections have significant content (threshold can be adjusted)
                    if (leftSection > totalWidth * 0.1) // If left section has significant content
                    {
                        SetParameter(instance, "W1", sectionWidth, "left-width");
                    }
                    else
                    {
                        SetParameter(instance, "W1", 0.0, "left-width-empty");
                    }
                    
                    if (middleSection > totalWidth * 0.1) // If middle section has significant content
                    {
                        SetParameter(instance, "W2", sectionWidth, "middle-width");
                    }
                    else
                    {
                        SetParameter(instance, "W2", 0.0, "middle-width-empty");
                    }
                    
                    if (rightSection > totalWidth * 0.1) // If right section has significant content
                    {
                        SetParameter(instance, "W3", sectionWidth, "right-width");
                    }
                    else
                    {
                        SetParameter(instance, "W3", 0.0, "right-width-empty");
                    }
                    
                    DebugLogger.Log($"Distributed width: W1={(leftSection > totalWidth * 0.1 ? sectionWidth : 0)}, W2={(middleSection > totalWidth * 0.1 ? sectionWidth : 0)}, W3={(rightSection > totalWidth * 0.1 ? sectionWidth : 0)}");
                }
                
                // Distribute height parameters (H1, H2) based on Y-axis distribution
                if (HasParameter(instance, "H1") && HasParameter(instance, "H2"))
                {
                    double topHeight = clusterGeometry.MidY - clusterGeometry.MinY;
                    double bottomHeight = clusterGeometry.MaxY - clusterGeometry.MidY;
                    
                    SetParameter(instance, "H1", topHeight, "top-height");
                    SetParameter(instance, "H2", bottomHeight, "bottom-height");
                    DebugLogger.Log($"Distributed height: H1={topHeight}, H2={bottomHeight}");
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Log($"Error distributing vertical parameters: {ex.Message}");
            }
        }

        // Data structures for cluster analysis
        private class ClusterGeometry
        {
            public double MinX { get; set; }
            public double MaxX { get; set; }
            public double MinY { get; set; }
            public double MaxY { get; set; }
            public double MidX => (MinX + MaxX) / 2.0;
            public double MidY => (MinY + MaxY) / 2.0;
        }

        private class XAxisDistribution
        {
            public double LeftSection { get; set; }
            public double MiddleSection { get; set; }
            public double RightSection { get; set; }
        }

        private class YAxisDistribution
        {
            public double TopSection { get; set; }
            public double MiddleSection { get; set; }
            public double BottomSection { get; set; }
        }

        /// <summary>
        /// Stores the cluster geometry for use in parameter distribution
        /// </summary>
        private void StoreClusterGeometry(ClusterGeometry geometry)
        {
            _storedClusterGeometry = geometry;
            DebugLogger.Log($"Stored cluster geometry: MinX={geometry.MinX}, MaxX={geometry.MaxX}, MinY={geometry.MinY}, MaxY={geometry.MaxY}");
        }

        /// <summary>
        /// Gets the stored cluster geometry
        /// </summary>
        private ClusterGeometry? GetStoredClusterGeometry()
        {
            return _storedClusterGeometry;
        }

        /// <summary>
        /// Gets the selected instances from the current transaction context
        /// </summary>
        private List<FamilyInstance> GetSelectedInstancesFromTransaction()
        {
            try
            {
                // Since we're in a transaction and the original instances are deleted,
                // we need to get the information from the stored data
                // This is a fallback method - in practice, we'll use the stored bounding box data
                return null;
            }
            catch
            {
                return null;
            }
        }

        /// <summary>
        /// Analyzes the geometry of the cluster based on stored bounding box information
        /// </summary>
        private ClusterGeometry AnalyzeClusterGeometry(List<FamilyInstance> instances)
        {
            try
            {
                if (instances == null || instances.Count == 0)
                {
                    // Return default geometry if no instances available
                    return new ClusterGeometry
                    {
                        MinX = 0, MaxX = 1,
                        MinY = 0, MaxY = 1
                    };
                }

                var minX = instances.Min(i => i.get_BoundingBox(null)?.Min.X ?? 0);
                var maxX = instances.Max(i => i.get_BoundingBox(null)?.Max.X ?? 1);
                var minY = instances.Min(i => i.get_BoundingBox(null)?.Min.Y ?? 0);
                var maxY = instances.Max(i => i.get_BoundingBox(null)?.Max.Y ?? 1);

                return new ClusterGeometry
                {
                    MinX = minX,
                    MaxX = maxX,
                    MinY = minY,
                    MaxY = maxY
                };
            }
            catch
            {
                return new ClusterGeometry
                {
                    MinX = 0, MaxX = 1,
                    MinY = 0, MaxY = 1
                };
            }
        }

        /// <summary>
        /// Analyzes the distribution of clusters along the X-axis
        /// </summary>
        private XAxisDistribution AnalyzeXAxisDistribution(List<FamilyInstance> instances, double minX, double maxX)
        {
            try
            {
                double totalWidth = maxX - minX;
                double sectionWidth = totalWidth / 3.0;
                
                double leftBound = minX + sectionWidth;
                double rightBound = maxX - sectionWidth;
                
                var leftSection = instances.Count(i => 
                    (i.get_BoundingBox(null)?.Min.X ?? 0) < leftBound);
                var middleSection = instances.Count(i => 
                    (i.get_BoundingBox(null)?.Min.X ?? 0) >= leftBound && 
                    (i.get_BoundingBox(null)?.Max.X ?? 1) <= rightBound);
                var rightSection = instances.Count(i => 
                    (i.get_BoundingBox(null)?.Max.X ?? 1) > rightBound);
                
                return new XAxisDistribution
                {
                    LeftSection = (double)leftSection / instances.Count,
                    MiddleSection = (double)middleSection / instances.Count,
                    RightSection = (double)rightSection / instances.Count
                };
            }
            catch
            {
                return new XAxisDistribution
                {
                    LeftSection = 0.33,
                    MiddleSection = 0.34,
                    RightSection = 0.33
                };
            }
        }

        /// <summary>
        /// Analyzes the distribution of clusters along the Y-axis
        /// </summary>
        private YAxisDistribution AnalyzeYAxisDistribution(List<FamilyInstance> instances, double minY, double maxY)
        {
            try
            {
                double totalHeight = maxY - minY;
                double sectionHeight = totalHeight / 3.0;
                
                double topBound = minY + sectionHeight;
                double bottomBound = maxY - sectionHeight;
                
                var topSection = instances.Count(i => 
                    (i.get_BoundingBox(null)?.Min.Y ?? 0) < topBound);
                var middleSection = instances.Count(i => 
                    (i.get_BoundingBox(null)?.Min.Y ?? 0) >= topBound && 
                    (i.get_BoundingBox(null)?.Max.Y ?? 1) <= bottomBound);
                var bottomSection = instances.Count(i => 
                    (i.get_BoundingBox(null)?.Max.Y ?? 1) > bottomBound);
                
                return new YAxisDistribution
                {
                    TopSection = (double)topSection / instances.Count,
                    MiddleSection = (double)middleSection / instances.Count,
                    BottomSection = (double)bottomSection / instances.Count
                };
            }
            catch
            {
                return new YAxisDistribution
                {
                    TopSection = 0.33,
                    MiddleSection = 0.34,
                    BottomSection = 0.33
                };
            }
        }
    }
}