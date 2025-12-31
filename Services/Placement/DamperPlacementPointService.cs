using System;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Models;
using JSE_RevitAddin_MEP_OPENINGS.Services.Configuration;
using JSE_RevitAddin_MEP_OPENINGS.Helpers;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.Placement
{
    /// <summary>
    /// ✅ SRP COMPLIANCE: Dedicated service for damper placement point adjustment ONLY.
    /// Single Responsibility: Adjust damper placement point from centroid to wall centerline if offset.
    /// 
    /// This service ONLY handles:
    /// - Check if damper centroid is offset from wall centerline
    /// - If offset > threshold, move placement point to wall centerline
    /// - That's it - no other adjustments, no damper offsets, no complex logic
    /// </summary>
    public class DamperPlacementPointService
    {
        private readonly Document _doc;

        public DamperPlacementPointService(Document doc)
        {
            _doc = doc ?? throw new ArgumentNullException(nameof(doc));
        }

        /// <summary>
        /// ✅ SRP: Adjust damper placement point from centroid to wall centerline if offset.
        /// This is the ONLY method that moves damper sleeves - all other classes delegate to this.
        /// </summary>
        /// <param name="zone">The clash zone (must be Duct Accessories category)</param>
        /// <param name="centroid">The damper centroid (original placement point)</param>
        /// <returns>Adjusted placement point (wall centerline if offset, otherwise original centroid)</returns>
        public XYZ AdjustDamperPlacementPoint(ClashZone zone, XYZ centroid)
        {
            if (zone == null)
                throw new ArgumentNullException(nameof(zone));
            
            if (centroid == null)
                throw new ArgumentNullException(nameof(centroid));

            // ✅ ONLY handle Duct Accessories (dampers)
            if (!string.Equals(zone.MepElementCategory, "Duct Accessories", StringComparison.OrdinalIgnoreCase))
            {
                // Not a damper - return original point unchanged
                return centroid;
            }

            // ✅ ONLY handle walls and framing (floors don't need centerline adjustment)
            if (string.IsNullOrEmpty(zone.StructuralElementType) ||
                (!zone.StructuralElementType.Contains("Wall", StringComparison.OrdinalIgnoreCase) &&
                 !zone.StructuralElementType.Contains("Framing", StringComparison.OrdinalIgnoreCase)))
            {
                // Not a wall/framing - return original point unchanged
                return centroid;
            }

            // ✅ TRACE LOGGING: Diagnosing why Ray Trace might be skipped
            try 
            {
                 string traceMsg = $"[{DateTime.Now:HH:mm:ss.fff}] [TRACE] Zone {zone.Id}: StructId={zone.StructuralElementIdValue}, ";
                 
                 Element hostElement = null;
                 if (zone.StructuralElementIdValue > 0)
                 {
                      // 1. Try active document
                      hostElement = _doc.GetElement(new ElementId(zone.StructuralElementIdValue));
                      
                      // 2. If not found, try linked documents
                      if (hostElement == null)
                      {
                           traceMsg += "Not in HostDoc, checking Links... ";
                           var failedLinks = new System.Collections.Generic.List<string>();
                           
                           var links = new FilteredElementCollector(_doc)
                               .OfClass(typeof(RevitLinkInstance))
                               .Cast<RevitLinkInstance>();
                               
                           foreach (var link in links)
                           {
                               try 
                               {
                                   var linkDoc = link.GetLinkDocument();
                                   if (linkDoc != null)
                                   {
                                       var verify = linkDoc.GetElement(new ElementId(zone.StructuralElementIdValue));
                                       if (verify != null)
                                       {
                                           hostElement = verify;
                                            traceMsg += $"Found in Link '{linkDoc.Title}', ";
                                           break;
                                       }
                                   }
                               }
                               catch { /* Ignore link access errors */ }
                           }
                      }
                      
                      traceMsg += $"HostFound={(hostElement != null)}, HostType={hostElement?.GetType().Name}, ";
                 }
                 else
                 {
                     traceMsg += "StructId INVALID/ZERO, ";
                 }

                 SafeFileLogger.SafeAppendText("damper_placement_trace.log", traceMsg + "\n");
                 
                  // ✅ STEP 1: Get the Insertion Point (True Center of Body)
                  // We do NOT use BBOX Center because connectors can skew the bounding box of the family.
                  // Doc Section 6.0.1.1: Use Damper's Insertion Point (LocationPoint) as center.
                  XYZ sourcePoint = null;
                  
                  // Need to check the *MEP Element* (damper) for its location point
                  string damperTrace = "";
                  Element damperElement = null;
                  
                  // Fix: zone.MepElementId is of type ElementId, so compare IntegerValue > 0
                  if (zone.MepElementId?.IntegerValue > 0) 
                  {
                      // Fix: zone.MepElementId is already ElementId, do not wrap in new ElementId()
                      damperElement = _doc.GetElement(zone.MepElementId);
                  }
                  
                  if (damperElement == null && zone.MepElementId?.IntegerValue > 0)
                  {
                      // Double check logic or fallback if needed
                  }
                  
                  // Use the 'mepElement' passed into the method context if available, else retrieve
                  // Actually, this method only takes 'zone' and 'centroid'. 
                  // We should trust 'centroid' meant "Initial Placement Point", but strictly we want Insertion Point.
                  // Retrieve the Damper Element from the Zone ID.
                  
                  if (damperElement is FamilyInstance fi && fi.Location is LocationPoint lp)
                  {
                      sourcePoint = lp.Point;
                      damperTrace = $"Using LocationPoint (Body Center): ({sourcePoint.X:F3}, {sourcePoint.Y:F3}, {sourcePoint.Z:F3})";
                  }
                  else if (damperElement is FamilyInstance fi2)
                  {
                       sourcePoint = fi2.GetTransform().Origin;
                       damperTrace = $"Using Transform Origin (Fallback): ({sourcePoint.X:F3}, {sourcePoint.Y:F3}, {sourcePoint.Z:F3})";
                  }
                  else
                  {
                       sourcePoint = centroid; // Ultimate fallback 
                       damperTrace = $"Using BBox Centroid (Last Resort): ({sourcePoint.X:F3}, {sourcePoint.Y:F3}, {sourcePoint.Z:F3})";
                  }
                  SafeFileLogger.SafeAppendText("damper_placement_trace.log", $"[{DateTime.Now:HH:mm:ss.fff}] [TRACE] SOURCE POINT: {damperTrace}\n");

                  // ✅ STEP 2: Use Location Curve Projection (User Request: "Use Wall Center Line")
                  if (hostElement != null && hostElement is Wall wall)
                  {
                       SafeFileLogger.SafeAppendText("damper_placement_trace.log", $"[{DateTime.Now:HH:mm:ss.fff}] [TRACE] 👉 ENTERING LOCATION CURVE LOGIC\n");

                       // Handle Linked Transforms
                       Transform linkTransform = Transform.Identity;
                       if (wall.Document.Title != _doc.Title)
                       {
                            // Find the Link Instance that contains this wall
                            var links = new FilteredElementCollector(_doc)
                                .OfClass(typeof(RevitLinkInstance))
                                .Cast<RevitLinkInstance>();
                            
                            foreach (var link in links)
                            {
                                var linkDoc = link.GetLinkDocument();
                                if (linkDoc != null && linkDoc.Title == wall.Document.Title)
                                {
                                    linkTransform = link.GetTotalTransform();
                                    SafeFileLogger.SafeAppendText("damper_placement_trace.log", $"[{DateTime.Now:HH:mm:ss.fff}] [TRACE] Found Link Transform: {link.Name}\n");
                                    break;
                                }
                            }
                       }
                       
                       LocationCurve locCurve = wall.Location as LocationCurve;
                       if (locCurve != null && locCurve.Curve != null)
                       {
                            // Transform the curve to Host Coordinates
                            Curve hostCurve = locCurve.Curve.CreateTransformed(linkTransform);
                            
                            // Project Source Point onto the Wall Curve
                            IntersectionResult result = hostCurve.Project(sourcePoint);
                            
                            if (result != null)
                            {
                                XYZ projectedPoint = result.XYZPoint;
                                
                                // The projected point is ON the wall centerline.
                                // However, the user might want the Vertical (Z) of the Damper kept.
                                // Usually, Wall Location Line is at Base Constraint (Z=0 or Level).
                                // So ProjectedPoint.Z will be at wall base.
                                // We want: X,Y from Wall Center, Z from Damper.
                                
                                XYZ finalPoint = new XYZ(projectedPoint.X, projectedPoint.Y, sourcePoint.Z);
                                
                                SafeFileLogger.SafeAppendText("damper_placement_trace.log", 
                                   $"[{DateTime.Now:HH:mm:ss.fff}] [TRACE] ✅ CURVE PROJECTION SUCCESS: " +
                                   $"Source=({sourcePoint.X:F3}, {sourcePoint.Y:F3}, {sourcePoint.Z:F3}) -> " +
                                   $"Projected=({projectedPoint.X:F3}, {projectedPoint.Y:F3}, {projectedPoint.Z:F3}) -> " +
                                   $"Final=({finalPoint.X:F3}, {finalPoint.Y:F3}, {finalPoint.Z:F3})\n");
                                
                                return finalPoint;
                            }
                           else
                           {
                               SafeFileLogger.SafeAppendText("damper_placement_trace.log", $"[{DateTime.Now:HH:mm:ss.fff}] [TRACE] ⚠️ PROJECTION FAILED (Point too far?)\n");
                           }
                      }
                      else
                      {
                           SafeFileLogger.SafeAppendText("damper_placement_trace.log", $"[{DateTime.Now:HH:mm:ss.fff}] [TRACE] ❌ NO LOCATION CURVE\n");
                      }
                 }
                 else
                 {
                      SafeFileLogger.SafeAppendText("damper_placement_trace.log", $"[{DateTime.Now:HH:mm:ss.fff}] [TRACE] ❌ SKIPPING LOGIC (Host is null or not Wall). Falling back to saved points.\n");
                 }
            }
            catch (Exception ex)
            {
                 SafeFileLogger.SafeAppendText("damper_placement_trace.log", $"[{DateTime.Now:HH:mm:ss.fff}] [TRACE] 💥 EXCEPTION: {ex.Message}\n");
            }

            // ✅ FALLBACK: Use saved WallCenterlinePoint if host not found (e.g. linked element unloaded)
            bool hasWallCenterline = (zone.WallCenterlinePointX != 0.0 || zone.WallCenterlinePointY != 0.0 || zone.WallCenterlinePointZ != 0.0);
            
            if (!hasWallCenterline)
            {
                if (!DeploymentConfiguration.DeploymentMode)
                {
                    SafeFileLogger.SafeAppendText("placement_debug.log",
                        $"[{DateTime.Now:HH:mm:ss.fff}] [DamperPlacementPoint] ⚠️ NO HOST & NO SAVED CENTERLINE: Zone {zone.Id} - KEEPING ORIGINAL CENTROID\n");
                }
                return centroid;
            }

            XYZ wallCenterline = new XYZ(zone.WallCenterlinePointX, zone.WallCenterlinePointY, zone.WallCenterlinePointZ);
            XYZ adjustedPoint = centroid;
            
            if (zone.HostOrientation == "X")
            {
                adjustedPoint = new XYZ(centroid.X, zone.WallCenterlinePointY, centroid.Z);
            }
            else if (zone.HostOrientation == "Y")
            {
                adjustedPoint = new XYZ(zone.WallCenterlinePointX, centroid.Y, centroid.Z);
            }
            else
            {
                adjustedPoint = wallCenterline;
            }
            
            return adjustedPoint;
        }
    }
}

