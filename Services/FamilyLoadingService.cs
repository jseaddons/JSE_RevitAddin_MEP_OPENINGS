using System;
using System.IO;
using System.Linq;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace JSE_RevitAddin_MEP_OPENINGS.Services
{
    /// <summary>
    /// Service for automatically loading opening families from the Resources folder
    /// </summary>
    public class FamilyLoadingService
    {
        private readonly Document _document;
        private readonly string _resourcesPath;

        public FamilyLoadingService(Document document)
        {
            _document = document;
            
            // Get the Resources folder path relative to the add-in location
            var addinLocation = System.Reflection.Assembly.GetExecutingAssembly().Location;
            var addinDirectory = Path.GetDirectoryName(addinLocation);
            _resourcesPath = Path.Combine(addinDirectory ?? "", "Resources");
        }

        /// <summary>
        /// Loads all required opening families from the Resources folder
        /// </summary>
        public bool LoadAllRequiredFamilies()
        {
            try
            {
                DebugLogger.Info("[FamilyLoadingService] Starting automatic family loading from Resources folder");
                
                if (!Directory.Exists(_resourcesPath))
                {
                    DebugLogger.Error($"[FamilyLoadingService] Resources folder not found at: {_resourcesPath}");
                    return false;
                }

                // List of required universal opening families
                var requiredFamilies = new[]
                {
                    "OpeningOnWall-Rectangular.rfa",
                    "OpeningOnWall-Circular.rfa", 
                    "OpeningOnSlab-Rectangular.rfa",
                    "OpeningOnSlab-Circular.rfa"
                };

                int loadedCount = 0;
                int alreadyLoadedCount = 0;

                foreach (var familyFile in requiredFamilies)
                {
                    var familyPath = Path.Combine(_resourcesPath, familyFile);
                    
                    if (!File.Exists(familyPath))
                    {
                        DebugLogger.Warning($"[FamilyLoadingService] Family file not found: {familyFile}");
                        continue;
                    }

                    // Check if family is already loaded
                    var familyName = Path.GetFileNameWithoutExtension(familyFile);
                    if (IsFamilyAlreadyLoaded(familyName))
                    {
                        DebugLogger.Info($"[FamilyLoadingService] Family already loaded: {familyName}");
                        alreadyLoadedCount++;
                        continue;
                    }

                    // Load the family
                    if (LoadFamilyFromFile(familyPath))
                    {
                        DebugLogger.Info($"[FamilyLoadingService] Successfully loaded family: {familyName}");
                        loadedCount++;
                    }
                    else
                    {
                        DebugLogger.Error($"[FamilyLoadingService] Failed to load family: {familyName}");
                    }
                }

                DebugLogger.Info($"[FamilyLoadingService] Family loading complete. Loaded: {loadedCount}, Already loaded: {alreadyLoadedCount}");
                return loadedCount > 0 || alreadyLoadedCount == requiredFamilies.Length;
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[FamilyLoadingService] Error loading families: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Checks if a family is already loaded in the document
        /// </summary>
        private bool IsFamilyAlreadyLoaded(string familyName)
        {
            try
            {
                var families = new FilteredElementCollector(_document)
                    .OfClass(typeof(Family))
                    .Cast<Family>()
                    .Where(f => f.Name == familyName)
                    .ToList();

                return families.Any();
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[FamilyLoadingService] Error checking if family is loaded: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Loads a family from file path
        /// </summary>
        private bool LoadFamilyFromFile(string familyPath)
        {
            try
            {
                using (var t = new Transaction(_document, "Load Opening Family"))
                {
                    if (t.Start() == TransactionStatus.Started)
                    {
                        bool loaded = _document.LoadFamily(familyPath);
                        t.Commit();
                        return loaded;
                    }
                    else
                    {
                        DebugLogger.Error($"[FamilyLoadingService] Failed to start transaction for loading family");
                        return false;
                    }
                }
            }
            catch (Exception ex)
            {
                DebugLogger.Error($"[FamilyLoadingService] Error loading family from {familyPath}: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Gets the Resources folder path
        /// </summary>
        public string ResourcesPath => _resourcesPath;

        /// <summary>
        /// Checks if Resources folder exists and contains required families
        /// </summary>
        public bool IsResourcesFolderValid()
        {
            if (!Directory.Exists(_resourcesPath))
            {
                return false;
            }

            var requiredFamilies = new[]
            {
                "OpeningOnWall-Rectangular.rfa",
                "OpeningOnWall-Circular.rfa", 
                "OpeningOnSlab-Rectangular.rfa",
                "OpeningOnSlab-Circular.rfa"
            };

            return requiredFamilies.All(family => File.Exists(Path.Combine(_resourcesPath, family)));
        }
    }
}