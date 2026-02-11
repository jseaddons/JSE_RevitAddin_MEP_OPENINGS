using System;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.DB;
using JSE_RevitAddin_MEP_OPENINGS.Utils;

namespace JSE_RevitAddin_MEP_OPENINGS.Services.MultiFloor
{
    /// <summary>
    /// Caches resources that are shared across all floors
    /// Loaded once at startup, accessed by all parallel threads
    /// </summary>
    public class SharedResourceCache
    {
        private readonly Document _doc;
        
        public Dictionary<string, FamilySymbol> Symbols { get; private set; }
        public Dictionary<ElementId, Wall> Walls { get; private set; }
        public Dictionary<ElementId, Floor> Floors { get; private set; }
        public bool IsInitialized { get; private set; }
        
        public SharedResourceCache(Document doc)
        {
            _doc = doc;
            Symbols = new Dictionary<string, FamilySymbol>();
            Walls = new Dictionary<ElementId, Wall>();
            Floors = new Dictionary<ElementId, Floor>();
        }
        
        /// <summary>
        /// Load all shared resources once
        /// </summary>
        public void Initialize()
        {
            if (IsInitialized) return;
            
            SafeFileLogger.SafeAppendText("cache.log",
                $"[{DateTime.Now}] 🔧 Initializing shared resource cache...\n");
            
            // Load and activate sleeve symbols
            var symbolCollector = new FilteredElementCollector(_doc)
                .OfClass(typeof(FamilySymbol))
                .Cast<FamilySymbol>();
            
            foreach (var symbol in symbolCollector)
            {
                if (symbol.FamilyName.Contains("Sleeve") || 
                    symbol.FamilyName.Contains("Opening"))
                {
                    if (!symbol.IsActive)
                    {
                        symbol.Activate();
                    }
                    Symbols[symbol.Name] = symbol;
                }
            }
            
            // Cache walls (read-only access, safe for parallel use)
            var wallCollector = new FilteredElementCollector(_doc)
                .OfClass(typeof(Wall))
                .Cast<Wall>();
            
            foreach (var w in wallCollector)
            {
                Walls[w.Id] = w;
            }
            
            // Cache floors
            var floorCollector = new FilteredElementCollector(_doc)
                .OfClass(typeof(Floor))
                .Cast<Floor>();
            
            foreach (var f in floorCollector)
            {
                Floors[f.Id] = f;
            }
            
            IsInitialized = true;
            
            SafeFileLogger.SafeAppendText("cache.log",
                $"[{DateTime.Now}] ✅ Cache initialized: {Symbols.Count} symbols, {Walls.Count} walls, {Floors.Count} floors\n");
        }
        
        public FamilySymbol? GetSymbol(string name)
        {
            if (!IsInitialized)
                throw new InvalidOperationException("Cache not initialized. Call Initialize() first.");
            
            return Symbols.TryGetValue(name, out var symbol) ? symbol : null;
        }
    }
}
