-- ============================================================================
-- CROSS-CATEGORY COMBINED SLEEVES - DATABASE SCHEMA
-- ============================================================================
-- Purpose: Support combined sleeves that group individual and cluster sleeves
--          across different MEP categories (Pipes, Ducts, Cable Trays)
-- Author: Agent A - Database & Data Layer
-- Date: 2025-12-16
-- ============================================================================

-- ============================================================================
-- TABLE: CombinedSleeves
-- ============================================================================
-- Stores cross-category combined sleeve data
-- A combined sleeve is a single sleeve that encompasses multiple individual
-- and/or cluster sleeves from different categories
-- ============================================================================

CREATE TABLE IF NOT EXISTS CombinedSleeves (
    -- Primary Key
    CombinedSleeveId INTEGER PRIMARY KEY AUTOINCREMENT,
    
    -- Revit Element Reference
    CombinedInstanceId INTEGER NOT NULL UNIQUE, -- Revit ElementId.IntegerValue
    
    -- Foreign Keys
    ComboId INTEGER NOT NULL,
    FilterId INTEGER NOT NULL,
    
    -- Categories (comma-separated or JSON array)
    -- Example: "Pipes,Ducts" or "Pipes,Ducts,Cable Trays"
    Categories TEXT NOT NULL,
    
    -- Bounding Box (World Coordinates)
    BoundingBoxMinX REAL NOT NULL,
    BoundingBoxMinY REAL NOT NULL,
    BoundingBoxMinZ REAL NOT NULL,
    BoundingBoxMaxX REAL NOT NULL,
    BoundingBoxMaxY REAL NOT NULL,
    BoundingBoxMaxZ REAL NOT NULL,
    
    -- Dimensions (in feet)
    CombinedWidth REAL NOT NULL,
    CombinedHeight REAL NOT NULL,
    CombinedDepth REAL NOT NULL,
    
    -- Placement (World Coordinates)
    PlacementX REAL NOT NULL,
    PlacementY REAL NOT NULL,
    PlacementZ REAL NOT NULL,
    RotationAngleDeg REAL DEFAULT 0.0,
    
    -- Host Information
    HostType TEXT, -- 'Wall', 'Floor', 'Roof', 'Framing'
    HostOrientation TEXT, -- 'X', 'Y', 'Z'
    
    -- Corner Coordinates (World Space) - For Parameter Transfer
    -- Calculated after placement, stored for downstream processes
    Corner1X REAL DEFAULT 0.0,
    Corner1Y REAL DEFAULT 0.0,
    Corner1Z REAL DEFAULT 0.0,
    Corner2X REAL DEFAULT 0.0,
    Corner2Y REAL DEFAULT 0.0,
    Corner2Z REAL DEFAULT 0.0,
    Corner3X REAL DEFAULT 0.0,
    Corner3Y REAL DEFAULT 0.0,
    Corner3Z REAL DEFAULT 0.0,
    Corner4X REAL DEFAULT 0.0,
    Corner4Y REAL DEFAULT 0.0,
    Corner4Z REAL DEFAULT 0.0,
    
    -- Timestamps
    CreatedAt DATETIME DEFAULT CURRENT_TIMESTAMP,
    UpdatedAt DATETIME DEFAULT CURRENT_TIMESTAMP,
    
    -- Foreign Key Constraints
    FOREIGN KEY (ComboId) REFERENCES FileCombos(ComboId) ON DELETE CASCADE,
    FOREIGN KEY (FilterId) REFERENCES Filters(FilterId) ON DELETE CASCADE
);

-- ============================================================================
-- TABLE: CombinedSleeveConstituents
-- ============================================================================
-- Junction table linking combined sleeves to their constituent sleeves
-- A constituent can be either an individual sleeve (ClashZone) or a cluster sleeve
-- ============================================================================

CREATE TABLE IF NOT EXISTS CombinedSleeveConstituents (
    -- Primary Key
    ConstituentId INTEGER PRIMARY KEY AUTOINCREMENT,
    
    -- Foreign Key to CombinedSleeves
    CombinedSleeveId INTEGER NOT NULL,
    
    -- Constituent Type: 'Individual' or 'Cluster'
    ConstituentType TEXT NOT NULL CHECK (ConstituentType IN ('Individual', 'Cluster')),
    
    -- For Individual Sleeves (ClashZones)
    ClashZoneId INTEGER, -- FK to ClashZones.ClashZoneId
    ClashZoneGuid TEXT, -- For deterministic lookup
    
    -- For Cluster Sleeves
    ClusterSleeveId INTEGER, -- FK to ClusterSleeves.ClusterSleeveId
    ClusterInstanceId INTEGER, -- Revit ElementId.IntegerValue
    
    -- Category of this constituent
    Category TEXT NOT NULL, -- 'Pipes', 'Ducts', 'Cable Trays', 'Conduits'
    
    -- Timestamp
    CreatedAt DATETIME DEFAULT CURRENT_TIMESTAMP,
    
    -- Foreign Key Constraints
    FOREIGN KEY (CombinedSleeveId) REFERENCES CombinedSleeves(CombinedSleeveId) ON DELETE CASCADE,
    FOREIGN KEY (ClashZoneId) REFERENCES ClashZones(ClashZoneId) ON DELETE SET NULL,
    FOREIGN KEY (ClusterSleeveId) REFERENCES ClusterSleeves(ClusterSleeveId) ON DELETE SET NULL,
    
    -- Constraint: Either ClashZoneId OR ClusterSleeveId must be set, not both
    CHECK (
        (ConstituentType = 'Individual' AND ClashZoneId IS NOT NULL AND ClusterSleeveId IS NULL) OR
        (ConstituentType = 'Cluster' AND ClusterSleeveId IS NOT NULL AND ClashZoneId IS NULL)
    )
);

-- ============================================================================
-- INDEXES
-- ============================================================================
-- Optimize query performance for common access patterns
-- ============================================================================

-- Index for querying constituents by combined sleeve
CREATE INDEX IF NOT EXISTS idx_combined_constituents_combined 
ON CombinedSleeveConstituents(CombinedSleeveId);

-- Index for querying constituents by individual sleeve
CREATE INDEX IF NOT EXISTS idx_combined_constituents_clash 
ON CombinedSleeveConstituents(ClashZoneId);

-- Index for querying constituents by cluster sleeve
CREATE INDEX IF NOT EXISTS idx_combined_constituents_cluster 
ON CombinedSleeveConstituents(ClusterSleeveId);

-- Index for querying combined sleeves by combo
CREATE INDEX IF NOT EXISTS idx_combined_sleeves_combo 
ON CombinedSleeves(ComboId);

-- Index for querying combined sleeves by filter
CREATE INDEX IF NOT EXISTS idx_combined_sleeves_filter 
ON CombinedSleeves(FilterId);

-- Index for querying combined sleeves by instance ID
CREATE INDEX IF NOT EXISTS idx_combined_sleeves_instance 
ON CombinedSleeves(CombinedInstanceId);

-- Index for querying constituents by category
CREATE INDEX IF NOT EXISTS idx_combined_constituents_category 
ON CombinedSleeveConstituents(Category);

-- ============================================================================
-- RESOLUTION FLAGS (Add to existing tables)
-- ============================================================================
-- Track which sleeves have been combined to avoid duplicates
-- ============================================================================

-- Add IsCombinedResolved flag to ClashZones table (if not exists)
-- This will be executed as part of schema migration
-- ALTER TABLE ClashZones ADD COLUMN IsCombinedResolved INTEGER DEFAULT 0;

-- Add IsCombinedResolved flag to ClusterSleeves table (if not exists)
-- This will be executed as part of schema migration
-- ALTER TABLE ClusterSleeves ADD COLUMN IsCombinedResolved INTEGER DEFAULT 0;

-- ============================================================================
-- NOTES
-- ============================================================================
-- 1. Categories field stores comma-separated list for simplicity
--    Alternative: JSON array for more complex queries
-- 2. Corner coordinates follow same pattern as ClashZones and ClusterSleeves
-- 3. Constituent type constraint ensures data integrity
-- 4. Cascade deletes ensure referential integrity
-- 5. Indexes optimize common query patterns (by combo, filter, constituents)
-- ============================================================================
