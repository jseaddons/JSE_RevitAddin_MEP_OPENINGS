using System;
#if !NET8_0_OR_GREATER
using System.Data.SQLite;
#endif
using System.Text.Json;
using JSE_RevitAddin_MEP_OPENINGS.Data;
using JSE_RevitAddin_MEP_OPENINGS.Models;

namespace JSE_RevitAddin_MEP_OPENINGS.Data.Repositories
{
    /// <summary>
    /// Repository for persisting opening conditions (clearances, opening types, etc.) in SQLite.
    /// </summary>
    public class ConditionRepository
    {
        private readonly SleeveDbContext _context;
        private readonly Action<string> _logger;

        public ConditionRepository(SleeveDbContext context, Action<string>? logger = null)
        {
            _context = context ?? throw new ArgumentNullException(nameof(context));
            _logger = logger ?? (_ => { });
        }

        /// <summary>
        /// ✅ FIX: Normalizes CombinedKey by stripping .xml extension to prevent duplicate rows
        /// </summary>
        private string NormalizeCombinedKey(string combinedKey)
        {
            if (string.IsNullOrWhiteSpace(combinedKey))
                return combinedKey;

            // Strip .xml extension if present anywhere in the key
            // Handle cases like "Ventilation_duct_accessories.xml_du..." or "Ventilation_duct_accessories"
            string normalized = combinedKey;
            
            // Check if key contains .xml (could be in middle or end)
            int xmlIndex = normalized.IndexOf(".xml", StringComparison.OrdinalIgnoreCase);
            if (xmlIndex >= 0)
            {
                // Remove .xml and anything after it that looks like a file extension pattern
                // But preserve the category suffix (e.g., "_duct_accessories")
                normalized = normalized.Substring(0, xmlIndex);
                
                // If we removed .xml from the middle, we might have broken the structure
                // Reconstruct: filterName_categorySuffix
                // Example: "Ventilation_duct_accessories.xml_du" -> "Ventilation_duct_accessories"
                // The category suffix should already be there before .xml
            }
            
            return normalized;
        }

        /// <summary>
        /// ⚠️⚠️⚠️ PROTECTED METHOD - DO NOT MODIFY WITHOUT EXTENSIVE TESTING ⚠️⚠️⚠️
        /// 
        /// Upserts (inserts or updates) opening conditions in SQLite database.
        /// 
        /// ✅ WORKING AS OF 2025-11-15: Conditions are successfully populating in SQLite database
        /// ✅ FIXED 2025-12-09: Normalizes CombinedKey to prevent duplicate rows with .xml suffix
        /// 
        /// ⚠️ CRITICAL VALIDATION CHECKS (DO NOT REMOVE):
        /// 1. FilterId must be > 0 (indicates filter was successfully registered)
        /// 2. CombinedKey must not be null or empty (required for unique identification)
        /// 3. Conditions object must not be null (required for data persistence)
        /// 
        /// 🔒 LOCKED BEHAVIOR:
        /// - All validation failures result in early return (no data saved)
        /// - All validation failures are logged
        /// - Existing conditions are updated, new conditions are inserted
        /// - All operations are logged for debugging
        /// - CombinedKey is normalized to prevent duplicates
        /// 
        /// ⚠️ DO NOT:
        /// - Remove validation checks
        /// - Remove early returns
        /// - Remove logging statements
        /// - Change method signature
        /// - Bypass GetConditionId check
        /// </summary>
        public void UpsertConditions(int filterId, string combinedKey, string normalizedCategory, OpeningConditions conditions)
        {
            // ⚠️ CRITICAL VALIDATION #1: FilterId must be valid (indicates filter registration succeeded)
            if (filterId <= 0)
            {
                _logger($"[SQLite] ❌ UpsertConditions skipped: FilterId={filterId} (must be > 0)");
                return; // ⚠️ DO NOT remove this early return - prevents invalid data from being saved
            }
            
            // ⚠️ CRITICAL VALIDATION #2: CombinedKey must be valid (required for unique identification)
            if (string.IsNullOrWhiteSpace(combinedKey))
            {
                _logger($"[SQLite] ❌ UpsertConditions skipped: CombinedKey is null or empty");
                return; // ⚠️ DO NOT remove this early return - prevents invalid data from being saved
            }
            
            // ⚠️ CRITICAL VALIDATION #3: Conditions object must not be null
            if (conditions == null)
            {
                _logger($"[SQLite] ❌ UpsertConditions skipped: Conditions is null");
                return; // ⚠️ DO NOT remove this early return - prevents null reference exceptions
            }

            // ✅ FIX: Normalize CombinedKey to prevent duplicate rows
            string normalizedKey = NormalizeCombinedKey(combinedKey);
            if (normalizedKey != combinedKey)
            {
                _logger($"[SQLite] ✅ Normalized CombinedKey: '{combinedKey}' -> '{normalizedKey}'");
            }

            // ✅ FIX: Clean up any duplicate rows with .xml suffix before lookup
            CleanupDuplicateConditions(normalizedKey);

            int? existingId = GetConditionId(normalizedKey);

            if (existingId.HasValue)
            {
                _logger($"[SQLite] Updating existing conditions for '{normalizedKey}' (ConditionId={existingId.Value}, FilterId={filterId})");
                UpdateConditions(existingId.Value, filterId, normalizedKey, normalizedCategory, conditions);
            }
            else
            {
                _logger($"[SQLite] Inserting new conditions for '{normalizedKey}' (FilterId={filterId}, Category={normalizedCategory})");
                InsertConditions(filterId, normalizedKey, normalizedCategory, conditions);
            }
        }

        /// <summary>
        /// ⚠️ PROTECTED METHOD - DO NOT MODIFY SQL STATEMENT WITHOUT TESTING ⚠️
        /// 
        /// Inserts new opening conditions into SQLite database.
        /// 
        /// ✅ WORKING AS OF 2025-11-15: Conditions are successfully populating in SQLite database
        /// 
        /// ⚠️ DO NOT modify the SQL INSERT statement without:
        /// 1. Testing with actual database
        /// 2. Verifying all columns match the database schema
        /// 3. Ensuring AddCommonParameters provides all required values
        /// </summary>
        private void InsertConditions(int filterId, string combinedKey, string normalizedCategory, OpeningConditions conditions)
        {
            using (var cmd = _context.Connection.CreateCommand())
            {
                // ⚠️ PROTECTED SQL: DO NOT modify without testing against actual database schema
                cmd.CommandText = @"
                    INSERT INTO Conditions (
                        FilterId,
                        CombinedKey,
                        Category,
                        RectNormal,
                        RectInsulated,
                        RoundNormal,
                        RoundInsulated,
                        PipesNormal,
                        PipesInsulated,
                        UseNominalDiameterForPipes,
                        CableTrayTop,
                        CableTrayTopInsulated,
                        CableTrayOther,
                        CableTrayOtherInsulated,
                        DuctAccessoryMepNormal,
                        DuctAccessoryMepInsulated,
                        DuctAccessoryOtherNormal,
                        DuctAccessoryOtherInsulated,
                        OpeningPrefs,
                        HorizontalLevel,
                        VerticalLevel,
                        CreationMode,
                        JoinOpeningsDistanceMm,
                        IgnoreArchitecturalFloors,
                        CircularToRectangularThresholdMm,
                        MinWallThicknessMm,
                        UpdatedAt
                    ) VALUES (
                        @FilterId,
                        @CombinedKey,
                        @Category,
                        @RectNormal,
                        @RectInsulated,
                        @RoundNormal,
                        @RoundInsulated,
                        @PipesNormal,
                        @PipesInsulated,
                        @UseNominalDiameterForPipes,
                        @CableTrayTop,
                        @CableTrayTopInsulated,
                        @CableTrayOther,
                        @CableTrayOtherInsulated,
                        @DuctAccessoryMepNormal,
                        @DuctAccessoryMepInsulated,
                        @DuctAccessoryOtherNormal,
                        @DuctAccessoryOtherInsulated,
                        @OpeningPrefs,
                        @HorizontalLevel,
                        @VerticalLevel,
                        @CreationMode,
                        @JoinOpeningsDistanceMm,
                        @IgnoreArchitecturalFloors,
                        @CircularToRectangularThresholdMm,
                        @MinWallThicknessMm,
                        CURRENT_TIMESTAMP
                    )";

                AddCommonParameters(cmd, filterId, combinedKey, normalizedCategory, conditions);
                cmd.ExecuteNonQuery();
                _logger($"[SQLite] ✅ Inserted conditions for '{combinedKey}'.");
            }
        }

        /// <summary>
        /// ⚠️ PROTECTED METHOD - DO NOT MODIFY SQL STATEMENT WITHOUT TESTING ⚠️
        /// 
        /// Updates existing opening conditions in SQLite database.
        /// 
        /// ✅ WORKING AS OF 2025-11-15: Conditions are successfully populating in SQLite database
        /// 
        /// ⚠️ DO NOT modify the SQL UPDATE statement without:
        /// 1. Testing with actual database
        /// 2. Verifying all columns match the database schema
        /// 3. Ensuring AddCommonParameters provides all required values
        /// </summary>
        private void UpdateConditions(int conditionId, int filterId, string combinedKey, string normalizedCategory, OpeningConditions conditions)
        {
            using (var cmd = _context.Connection.CreateCommand())
            {
                // ⚠️ PROTECTED SQL: DO NOT modify without testing against actual database schema
                cmd.CommandText = @"
                    UPDATE Conditions SET
                        FilterId = @FilterId,
                        CombinedKey = @CombinedKey,
                        Category = @Category,
                        RectNormal = @RectNormal,
                        RectInsulated = @RectInsulated,
                        RoundNormal = @RoundNormal,
                        RoundInsulated = @RoundInsulated,
                        PipesNormal = @PipesNormal,
                        PipesInsulated = @PipesInsulated,
                        UseNominalDiameterForPipes = @UseNominalDiameterForPipes,
                        CableTrayTop = @CableTrayTop,
                        CableTrayTopInsulated = @CableTrayTopInsulated,
                        CableTrayOther = @CableTrayOther,
                        CableTrayOtherInsulated = @CableTrayOtherInsulated,
                        DuctAccessoryMepNormal = @DuctAccessoryMepNormal,
                        DuctAccessoryMepInsulated = @DuctAccessoryMepInsulated,
                        DuctAccessoryOtherNormal = @DuctAccessoryOtherNormal,
                        DuctAccessoryOtherInsulated = @DuctAccessoryOtherInsulated,
                        OpeningPrefs = @OpeningPrefs,
                        HorizontalLevel = @HorizontalLevel,
                        VerticalLevel = @VerticalLevel,
                        CreationMode = @CreationMode,
                        JoinOpeningsDistanceMm = @JoinOpeningsDistanceMm,
                        IgnoreArchitecturalFloors = @IgnoreArchitecturalFloors,
                        CircularToRectangularThresholdMm = @CircularToRectangularThresholdMm,
                        MinWallThicknessMm = @MinWallThicknessMm,
                        UpdatedAt = CURRENT_TIMESTAMP
                    WHERE ConditionId = @ConditionId";

                cmd.Parameters.AddWithValue("@ConditionId", conditionId);
                AddCommonParameters(cmd, filterId, combinedKey, normalizedCategory, conditions);
                cmd.ExecuteNonQuery();
                _logger($"[SQLite] ✅ Updated conditions for '{combinedKey}'.");
            }
        }

        /// <summary>
        /// ⚠️ PROTECTED METHOD - DO NOT MODIFY WITHOUT TESTING ⚠️
        /// 
        /// Adds common parameters to SQLite command for both INSERT and UPDATE operations.
        /// 
        /// ✅ WORKING AS OF 2025-11-15: Conditions are successfully populating in SQLite database
        /// 
        /// ⚠️ DO NOT modify parameter names or values without:
        /// 1. Verifying they match the SQL INSERT/UPDATE statements
        /// 2. Testing with actual database
        /// 3. Ensuring all required columns are populated
        /// 
        /// This method is used by BOTH InsertConditions and UpdateConditions - any changes affect both operations.
        /// 
        /// 🔒 LOCKED PARAMETERS (DO NOT REMOVE):
        /// - @FilterId, @CombinedKey, @Category (required for identification)
        /// - All clearance parameters (RectNormal, RectInsulated, RoundNormal, etc.)
        /// - OpeningPrefs (JSON serialized)
        /// - LevelConstraints (HorizontalLevel, VerticalLevel)
        /// - CreationMode
        /// </summary>
        private void AddCommonParameters(SQLiteCommand cmd, int filterId, string combinedKey, string normalizedCategory, OpeningConditions conditions)
        {
            // ⚠️ PROTECTED: DO NOT remove or modify these parameter assignments
            // They must match the SQL INSERT/UPDATE statements exactly
            cmd.Parameters.AddWithValue("@FilterId", filterId);
            cmd.Parameters.AddWithValue("@CombinedKey", combinedKey);
            cmd.Parameters.AddWithValue("@Category", normalizedCategory ?? string.Empty);

            var clearance = conditions.ClearanceSettings ?? new ClearanceSettings();
            cmd.Parameters.AddWithValue("@RectNormal", clearance.RectangularNormal);
            cmd.Parameters.AddWithValue("@RectInsulated", clearance.RectangularInsulated);
            cmd.Parameters.AddWithValue("@RoundNormal", clearance.RoundNormal);
            cmd.Parameters.AddWithValue("@RoundInsulated", clearance.RoundInsulated);
            cmd.Parameters.AddWithValue("@PipesNormal", clearance.PipesNormal);
            cmd.Parameters.AddWithValue("@PipesInsulated", clearance.PipesInsulated);
            cmd.Parameters.AddWithValue("@UseNominalDiameterForPipes", clearance.UseNominalDiameterForPipes ? 1 : 0);
            cmd.Parameters.AddWithValue("@CableTrayTop", clearance.CableTrayTop);
            cmd.Parameters.AddWithValue("@CableTrayTopInsulated", clearance.CableTrayTopInsulated);
            cmd.Parameters.AddWithValue("@CableTrayOther", clearance.CableTrayOther);
            cmd.Parameters.AddWithValue("@CableTrayOtherInsulated", clearance.CableTrayOtherInsulated);
            cmd.Parameters.AddWithValue("@DuctAccessoryMepNormal", clearance.DuctAccessoryMepNormal);
            cmd.Parameters.AddWithValue("@DuctAccessoryMepInsulated", clearance.DuctAccessoryMepInsulated);
            cmd.Parameters.AddWithValue("@DuctAccessoryOtherNormal", clearance.DuctAccessoryOtherNormal);
            cmd.Parameters.AddWithValue("@DuctAccessoryOtherInsulated", clearance.DuctAccessoryOtherInsulated);

            var openingPrefs = conditions.OpeningTypePreferences ?? new OpeningTypePreferences();
            
            // ✅ NORMALIZE: Trim and normalize property values before serialization
            if (openingPrefs != null)
            {
                openingPrefs.RoundDucts = NormalizeOpeningType(openingPrefs.RoundDucts);
                openingPrefs.Pipes = NormalizeOpeningType(openingPrefs.Pipes);
            }
            
            // ✅ NORMALIZE: Serialize with compact formatting (no indentation) and trim whitespace
            var options = new JsonSerializerOptions { WriteIndented = false };
            string openingPrefsJson = JsonSerializer.Serialize(openingPrefs, options)?.Trim() ?? string.Empty;
            cmd.Parameters.AddWithValue("@OpeningPrefs", openingPrefsJson);

            var levelConstraints = conditions.LevelConstraints ?? new LevelConstraints();
            cmd.Parameters.AddWithValue("@HorizontalLevel", levelConstraints.HorizontalLevel ?? string.Empty);
            cmd.Parameters.AddWithValue("@VerticalLevel", levelConstraints.VerticalLevel ?? string.Empty);

            cmd.Parameters.AddWithValue("@CreationMode", conditions.CreationMode ?? string.Empty);

            // Per-filter/category behavioral settings (backed primarily by Conditions table)
            // Join distance: millimeters, mirrors SettingsModel.JoinOpeningsDistance but stored per filter/category.
            cmd.Parameters.AddWithValue("@JoinOpeningsDistanceMm", conditions.JoinOpeningsDistanceMm);
            // Ignore architectural floors flag: stored as integer 0/1.
            cmd.Parameters.AddWithValue("@IgnoreArchitecturalFloors", conditions.IgnoreArchitecturalFloors ? 1 : 0);
            // Circular-to-rectangular threshold: prefer explicit property, fall back to sizing settings if set.
            double thresholdMm = conditions.CircularToRectangularThresholdMm;
            if (thresholdMm <= 0 && conditions.SizingSettings != null)
            {
                thresholdMm = conditions.SizingSettings.CircularToRectangularThresholdMm;
            }
            cmd.Parameters.AddWithValue("@CircularToRectangularThresholdMm", thresholdMm);
            cmd.Parameters.AddWithValue("@MinWallThicknessMm", conditions.MinWallThicknessMm);
        }

        public OpeningConditions GetConditions(string combinedKey)
        {
            if (string.IsNullOrWhiteSpace(combinedKey))
                return null;

            // ✅ FIX: Normalize CombinedKey before database lookup
            string normalizedKey = NormalizeCombinedKey(combinedKey);

            using (var cmd = _context.Connection.CreateCommand())
            {
                    cmd.CommandText = @"
                    SELECT 
                        c.FilterId,
                        c.CombinedKey,
                        c.Category,
                        c.RectNormal,
                        c.RectInsulated,
                        c.RoundNormal,
                        c.RoundInsulated,
                        c.PipesNormal,
                        c.PipesInsulated,
                        c.UseNominalDiameterForPipes,
                        c.CableTrayTop,
                        c.CableTrayTopInsulated,
                        c.CableTrayOther,
                        c.CableTrayOtherInsulated,
                        c.DuctAccessoryMepNormal,
                        c.DuctAccessoryMepInsulated,
                        c.DuctAccessoryOtherNormal,
                        c.DuctAccessoryOtherInsulated,
                        c.OpeningPrefs,
                        c.HorizontalLevel,
                        c.VerticalLevel,
                        c.CreationMode,
                        c.JoinOpeningsDistanceMm,
                        c.IgnoreArchitecturalFloors,
                        c.CircularToRectangularThresholdMm,
                        c.MinWallThicknessMm,
                        f.FilterName,
                        f.Category AS FilterCategory,
                        c.UpdatedAt
                    FROM Conditions c
                    INNER JOIN Filters f ON c.FilterId = f.FilterId
                    WHERE c.CombinedKey = @CombinedKey
                    LIMIT 1";
                cmd.Parameters.AddWithValue("@CombinedKey", normalizedKey);

                using (var reader = cmd.ExecuteReader())
                {
                    if (!reader.Read())
                        return null;

                    var conditions = new OpeningConditions
                    {
                        FilterName = reader["FilterName"]?.ToString() ?? string.Empty,
                        Category = reader["FilterCategory"]?.ToString() ?? string.Empty,
                        LastModified = reader["UpdatedAt"] is DBNull
                            ? DateTime.Now
                            : Convert.ToDateTime(reader["UpdatedAt"]),
                        ClearanceSettings = new ClearanceSettings
                        {
                            RectangularNormal = GetDouble(reader, "RectNormal"),
                            RectangularInsulated = GetDouble(reader, "RectInsulated"),
                            RoundNormal = GetDouble(reader, "RoundNormal"),
                            RoundInsulated = GetDouble(reader, "RoundInsulated"),
                            PipesNormal = GetDouble(reader, "PipesNormal"),
                            PipesInsulated = GetDouble(reader, "PipesInsulated"),
                            UseNominalDiameterForPipes = reader["UseNominalDiameterForPipes"] is DBNull
                                ? false
                                : (Convert.ToInt32(reader["UseNominalDiameterForPipes"]) != 0),
                            CableTrayTop = GetDouble(reader, "CableTrayTop"),
                            CableTrayTopInsulated = GetDouble(reader, "CableTrayTopInsulated"),
                            CableTrayOther = GetDouble(reader, "CableTrayOther"),
                            CableTrayOtherInsulated = GetDouble(reader, "CableTrayOtherInsulated"),
                            DuctAccessoryMepNormal = GetDouble(reader, "DuctAccessoryMepNormal"),
                            DuctAccessoryMepInsulated = GetDouble(reader, "DuctAccessoryMepInsulated"),
                            DuctAccessoryOtherNormal = GetDouble(reader, "DuctAccessoryOtherNormal"),
                            DuctAccessoryOtherInsulated = GetDouble(reader, "DuctAccessoryOtherInsulated")
                        },
                        JoinOpeningsDistanceMm = GetDouble(reader, "JoinOpeningsDistanceMm"),
                        IgnoreArchitecturalFloors = (reader["IgnoreArchitecturalFloors"] is DBNull
                            ? 0
                            : Convert.ToInt32(reader["IgnoreArchitecturalFloors"])) != 0,
                        CircularToRectangularThresholdMm = GetDouble(reader, "CircularToRectangularThresholdMm"),
                        MinWallThicknessMm = GetDouble(reader, "MinWallThicknessMm")
                    };

                    // Also hydrate SizingSettings from the stored threshold so sizing/planning
                    // code that only looks at SizingSettings still sees the correct value.
                    conditions.SizingSettings = conditions.SizingSettings ?? new SizingSettings();
                    conditions.SizingSettings.CircularToRectangularThresholdMm = conditions.CircularToRectangularThresholdMm;

                    var openingPrefsJson = reader["OpeningPrefs"]?.ToString()?.Trim();
                    if (!string.IsNullOrWhiteSpace(openingPrefsJson))
                    {
                        try
                        {
                            conditions.OpeningTypePreferences = JsonSerializer.Deserialize<OpeningTypePreferences>(openingPrefsJson)
                                ?? new OpeningTypePreferences();
                            
                            // ✅ NORMALIZE: Normalize property values after deserialization
                            if (conditions.OpeningTypePreferences != null)
                            {
                                conditions.OpeningTypePreferences.RoundDucts = NormalizeOpeningType(conditions.OpeningTypePreferences.RoundDucts);
                                conditions.OpeningTypePreferences.Pipes = NormalizeOpeningType(conditions.OpeningTypePreferences.Pipes);
                            }
                        }
                        catch
                        {
                            conditions.OpeningTypePreferences = new OpeningTypePreferences();
                        }
                    }

                    conditions.LevelConstraints = new LevelConstraints
                    {
                        HorizontalLevel = reader["HorizontalLevel"]?.ToString() ?? "Host Level",
                        VerticalLevel = reader["VerticalLevel"]?.ToString() ?? "Host Level"
                    };

                    conditions.CreationMode = reader["CreationMode"]?.ToString() ?? "Opening";
                    return conditions;
                }
            }
        }

        /// <summary>
        /// ✅ FIX: Gets ConditionId, checking both normalized key and key with .xml suffix
        /// This handles legacy duplicate rows that may exist
        /// </summary>
        private int? GetConditionId(string combinedKey)
        {
            // ✅ FIX: Normalize the key first
            string normalizedKey = NormalizeCombinedKey(combinedKey);
            
            using (var cmd = _context.Connection.CreateCommand())
            {
                // ✅ FIX: Check for normalized key first (preferred)
                cmd.CommandText = @"
                    SELECT ConditionId FROM Conditions
                    WHERE CombinedKey = @CombinedKey
                    LIMIT 1";
                cmd.Parameters.AddWithValue("@CombinedKey", normalizedKey);

                var result = cmd.ExecuteScalar();
                if (result != null && int.TryParse(result.ToString(), out int conditionId))
                {
                    return conditionId;
                }
            }

            // ✅ FIX: If not found with normalized key, check for key with .xml suffix (legacy)
            // This handles cases where duplicate exists with .xml suffix
            if (normalizedKey != combinedKey)
            {
                using (var cmd = _context.Connection.CreateCommand())
                {
                    cmd.CommandText = @"
                        SELECT ConditionId FROM Conditions
                        WHERE CombinedKey = @CombinedKey
                        LIMIT 1";
                    cmd.Parameters.AddWithValue("@CombinedKey", combinedKey);

                    var result = cmd.ExecuteScalar();
                    if (result != null && int.TryParse(result.ToString(), out int conditionId))
                    {
                        _logger($"[SQLite] ⚠️ Found legacy row with .xml suffix for '{combinedKey}', will be cleaned up");
                        return conditionId;
                    }
                }
            }

            return null;
        }

        /// <summary>
        /// ✅ FIX: Removes duplicate condition rows that have .xml suffix in CombinedKey
        /// Keeps only the normalized version (without .xml)
        /// </summary>
        private void CleanupDuplicateConditions(string normalizedKey)
        {
            try
            {
                using (var cmd = _context.Connection.CreateCommand())
                {
                    // Find all rows that start with normalizedKey but have .xml in them
                    // Example: normalizedKey = "Ventilation_duct_accessories"
                    // Find: "Ventilation_duct_accessories.xml_du..." or similar
                    cmd.CommandText = @"
                        SELECT ConditionId, CombinedKey FROM Conditions
                        WHERE CombinedKey LIKE @Pattern
                          AND CombinedKey != @NormalizedKey";
                    cmd.Parameters.AddWithValue("@Pattern", normalizedKey + ".xml%");
                    cmd.Parameters.AddWithValue("@NormalizedKey", normalizedKey);

                    var duplicatesToDelete = new List<int>();
                    using (var reader = cmd.ExecuteReader())
                    {
                        while (reader.Read())
                        {
                            int conditionId = reader.GetInt32(0);
                            string duplicateKey = reader.GetString(1);
                            duplicatesToDelete.Add(conditionId);
                            _logger($"[SQLite] 🧹 Found duplicate condition row: ConditionId={conditionId}, CombinedKey='{duplicateKey}' (will be deleted)");
                        }
                    }

                    // Delete duplicate rows
                    if (duplicatesToDelete.Count > 0)
                    {
                        using (var deleteCmd = _context.Connection.CreateCommand())
                        {
                            string idList = string.Join(",", duplicatesToDelete);
                            deleteCmd.CommandText = $@"
                                DELETE FROM Conditions
                                WHERE ConditionId IN ({idList})";
                            
                            int deleted = deleteCmd.ExecuteNonQuery();
                            _logger($"[SQLite] ✅ Cleaned up {deleted} duplicate condition row(s) for normalized key '{normalizedKey}'");
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                // Log but don't fail - cleanup is best effort
                _logger($"[SQLite] ⚠️ Error during duplicate cleanup for '{normalizedKey}': {ex.Message}");
            }
        }

        /// <summary>
        /// ✅ NORMALIZE: Normalizes opening type text (trim, capitalize first letter)
        /// Ensures consistent formatting: "Circular", "Rectangular", etc.
        /// </summary>
        private static string NormalizeOpeningType(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return "Circular"; // Default
            
            // Trim whitespace and normalize casing (first letter uppercase, rest lowercase)
            var trimmed = value.Trim();
            if (trimmed.Length == 0)
                return "Circular";
            
            // Capitalize first letter, lowercase the rest
            return char.ToUpperInvariant(trimmed[0]) + trimmed.Substring(1).ToLowerInvariant();
        }
        
        private static double GetDouble(SQLiteDataReader reader, string columnName)
        {
            var ordinal = reader.GetOrdinal(columnName);
            if (reader.IsDBNull(ordinal))
                return 0.0;

            return Convert.ToDouble(reader.GetValue(ordinal));
        }
    }
}

