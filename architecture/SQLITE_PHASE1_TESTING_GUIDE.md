# SQLite Phase 1 Testing Guide

## 🧪 How to Test Dual-Write Mode

After building and running refresh, here's what you should see and how to verify everything is working.

---

## 📋 Step-by-Step Testing Process

### **Step 1: Build the Project**
1. Build the solution in Visual Studio
2. Ensure there are no compilation errors
3. Deploy to Revit (or run from Visual Studio)

### **Step 2: Run Refresh Operation**
1. Open Revit with your test project
2. Open the MEP Openings add-in
3. Select:
   - **Filter**: Choose a filter (e.g., "ALL" or "Plumbing")
   - **MEP Categories**: Select at least one category (e.g., "Pipes", "Ducts")
   - **Reference Files**: Select linked MEP files
   - **Host Files**: Select linked host files (Walls, Floors)
4. Click **"Refresh"** or **"Process Clash Zone"**

### **Step 3: Check Logs for SQLite Activity**

#### **What to Look For in Debug Logs:**

Open the debug log file (usually in `%APPDATA%\JSE_MEP_Openings\Logs\` or check `SafeFileLogger` output):

**✅ Success Indicators:**
```
[SQLite] ✅ Schema created/verified for database: C:\Users\{username}\AppData\Roaming\JSE_MEP_Openings\Projects\{ProjectName}\Filters\{ProjectName}_SleevePersistence.db
[CLASH-ZONE-PERSISTENCE] ✅ SQLite dual-write enabled
[SQLite] ✅ Inserted/updated {count} clash zones for filter '{filterName}', category '{category}'
[CLASH-ZONE-PERSISTENCE] ✅ Dual-write to SQLite: {count} zones for '{category}'
```

**⚠️ Warning (Non-Critical):**
```
[CLASH-ZONE-PERSISTENCE] ⚠️ SQLite initialization failed, continuing with XML-only: {error}
```
*This is OK - XML writes will still work, but SQLite won't be written to.*

**❌ Error (Should Not See):**
```
[CLASH-ZONE-PERSISTENCE] ⚠️ SQLite dual-write failed (XML write succeeded): {error}
```
*This is OK - XML writes succeeded, SQLite write failed (non-blocking).*

---

## 🔍 Step 4: Verify Database File Created

### **Database Location:**
```
C:\Users\{username}\AppData\Roaming\JSE_MEP_Openings\Projects\{ProjectName}\Filters\{ProjectName}_SleevePersistence.db
```

**Example:**
```
C:\Users\jse2084\AppData\Roaming\JSE_MEP_Openings\Projects\Shared Test Model -00001\Filters\Shared Test Model -00001_SleevePersistence.db
```

### **How to Check:**
1. Navigate to the Filters directory for your project
2. Look for a `.db` file with your project name
3. File should be created **after** first refresh operation
4. File size should be > 0 bytes (even if small)

---

## 🗄️ Step 5: Inspect SQLite Database

### **Option A: Using DB Browser for SQLite (Recommended)**

1. **Download**: https://sqlitebrowser.org/
2. **Open**: File → Open Database → Select `{ProjectName}_SleevePersistence.db`
3. **Browse Data** tab → Select tables to view

### **Option B: Using SQLiteStudio**

1. **Download**: https://sqlitestudio.pl/
2. **Open**: Database → Add Database → Select `.db` file
3. **Browse**: Expand database → Tables → View data

### **Option C: Using VS Code Extension**

1. Install "SQLite Viewer" extension in VS Code
2. Right-click `.db` file → "Open Database"
3. Browse tables and run queries

---

## ✅ What to Verify in Database

### **1. Check Tables Exist**

Run this query:
```sql
SELECT name FROM sqlite_master WHERE type='table';
```

**Expected Tables:**
- `Filters`
- `FileCombos`
- `ClashZones`
- `SleeveEvents`
- `Conditions`
- `SchemaMigrations`

### **2. Check Filters Table**

```sql
SELECT * FROM Filters;
```

**Expected:**
- Should have rows for each filter + category combination
- Example: `FilterName='ALL', Category='Pipes'`
- `CreatedAt` and `UpdatedAt` should have timestamps

### **3. Check FileCombos Table**

```sql
SELECT * FROM FileCombos;
```

**Expected:**
- Should have rows for each file combo (LinkedFile + HostFile)
- Example: `LinkedFileKey='PH-00001', HostFileKey='AR-00001'`
- `FilterId` should reference `Filters` table

### **4. Check ClashZones Table**

```sql
SELECT COUNT(*) as TotalZones, 
       COUNT(DISTINCT ComboId) as UniqueCombos,
       COUNT(CASE WHEN SleeveState = 0 THEN 1 END) as Unprocessed,
       COUNT(CASE WHEN SleeveState = 1 THEN 1 END) as IndividualPlaced,
       COUNT(CASE WHEN SleeveState = 2 THEN 1 END) as ClusterPlaced
FROM ClashZones;
```

**Expected:**
- `TotalZones` should match the number of clash zones from XML
- `UniqueCombos` should match number of file combos
- Most zones should be `SleeveState = 0` (Unprocessed) initially

### **5. Compare with XML Files**

**Check XML Count:**
- Open `{filterName}_{category}.xml` files
- Count clash zones in XML
- Compare with SQLite `ClashZones` table count

**They should match!** ✅

---

## 📊 Sample Verification Queries

### **Get Clash Zones by Category:**
```sql
SELECT 
    f.FilterName,
    f.Category,
    COUNT(cz.ClashZoneId) as ZoneCount
FROM Filters f
LEFT JOIN FileCombos fc ON f.FilterId = fc.FilterId
LEFT JOIN ClashZones cz ON fc.ComboId = cz.ComboId
GROUP BY f.FilterName, f.Category
ORDER BY f.FilterName, f.Category;
```

### **Get File Combos:**
```sql
SELECT 
    f.FilterName,
    f.Category,
    fc.LinkedFileKey,
    fc.HostFileKey,
    COUNT(cz.ClashZoneId) as ZoneCount
FROM FileCombos fc
JOIN Filters f ON fc.FilterId = f.FilterId
LEFT JOIN ClashZones cz ON fc.ComboId = cz.ComboId
GROUP BY f.FilterName, f.Category, fc.LinkedFileKey, fc.HostFileKey
ORDER BY f.FilterName, f.Category;
```

### **Check Sleeve States:**
```sql
SELECT 
    CASE SleeveState 
        WHEN 0 THEN 'Unprocessed'
        WHEN 1 THEN 'IndividualPlaced'
        WHEN 2 THEN 'ClusterPlaced'
        ELSE 'Unknown'
    END as State,
    COUNT(*) as Count
FROM ClashZones
GROUP BY SleeveState
ORDER BY SleeveState;
```

---

## 🎯 Success Criteria

### **✅ Phase 1 is Working If:**

1. ✅ Database file is created after refresh
2. ✅ Logs show `[SQLite] ✅ Dual-write to SQLite` messages
3. ✅ All expected tables exist in database
4. ✅ Clash zone counts in SQLite match XML files
5. ✅ No errors that prevent XML writes
6. ✅ Database file size grows after each refresh

### **⚠️ Non-Critical Warnings (OK):**

- SQLite initialization warnings (XML still works)
- SQLite write failures (XML writes succeeded)
- These are logged but don't affect XML operations

### **❌ Critical Errors (Not OK):**

- XML writes failing
- Database corruption
- Missing tables in database
- Zero clash zones in database when XML has zones

---

## 🔧 Troubleshooting

### **Problem: Database file not created**

**Check:**
1. Logs for SQLite initialization errors
2. File system permissions on Filters directory
3. Project name contains invalid characters (should be sanitized)

**Solution:**
- Check logs for specific error messages
- Ensure Filters directory is writable
- Database file name is sanitized from project title

### **Problem: Zero clash zones in database**

**Check:**
1. Logs for `[SQLite] ✅ Inserted/updated` messages
2. Verify XML files have clash zones
3. Check if dual-write is actually being called

**Solution:**
- Ensure refresh actually detected intersections
- Check that `ClashZonePersistenceService` is being called
- Verify SQLite repository is initialized

### **Problem: Database file exists but is empty**

**Check:**
1. Database file size (should be > 0)
2. Tables exist but have no data
3. Logs for SQLite write errors

**Solution:**
- Check transaction commits
- Verify clash zones are being passed to repository
- Check for SQLite constraint violations

---

## 📝 Testing Checklist

- [ ] Build project successfully
- [ ] Run refresh operation
- [ ] Check logs for SQLite messages
- [ ] Verify database file created
- [ ] Open database in SQLite browser
- [ ] Verify all tables exist
- [ ] Check Filters table has data
- [ ] Check FileCombos table has data
- [ ] Check ClashZones table has data
- [ ] Compare counts with XML files
- [ ] Run verification queries
- [ ] Test multiple refresh operations
- [ ] Verify database grows with new data

---

## 🚀 Next Steps After Successful Testing

Once Phase 1 is verified working:

1. **Monitor for a few days** - Ensure stability
2. **Compare data** - Regularly compare XML vs SQLite counts
3. **Check performance** - Ensure dual-write doesn't slow down refresh
4. **Document any issues** - Log any SQLite errors for Phase 2 fixes
5. **Prepare for Phase 2** - When ready, begin cutover to SQLite as primary store

---

## 📞 Support

If you encounter issues:
1. Check logs first (`refresh.log` and debug logs)
2. Verify database file exists and is accessible
3. Check SQLite browser can open the database
4. Compare XML vs SQLite data counts
5. Review error messages in logs

**Remember**: XML is still the operational store. SQLite failures don't affect XML operations.

