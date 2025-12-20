-- Emergency cleanup script for orphaned CombinedSleeves_Old table
-- Run this SQL against your SleevePersistence.db database

-- Step 1: Check if CombinedSleeves_Old exists
SELECT name FROM sqlite_master WHERE type='table' AND name='CombinedSleeves_Old';

-- Step 2: If it exists, check if it has data
SELECT COUNT(*) as RecordCount FROM CombinedSleeves_Old;

-- Step 3: Check current CombinedSleeves table
SELECT COUNT(*) as RecordCount FROM CombinedSleeves;

-- Step 4: If CombinedSleeves_Old has data but CombinedSleeves doesn't, restore it
-- (Uncomment if needed)
-- INSERT INTO CombinedSleeves SELECT * FROM CombinedSleeves_Old;

-- Step 5: Drop the orphaned table
DROP TABLE IF EXISTS CombinedSleeves_Old;

-- Step 6: Verify cleanup
SELECT name FROM sqlite_master WHERE type='table' AND name LIKE '%CombinedSleeves%';
