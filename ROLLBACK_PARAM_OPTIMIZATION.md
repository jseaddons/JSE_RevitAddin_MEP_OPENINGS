# Rollback: Parameter optimization (definition cache, group-by-size, rotation skip)

If the parameter optimization changes fail, restore the previous state:

**Checkpoint commit (current state before optimization):**  
`bcb5018c73b5e5bb4aaabb8820524e8b00b01f58`

## Option A – Create branch at checkpoint (if not done yet), then reset to it later

```powershell
cd "c:\Jse_Developments\JSE_MEPOPENING_23"
git branch checkpoint-before-param-optimization bcb5018c73b5e5bb4aaabb8820524e8b00b01f58
```

To roll back later:

```powershell
git checkout checkpoint-before-param-optimization
# or to discard local changes and match that branch exactly:
git reset --hard checkpoint-before-param-optimization
```

## Option B – Reset current branch to checkpoint commit (discard optimization changes)

```powershell
cd "c:\Jse_Developments\JSE_MEPOPENING_23"
git checkout -- Services/Placement/SleeveParameterService.cs
# or full reset to checkpoint:
git reset --hard bcb5018c73b5e5bb4aaabb8820524e8b00b01f58
```

**Note:** If the repo had an `index.lock` or ref lock, run the branch/tag commands when no other Git process is running.
