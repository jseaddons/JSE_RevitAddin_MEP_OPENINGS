# Performance Analysis - 2025-11-24 16:14:08

## Summary
Analysis of performance logs from latest refresh and placement operations.

---

## 1. REFRESH PERFORMANCE (16-14-08)

### Overall Metrics
- **Total Time**: 54,330ms (54.3 seconds)
- **Total Clash Zones**: 378
- **Clash Zones/Second**: 7 (target: 500+) ⚠️ **BELOW TARGET**
- **Memory Usage**: 136.43 MB → 135.93 MB (Δ: -0.51 MB)

### Operation Breakdown

| Operation | Time | % of Total | Items/s |
|-----------|------|------------|---------|
| **6. Intersection Processing** | 43,655ms | **80.4%** | 9/s |
| RunDetection | 43,633ms | 80.3% | - |
| **5B. Flag Reset** | 5,109ms | **9.4%** | - |
| **9. Save** | 3,293ms | 6.1% | 115/s |
| 8. Parameter Capture | 1,544ms | 2.8% | 245/s |
| 10. Cleanup | 241ms | 0.4% | - |
| 4. Flag Sync | 222ms | 0.4% | - |
| 2. XML Loading | 165ms | 0.3% | 30/s |
| Others | 25ms | <0.1% | - |

### Key Findings
1. **Intersection Processing is the biggest bottleneck** (80.4% of total time)
   - 43,655ms for 378 zones = 115.5ms per zone
   - This is the primary optimization target

2. **Flag Reset is still slow** (9.4% of total time)
   - 5,109ms - should be faster with batch updates
   - Need to verify batch update logs are present

3. **Save operation is reasonable** (6.1% of total time)
   - 3,293ms for 378 zones = 8.7ms per zone
   - 115 zones/second is good

---

## 2. PLACEMENT PERFORMANCE - DUCTS (16-15-54)

### Overall Metrics
- **Total Time**: 27,097ms (27.1 seconds)
- **Individual Sleeves**: 119 placed
- **Clusters**: 0
- **Individual Sleeves/Second**: 4 (target: 50+) ⚠️ **BELOW TARGET**
- **Memory Usage**: 139.88 MB → 173.32 MB (Δ: +33.44 MB)

### Operation Breakdown

| Operation | Time | % of Total | Items/s |
|-----------|------|------------|---------|
| Individual Sleeve Placement | 26,842ms | 99.1% | 4/s |
| Cluster Sleeve Placement | 212ms | 0.8% | - |

### Key Findings
1. **Individual sleeve placement is very slow**
   - 26,842ms for 119 sleeves = 225.6ms per sleeve
   - Target is <20ms per sleeve (50+ sleeves/second)
   - **11× slower than target**

2. **No clusters formed** - all sleeves placed individually

---

## 3. PLACEMENT PERFORMANCE - PIPES (16-17-07)

### Overall Metrics
- **Total Time**: 49,991ms (50.0 seconds)
- **Individual Sleeves**: 105 placed
- **Clusters**: 39 placed
- **Individual Sleeves/Second**: 2 (target: 50+) ⚠️ **BELOW TARGET**
- **Clusters/Second**: 1 (target: 10+) ⚠️ **BELOW TARGET**
- **Memory Usage**: 193.37 MB → 143.52 MB (Δ: -49.85 MB)

### Operation Breakdown

| Operation | Time | % of Total | Items/s |
|-----------|------|------------|---------|
| Individual Sleeve Placement | 29,312ms | 58.6% | 4/s |
| Cluster Sleeve Placement | 20,637ms | 41.3% | 2/s |

### Key Findings
1. **Individual sleeve placement is slow**
   - 29,312ms for 105 sleeves = 279.1ms per sleeve
   - **14× slower than target**

2. **Cluster placement is also slow**
   - 20,637ms for 39 clusters = 529.2ms per cluster
   - Target is <100ms per cluster (10+ clusters/second)
   - **5× slower than target**

---

## 4. CLUSTER PLACEMENT DETAILED - PIPES (16-17-36)

### Overall Metrics
- **Total Time**: 15,813ms (15.8 seconds)
- **Clusters**: 39 placed
- **Clusters/Second**: 2 (target: 10+) ⚠️ **BELOW TARGET**

### Operation Breakdown

| Operation | Time | % of Total | Avg per Cluster | Items/s |
|-----------|------|------------|-----------------|---------|
| **Place Clusters Loop** | 14,210ms | **89.9%** | 364.4ms | 3/s |
| ├─ Place Cluster Sleeve | 13,853ms | 87.6% | **355.2ms** | - |
| ├─ Calculate Rotated Bounding Box | 6,745ms | 42.7% | **172.9ms** | - |
| └─ Determine Rotation Angle | 0ms | 0.0% | 0.0ms | - |
| Save Cluster Data to Database | 998ms | 6.3% | 25.6ms | 39/s |
| Cleanup Individual Sleeves | 132ms | 0.8% | 3.4ms | - |
| Load Clash Zones from Database | 80ms | 0.5% | 2.1ms | 1800/s |
| Form Clusters | 10ms | 0.1% | 0.3ms | 3900/s |
| Others | 383ms | 2.4% | - | - |

### Key Findings
1. **Place Cluster Sleeve is the biggest bottleneck** (87.6% of total time)
   - 13,853ms for 39 clusters = 355.2ms per cluster
   - This includes parameter writes, rotation, metadata setting

2. **Calculate Rotated Bounding Box is slow** (42.7% of total time)
   - 6,745ms for 39 clusters = 172.9ms per cluster
   - This is called within Place Cluster Sleeve, so it's part of the 355.2ms

3. **Determine Rotation Angle is fast** (0ms)
   - Rotation angle determination is instant (likely cached or pre-calculated)

4. **Save Cluster Data is reasonable** (6.3% of total time)
   - 998ms for 39 clusters = 25.6ms per cluster
   - 39 clusters/second is good

---

## 5. PERFORMANCE COMPARISON

### Refresh Performance
- **Previous runs**: Flag Reset was 3,895ms (similar to current 5,109ms)
- **Intersection Processing**: Still the biggest bottleneck at 80.4%
- **Flag Reset**: Still slow at 9.4% despite batch updates

### Placement Performance
- **Individual Sleeves**: 225-279ms per sleeve (target: <20ms)
  - **11-14× slower than target**
- **Clusters**: 355ms per cluster (target: <100ms)
  - **3.5× slower than target**

---

## 6. RECOMMENDATIONS

### High Priority
1. **Optimize Intersection Processing** (80.4% of refresh time)
   - Current: 115.5ms per zone
   - Target: <10ms per zone
   - This is the #1 optimization target

2. **Optimize Individual Sleeve Placement** (99.1% of placement time)
   - Current: 225-279ms per sleeve
   - Target: <20ms per sleeve
   - Check if parameter batching is working correctly

3. **Optimize Cluster Placement** (87.6% of cluster time)
   - Current: 355ms per cluster
   - Target: <100ms per cluster
   - Focus on:
     - Rotated Bounding Box calculation (172.9ms per cluster)
     - Parameter writes (verify batch parameter writes are working)

### Medium Priority
4. **Investigate Flag Reset** (9.4% of refresh time)
   - Current: 5,109ms
   - Should be faster with batch updates
   - Verify batch update logs are present

5. **Optimize Rotated Bounding Box Calculation** (42.7% of cluster time)
   - Current: 172.9ms per cluster
   - Consider caching or pre-calculation

### Low Priority
6. **Memory Usage** - Currently good (no leaks detected)

---

## 7. NEXT STEPS

1. **Verify batch parameter writes are working** for cluster placement
2. **Check if R-tree is being used** for Flag Reset (should see batch update logs)
3. **Profile Intersection Processing** to identify specific bottlenecks
4. **Profile Individual Sleeve Placement** to identify slow operations
5. **Profile Rotated Bounding Box Calculation** to optimize

---

## 8. BUILD INFORMATION

- **Build Timestamp**: 2025-11-24 16:13:34
- **Assembly**: JSE_RevitAddin_MEP_OPENINGS.dll
- **Logs Analyzed**: 
  - Refresh: 16-14-08
  - Placement Ducts: 16-15-54
  - Placement Pipes: 16-17-07
  - Cluster Placement Pipes: 16-17-36

