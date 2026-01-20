#!/usr/bin/env python3
"""
Simple C# Null Reference Warnings Fixer
Systematically fixes CS8602, CS8604, CS8600, CS8625, CS8629 warnings
"""

import os
import re
import sys
from pathlib import Path
from typing import List, Tuple, Optional

class NullReferenceFixer:
    def __init__(self, project_root: str):
        self.project_root = Path(project_root)
        self.files_to_fix = {
            "ClashZoneService.cs": {
                "CS8602": [2479, 2501, 2524, 2532, 2555, 2577, 2633, 2673, 2677, 2712, 2735, 2769, 2806, 2845, 2945, 2984],
                "CS8600": [2644, 2739, 2806, 2945, 3494, 4491]
            },
            "UniversalClusterService.cs": {
                "CS8602": [1711, 1977, 1979, 1983, 3987, 3988, 5418, 5469, 6380, 6382],
                "CS8600": [615, 1057, 1062, 1472, 1515, 1530, 1573, 1592, 1644, 1663, 1678, 1859, 1915, 1977, 1979, 1983, 4980, 5339, 6201, 6499, 6782, 7794, 7811],
                "CS8604": [198, 627, 860, 879, 935, 4637, 4753, 6069, 6138, 6522, 7737],
                "CS8629": [3845, 3850, 3855, 3860, 3865, 3870, 3875, 3880, 3885, 3890, 3895, 3900, 3905, 3910, 3915, 3920, 3925, 3930, 7901]
            },
            "UniversalSleevePlacerService.cs": {
                "CS8602": [327, 807, 922, 924, 1058, 1532, 2849, 2855, 2864, 2870],
                "CS8600": [1240, 1252, 1269, 1333, 3723]
            },
            "EmergencyMainDialog.cs": {
                "CS8602": [3743, 3750, 4167, 4238],
                "CS8604": [3743, 3750, 4167, 4238]
            },
            "FlagManager.cs": {
                "CS8625": [1441, 1738, 1876, 2121]
            },
            "ParameterTransferService.cs": {
                "CS8603": [2060, 2072, 2079]
            }
        }
        self.stats = {"total": 0, "fixed": 0, "skipped": 0}

    def read_file(self, file_path: str) -> List[str]:
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                return f.readlines()
        except Exception as e:
            print("[ERROR] Reading " + file_path + ": " + str(e))
            return []

    def write_file(self, file_path: str, lines: List[str]) -> bool:
        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                f.writelines(lines)
            return True
        except Exception as e:
            print("[ERROR] Writing " + file_path + ": " + str(e))
            return False

    def fix_cs8602_dereference(self, line: str) -> Optional[str]:
        """Fix dereference of possibly null reference"""
        original = line
        
        # Category.Name pattern
        if '.Category.Name' in line and '?.Category' not in line:
            line = line.replace('.Category.Name', '?.Category?.Name ?? "Unknown"')
        
        # Location.Position pattern
        elif '.Location.Position' in line and '?.Location' not in line:
            line = line.replace('.Location.Position', '?.Location?.Position')
        
        # Generic chained property (obj.prop.value)
        elif re.search(r'(\w+)\.(\w+)\.' + r'(\w+)', line) and '?.' not in line:
            match = re.search(r'(\w+)\.(\w+)\.(\w+)', line)
            if match and '?.' not in match.group(0):
                obj, prop1, prop2 = match.groups()
                old_pattern = obj + '.' + prop1 + '.' + prop2
                new_pattern = obj + '?.' + prop1 + '?.' + prop2
                line = line.replace(old_pattern, new_pattern)
        
        return line if line != original else None

    def fix_cs8604_null_argument(self, line: str) -> Optional[str]:
        """Fix null reference argument"""
        if 'null,' in line or 'null)' in line:
            return line.replace('null,', 'string.Empty,').replace('null)', 'string.Empty)')
        return None

    def fix_cs8600_null_to_nonnull(self, line: str) -> Optional[str]:
        """Fix converting null to non-nullable"""
        if re.search(r'string\s+\w+\s*=\s*\w+\s*\?\.', line):
            return re.sub(r'(string\s+\w+\s*=\s*)(\w+\s*\?\.)', r'\1\2 ?? string.Empty', line)
        elif re.search(r'double\s+\w+\s*=\s*\w+\s*\?\.', line):
            return re.sub(r'(double\s+\w+\s*=\s*)(\w+\s*\?\.)', r'\1\2 ?? 0.0', line)
        return None

    def fix_cs8625_null_literal(self, line: str) -> Optional[str]:
        """Fix null literal to non-nullable"""
        if 'null!' in line:
            return line.replace('null!', 'null')
        return None

    def fix_cs8629_nullable_value_type(self, line: str) -> Optional[str]:
        """Fix nullable value type may be null"""
        if re.search(r'(double|int|decimal|float)\s+\w+\s*=\s*\w+\s*\?\.', line):
            if ';' in line and '??' not in line:
                return line.replace(';', ' ?? 0.0;')
        return None

    def process_file(self, file_name: str, warnings_dict: dict) -> Tuple[int, int]:
        """Process single file"""
        # Try different locations
        paths_to_try = [
            self.project_root / "Services" / file_name,
            self.project_root / "Views" / file_name,
            self.project_root / file_name
        ]
        
        file_path = None
        for p in paths_to_try:
            if p.exists():
                file_path = p
                break
        
        if not file_path:
            print("[SKIP] File not found: " + file_name)
            return 0, 0
        
        print("[FILE] Processing: " + file_name)
        lines = self.read_file(str(file_path))
        if not lines:
            return 0, 0
        
        fixed_count = 0
        skipped_count = 0
        modified_lines = lines.copy()
        
        for warning_code, line_numbers in warnings_dict.items():
            for line_num in line_numbers:
                idx = line_num - 1
                if idx >= len(lines):
                    continue
                
                line = lines[idx]
                fixed_line = None
                
                if warning_code == "CS8602":
                    fixed_line = self.fix_cs8602_dereference(line)
                elif warning_code == "CS8604":
                    fixed_line = self.fix_cs8604_null_argument(line)
                elif warning_code == "CS8600":
                    fixed_line = self.fix_cs8600_null_to_nonnull(line)
                elif warning_code == "CS8625":
                    fixed_line = self.fix_cs8625_null_literal(line)
                elif warning_code == "CS8629":
                    fixed_line = self.fix_cs8629_nullable_value_type(line)
                
                if fixed_line and fixed_line != line:
                    modified_lines[idx] = fixed_line
                    print("   [FIXED] Line " + str(line_num) + " (" + warning_code + ")")
                    print("      Before: " + line.strip()[:70])
                    print("      After:  " + fixed_line.strip()[:70])
                    fixed_count += 1
                else:
                    skipped_count += 1
        
        # Write if changes made
        if fixed_count > 0:
            if self.write_file(str(file_path), modified_lines):
                print("   [SAVED] " + file_name)
            else:
                fixed_count = 0
        
        self.stats["fixed"] += fixed_count
        self.stats["skipped"] += skipped_count
        self.stats["total"] += len([v for vals in warnings_dict.values() for v in vals])
        
        return fixed_count, skipped_count

    def run(self) -> None:
        """Run fixer on all files"""
        print("=" * 80)
        print("C# Null Reference Warnings Fixer")
        print("=" * 80)
        print("[INFO] Project Root: " + str(self.project_root))
        print("[INFO] Target Files: " + str(len(self.files_to_fix)))
        print("")
        
        total_fixed = 0
        total_skipped = 0
        
        for file_name, warnings in self.files_to_fix.items():
            fixed, skipped = self.process_file(file_name, warnings)
            total_fixed += fixed
            total_skipped += skipped
            print("")
        
        print("=" * 80)
        print("SUMMARY")
        print("=" * 80)
        print("[STATS] Total Fixed: " + str(total_fixed))
        print("[STATS] Total Skipped: " + str(total_skipped))
        print("[STATS] Total Targeted: " + str(self.stats["total"]))
        if (total_fixed + total_skipped) > 0:
            success_rate = (total_fixed / (total_fixed + total_skipped) * 100)
            print("[STATS] Success Rate: " + str(round(success_rate, 1)) + "%")
        print("=" * 80)

def main():
    project_root = r"c:\JSE_CSharp_Projects\JSE_MEPOPENING_23"
    
    if len(sys.argv) > 1:
        project_root = sys.argv[1]
    
    fixer = NullReferenceFixer(project_root)
    fixer.run()

if __name__ == "__main__":
    main()
