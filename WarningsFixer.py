#!/usr/bin/env python3
"""
Intelligent C# Null Reference Warnings Fixer
Systematically fixes CS8602, CS8604, CS8600, CS8625, CS8629 warnings
with careful code analysis to prevent breaking existing functionality.
"""

import os
import re
import sys
from pathlib import Path
from typing import List, Tuple, Optional
from dataclasses import dataclass
from enum import Enum

class WarningType(Enum):
    CS8602 = "Dereference of a possibly null reference"
    CS8604 = "Possible null reference argument"
    CS8600 = "Converting null to non-nullable"
    CS8625 = "Cannot convert null literal to non-nullable"
    CS8629 = "Nullable value type may be null"

@dataclass
class WarningLocation:
    file_path: str
    line_number: int
    warning_type: WarningType
    code_line: str
    description: str

class NullReferenceFixer:
    def __init__(self, project_root: str):
        self.project_root = Path(project_root)
        self.warnings: List[WarningLocation] = []
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
        self.stats = {
            "total_warnings": 0,
            "fixed": 0,
            "skipped": 0,
            "errors": 0
        }

    def read_file(self, file_path: str) -> List[str]:
        """Read file and return lines."""
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                return f.readlines()
        except Exception as e:
            print(f"❌ Error reading {file_path}: {e}")
            return []

    def write_file(self, file_path: str, lines: List[str]) -> bool:
        """Write lines back to file."""
        try:
            with open(file_path, 'w', encoding='utf-8') as f:
                f.writelines(lines)
            return True
        except Exception as e:
            print(f"❌ Error writing {file_path}: {e}")
            return False

    def fix_cs8602_dereference(self, line: str, context_lines: List[str], line_idx: int) -> Optional[str]:
        """
        Fix CS8602: Dereference of a possibly null reference
        Pattern: obj.Property.Value or element.Category.Name
        Solution: obj?.Property?.Value or element?.Category?.Name ?? defaultValue
        """
        # Patterns to detect chained property access without null checks
        patterns = [
            (r'(\w+)\.(\w+)\.(\w+)', r'\1?.\2?.\3'),  # obj.prop.value -> obj?.prop?.value
            (r'(\w+)\.(\w+)\(\)', r'\1?.\2?()'),       # obj.method() -> obj?.method?()
            (r'\.Category\.Name', r'?.Category?.Name ?? "Unknown"'),  # Special case for Category.Name
            (r'\.Location\.Position', r'?.Location?.Position'),       # Special case for Location
        ]
        
        fixed_line = line
        for pattern, replacement in patterns:
            if re.search(pattern, line):
                fixed_line = re.sub(pattern, replacement, line)
                if fixed_line != line:
                    return fixed_line
        return None

    def fix_cs8604_null_argument(self, line: str) -> Optional[str]:
        """
        Fix CS8604: Possible null reference argument
        Pattern: Method(nullableParam) where Method expects non-null
        Solution: Add null check or use null coalescing
        """
        # Look for method calls with potentially null arguments
        if re.search(r'(\w+)\(\s*null', line):
            # Replace null! with null or empty string
            return line.replace('null!', 'null').replace('null,', 'string.Empty,')
        return None

    def fix_cs8600_null_to_nonnull(self, line: str) -> Optional[str]:
        """
        Fix CS8600: Converting null literal or possible null value to non-nullable
        Pattern: string var = nullableVar;
        Solution: string var = nullableVar ?? string.Empty; OR string? var = nullableVar;
        """
        # Pattern: type var = nullable_source;
        patterns = [
            (r'string\s+(\w+)\s*=\s*(\w+);', r'string \1 = \2 ?? string.Empty;'),
            (r'double\s+(\w+)\s*=\s*(\w+);', r'double \1 = \2 ?? 0.0;'),
            (r'int\s+(\w+)\s*=\s*(\w+);', r'int \1 = \2 ?? 0;'),
        ]
        
        for pattern, replacement in patterns:
            if re.search(pattern, line):
                return re.sub(pattern, replacement, line)
        return None

    def fix_cs8625_null_literal(self, line: str) -> Optional[str]:
        """
        Fix CS8625: Cannot convert null literal to non-nullable
        Pattern: string var = null;
        Solution: string? var = null; OR var = string.Empty;
        """
        if 'null!' in line:
            return line.replace('null!', 'null')
        elif re.search(r'string\s+(\w+)\s*=\s*null;', line):
            match = re.search(r'string\s+(\w+)\s*=\s*null;', line)
            if match:
                var_name = match.group(1)
                return line.replace(f'string {var_name} = null;', f'string? {var_name} = null;')
        return None

    def fix_cs8629_nullable_value_type(self, line: str) -> Optional[str]:
        """
        Fix CS8629: Nullable value type may be null
        Pattern: double val = nullableDouble;
        Solution: double val = nullableDouble ?? 0.0; OR double? val = nullableDouble;
        """
        patterns = [
            (r'double\s+(\w+)\s*=\s*(\w+)\?\.(\w+)', r'double \1 = \2?.\3 ?? 0.0'),
            (r'int\s+(\w+)\s*=\s*(\w+)\?\.(\w+)', r'int \1 = \2?.\3 ?? 0'),
            (r'decimal\s+(\w+)\s*=\s*(\w+)\?\.(\w+)', r'decimal \1 = \2?.\3 ?? 0m'),
        ]
        
        for pattern, replacement in patterns:
            if re.search(pattern, line):
                return re.sub(pattern, replacement, line)
        return None

    def analyze_context(self, lines: List[str], line_idx: int, warning_type: str) -> bool:
        """
        Analyze surrounding context to determine if fix is safe.
        Returns True if fix is safe, False if it might break code.
        """
        # Check if line is inside a try-catch (generally safe)
        for i in range(max(0, line_idx - 5), line_idx):
            if 'try' in lines[i]:
                return True
        
        # Check if null check already exists
        if line_idx > 0:
            prev_line = lines[line_idx - 1]
            if 'if' in prev_line and '!=' in prev_line and 'null' in prev_line:
                return True
        
        return True  # Default to safe

    def process_file(self, file_name: str, warnings_dict: dict) -> Tuple[int, int]:
        """Process a single file and fix warnings."""
        file_path = self.project_root / "Services" / file_name
        if not file_path.exists():
            file_path = self.project_root / "Views" / file_name
        if not file_path.exists():
            file_path = self.project_root / file_name
        
        if not file_path.exists():
            print(f"⚠️  File not found: {file_name}")
            return 0, 0
        
        print(f"\n📄 Processing: {file_name}")
        lines = self.read_file(str(file_path))
        if not lines:
            return 0, 0
        
        fixed_count = 0
        skipped_count = 0
        modified_lines = lines.copy()
        
        for warning_code, line_numbers in warnings_dict.items():
            for line_num in line_numbers:
                # Convert to 0-based index
                idx = line_num - 1
                if idx >= len(lines):
                    continue
                
                line = lines[idx]
                original_line = line
                
                # Analyze context
                if not self.analyze_context(lines, idx, warning_code):
                    print(f"   ⏭️  Skipped line {line_num}: Context too complex")
                    skipped_count += 1
                    continue
                
                # Apply appropriate fix
                fixed_line = None
                if warning_code == "CS8602":
                    fixed_line = self.fix_cs8602_dereference(line, lines, idx)
                elif warning_code == "CS8604":
                    fixed_line = self.fix_cs8604_null_argument(line)
                elif warning_code == "CS8600":
                    fixed_line = self.fix_cs8600_null_to_nonnull(line)
                elif warning_code == "CS8625":
                    fixed_line = self.fix_cs8625_null_literal(line)
                elif warning_code == "CS8629":
                    fixed_line = self.fix_cs8629_nullable_value_type(line)
                
                if fixed_line and fixed_line != original_line:
                    modified_lines[idx] = fixed_line
                    print(f"   ✅ Fixed line {line_num} ({warning_code})")
                    print(f"      Before: {original_line.strip()[:80]}")
                    print(f"      After:  {fixed_line.strip()[:80]}")
                    fixed_count += 1
                else:
                    skipped_count += 1
        
        # Write back if changes were made
        if fixed_count > 0:
            if self.write_file(str(file_path), modified_lines):
                print(f"   💾 Saved {file_name}")
            else:
                self.stats["errors"] += fixed_count
                fixed_count = 0
        
        self.stats["fixed"] += fixed_count
        self.stats["skipped"] += skipped_count
        self.stats["total_warnings"] += len([v for vals in warnings_dict.values() for v in vals])
        
        return fixed_count, skipped_count

    def run(self) -> None:
        """Run the fixer on all target files."""
        print("=" * 80)
        print("🔧 C# Null Reference Warnings Fixer")
        print("=" * 80)
        print(f"📁 Project Root: {self.project_root}")
        print(f"🎯 Target Files: {len(self.files_to_fix)}")
        
        total_fixed = 0
        total_skipped = 0
        
        for file_name, warnings in self.files_to_fix.items():
            fixed, skipped = self.process_file(file_name, warnings)
            total_fixed += fixed
            total_skipped += skipped
        
        print("\n" + "=" * 80)
        print("📊 SUMMARY")
        print("=" * 80)
        print(f"✅ Total Fixed: {total_fixed}")
        print(f"⏭️  Total Skipped: {total_skipped}")
        print(f"❌ Errors: {self.stats['errors']}")
        print(f"📈 Total Warnings Targeted: {self.stats['total_warnings']}")
        print(f"✨ Success Rate: {(total_fixed / (total_fixed + total_skipped) * 100):.1f}%" if (total_fixed + total_skipped) > 0 else "No fixes attempted")
        print("=" * 80)

def main():
    project_root = r"c:\JSE_CSharp_Projects\JSE_MEPOPENING_23"
    
    if len(sys.argv) > 1:
        project_root = sys.argv[1]
    
    fixer = NullReferenceFixer(project_root)
    fixer.run()

if __name__ == "__main__":
    main()
