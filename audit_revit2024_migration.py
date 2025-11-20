#!/usr/bin/env python3
"""
Revit 2024 Migration Audit & Fix Script
Carefully audits and fixes migration issues one file at a time.

What it audits:
1. Direct UnitUtils.Convert* calls that should use RevitUnitConversionService
2. BuiltInCategory/BuiltInParameter casts to int that need ElementId comparisons
3. UnitUtils usage without exception handling

Usage:
    python audit_revit2024_migration.py --file Services/SomeService.cs
    python audit_revit2024_migration.py --audit-only  # Just report issues
    python audit_revit2024_migration.py --fix Services/SomeService.cs  # Apply fixes
"""

import re
import os
import sys
import argparse
import shutil
from pathlib import Path
from typing import List, Tuple, Dict
from dataclasses import dataclass
from datetime import datetime

@dataclass
class MigrationIssue:
    """Represents a migration issue found in code"""
    file_path: str
    line_number: int
    issue_type: str  # 'UnitUtils', 'BuiltInCast', etc.
    severity: str  # 'warning', 'error'
    description: str
    original_code: str
    suggested_fix: str = ""

class Revit2024MigrationAuditor:
    """Audits C# files for Revit 2024 migration issues"""
    
    # Patterns to find issues
    UNIT_UTILS_PATTERNS = [
        (r'UnitUtils\.ConvertFromInternalUnits\s*\([^,]+,\s*UnitTypeId\.Millimeters\)', 
         'UnitUtils.ConvertFromInternalUnits with Millimeters',
         'Use RevitUnitConversionService.Instance.FromInternalMillimeters()'),
        (r'UnitUtils\.ConvertToInternalUnits\s*\([^,]+,\s*UnitTypeId\.Millimeters\)',
         'UnitUtils.ConvertToInternalUnits with Millimeters',
         'Use RevitUnitConversionService.Instance.ToInternalMillimeters()'),
        (r'UnitUtils\.ConvertFromInternalUnits\s*\([^,]+,\s*UnitTypeId\.Feet\)',
         'UnitUtils.ConvertFromInternalUnits with Feet',
         'Use RevitUnitConversionService.Instance.FromInternalFeet()'),
        (r'UnitUtils\.ConvertToInternalUnits\s*\([^,]+,\s*UnitTypeId\.Feet\)',
         'UnitUtils.ConvertToInternalUnits with Feet',
         'Use RevitUnitConversionService.Instance.ToInternalFeet()'),
    ]
    
    BUILTIN_CAST_PATTERNS = [
        (r'\(int\)\s*BuiltInCategory\.', 'BuiltInCategory cast to int', 
         'Use ElementId comparison: new ElementId((int)BuiltInCategory.XXX)'),
        (r'\(int\)\s*BuiltInParameter\.', 'BuiltInParameter cast to int',
         'Use ElementId comparison: new ElementId((int)BuiltInParameter.XXX)'),
        (r'Category\.Id\.IntegerValue\s*==\s*\(int\)BuiltInCategory\.',
         'Category.Id.IntegerValue == (int)BuiltInCategory comparison',
         'Use ElementId.Equals() or direct comparison without cast'),
    ]
    
    # Files to skip (already migrated or system files)
    SKIP_PATTERNS = [
        r'.*\.backup',
        r'.*\.bak',
        r'.*\.tmp',
        r'.*RevitUnitConversionService\.cs',  # This is the service itself
        r'.*bin/.*',
        r'.*obj/.*',
    ]
    
    def __init__(self, root_dir: str = '.'):
        self.root_dir = Path(root_dir)
        self.issues: List[MigrationIssue] = []
        self.backup_dir = self.root_dir / '_BACKUPS' / 'migration_audit'
        self.backup_dir.mkdir(parents=True, exist_ok=True)
    
    def should_skip_file(self, file_path: Path) -> bool:
        """Check if file should be skipped"""
        rel_path = str(file_path.relative_to(self.root_dir))
        for pattern in self.SKIP_PATTERNS:
            if re.search(pattern, rel_path, re.IGNORECASE):
                return True
        return False
    
    def audit_file(self, file_path: Path, fix: bool = False) -> List[MigrationIssue]:
        """Audit a single C# file for migration issues"""
        if not file_path.exists():
            print(f"❌ File not found: {file_path}")
            return []
        
        if self.should_skip_file(file_path):
            print(f"⏭️  Skipping: {file_path}")
            return []
        
        print(f"\n🔍 Auditing: {file_path}")
        issues = []
        
        try:
            with open(file_path, 'r', encoding='utf-8', errors='ignore') as f:
                lines = f.readlines()
                content = ''.join(lines)
        except Exception as e:
            print(f"❌ Error reading file: {e}")
            return []
        
        # Check if file already uses RevitUnitConversionService
        uses_service = 'RevitUnitConversionService' in content or 'IRevitUnitConversionService' in content
        
        # Audit UnitUtils patterns
        for pattern, description, suggestion in self.UNIT_UTILS_PATTERNS:
            for match in re.finditer(pattern, content):
                line_num = content[:match.start()].count('\n') + 1
                line_content = lines[line_num - 1].strip()
                
                # Skip if already wrapped in try-catch or using service
                context_start = max(0, match.start() - 100)
                context_end = min(len(content), match.end() + 100)
                context = content[context_start:context_end]
                
                # Skip if in a try-catch or already using service
                if uses_service or 'try' in context.lower()[:200]:
                    severity = 'warning'
                else:
                    severity = 'error'
                
                issues.append(MigrationIssue(
                    file_path=str(file_path),
                    line_number=line_num,
                    issue_type='UnitUtils',
                    severity=severity,
                    description=description,
                    original_code=line_content[:80] + ('...' if len(line_content) > 80 else ''),
                    suggested_fix=suggestion
                ))
        
        # Audit BuiltIn casts
        for pattern, description, suggestion in self.BUILTIN_CAST_PATTERNS:
            for match in re.finditer(pattern, content):
                line_num = content[:match.start()].count('\n') + 1
                line_content = lines[line_num - 1].strip()
                
                issues.append(MigrationIssue(
                    file_path=str(file_path),
                    line_number=line_num,
                    issue_type='BuiltInCast',
                    severity='warning',
                    description=description,
                    original_code=line_content[:80] + ('...' if len(line_content) > 80 else ''),
                    suggested_fix=suggestion
                ))
        
        self.issues.extend(issues)
        return issues
    
    def print_report(self, issues: List[MigrationIssue] = None):
        """Print audit report"""
        if issues is None:
            issues = self.issues
        
        if not issues:
            print("\n✅ No migration issues found!")
            return
        
        print(f"\n📊 Migration Audit Report")
        print(f"{'='*80}")
        print(f"Total issues found: {len(issues)}")
        
        by_type = {}
        by_severity = {'error': 0, 'warning': 0}
        
        for issue in issues:
            by_type.setdefault(issue.issue_type, []).append(issue)
            by_severity[issue.severity] += 1
        
        print(f"\nBy Severity:")
        print(f"  🔴 Errors: {by_severity['error']}")
        print(f"  🟡 Warnings: {by_severity['warning']}")
        
        print(f"\nBy Type:")
        for issue_type, type_issues in by_type.items():
            print(f"  {issue_type}: {len(type_issues)}")
        
        print(f"\n📋 Detailed Issues:")
        print(f"{'='*80}")
        
        current_file = None
        for issue in sorted(issues, key=lambda x: (x.file_path, x.line_number)):
            if issue.file_path != current_file:
                current_file = issue.file_path
                print(f"\n📄 {Path(issue.file_path).name}")
            
            icon = "🔴" if issue.severity == 'error' else "🟡"
            print(f"  {icon} Line {issue.line_number}: {issue.description}")
            print(f"     Code: {issue.original_code}")
            if issue.suggested_fix:
                print(f"     💡 Fix: {issue.suggested_fix}")
    
    def create_backup(self, file_path: Path) -> Path:
        """Create backup of file"""
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        backup_path = self.backup_dir / f"{file_path.stem}_{timestamp}{file_path.suffix}"
        shutil.copy2(file_path, backup_path)
        print(f"📦 Backup created: {backup_path}")
        return backup_path
    
    def apply_fixes(self, file_path: Path, issues: List[MigrationIssue]) -> bool:
        """Apply fixes to a file (with backup)"""
        if not issues:
            return True
        
        # Create backup
        self.create_backup(file_path)
        
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                content = f.read()
            
            original_content = content
            changes_made = 0
            
            # Apply UnitUtils fixes
            unit_utils_issues = [i for i in issues if i.issue_type == 'UnitUtils']
            
            # Check if service needs to be imported
            needs_using = 'using JSE_RevitAddin_MEP_OPENINGS.Services;' not in content
            needs_service_field = 'RevitUnitConversionService' not in content
            
            if unit_utils_issues and needs_service_field:
                # Add service field (if class has fields)
                # This is a simplified approach - manual review recommended
                print("⚠️  Note: File may need RevitUnitConversionService field added manually")
            
            # Replace UnitUtils.ConvertFromInternalUnits(..., UnitTypeId.Millimeters)
            pattern1 = r'UnitUtils\.ConvertFromInternalUnits\s*\(([^,]+),\s*UnitTypeId\.Millimeters\)'
            def replace_mm_from(match):
                var = match.group(1).strip()
                return f'RevitUnitConversionService.Instance.FromInternalMillimeters({var})'
            content = re.sub(pattern1, replace_mm_from, content)
            
            # Replace UnitUtils.ConvertToInternalUnits(..., UnitTypeId.Millimeters)
            pattern2 = r'UnitUtils\.ConvertToInternalUnits\s*\(([^,]+),\s*UnitTypeId\.Millimeters\)'
            def replace_mm_to(match):
                var = match.group(1).strip()
                return f'RevitUnitConversionService.Instance.ToInternalMillimeters({var})'
            content = re.sub(pattern2, replace_mm_to, content)
            
            # Count changes
            if content != original_content:
                changes_made = len(re.findall(r'RevitUnitConversionService\.Instance', content))
                
                # Add using statement if needed
                if needs_using and 'RevitUnitConversionService' in content:
                    # Find last using statement
                    using_match = list(re.finditer(r'^using\s+[^;]+;', content, re.MULTILINE))
                    if using_match:
                        last_using = using_match[-1]
                        insert_pos = last_using.end()
                        content = (content[:insert_pos] + 
                                 '\nusing JSE_RevitAddin_MEP_OPENINGS.Services;' +
                                 content[insert_pos:])
                
                # Write changes
                with open(file_path, 'w', encoding='utf-8') as f:
                    f.write(content)
                
                print(f"✅ Applied {changes_made} fixes to {file_path.name}")
                return True
            else:
                print(f"ℹ️  No changes needed for {file_path.name}")
                return False
                
        except Exception as e:
            print(f"❌ Error applying fixes: {e}")
            return False


def main():
    parser = argparse.ArgumentParser(description='Audit Revit 2024 migration issues')
    parser.add_argument('--file', type=str, help='Single file to audit')
    parser.add_argument('--fix', action='store_true', help='Apply fixes (creates backup)')
    parser.add_argument('--yes', action='store_true', help='Skip confirmation prompt (use with --fix)')
    parser.add_argument('--audit-only', action='store_true', help='Only audit, don\'t fix')
    parser.add_argument('--services-dir', type=str, default='Services', 
                       help='Directory to audit (default: Services)')
    parser.add_argument('--recursive', action='store_true', 
                       help='Recursively audit directory')
    
    args = parser.parse_args()
    
    auditor = Revit2024MigrationAuditor()
    
    if args.file:
        # Single file mode
        file_path = Path(args.file)
        if not file_path.exists():
            file_path = Path(args.services_dir) / args.file
        
        issues = auditor.audit_file(file_path, fix=False)
        auditor.print_report(issues)
        
        if args.fix and issues:
            print(f"\n🔧 Applying fixes to {file_path.name}...")
            if args.yes:
                print("⚠️  Auto-confirming (backup will be created)...")
                auditor.apply_fixes(file_path, issues)
            else:
                response = input("⚠️  This will modify the file (backup will be created). Continue? (yes/no): ")
                if response.lower() == 'yes':
                    auditor.apply_fixes(file_path, issues)
                else:
                    print("❌ Fix cancelled by user")
    else:
        # Directory audit mode
        services_dir = Path(args.services_dir)
        if not services_dir.exists():
            print(f"❌ Directory not found: {services_dir}")
            return
        
        print(f"🔍 Auditing directory: {services_dir}")
        print(f"   Mode: {'Fix' if args.fix else 'Audit Only'}")
        
        cs_files = list(services_dir.rglob('*.cs')) if args.recursive else list(services_dir.glob('*.cs'))
        
        print(f"\nFound {len(cs_files)} C# files")
        
        for cs_file in cs_files:
            issues = auditor.audit_file(cs_file, fix=False)
        
        auditor.print_report()
        
        if args.fix and auditor.issues:
            print(f"\n⚠️  Fix mode would modify {len(auditor.issues)} files")
            print("   Use --file <filename> to fix files one at a time")


if __name__ == '__main__':
    main()

