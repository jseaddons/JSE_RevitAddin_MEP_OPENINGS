#!/usr/bin/env python3
"""
Advanced C# Null Reference Warnings Analyzer
Reads build output and generates targeted fixes with context awareness
"""

import os
import re
from pathlib import Path
from typing import List, Dict, Tuple, Optional
from collections import defaultdict

class WarningsAnalyzer:
    def __init__(self, project_root):
        self.project_root = Path(project_root) if isinstance(project_root, str) else project_root
        self.warnings_by_file = defaultdict(list)
        self.patterns = {
            "CS8602": {
                "name": "Dereference of possibly null reference",
                "regex": r"(?:\.(?:Category|Location|Parent|Type|Definition|Document)\b)",
                "solution": "Use null-conditional operator (?.)"
            },
            "CS8604": {
                "name": "Possible null reference argument",
                "regex": r"(?:\w+\(.*null.*\))",
                "solution": "Check for null before passing to method"
            },
            "CS8600": {
                "name": "Converting null to non-nullable",
                "regex": r"(?:=\s*(?:\w+\s*\?\.\s*\w+|null)(?:\s*??|;))",
                "solution": "Use null coalescing (??) or nullable type (?)"
            },
            "CS8625": {
                "name": "Cannot convert null literal to non-nullable",
                "regex": r"(?:\w+\s*=\s*null[!]?;)",
                "solution": "Use string.Empty or make type nullable"
            },
            "CS8629": {
                "name": "Nullable value type may be null",
                "regex": r"(?:double|int|decimal|float)\s+\w+\s*=\s*(?:\w+\s*\?\.|\w+\?\s*\.|\w+\?\.)",
                "solution": "Use null coalescing (??) to provide default value"
            }
        }

    def parse_build_output(self, output_file: str) -> None:
        """Parse build output file for warnings."""
        try:
            with open(output_file, 'r', encoding='utf-8', errors='ignore') as f:
                for line in f:
                    # Pattern: filepath(line,col): error CSxxxx: message
                    match = re.search(r'([^:]+)\((\d+),\d+\):\s*warning\s+(CS\d+):\s*(.+)', line)
                    if match:
                        file_path, line_num, code, message = match.groups()
                        file_name = Path(file_path).name
                        self.warnings_by_file[file_name].append({
                            'line': int(line_num),
                            'code': code,
                            'message': message.strip()
                        })
        except Exception as e:
            print(f"❌ Error parsing build output: {e}")

    def extract_context(self, file_path: str, line_num: int, context_lines: int = 3) -> Tuple[str, List[str]]:
        """Extract code context around warning line."""
        try:
            with open(file_path, 'r', encoding='utf-8') as f:
                lines = f.readlines()
            
            start = max(0, line_num - context_lines - 1)
            end = min(len(lines), line_num + context_lines)
            
            code_line = lines[line_num - 1] if line_num <= len(lines) else ""
            context = lines[start:end]
            
            return code_line.strip(), context
        except:
            return "", []

    def suggest_fix(self, warning_code: str, code_line: str, context: List[str]) -> Optional[str]:
        """Suggest a specific fix based on code context."""
        suggestions = {
            "CS8602": self._suggest_cs8602_fix,
            "CS8604": self._suggest_cs8604_fix,
            "CS8600": self._suggest_cs8600_fix,
            "CS8625": self._suggest_cs8625_fix,
            "CS8629": self._suggest_cs8629_fix,
        }
        
        if warning_code in suggestions:
            return suggestions[warning_code](code_line, context)
        return None

    def _suggest_cs8602_fix(self, line: str, context: List[str]) -> Optional[str]:
        """Suggest fix for dereference of possibly null."""
        if '.Category.Name' in line:
            return line.replace('.Category.Name', '?.Category?.Name ?? "Unknown"')
        elif '.Location.Position' in line:
            return line.replace('.Location.Position', '?.Location?.Position')
        elif '.GetParameters()' in line:
            return line.replace('.GetParameters()', '?.GetParameters()')
        
        # Generic chained property access
        match = re.search(r'(\w+)\.(\w+)\.', line)
        if match:
            obj, prop = match.groups()
            return re.sub(rf'{re.escape(obj)}\.{re.escape(prop)}\.', f'{obj}?.{prop}?.', line)
        return None

    def _suggest_cs8604_fix(self, line: str, context: List[str]) -> Optional[str]:
        """Suggest fix for null reference argument."""
        # Look for method calls with null
        if 'null,' in line or 'null)' in line:
            return line.replace('null,', 'string.Empty,').replace('null)', 'string.Empty)')
        return None

    def _suggest_cs8600_fix(self, line: str, context: List[str]) -> Optional[str]:
        """Suggest fix for null to non-nullable conversion."""
        if re.search(r'string\s+\w+\s*=\s*\w+\s*\?\.', line):
            return re.sub(r'(string\s+\w+\s*=\s*)(\w+\s*\?\.)', r'\1\2 ?? string.Empty', line)
        elif re.search(r'double\s+\w+\s*=\s*\w+\s*\?\.', line):
            return re.sub(r'(double\s+\w+\s*=\s*)(\w+\s*\?\.)', r'\1\2 ?? 0.0', line)
        return None

    def _suggest_cs8625_fix(self, line: str, context: List[str]) -> Optional[str]:
        """Suggest fix for null literal to non-nullable."""
        if 'null!' in line:
            return line.replace('null!', 'null')
        
        match = re.search(r'(string|int|double|decimal|bool)\s+(\w+)\s*=\s*null', line)
        if match:
            type_name, var_name = match.groups()
            defaults = {'string': 'string.Empty', 'int': '0', 'double': '0.0', 'decimal': '0m', 'bool': 'false'}
            default_val = defaults.get(type_name, type_name + '.Empty')
            return re.sub(r'string\s+' + var_name + r'\s*=\s*null', f'string? {var_name} = null', line)
        return None

    def _suggest_cs8629_fix(self, line: str, context: List[str]) -> Optional[str]:
        """Suggest fix for nullable value type may be null."""
        if '?.' in line and re.search(r'(double|int|decimal|float)\s+\w+\s*=', line):
            # Add null coalescing
            match = re.search(r'((?:double|int|decimal|float)\s+\w+\s*=\s*[^;]+)', line)
            if match and '??' not in match.group(1):
                type_match = re.search(r'(double|int|decimal|float)', line)
                if type_match:
                    type_name = type_match.group(1)
                    defaults = {'double': '0.0', 'int': '0', 'decimal': '0m', 'float': '0f'}
                    default = defaults.get(type_name, '0')
                    
                    # Insert ?? default before semicolon
                    if ';' in line:
                        return line.replace(';', f' ?? {default};')
        return None

    def generate_report(self, output_file: str = "WARNINGS_FIXES_REPORT.md") -> None:
        """Generate detailed report with suggested fixes."""
        report = []
        report.append("# Null Reference Warnings Fix Report\n")
        report.append(f"Generated: {os.popen('date /t').read().strip()}\n")
        report.append("---\n\n")
        
        total_warnings = 0
        
        for file_name, warnings in sorted(self.warnings_by_file.items()):
            report.append(f"## {file_name}\n\n")
            report.append(f"**Total Warnings: {len(warnings)}**\n\n")
            
            for warning in sorted(warnings, key=lambda w: w['line']):
                file_path = self.project_root / "Services" / file_name
                if not file_path.exists():
                    file_path = self.project_root / "Views" / file_name
                if not file_path.exists():
                    file_path = self.project_root / file_name
                
                code_line, context = self.extract_context(str(file_path), warning['line'])
                
                report.append(f"### Line {warning['line']}: {warning['code']}\n\n")
                report.append(f"**Error:** {warning['message']}\n\n")
                report.append(f"**Current Code:**\n```csharp\n{code_line}\n```\n\n")
                
                suggestion = self.suggest_fix(warning['code'], code_line, context)
                if suggestion:
                    report.append(f"**Suggested Fix:**\n```csharp\n{suggestion}\n```\n\n")
                
                pattern_info = self.patterns.get(warning['code'], {})
                report.append(f"**Solution Strategy:** {pattern_info.get('solution', 'Review context')}\n\n")
                report.append("---\n\n")
                
                total_warnings += 1
        
        report.append(f"\n## Summary\n\n")
        report.append(f"- **Total Warnings:** {total_warnings}\n")
        report.append(f"- **Files Affected:** {len(self.warnings_by_file)}\n")
        report.append(f"- **Warning Breakdown:**\n")
        
        # Count by code
        by_code = defaultdict(int)
        for warnings in self.warnings_by_file.values():
            for w in warnings:
                by_code[w['code']] += 1
        
        for code in sorted(by_code.keys()):
            count = by_code[code]
            pattern = self.patterns.get(code, {})
            report.append(f"  - **{code}**: {count} ({pattern.get('name', 'Unknown')})\n")
        
        # Write report
        output_path = self.project_root / output_file
        with open(output_path, 'w', encoding='utf-8') as f:
            f.writelines(report)
        
        print(f"✅ Report generated: {output_path}")

def main():
    import sys
    
    project_root = r"c:\JSE_CSharp_Projects\JSE_MEPOPENING_23"
    
    if len(sys.argv) > 1:
        project_root = sys.argv[1]
    
    project_root = Path(project_root)
    analyzer = WarningsAnalyzer(project_root)
    
    # Try to find build output
    build_output_paths = [
        project_root / "build_warnings_full.txt",
        project_root / "build_output.txt",
    ]
    
    for build_output in build_output_paths:
        if build_output.exists():
            print(f"📖 Parsing build output: {build_output}")
            analyzer.parse_build_output(str(build_output))
            break
    
    if analyzer.warnings_by_file:
        analyzer.generate_report()
    else:
        print("⚠️  No build output found. Please run build first.")

if __name__ == "__main__":
    main()
