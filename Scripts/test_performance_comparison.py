"""
Performance Comparison Test Script for NewSleevePlacerService
Compares legacy UniversalSleevePlacerService vs refactored NewSleevePlacerService

Usage:
    python Scripts/test_performance_comparison.py

Requirements:
    - Revit add-in must be deployed
    - OptimizationFlags.UseNewSleevePlacerService must be configurable
    - placement_performance.log must be generated after each test run
"""

import time
import subprocess
import json
import os
import re
from pathlib import Path
from datetime import datetime

# Configuration
REVIT_VERSION = "2024"  # Change to "2023" if needed
LOG_DIR = Path(os.path.expanduser("~")) / "AppData" / "Roaming" / "JSE_MEP_Openings" / "Logs"
PERFORMANCE_LOG = LOG_DIR / "placement_performance.log"
SUMMARY_LOG = LOG_DIR / "placement_summary.log"
REPORT_FILE = "performance_comparison_report.json"


def parse_performance_log(log_path):
    """Parse placement_performance.log and extract timing metrics"""
    if not log_path.exists():
        return None
    
    metrics = {
        "operations": {},
        "total_time": 0,
        "operations_count": 0
    }
    
    try:
        with open(log_path, 'r', encoding='utf-8') as f:
            lines = f.readlines()
        
        # Parse timing entries
        for line in lines:
            # Format: [HH:mm:ss] OperationName: XXXms (count items, YYYms/item)
            match = re.search(r'\[(\d{2}:\d{2}:\d{2})\] (\w+): (\d+)ms', line)
            if match:
                operation = match.group(2)
                time_ms = int(match.group(3))
                
                if operation not in metrics["operations"]:
                    metrics["operations"][operation] = {
                        "total_time": 0,
                        "count": 0,
                        "avg_time": 0
                    }
                
                metrics["operations"][operation]["total_time"] += time_ms
                metrics["operations"][operation]["count"] += 1
                metrics["total_time"] += time_ms
                metrics["operations_count"] += 1
        
        # Calculate averages
        for op_name, op_data in metrics["operations"].items():
            if op_data["count"] > 0:
                op_data["avg_time"] = op_data["total_time"] / op_data["count"]
        
    except Exception as e:
        print(f"Error parsing performance log: {e}")
        return None
    
    return metrics


def parse_summary_log(log_path):
    """Parse placement_summary.log and extract summary metrics"""
    if not log_path.exists():
        return None
    
    summary = {
        "total_zones": 0,
        "placed": 0,
        "skipped": 0,
        "errors": 0,
        "success_rate": 0.0
    }
    
    try:
        with open(log_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Extract metrics from summary
        total_match = re.search(r'Total Zones: (\d+)', content)
        placed_match = re.search(r'Placed: (\d+)', content)
        skipped_match = re.search(r'Skipped: (\d+)', content)
        errors_match = re.search(r'Errors: (\d+)', content)
        rate_match = re.search(r'Success Rate: ([\d.]+)%', content)
        
        if total_match:
            summary["total_zones"] = int(total_match.group(1))
        if placed_match:
            summary["placed"] = int(placed_match.group(1))
        if skipped_match:
            summary["skipped"] = int(skipped_match.group(1))
        if errors_match:
            summary["errors"] = int(errors_match.group(1))
        if rate_match:
            summary["success_rate"] = float(rate_match.group(1))
    
    except Exception as e:
        print(f"Error parsing summary log: {e}")
        return None
    
    return summary


def update_optimization_flag(value):
    """
    Update OptimizationFlags.UseNewSleevePlacerService
    Note: This requires modifying the code or using a configuration file
    For now, this is a placeholder - actual implementation depends on your configuration system
    """
    print(f"⚠️  NOTE: Manually set OptimizationFlags.UseNewSleevePlacerService = {value} before running test")
    print(f"   Location: Services/OptimizationFlags.cs")
    return True


def clear_logs():
    """Clear performance and summary logs before test run"""
    if PERFORMANCE_LOG.exists():
        PERFORMANCE_LOG.unlink()
        print(f"Cleared {PERFORMANCE_LOG}")
    
    if SUMMARY_LOG.exists():
        SUMMARY_LOG.unlink()
        print(f"Cleared {SUMMARY_LOG}")


def run_test(config_name, use_new_service):
    """
    Run placement with specific flag configuration
    
    Args:
        config_name: Name of configuration (e.g., "Legacy", "Refactored")
        use_new_service: Boolean - whether to use NewSleevePlacerService
    
    Returns:
        Dictionary with performance metrics
    """
    print(f"\n{'='*60}")
    print(f"Testing {config_name} Service (UseNewSleevePlacerService = {use_new_service})...")
    print(f"{'='*60}")
    
    # Update flag (placeholder - requires manual configuration)
    update_optimization_flag(use_new_service)
    
    # Clear previous logs
    clear_logs()
    
    # Wait for user to run placement in Revit
    input(f"\n⏸️  Please run sleeve placement in Revit, then press ENTER to continue...")
    
    # Parse logs
    print("\n📊 Parsing performance logs...")
    performance_metrics = parse_performance_log(PERFORMANCE_LOG)
    summary_metrics = parse_summary_log(SUMMARY_LOG)
    
    if not performance_metrics:
        print("⚠️  Warning: Could not parse performance log")
        performance_metrics = {"operations": {}, "total_time": 0, "operations_count": 0}
    
    if not summary_metrics:
        print("⚠️  Warning: Could not parse summary log")
        summary_metrics = {"total_zones": 0, "placed": 0, "skipped": 0, "errors": 0, "success_rate": 0.0}
    
    # Combine metrics
    combined = {
        "config_name": config_name,
        "use_new_service": use_new_service,
        "performance": performance_metrics,
        "summary": summary_metrics,
        "timestamp": datetime.now().isoformat()
    }
    
    print(f"✅ Test complete: {summary_metrics['placed']} sleeves placed in {performance_metrics.get('total_time', 0)}ms")
    
    return combined


def compare_performance():
    """Compare legacy vs refactored performance"""
    print("\n" + "="*60)
    print("PERFORMANCE COMPARISON TEST")
    print("="*60)
    print(f"Log Directory: {LOG_DIR}")
    print(f"Report File: {REPORT_FILE}")
    print("\n⚠️  IMPORTANT: This script requires manual configuration:")
    print("   1. Set OptimizationFlags.UseNewSleevePlacerService = false")
    print("   2. Run placement in Revit")
    print("   3. Press ENTER when done")
    print("   4. Set OptimizationFlags.UseNewSleevePlacerService = true")
    print("   5. Run placement in Revit again")
    print("   6. Press ENTER when done")
    print("\n" + "="*60)
    
    # Test Legacy Service
    legacy_metrics = run_test("Legacy", False)
    
    # Test Refactored Service
    refactored_metrics = run_test("Refactored", True)
    
    # Calculate speedup
    legacy_time = legacy_metrics["performance"].get("total_time", 1)
    refactored_time = refactored_metrics["performance"].get("total_time", 1)
    
    if legacy_time > 0 and refactored_time > 0:
        speedup = legacy_time / refactored_time
    else:
        speedup = 1.0
    
    # Generate comparison report
    report = {
        "test_date": datetime.now().isoformat(),
        "legacy": legacy_metrics,
        "refactored": refactored_metrics,
        "comparison": {
            "speedup": speedup,
            "time_saved_ms": legacy_time - refactored_time,
            "time_saved_percent": ((legacy_time - refactored_time) / legacy_time * 100) if legacy_time > 0 else 0,
            "placed_difference": refactored_metrics["summary"]["placed"] - legacy_metrics["summary"]["placed"],
            "success_rate_improvement": refactored_metrics["summary"]["success_rate"] - legacy_metrics["summary"]["success_rate"]
        }
    }
    
    # Save report
    with open(REPORT_FILE, 'w', encoding='utf-8') as f:
        json.dump(report, f, indent=2)
    
    # Print summary
    print("\n" + "="*60)
    print("PERFORMANCE COMPARISON RESULTS")
    print("="*60)
    print(f"Legacy Service:")
    print(f"  Total Time: {legacy_time}ms")
    print(f"  Placed: {legacy_metrics['summary']['placed']} sleeves")
    print(f"  Success Rate: {legacy_metrics['summary']['success_rate']:.2f}%")
    print(f"\nRefactored Service:")
    print(f"  Total Time: {refactored_time}ms")
    print(f"  Placed: {refactored_metrics['summary']['placed']} sleeves")
    print(f"  Success Rate: {refactored_metrics['summary']['success_rate']:.2f}%")
    print(f"\n📈 Improvement:")
    print(f"  Speedup: {speedup:.2f}x")
    print(f"  Time Saved: {legacy_time - refactored_time}ms ({(legacy_time - refactored_time) / legacy_time * 100:.1f}%)")
    print(f"\n📄 Full report saved to: {REPORT_FILE}")
    print("="*60)


if __name__ == "__main__":
    try:
        compare_performance()
    except KeyboardInterrupt:
        print("\n\n⚠️  Test interrupted by user")
    except Exception as e:
        print(f"\n\n❌ Error: {e}")
        import traceback
        traceback.print_exc()

