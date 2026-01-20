import argparse
import os
import sys
import re

REPLACEMENTS = [
    (re.compile(r"\.ClashZoneStorage\s*\.\s*ClashZones"), ".ClashZoneStorage.AllZones"),
    (re.compile(r"\.ClashZoneStorage\s*\?\.\s*ClashZones"), ".ClashZoneStorage?.AllZones"),
    (re.compile(r"\?\.ClashZoneStorage\s*\.\s*ClashZones"), "?.ClashZoneStorage.AllZones"),
    (re.compile(r"\?\.ClashZoneStorage\s*\?\.\s*ClashZones"), "?.ClashZoneStorage?.AllZones"),
]


def process_file(path: str) -> bool:
    try:
        with open(path, "r", encoding="utf-8") as f:
            original = f.read()
    except (OSError, UnicodeDecodeError) as exc:
        print(f"[SKIP] {path}: {exc}")
        return False

    updated = original
    for pattern, replacement in REPLACEMENTS:
        updated = pattern.sub(replacement, updated)
    if updated == original:
        return False

    with open(path, "w", encoding="utf-8") as f:
        f.write(updated)

    print(f"[UPDATED] {path}")
    return True


def iter_files(root: str):
    if os.path.isfile(root):
        yield root
        return

    for dirpath, _, filenames in os.walk(root):
        for name in filenames:
            if name.endswith(".cs"):
                yield os.path.join(dirpath, name)


def main(argv=None):
    parser = argparse.ArgumentParser(description="Replace .ClashZoneStorage.ClashZones with .ClashZoneStorage.AllZones")
    parser.add_argument("paths", nargs="+", help="File or directory paths to process")
    args = parser.parse_args(argv)

    updated_any = False
    for path in args.paths:
        for file_path in iter_files(path):
            if process_file(file_path):
                updated_any = True

    if not updated_any:
        print("No matches found.")


if __name__ == "__main__":
    sys.exit(main())
