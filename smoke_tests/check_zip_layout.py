from __future__ import annotations

import argparse
import sys
import zipfile
from pathlib import Path


def check_zip_layout(zip_path: Path) -> list[str]:
    errors: list[str] = []
    if not zip_path.exists():
        return [f"Zip not found: {zip_path}"]

    with zipfile.ZipFile(zip_path) as zf:
        names = [name.replace("\\", "/") for name in zf.namelist() if not name.endswith("/")]

    if not names:
        return ["Zip is empty"]

    roots = {name.split("/", 1)[0] for name in names}
    if roots != {"pro-excel-gen"}:
        errors.append(f"Expected single root 'pro-excel-gen', got: {sorted(roots)}")

    forbidden_prefix = "pro-excel-gen/pro-excel-gen/"
    offenders = [name for name in names if name.startswith(forbidden_prefix)]
    if offenders:
        errors.append(f"Nested skill root detected: {forbidden_prefix}")

    required = {
        "pro-excel-gen/SKILL.md",
        "pro-excel-gen/requirements.txt",
        "pro-excel-gen/pro_excel_gen/__init__.py",
        "pro-excel-gen/smoke_tests/run_smoke_tests.py",
    }
    missing = sorted(required - set(names))
    if missing:
        errors.append(f"Missing required files: {missing}")

    return errors


def main() -> int:
    parser = argparse.ArgumentParser(description="Validate Pro Excel Gen zip root layout.")
    parser.add_argument("zip_path", type=Path)
    args = parser.parse_args()

    errors = check_zip_layout(args.zip_path)
    if errors:
        for error in errors:
            print(f"FAIL {error}")
        return 1

    print(f"PASS zip layout {args.zip_path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
