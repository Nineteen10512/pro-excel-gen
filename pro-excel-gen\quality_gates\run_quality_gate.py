from __future__ import annotations

import argparse
import json
import sys
import zipfile
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

from pro_excel_gen.quality import inspect_workbook
from pro_excel_gen.theme_law import audit_theme_contrast
from pro_excel_gen.delivery_kernel import render_sheet_preview
from pro_excel_gen.quality import scan_formula_issues
from pro_excel_gen.workbook_controls import inspect_visual_layout


def _issue(level: str, code: str, message: str) -> dict:
    return {"level": level, "code": code, "message": message}


def inspect_file(path: Path) -> dict:
    report = {"path": str(path), "kind": "xlsx", "errors": [], "warnings": [], "metrics": {}}
    if path.suffix.lower() != ".xlsx":
        report["errors"].append(_issue("error", "unsupported_file", f"Unsupported file type: {path.suffix}"))
        return report
    if not path.exists():
        report["errors"].append(_issue("error", "missing_file", "Workbook file not found"))
        return report

    try:
        with zipfile.ZipFile(path) as zf:
            names = set(zf.namelist())
        if "xl/workbook.xml" not in names:
            report["errors"].append(_issue("error", "missing_workbook_xml", "Missing xl/workbook.xml"))
            return report
    except zipfile.BadZipFile:
        report["errors"].append(_issue("error", "bad_zip", "File is not a valid XLSX zip package"))
        return report

    try:
        metrics = inspect_workbook(path)
        metrics["theme_audit"] = audit_theme_contrast(path)
        metrics["formula_scan"] = scan_formula_issues(path)
        metrics["visual_layout"] = inspect_visual_layout(path)
    except Exception as exc:
        report["errors"].append(_issue("error", "open_failed", f"Workbook inspection failed: {exc}"))
        return report

    report["metrics"] = metrics
    if metrics["sheet_count"] == 0:
        report["errors"].append(_issue("error", "empty_workbook", "Workbook has no sheets"))
    if metrics["formula_error_hits"]:
        report["warnings"].append(_issue("warning", "formula_error_text", "Workbook contains formula error text"))
    if metrics["formula_scan"]["error_count"]:
        report["warnings"].append(_issue("warning", "formula_scan_issues", "Formula scan found error tokens or bad references"))
    if metrics["placeholder_hits"]:
        report["warnings"].append(_issue("warning", "placeholder_text", "Workbook contains placeholder-like text"))
    if metrics["media_count"] and metrics["drawing_count"] == 0:
        report["warnings"].append(_issue("warning", "media_without_drawings", "Media present but drawing XML missing"))
    if not metrics["theme_audit"]["contrast_pass"]:
        report["warnings"].append(_issue("warning", "low_contrast", "Workbook contains low-contrast text/background pairs"))
    if not metrics["visual_layout"]["passed"]:
        report["warnings"].append(_issue("warning", "visual_layout", "Visual layout audit found possible clipping or blank charts"))
    try:
        preview = render_sheet_preview(path)
        report["metrics"]["preview"] = preview
    except Exception as exc:
        report["warnings"].append(_issue("warning", "preview_render_failed", f"Preview render failed: {exc}"))
    return report


def main() -> int:
    parser = argparse.ArgumentParser(description="Run delivery quality gate for PRO-EXCEL workbooks.")
    parser.add_argument("files", nargs="+", type=Path)
    parser.add_argument("--json-report", type=Path)
    parser.add_argument("--strict", action="store_true")
    args = parser.parse_args()

    reports = [inspect_file(path) for path in args.files]
    summary = {
        "files": reports,
        "error_count": sum(len(item["errors"]) for item in reports),
        "warning_count": sum(len(item["warnings"]) for item in reports),
    }
    if args.json_report:
        args.json_report.parent.mkdir(parents=True, exist_ok=True)
        args.json_report.write_text(json.dumps(summary, ensure_ascii=False, indent=2), encoding="utf-8")

    for report in reports:
        status = "FAIL" if report["errors"] else "WARN" if report["warnings"] else "PASS"
        print(f"{status} {report['kind']} {report['path']}")
        for err in report["errors"]:
            print(f"  ERROR {err['code']}: {err['message']}")
        for warn in report["warnings"]:
            print(f"  WARN {warn['code']}: {warn['message']}")

    if summary["error_count"]:
        return 1
    if args.strict and summary["warning_count"]:
        return 1
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
