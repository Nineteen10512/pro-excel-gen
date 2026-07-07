from __future__ import annotations

from pathlib import Path

from .chart_bridge import (
    embed_microchart_column,
    export_chart_data_to_sheet,
    insert_chart_into_table_region,
)
from .formula_planner import autofill_formulas
from .normalizer import normalize_sheet_ui, normalize_workbook_ui
from .theme_law import enforce_theme_law
from .translation import apply_translation_map


def apply_operations_with_audit(workbook_path: str, operations: list[dict], output_path: str | None = None, *, theme: str = "corporate_formal") -> dict:
    target = output_path or workbook_path
    current = workbook_path
    audit = []
    for operation in operations:
        op_type = operation.get("type")
        if op_type == "normalize_sheet_ui":
            result = normalize_sheet_ui(current, output_path=target, sheet_name=operation.get("sheet"), theme=theme)
        elif op_type == "normalize_workbook_ui":
            result = normalize_workbook_ui(current, output_path=target, theme=theme)
        elif op_type == "autofill_formulas":
            result = autofill_formulas(current, output_path=target, sheet_name=operation.get("sheet"), table_name=operation.get("table"))
        elif op_type == "insert_chart_into_table_region":
            result = insert_chart_into_table_region(
                current,
                output_path=target,
                sheet_name=operation["sheet"],
                table_name=operation["table"],
                categories_column=operation["categories_column"],
                series_columns=operation["series_columns"],
                chart_type=operation.get("chart_type", "column"),
                title=operation.get("title", "Chart"),
                anchor=operation.get("anchor", "H2"),
                theme=theme,
            )
        elif op_type == "embed_microchart_column":
            result = embed_microchart_column(
                current,
                output_path=target,
                sheet_name=operation["sheet"],
                source_column=operation["source_column"],
                target_header=operation.get("target_header", "Trend"),
                start_row=operation.get("start_row", 2),
                end_row=operation.get("end_row"),
            )
        elif op_type == "export_chart_data_to_sheet":
            result = export_chart_data_to_sheet(current, output_path=target, sheet_name=operation.get("sheet_name", "ChartData"))
        elif op_type == "enforce_theme_law":
            result = enforce_theme_law(current, output_path=target, theme=theme)
        else:
            result = {"skipped": True, "reason": f"Unsupported operation: {op_type}"}
        audit.append({"operation": op_type, "result": result})
        current = target
    return {"output_path": str(target), "operations": audit}


def edit_existing_workbook(workbook_path: str, output_path: str, operations: list[dict], *, theme: str = "corporate_formal") -> dict:
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    return apply_operations_with_audit(workbook_path, operations, output_path=output_path, theme=theme)


def translate_and_normalize_workbook(
    workbook_path: str,
    translations,
    *,
    output_path: str,
    target_lang: str | None = None,
    theme: str = "corporate_formal",
) -> dict:
    translated = apply_translation_map(workbook_path, translations, output_path=output_path, target_lang=target_lang)
    normalized = normalize_workbook_ui(translated, output_path=output_path, theme=theme)
    return {"output_path": output_path, "translation_output": translated, "normalization": normalized}
