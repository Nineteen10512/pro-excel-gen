# PRO-EXCEL v1.2.0

Professional Excel workbook generation, translation, chart export, formula-first modeling, full-sheet UI normalization, table-aware charting, chart-to-data extraction, and image-fidelity-preserving safe edits.

## v1.2.0 Update Log

### Official-grade delivery kernel

1. Import and source handling
   - CSV, TSV, and XLSX import through `import_tabular_data()`
   - Typed value coercion for numeric and percentage-like source values
   - `create_workbook_from_source()` builds stable workbooks from external data files
2. Analysis workflow
   - `analyze_table()` returns row, column, and numeric summary metadata
   - `add_analysis_sheet()` creates a formula-backed analysis sheet with `COUNT`, `SUM`, `AVERAGE`, `MIN`, and `MAX`
3. Visualization workflow
   - Source-driven chart creation is wired into imported workbooks
   - Chart source ranges remain traceable to worksheet cells
   - Visual preview rendering via `render_sheet_preview()`
4. Formula and recalculation discipline
   - `scan_formula_issues()` scans visible formula errors and bad references
   - `recalculate_workbook()` marks workbooks for full recalculation and uses LibreOffice when available
5. Official-style delivery verification
   - `verify_delivery_workbook()` runs structure, formula, contrast, visual layout, and preview checks
   - `export_xlsx()` copies and verifies final `.xlsx` output
   - `finalize_delivery_workbook()` embeds an optional `DeliveryAudit` sheet, exports, and verifies
6. Workbook control polish
   - `add_dropdown_validations()` adds editable data validation dropdowns
   - `apply_semantic_conditional_formats()` applies theme-aware conditional formats
   - `inspect_visual_layout()` detects possible clipping and blank charts
   - `create_delivery_audit_sheet()` writes compact QA evidence into the workbook

## v1.1.0 Update Log

### Medium modular refactor

1. Formula-first workbook generation
   - New workbook generation supports `formula_plan`
   - Derived columns can be formulas instead of static values
   - Modern formula mode is default, with compatibility downgrade helpers
   - Public APIs: `plan_formulas()`, `autofill_formulas()`, `downgrade_formula_mode()`
2. Unified workbook UI
   - Full-sheet UI normalization for existing workbooks
   - Standardized headers, borders, widths, row heights, filters, freeze panes, and number formats
   - Public APIs: `normalize_sheet_ui()`, `normalize_workbook_ui()`, `upgrade_sheet_layout()`
3. Theme law and contrast self-check
   - Generated and repaired formatting use theme tokens
   - Background/text contrast audit and repair are first-class
   - Public APIs: `audit_theme_contrast()`, `enforce_theme_law()`, `contrast_ratio()`
4. Chart bridge
   - Build native Excel charts from ranges and tables
   - Insert table-local charts
   - Extract workbook-native chart source data into structured rows
   - Export chart source data to a new sheet
   - Public APIs: `build_chart_from_range()`, `build_chart_from_table()`, `insert_chart_into_table_region()`, `extract_chart_source_data()`, `export_chart_data_to_sheet()`
5. Table microcharts
   - `embed_microchart_column()` inserts formula-driven in-table microchart bars
6. Existing workbook operations
   - `edit_existing_workbook()` and `apply_operations_with_audit()` provide operation-level editing and audit output
   - `translate_and_normalize_workbook()` chains translation with UI normalization and theme checks
7. Image chart reconstruction boundary
   - `infer_chart_data_from_image()` is auxiliary only and returns confidence metadata with `review_required=True`
8. Expanded smoke coverage
   - Smoke suite covers formula-first generation, theme law, contrast repair, chart bridge, chart-to-data, existing workbook edits, image fidelity, and quality gate

## v1.0.0 Update Log

### Initial release

1. Workbook generation
   - Native `.xlsx` generation for dashboards, trackers, models, and report sheets
   - Theme-aware formatting, freeze panes, conditional formatting, tables, formulas, comments, and image insertion
2. Translation workflow
   - `collect_translation_segments()`
   - `build_translation_prompt()`
   - `assess_translation_risk()`
   - `apply_translation_map()`
3. Hard translation guardrails
   - Translated output must not add any source-absent fact, explanation, or background detail
   - Formulas, numbers, units, dates, and identifiers stay stable unless user explicitly asks to localize them
4. Table translation layer
   - Growth-aware column widening
   - Row-height repair for expanded translated text
   - Conservative formatting repair to keep workbook readable
5. Image fidelity path
   - Existing workbook translation preserves media package parts and drawing relationships
   - Default smoke suite includes image-preservation regression
6. Chart support
   - Native workbook chart insertion for generated sheets
   - `render_chart_to_png()` for reusable chart assets
7. Quality discipline
   - Delivery quality gate for `.xlsx`
   - Default smoke suite covers generation, formulas, translation, chart PNG export, quality gate, and image fidelity

## Public APIs

```python
from pro_excel_gen import (
    generate,
    generate_from_rows,
    build_kpi_dashboard,
    build_chart_sheet,
    render_chart_to_png,
    import_tabular_data,
    analyze_table,
    create_workbook_from_source,
    add_analysis_sheet,
    recalculate_workbook,
    render_sheet_preview,
    verify_delivery_workbook,
    export_xlsx,
    finalize_delivery_workbook,
    scan_formula_issues,
    add_dropdown_validations,
    apply_semantic_conditional_formats,
    inspect_visual_layout,
    create_delivery_audit_sheet,
    plan_formulas,
    autofill_formulas,
    downgrade_formula_mode,
    normalize_workbook_ui,
    normalize_sheet_ui,
    enforce_theme_law,
    audit_theme_contrast,
    build_chart_from_table,
    build_chart_from_range,
    insert_chart_into_table_region,
    embed_microchart_column,
    extract_chart_source_data,
    export_chart_data_to_sheet,
    infer_chart_data_from_image,
    edit_existing_workbook,
    apply_operations_with_audit,
    translate_and_normalize_workbook,
    list_themes,
    list_templates,
    recommend_template,
    collect_translation_segments,
    build_translation_prompt,
    assess_translation_risk,
    apply_translation_map,
)
```

## Examples

Generate formula-first workbook:

```python
from pro_excel_gen import generate

spec = {
    "meta": {"formula_mode": "modern"},
    "sheets": [{
        "title": "FormulaFirst",
        "columns": [
            {"header": "Product"},
            {"header": "Revenue", "format": "currency"},
            {"header": "Cost", "format": "currency"},
        ],
        "rows": [
            {"Product": "Alpha", "Revenue": 100, "Cost": 60},
            {"Product": "Beta", "Revenue": 300, "Cost": 120},
        ],
        "formula_plan": [
            {"column": "Gross Profit", "kind": "delta", "numerator_column": "Revenue", "denominator_column": "Cost", "format": "currency"},
            {"column": "Revenue Share", "kind": "percent_of_total", "source_column": "Revenue", "format": "percent"},
        ],
        "table": {"name": "FormulaTable"},
    }],
}
generate(spec, "formula_first.xlsx")
```

Edit existing workbook and add table-local chart:

```python
from pro_excel_gen import edit_existing_workbook

edit_existing_workbook(
    "source.xlsx",
    "edited.xlsx",
    [
        {"type": "normalize_sheet_ui", "sheet": "Sales"},
        {
            "type": "insert_chart_into_table_region",
            "sheet": "Sales",
            "table": "SalesTable",
            "categories_column": "Month",
            "series_columns": ["Revenue"],
            "title": "Revenue Trend",
        },
    ],
)
```

Create stable workbook from CSV and verify delivery:

```python
from pro_excel_gen import create_workbook_from_source, finalize_delivery_workbook

create_workbook_from_source("source.csv", "working.xlsx", add_analysis=True, add_chart=True)
finalize_delivery_workbook("working.xlsx", "final.xlsx")
```

Translate existing workbook:

```python
from pro_excel_gen import collect_translation_segments, build_translation_prompt, apply_translation_map

segments = collect_translation_segments("source.xlsx")
prompt = build_translation_prompt(segments, source_lang="zh", target_lang="en")

# Call translation model, then map id -> translation.
apply_translation_map(
    "source.xlsx",
    translations,
    output_path="source.en.xlsx",
    target_lang="en",
    auto_format_tables=True,
)
```

Run smoke:

```powershell
& 'C:\Users\k1832\.cache\codex-runtimes\codex-primary-runtime\dependencies\python\python.exe' `
  'C:\Users\k1832\Documents\软件更新\pro-excel-gen\smoke_tests\run_smoke_tests.py'
```
