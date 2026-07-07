# Workbook Spec

`pro_excel_gen.generate()` accepts workbook specs shaped like:

```python
{
  "meta": {"title": "Board KPI Pack", "formula_mode": "modern"},
  "sheets": [
    {
      "title": "Dashboard",
      "freeze_panes": [1, 0],
      "columns": [
        {"header": "Metric", "width": 20},
        {"header": "Value", "format": "currency"},
      ],
      "rows": [
        {"Metric": "Revenue", "Value": 1280000},
        {"Metric": "Margin", "Value": 0.42},
      ],
      "formula_plan": [
        {"column": "Share", "kind": "percent_of_total", "source_column": "Value", "format": "percent"}
      ],
      "table": {"name": "DashboardTable", "style": "Table Style Medium 2"},
      "charts": [
        {
          "type": "column",
          "title": "Revenue trend",
          "anchor": "E2",
          "categories_column": "Metric",
          "series_columns": ["Value"],
        }
      ],
      "images": [
        {"path": "logo.png", "cell": "G1", "x_scale": 0.5, "y_scale": 0.5}
      ],
    }
  ]
}
```

## Formula-first rules

New workbook generation should use formulas for derived values whenever feasible.

Supported `formula_plan` kinds in v1.1:

- `delta`: `numerator_column - denominator_column`
- `ratio`: `numerator_column / denominator_column` with `IFERROR`
- `percent_of_total`: row value divided by column total
- `running_total`: cumulative total by row
- `lookup`: modern `XLOOKUP` with compatibility downgrade support

Use `formula_mode="modern"` by default. Use `downgrade_formula_mode()` or `formula_mode="compat"` when the workbook must avoid modern Excel functions.

## Existing workbook operations

`edit_existing_workbook()` accepts operation specs:

```python
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
  {"type": "embed_microchart_column", "sheet": "Sales", "source_column": 2},
]
```

Run `audit_theme_contrast()` after style-heavy edits. Use `enforce_theme_law()` to repair low-contrast text/background combinations.

## Official-style delivery workflow

For production workbook delivery:

```python
from pro_excel_gen import (
    create_workbook_from_source,
    add_dropdown_validations,
    apply_semantic_conditional_formats,
    verify_delivery_workbook,
    finalize_delivery_workbook,
)

create_workbook_from_source("source.csv", "working.xlsx", add_analysis=True, add_chart=True)
add_dropdown_validations(
    "working.xlsx",
    validations=[{"sheet": "Data", "range": "D2:D200", "options": ["Open", "Closed", "At risk"]}],
)
apply_semantic_conditional_formats(
    "working.xlsx",
    rules=[{"sheet": "Data", "range": "C2:C200", "type": "color_scale"}],
)
verify_delivery_workbook("working.xlsx", render=True)
finalize_delivery_workbook("working.xlsx", "final.xlsx")
```

Required final checks:

- workbook opens
- formula scan is clean
- visual preview renders
- contrast audit passes or repairs are reported
- charts are not blank
- final `.xlsx` is exported and verified

## Translation prompt rules

Use `build_translation_prompt()` default unless user gives stricter rules.

Hard rules:

1. Translate only source-present meaning.
2. Do not add facts, examples, summaries, or explanations absent from source workbook.
3. Keep formulas, numbers, units, identifiers, names, and dates unchanged unless explicitly asked.
4. Return one translation per segment id.
5. If source ambiguous, preserve ambiguity.

## Safe edit boundary

Prefer XML patch path for existing workbooks when user asks for translation or text-only localization. Avoid broad workbook rebuild when image, chart, or drawing fidelity matters.
