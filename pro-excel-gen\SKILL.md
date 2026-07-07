---
name: pro-excel-gen
description: Use when Codex needs to generate, translate, audit, or safely edit professional Excel workbooks (`.xlsx`) with formulas, tables, charts, images, and layout-sensitive spreadsheet content. Trigger for KPI dashboards, finance models, business trackers, workbook localization between Chinese and English, table translation, formula-safe workbook updates, and image-fidelity-preserving spreadsheet delivery.
---

# Pro Excel Gen

Generate new `.xlsx` workbooks, translate existing `.xlsx` files, normalize workbook UI, build/extract charts, and edit existing sheets without damaging formulas, tables, charts, or embedded images.

## Workflow

1. Detect mode first.
- New workbook: use `pro_excel_gen.generate()` or `generate_from_rows()`.
- Source import: use `import_tabular_data()` or `create_workbook_from_source()` for CSV, TSV, or XLSX inputs.
- Analysis: use `analyze_table()` and `add_analysis_sheet()` to create formula-backed analysis output.
- Formula-first workbook: use `formula_plan` so derived values become Excel formulas instead of static values.
- Existing workbook edit: use `edit_existing_workbook()` with operation specs.
- UI normalization: use `normalize_sheet_ui()` or `normalize_workbook_ui()`.
- Data validation and controls: use `add_dropdown_validations()` for editable categorical fields.
- Conditional formatting: use `apply_semantic_conditional_formats()` for theme-aware visual signals.
- Data to chart: use `build_chart_from_range()`, `build_chart_from_table()`, or `insert_chart_into_table_region()`.
- Chart to data: use `extract_chart_source_data()` or `export_chart_data_to_sheet()`.
- Translation: use `collect_translation_segments()`, `build_translation_prompt()`, `assess_translation_risk()`, then `apply_translation_map()`.
- Template pick: use `recommend_template()`.
- Chart asset export: use `render_chart_to_png()`.
- Data cleaning: use `clean_workbook()` for one-stop cleaning, or individual functions: `clean_missing()`, `clean_duplicates()`, `clean_outliers()`, `clean_whitespace()`, `normalize_headers()`, `type_infer_and_cast()`, `batch_fill()`.
- Statistical analysis: use `descriptive_stats()` for enhanced stats, `correlation_analysis()` for correlation matrix, `trend_analysis()` for moving average and period-over-period, `group_aggregate()` for group-by, `pivot_analysis()` for pivot tables. Use `add_statistics_sheet()` to write results into the workbook.
- Machine learning: use `auto_classify()` for automatic classification, `auto_regress()` for regression, `feature_importance()` for feature ranking, `model_report()` for formatted evaluation.

2. Prefer safe edit over broad rewrite.
- For existing workbooks, keep original package structure whenever possible.
- In translation flow, do not rebuild workbook through a writer if XML patching is enough.
- Preserve formulas, media parts, drawing parts, chart parts, workbook metadata, and nearby placement.

3. Enforce hard translation guardrails.
- Translate only source-present meaning.
- Do not add facts, explanations, examples, or context absent from source workbook.
- Keep numbers, formulas, units, dates, names, identifiers, and placeholders stable unless user explicitly asks to localize them.

4. Treat table translation as layout work, not text-only work.
- Estimate text growth before write-back.
- If translated cell text expands materially, widen columns and increase row heights conservatively.
- Keep workbook readable. Accuracy first. Style must stay clean.

5. Apply v1.1 formula-first and theme-law defaults.
- In new workbooks, use formulas for summaries, ratios, shares, deltas, running totals, and lookup outputs whenever feasible.
- Use `formula_mode="modern"` by default and downgrade only when compatibility is required.
- Use theme colors only for generated or repaired formatting.
- Run background/text contrast self-check; repair low-contrast combinations.
- For existing workbook edits, the user has approved full-sheet UI normalization when requested.

## Public APIs

Read package file `pro_excel_gen/__init__.py` for stable public surface. Core entrypoints:

- `generate()`
- `generate_from_rows()`
- `build_kpi_dashboard()`
- `build_chart_sheet()`
- `render_chart_to_png()`
- `import_tabular_data()`
- `analyze_table()`
- `create_workbook_from_source()`
- `add_analysis_sheet()`
- `recalculate_workbook()`
- `render_sheet_preview()`
- `verify_delivery_workbook()`
- `export_xlsx()`
- `finalize_delivery_workbook()`
- `scan_formula_issues()`
- `add_dropdown_validations()`
- `apply_semantic_conditional_formats()`
- `inspect_visual_layout()`
- `create_delivery_audit_sheet()`
- `plan_formulas()`
- `autofill_formulas()`
- `downgrade_formula_mode()`
- `normalize_sheet_ui()`
- `normalize_workbook_ui()`
- `upgrade_sheet_layout()`
- `audit_theme_contrast()`
- `enforce_theme_law()`
- `build_chart_from_range()`
- `build_chart_from_table()`
- `insert_chart_into_table_region()`
- `embed_microchart_column()`
- `extract_chart_source_data()`
- `export_chart_data_to_sheet()`
- `infer_chart_data_from_image()`
- `edit_existing_workbook()`
- `apply_operations_with_audit()`
- `translate_and_normalize_workbook()`
- `list_themes()`
- `list_templates()`
- `recommend_template()`
- `collect_translation_segments()`
- `build_translation_prompt()`
- `assess_translation_risk()`
- `apply_translation_map()`
- `TRANSLATION_SELF_AUDIT_RULES`
- `clean_missing()`
- `clean_duplicates()`
- `clean_outliers()`
- `clean_whitespace()`
- `normalize_headers()`
- `type_infer_and_cast()`
- `batch_fill()`
- `clean_workbook()`
- `descriptive_stats()`
- `correlation_analysis()`
- `trend_analysis()`
- `group_aggregate()`
- `pivot_analysis()`
- `add_statistics_sheet()`
- `auto_classify()`
- `auto_regress()`
- `feature_importance()`
- `model_report()`

## Translation Rules

Use prompt built by `build_translation_prompt()` unless user supplies stricter instructions.

During translation:

- Translate sheet names, textual cells, inline strings, and chart text.
- Do not translate formulas or cached numeric/date values.
- Preserve images and chart package parts exactly.
- Keep translated content near original information area.
- If layout risk high, auto-apply column width and row height repair.

## Verification

Run:

- `smoke_tests/run_smoke_tests.py`
- `smoke_tests/run_image_regression.py`
- `quality_gates/run_quality_gate.py <xlsx>`

Before final delivery, verify:

- Workbook opens.
- Source CSV/XLSX data is typed and imported cleanly.
- Formula cells still present.
- Formula scan has no bad references or visible error tokens.
- Recalculation is marked or completed.
- At least one sheet/range preview is rendered for visual QA when layout changed.
- Final `.xlsx` is exported and verified.
- New workbook derived values use formulas when feasible.
- Theme-law contrast audit passes or repairs are reported.
- Workbook-native chart source extraction works when chart-to-data is requested.
- Image/media count unchanged after translation.
- Table text translated.
- No source-absent information added.

## References

Load `references/workbook-spec.md` when you need semantic workbook spec examples, translation prompt guidance, or safe-edit boundaries.
