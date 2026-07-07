---
name: pro-excel-gen
description: Use when Codex needs to generate, translate, audit, analyze, visualize, or safely edit professional Excel workbooks (`.xlsx`, CSV/TSV sources) with formulas, tables, charts, images, pandas/ML helpers, and layout-sensitive content. Trigger for 中文 Excel 生成/翻译/美化/分析/图表/甘特图/销售漏斗, KPI dashboards, finance models, workbook localization, formula-safe updates, and image-fidelity delivery.
---

# Pro Excel Gen

Generate new `.xlsx` workbooks, translate existing `.xlsx` files, normalize workbook UI, build/extract charts, and edit existing sheets without damaging formulas, tables, charts, or embedded images.

## 中文快用

- 中文触发可直接说：生成 Excel、翻译 Excel、Excel 美化、表格转图表、图表转数据、销售漏斗、项目甘特图、CSV 转 xlsx、用公式建模、数据清洗、统计分析、机器学习向导。
- 中文输入优先保留业务语义；输出可以按用户要求中英互译。
- 任何翻译任务必须自审：不得加入源文件没有的信息。

## Dependencies

- Core workbook IO: `openpyxl`, `XlsxWriter`, `Pillow`.
- Visual/chart rendering: `matplotlib`, `numpy`.
- Data cleaning/statistics/source integration: `pandas`.
- Guided ML helpers: `scikit-learn`.
- If a runtime misses a required package, report the missing package and ask to install from `requirements.txt`; do not silently skip core Excel, visual, pandas, or ML checks.

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
- Machine learning: use `recommend_ml_task()` first. Then use `auto_classify()` for labels/categories, `auto_regress()` for continuous numbers, `feature_importance()` for feature ranking, `model_report()` for formatted evaluation.

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
- `recommend_ml_task()`
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

## ML Guide

- Use classification when target is a category: win/loss, churn/not churn, risk level, product segment, approval result.
- Use regression when target is a continuous number: revenue, cost, delivery days, quantity, score, forecast amount.
- Use `recommend_ml_task(rows, target_column=...)` before training. If it says `minimum_rows_ok=False`, warn user that model evidence is weak.
- Keep ML output as workbook assistance, not truth. Show feature importance and evaluation metrics; do not invent business conclusions beyond source data.

## Advanced Excel Boundaries

- Merged cells: preserve existing merged ranges when safe; avoid translating non-anchor cells inside merged ranges; repair widths/heights around merged titles.
- Pivot tables: this skill can create pivot-like analysis sheets via formulas/tables. It does not guarantee native Excel PivotTable cache regeneration after structural edits.
- VBA/macros: preserve `.xlsm` package parts only when using package-safe copying. Do not rewrite macro modules. Do not claim macro logic was audited unless explicitly inspected.
- External connections/Power Query: preserve package parts when possible, but do not refresh or validate external data connections unless Excel/LibreOffice automation succeeds.
- Protected sheets/workbooks: do not bypass protection. Ask user for password or operate only on accessible content.
- Complex chart objects: prefer native chart source extraction. Image-only chart reconstruction is auxiliary and must mark `review_required=True`.

## pandas Integration

- Use `import_tabular_data()` for simple CSV/TSV/XLSX ingestion.
- Use pandas before workbook generation when source needs joins, reshaping, type cleanup, or group-by analysis.
- After pandas transforms, pass `DataFrame.to_dict("records")` into `generate_from_rows()` or `create_workbook_from_source()` output flows.
- Keep formulas in final workbook when feasible; do not freeze every derived value as static data.

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
Load `references/ml-guide.md` for guided classification/regression selection.
Load `references/pandas-integration.md` for pandas-to-workbook examples.
Load `references/advanced-excel-boundaries.md` before editing merged cells, pivot tables, macros, protected sheets, or external connections.
