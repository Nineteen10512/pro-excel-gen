"""PRO-EXCEL v1.3.0 public package surface."""

__version__ = "1.3.0"

from .excel_jsx import (
    build_chart_sheet,
    build_kpi_dashboard,
    generate,
    generate_from_rows,
    list_templates,
    list_themes,
    recommend_template,
    render_chart_to_png,
)
from .quality import inspect_workbook
from .delivery_kernel import (
    add_analysis_sheet,
    analyze_table,
    create_workbook_from_source,
    export_xlsx,
    finalize_delivery_workbook,
    import_tabular_data,
    recalculate_workbook,
    render_sheet_preview,
    verify_delivery_workbook,
)
from .workbook_controls import (
    add_dropdown_validations,
    apply_semantic_conditional_formats,
    create_delivery_audit_sheet,
    inspect_visual_layout,
)
from .quality import scan_formula_issues
from .formula_planner import (
    autofill_formulas,
    downgrade_formula_mode,
    plan_formulas,
)
from .theme_law import (
    audit_theme_contrast,
    contrast_ratio,
    enforce_theme_law,
)
from .normalizer import (
    normalize_sheet_ui,
    normalize_workbook_ui,
    upgrade_sheet_layout,
)
from .chart_bridge import (
    build_chart_from_range,
    build_chart_from_table,
    embed_microchart_column,
    export_chart_data_to_sheet,
    extract_chart_source_data,
    infer_chart_data_from_image,
    insert_chart_into_table_region,
)
from .operations import (
    apply_operations_with_audit,
    edit_existing_workbook,
    translate_and_normalize_workbook,
)
from .translation import (
    TRANSLATION_SELF_AUDIT_RULES,
    apply_translation_map,
    assess_translation_risk,
    build_translation_prompt,
    collect_translation_segments,
)
from .data_cleaner import (
    batch_fill,
    clean_duplicates,
    clean_missing,
    clean_outliers,
    clean_whitespace,
    clean_workbook,
    normalize_headers,
    type_infer_and_cast,
)
from .statistics import (
    add_statistics_sheet,
    correlation_analysis,
    descriptive_stats,
    group_aggregate,
    pivot_analysis,
    trend_analysis,
)
from .ml_tools import (
    auto_classify,
    auto_regress,
    feature_importance,
    model_report,
)

__all__ = [
    "__version__",
    "generate",
    "generate_from_rows",
    "build_kpi_dashboard",
    "build_chart_sheet",
    "render_chart_to_png",
    "list_themes",
    "list_templates",
    "recommend_template",
    "inspect_workbook",
    "scan_formula_issues",
    "import_tabular_data",
    "analyze_table",
    "create_workbook_from_source",
    "add_analysis_sheet",
    "recalculate_workbook",
    "render_sheet_preview",
    "verify_delivery_workbook",
    "export_xlsx",
    "finalize_delivery_workbook",
    "add_dropdown_validations",
    "apply_semantic_conditional_formats",
    "inspect_visual_layout",
    "create_delivery_audit_sheet",
    "plan_formulas",
    "autofill_formulas",
    "downgrade_formula_mode",
    "audit_theme_contrast",
    "contrast_ratio",
    "enforce_theme_law",
    "normalize_sheet_ui",
    "normalize_workbook_ui",
    "upgrade_sheet_layout",
    "build_chart_from_range",
    "build_chart_from_table",
    "insert_chart_into_table_region",
    "embed_microchart_column",
    "extract_chart_source_data",
    "export_chart_data_to_sheet",
    "infer_chart_data_from_image",
    "apply_operations_with_audit",
    "edit_existing_workbook",
    "translate_and_normalize_workbook",
    "collect_translation_segments",
    "build_translation_prompt",
    "assess_translation_risk",
    "apply_translation_map",
    "TRANSLATION_SELF_AUDIT_RULES",
    "clean_missing",
    "clean_duplicates",
    "clean_outliers",
    "clean_whitespace",
    "normalize_headers",
    "type_infer_and_cast",
    "batch_fill",
    "clean_workbook",
    "descriptive_stats",
    "correlation_analysis",
    "trend_analysis",
    "group_aggregate",
    "pivot_analysis",
    "add_statistics_sheet",
    "auto_classify",
    "auto_regress",
    "feature_importance",
    "model_report",
]
