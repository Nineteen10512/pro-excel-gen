from __future__ import annotations

from pathlib import Path

import matplotlib
import numpy as np
import xlsxwriter

from .templates import list_templates as _list_templates
from .templates import recommend_template
from .themes import get_theme, list_themes
from .formula_planner import apply_formula_plan_to_sheet_spec

matplotlib.use("Agg")
import matplotlib.pyplot as plt  # noqa: E402


def _normalize_value(value):
    if isinstance(value, dict):
        return value
    return {"value": value}


def _resolve_sheet_title(sheet: dict, index: int) -> str:
    base = (sheet.get("title") or f"Sheet{index}").strip() or f"Sheet{index}"
    return base[:31]


def _cell_display_length(value) -> int:
    if value is None:
        return 0
    if isinstance(value, float):
        return len(f"{value:.2f}")
    return len(str(value))


def _build_formats(workbook: xlsxwriter.Workbook, theme_name: str) -> dict[str, xlsxwriter.format.Format]:
    theme = get_theme(theme_name)
    return {
        "title": workbook.add_format(
            {
                "bold": True,
                "font_size": 14,
                "font_color": theme.accent,
                "bottom": 1,
                "bottom_color": theme.grid,
            }
        ),
        "header": workbook.add_format(
            {
                "bold": True,
                "bg_color": theme.header_bg,
                "font_color": theme.header_fg,
                "border": 1,
                "border_color": theme.grid,
                "align": "center",
                "valign": "vcenter",
            }
        ),
        "text": workbook.add_format({"border": 1, "border_color": theme.grid, "valign": "top"}),
        "wrap": workbook.add_format({"border": 1, "border_color": theme.grid, "text_wrap": True, "valign": "top"}),
        "number": workbook.add_format({"border": 1, "border_color": theme.grid, "num_format": "#,##0"}),
        "currency": workbook.add_format({"border": 1, "border_color": theme.grid, "num_format": '"$"#,##0'}),
        "percent": workbook.add_format({"border": 1, "border_color": theme.grid, "num_format": "0.0%"}),
        "date": workbook.add_format({"border": 1, "border_color": theme.grid, "num_format": "yyyy-mm-dd"}),
        "note": workbook.add_format({"italic": True, "font_color": theme.accent_2}),
        "good": workbook.add_format({"border": 1, "border_color": theme.grid, "font_color": theme.good}),
        "warn": workbook.add_format({"border": 1, "border_color": theme.grid, "font_color": theme.warn}),
        "bad": workbook.add_format({"border": 1, "border_color": theme.grid, "font_color": theme.bad}),
    }


def _pick_format(formats: dict, fmt_name: str | None, fallback: str = "text"):
    return formats.get(fmt_name or fallback, formats[fallback])


def _write_cell(ws, row_idx: int, col_idx: int, cell_spec: dict, formats: dict) -> None:
    if "formula" in cell_spec:
        ws.write_formula(row_idx, col_idx, cell_spec["formula"], _pick_format(formats, cell_spec.get("format"), "number"))
        return

    value = cell_spec.get("value")
    fmt = _pick_format(formats, cell_spec.get("format"), "text")
    comment = cell_spec.get("comment")
    if isinstance(value, (int, float)) and cell_spec.get("format") in {"number", "currency", "percent"}:
        ws.write_number(row_idx, col_idx, value, fmt)
    else:
        ws.write(row_idx, col_idx, value, fmt)
    if comment:
        ws.write_comment(row_idx, col_idx, comment)


def _auto_widths(columns: list[dict], rows: list[dict]) -> list[float]:
    widths = []
    for column in columns:
        header = column.get("header", "")
        width = max(len(str(header)), int(column.get("width", 0) or 0), 10)
        key = column["header"]
        for row in rows:
            cell = _normalize_value(row.get(key))
            width = max(width, _cell_display_length(cell.get("value")))
        widths.append(min(max(width + 2, 10), 36))
    return widths


def _insert_images(ws, images: list[dict]) -> None:
    for image in images or []:
        options = {}
        for key in ("x_scale", "y_scale", "x_offset", "y_offset"):
            if key in image:
                options[key] = image[key]
        ws.insert_image(image.get("cell", "A1"), image["path"], options)


def _apply_conditional_formats(ws, conditional_formats: list[dict], data_row_end: int, col_index: dict[str, int]) -> None:
    for item in conditional_formats or []:
        target = item["range"]
        if isinstance(target, str) and target in col_index:
            col = col_index[target]
            target = xlsxwriter.utility.xl_range(1, col, data_row_end, col)
        ws.conditional_format(target, item["rule"])


def _add_table(ws, header_row: int, columns: list[dict], rows: list[dict], table_spec: dict | None) -> None:
    if not table_spec or not rows:
        return
    last_row = header_row + len(rows)
    last_col = len(columns) - 1
    ws.add_table(
        header_row,
        0,
        last_row,
        last_col,
        {
            "name": table_spec.get("name", "DataTable"),
            "style": table_spec.get("style", "Table Style Medium 2"),
            "columns": [{"header": column["header"]} for column in columns],
        },
    )


def _add_chart(ws, workbook: xlsxwriter.Workbook, chart_spec: dict, columns: list[dict], rows: list[dict], formats: dict, header_row: int = 0) -> None:
    if not rows:
        return
    headers = [column["header"] for column in columns]
    col_lookup = {name: idx for idx, name in enumerate(headers)}
    categories_col = col_lookup[chart_spec["categories_column"]]
    chart = workbook.add_chart({"type": chart_spec.get("type", "column")})
    first_data_row = header_row + 1
    last_data_row = header_row + len(rows)
    category_range = [ws.name, first_data_row, categories_col, last_data_row, categories_col]
    for series_name in chart_spec["series_columns"]:
        col = col_lookup[series_name]
        chart.add_series(
            {
                "name": [ws.name, header_row, col],
                "categories": category_range,
                "values": [ws.name, first_data_row, col, last_data_row, col],
            }
        )
    chart.set_title({"name": chart_spec.get("title", "Chart")})
    chart.set_legend({"position": chart_spec.get("legend_position", "bottom")})
    chart.set_size({"width": 560, "height": 320})
    chart.set_plotarea({"border": {"color": get_theme(chart_spec.get("theme")).grid}})
    anchor = chart_spec.get("anchor", "F2")
    ws.insert_chart(anchor, chart)


def generate(
    spec: dict,
    output_path: str,
    theme: str = "corporate_formal",
    lang: str | None = None,
    auto_style: bool = False,
    formula_mode: str | None = None,
) -> str:
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)
    workbook = xlsxwriter.Workbook(output_path)
    formats = _build_formats(workbook, theme)
    default_formula_mode = formula_mode or spec.get("meta", {}).get("formula_mode", "modern")

    for index, sheet in enumerate(spec.get("sheets", []), start=1):
        title = _resolve_sheet_title(sheet, index)
        ws = workbook.add_worksheet(title)
        if sheet.get("freeze_panes"):
            ws.freeze_panes(*sheet["freeze_panes"])

        if sheet.get("title_block"):
            ws.merge_range("A1:D1", sheet["title_block"], formats["title"])
            start_row = 2
        else:
            start_row = 0

        sheet = apply_formula_plan_to_sheet_spec(sheet, formula_mode=default_formula_mode, header_row=start_row)
        columns = sheet.get("columns") or []
        rows = sheet.get("rows") or []

        for col_idx, column in enumerate(columns):
            ws.write(start_row, col_idx, column["header"], formats["header"])

        row_offset = start_row + 1
        for row_idx, row in enumerate(rows, start=row_offset):
            for col_idx, column in enumerate(columns):
                cell_spec = _normalize_value(row.get(column["header"]))
                if "format" not in cell_spec and column.get("format"):
                    cell_spec["format"] = column["format"]
                if cell_spec.get("wrap") and "format" not in cell_spec:
                    cell_spec["format"] = "wrap"
                _write_cell(ws, row_idx, col_idx, cell_spec, formats)

        widths = _auto_widths(columns, rows)
        for col_idx, width in enumerate(widths):
            ws.set_column(col_idx, col_idx, width)

        _add_table(ws, start_row, columns, rows, sheet.get("table"))
        _insert_images(ws, sheet.get("images") or [])
        col_index = {column["header"]: idx for idx, column in enumerate(columns)}
        _apply_conditional_formats(ws, sheet.get("conditional_formats"), row_offset + len(rows) - 1, col_index)
        for chart_spec in sheet.get("charts") or []:
            _add_chart(ws, workbook, chart_spec, columns, rows, formats, start_row)
        for note in sheet.get("notes") or []:
            ws.write(note.get("cell", "A1"), note["text"], formats["note"])

    workbook.close()
    return output_path


def generate_from_rows(sheets: list[dict], output_path: str, theme: str = "corporate_formal") -> str:
    spec = {"sheets": []}
    for sheet in sheets:
        rows = sheet["rows"]
        headers = list(rows[0].keys()) if rows else []
        spec["sheets"].append(
            {
                "title": sheet.get("title", "Sheet1"),
                "columns": [{"header": header} for header in headers],
                "rows": rows,
                "table": {"name": f"{sheet.get('title', 'Sheet')[:12]}Table"},
            }
        )
    return generate(spec, output_path, theme=theme)


def build_chart_sheet(title: str, categories: list[str], series: list[dict], chart_type: str = "column") -> dict:
    headers = ["Category", *[item["name"] for item in series]]
    rows = []
    for idx, category in enumerate(categories):
        row = {"Category": category}
        for item in series:
            row[item["name"]] = item["values"][idx]
        rows.append(row)
    return {
        "title": title[:31],
        "columns": [{"header": headers[0]}, *[{"header": header, "format": "number"} for header in headers[1:]]],
        "rows": rows,
        "table": {"name": f"{title[:12].replace(' ', '')}Table"},
        "charts": [
            {
                "type": chart_type,
                "title": title,
                "anchor": "F2",
                "categories_column": "Category",
                "series_columns": headers[1:],
            }
        ],
    }


def build_kpi_dashboard(title: str, metrics: list[dict], series: dict | None = None) -> dict:
    summary_rows = []
    for item in metrics:
        summary_rows.append(
            {
                "Metric": item["metric"],
                "Value": {"value": item["value"], "format": item.get("format", "number")},
                "Owner": item.get("owner", ""),
                "Status": item.get("status", ""),
            }
        )
    sheets = [
        {
            "title": "Dashboard",
            "title_block": title,
            "freeze_panes": [3, 0],
            "columns": [
                {"header": "Metric", "width": 22},
                {"header": "Value", "format": "number", "width": 16},
                {"header": "Owner", "width": 16},
                {"header": "Status", "width": 14},
            ],
            "rows": summary_rows,
            "table": {"name": "DashboardTable", "style": "Table Style Medium 9"},
            "notes": [{"cell": "A2", "text": "Executive scorecard"}],
        }
    ]
    if series:
        sheets.append(build_chart_sheet("Trend", series["months"], series["series"], chart_type=series.get("chart_type", "line")))
    return {"meta": {"title": title}, "sheets": sheets}


def render_chart_to_png(chart_spec: dict, output_path: str, theme: str = "corporate_formal") -> str:
    profile = get_theme(theme)
    categories = chart_spec["categories"]
    series = chart_spec["series"]
    chart_type = chart_spec.get("type", "column")
    Path(output_path).parent.mkdir(parents=True, exist_ok=True)

    fig, ax = plt.subplots(figsize=(8, 4.5), dpi=160)
    fig.patch.set_facecolor(profile.background)
    ax.set_facecolor(profile.background)
    colors = [profile.accent, profile.accent_2, profile.accent_3, profile.good, profile.warn]

    x = np.arange(len(categories))
    if chart_type in {"column", "bar"}:
        width = 0.75 / max(len(series), 1)
        for idx, item in enumerate(series):
            offset = (idx - (len(series) - 1) / 2) * width
            if chart_type == "column":
                ax.bar(x + offset, item["values"], width=width, label=item["name"], color=colors[idx % len(colors)])
            else:
                ax.barh(x + offset, item["values"], height=width, label=item["name"], color=colors[idx % len(colors)])
    elif chart_type == "line":
        for idx, item in enumerate(series):
            ax.plot(x, item["values"], marker="o", linewidth=2.2, label=item["name"], color=colors[idx % len(colors)])
    elif chart_type == "pie":
        first = series[0]
        ax.pie(first["values"], labels=categories, colors=colors[: len(categories)], autopct="%1.0f%%")
    else:
        raise ValueError(f"Unsupported chart type: {chart_type}")

    ax.set_title(chart_spec.get("title", "Chart"), color=profile.accent, pad=12)
    if chart_type == "bar":
        ax.set_yticks(x, categories)
    elif chart_type != "pie":
        ax.set_xticks(x, categories)
        ax.grid(axis="y", color=profile.grid, linewidth=0.8)
    if chart_type != "pie":
        ax.spines["top"].set_visible(False)
        ax.spines["right"].set_visible(False)
        ax.spines["left"].set_color(profile.grid)
        ax.spines["bottom"].set_color(profile.grid)
        ax.legend(frameon=False, loc="best")
    plt.tight_layout()
    plt.savefig(output_path, bbox_inches="tight")
    plt.close(fig)
    return output_path


def list_templates() -> list[dict]:
    return _list_templates()
