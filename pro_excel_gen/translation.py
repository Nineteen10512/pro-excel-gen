from __future__ import annotations

import json
import math
import re
import zipfile
from collections import defaultdict
from pathlib import Path, PurePosixPath
from xml.etree import ElementTree as ET

from openpyxl.utils.cell import coordinate_from_string, column_index_from_string


NS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
NS_REL_DOC = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
NS_REL_PKG = "http://schemas.openxmlformats.org/package/2006/relationships"
NS_DRAW = "http://schemas.openxmlformats.org/drawingml/2006/main"
XML_SPACE = "{http://www.w3.org/XML/1998/namespace}space"


TRANSLATION_SELF_AUDIT_RULES = [
    "Translate only source-present meaning already present in the workbook.",
    "Do not add any fact, explanation, example, qualifier, summary, inferred intent, or background detail absent from source workbook.",
    "Preserve formulas, numbers, dates, units, names, identifiers, links, and placeholders unless user explicitly asked to localize them.",
    "Return one translation for each segment id and do not invent or drop ids.",
    "If source text is ambiguous, preserve ambiguity instead of filling gaps.",
]


def qn(namespace: str, tag: str) -> str:
    return f"{{{namespace}}}{tag}"


def _text_from_element(element: ET.Element | None, namespace: str = NS_MAIN) -> str:
    if element is None:
        return ""
    return "".join(node.text or "" for node in element.iterfind(f".//{qn(namespace, 't')}"))


def _set_text_in_element(element: ET.Element, text: str, namespace: str = NS_MAIN) -> None:
    for child in list(element):
        element.remove(child)
    t_el = ET.SubElement(element, qn(namespace, "t"))
    if text.startswith(" ") or text.endswith(" ") or "\n" in text:
        t_el.set(XML_SPACE, "preserve")
    t_el.text = text


def _load_parts(xlsx_path: str) -> dict[str, bytes]:
    with zipfile.ZipFile(xlsx_path) as zf:
        return {name: zf.read(name) for name in zf.namelist()}


def _natural_key(path: str) -> tuple[int, str]:
    match = re.search(r"(\d+)", Path(path).stem)
    return (int(match.group(1)) if match else 0, path)


def _parse_relationships(parts: dict[str, bytes], rels_path: str) -> dict[str, str]:
    xml = parts.get(rels_path)
    if not xml:
        return {}
    root = ET.fromstring(xml)
    mapping: dict[str, str] = {}
    for rel in root.findall(qn(NS_REL_PKG, "Relationship")):
        rid = rel.get("Id")
        target = rel.get("Target")
        if not rid or not target:
            continue
        base_dir = PurePosixPath(rels_path).parent.parent
        resolved = str((base_dir / target).as_posix())
        normalized = PurePosixPath(resolved)
        parts_list = []
        for part in normalized.parts:
            if part in {"", "."}:
                continue
            if part == "..":
                if parts_list:
                    parts_list.pop()
                continue
            parts_list.append(part)
        mapping[rid] = "/".join(parts_list)
    return mapping


def _workbook_sheet_map(parts: dict[str, bytes]) -> list[dict]:
    workbook_root = ET.fromstring(parts["xl/workbook.xml"])
    rels = _parse_relationships(parts, "xl/_rels/workbook.xml.rels")
    sheets = []
    for idx, sheet in enumerate(workbook_root.findall(f".//{qn(NS_MAIN, 'sheet')}"), start=1):
        rid = sheet.get(qn(NS_REL_DOC, "id"))
        sheets.append(
            {
                "index": idx,
                "name": sheet.get("name", f"Sheet{idx}"),
                "sheet_id": sheet.get("sheetId", str(idx)),
                "path": rels.get(rid, f"xl/worksheets/sheet{idx}.xml"),
                "rid": rid,
            }
        )
    return sheets


def _load_shared_strings(parts: dict[str, bytes]) -> list[str]:
    xml = parts.get("xl/sharedStrings.xml")
    if not xml:
        return []
    root = ET.fromstring(xml)
    values = []
    for item in root.findall(qn(NS_MAIN, "si")):
        values.append(_text_from_element(item))
    return values


def _cell_bounds(cell_ref: str) -> tuple[int, int]:
    column, row = coordinate_from_string(cell_ref)
    return row, column_index_from_string(column)


def _range_bounds(ref: str) -> tuple[int, int, int, int]:
    left, right = ref.split(":") if ":" in ref else (ref, ref)
    start_row, start_col = _cell_bounds(left)
    end_row, end_col = _cell_bounds(right)
    return start_row, start_col, end_row, end_col


def _cell_in_range(cell_ref: str, ref: str) -> bool:
    row, col = _cell_bounds(cell_ref)
    start_row, start_col, end_row, end_col = _range_bounds(ref)
    return start_row <= row <= end_row and start_col <= col <= end_col


def _sheet_tables(parts: dict[str, bytes], sheet_path: str) -> list[dict]:
    rel_name = Path(sheet_path).name + ".rels"
    rels_path = f"{Path(sheet_path).parent}/_rels/{rel_name}".replace("\\", "/")
    xml = parts.get(rels_path)
    if not xml:
        return []
    root = ET.fromstring(xml)
    table_paths = []
    for rel in root.findall(qn(NS_REL_PKG, "Relationship")):
        rel_type = rel.get("Type", "")
        if not rel_type.endswith("/table"):
            continue
        target = rel.get("Target", "")
        base_dir = PurePosixPath(sheet_path).parent
        normalized = PurePosixPath(base_dir / target)
        parts_list = []
        for part in normalized.parts:
            if part in {"", "."}:
                continue
            if part == "..":
                if parts_list:
                    parts_list.pop()
                continue
            parts_list.append(part)
        table_paths.append("/".join(parts_list))
    tables = []
    for path in table_paths:
        table_xml = parts.get(path)
        if not table_xml:
            continue
        root = ET.fromstring(table_xml)
        tables.append(
            {
                "name": root.get("displayName") or root.get("name") or Path(path).stem,
                "ref": root.get("ref", ""),
            }
        )
    return tables


def _chart_text_segments(parts: dict[str, bytes]) -> list[dict]:
    segments = []
    for path in sorted((name for name in parts if name.startswith("xl/charts/") and name.endswith(".xml")), key=_natural_key):
        root = ET.fromstring(parts[path])
        for index, node in enumerate(root.iterfind(f".//{qn(NS_DRAW, 't')}")):
            text = node.text or ""
            if not text.strip():
                continue
            segments.append(
                {
                    "id": f"chart:{Path(path).name}:text:{index}",
                    "kind": "chart_text",
                    "path": path,
                    "text": text,
                }
            )
    return segments


def _collect_segments_meta(parts: dict[str, bytes], *, skip_empty: bool = True) -> list[dict]:
    shared_strings = _load_shared_strings(parts)
    workbook_root = ET.fromstring(parts["xl/workbook.xml"])
    sheet_map = _workbook_sheet_map(parts)
    segments: list[dict] = []

    for idx, sheet in enumerate(workbook_root.findall(f".//{qn(NS_MAIN, 'sheet')}"), start=1):
        name = sheet.get("name", f"Sheet{idx}")
        if not skip_empty or name.strip():
            segments.append(
                {
                    "id": f"workbook:sheet:{idx}:name",
                    "kind": "sheet_name",
                    "path": "xl/workbook.xml",
                    "text": name,
                    "sheet_index": idx,
                }
            )

    for sheet in sheet_map:
        sheet_xml = parts.get(sheet["path"])
        if not sheet_xml:
            continue
        tables = _sheet_tables(parts, sheet["path"])
        root = ET.fromstring(sheet_xml)
        for cell in root.iterfind(f".//{qn(NS_MAIN, 'c')}"):
            cell_ref = cell.get("r")
            if not cell_ref:
                continue
            if cell.find(qn(NS_MAIN, "f")) is not None:
                continue
            cell_type = cell.get("t")
            text = None
            source = None
            if cell_type == "s":
                value = cell.findtext(qn(NS_MAIN, "v"))
                if value is None:
                    continue
                index = int(value)
                if index >= len(shared_strings):
                    continue
                text = shared_strings[index]
                source = {"type": "shared", "shared_index": index}
            elif cell_type == "inlineStr":
                text = _text_from_element(cell.find(qn(NS_MAIN, "is")))
                source = {"type": "inline"}
            elif cell_type == "str":
                text = cell.findtext(qn(NS_MAIN, "v")) or ""
                source = {"type": "string"}
            if text is None or (skip_empty and not text.strip()):
                continue
            row, col = _cell_bounds(cell_ref)
            table_name = None
            for table in tables:
                if table["ref"] and _cell_in_range(cell_ref, table["ref"]):
                    table_name = table["name"]
                    break
            segments.append(
                {
                    "id": f"sheet:{sheet['index']}:cell:{cell_ref}",
                    "kind": "cell",
                    "path": sheet["path"],
                    "sheet_name": sheet["name"],
                    "sheet_index": sheet["index"],
                    "cell": cell_ref,
                    "row": row,
                    "col": col,
                    "table": table_name,
                    "text": text,
                    "source": source,
                }
            )

    segments.extend(_chart_text_segments(parts))
    return segments


def build_translation_prompt(
    segments,
    *,
    source_lang: str | None = None,
    target_lang: str | None = None,
    bilingual: bool = False,
) -> str:
    payload = json.dumps(list(segments), ensure_ascii=False, indent=2)
    rules = "\n".join(f"{idx}. {rule}" for idx, rule in enumerate(TRANSLATION_SELF_AUDIT_RULES, start=1))
    return (
        "You are an Excel workbook translation engine.\n"
        f"Task: translate segments from {source_lang or 'source language'} to {target_lang or 'target language'}.\n"
        f"Mode: {'bilingual' if bilingual else 'replace'}.\n"
        "Hard rules:\n"
        f"{rules}\n\n"
        "Output format:\n"
        '- Return JSON only.\n'
        '- Return an array of objects: {"id": "...", "translation": "..."}.\n'
        "- Keep ids unchanged.\n"
        "- Do not emit commentary or extra keys.\n\n"
        f"Segments:\n{payload}\n"
    )


def collect_translation_segments(xlsx_path: str, *, skip_empty: bool = True) -> list[dict]:
    parts = _load_parts(xlsx_path)
    segments = _collect_segments_meta(parts, skip_empty=skip_empty)
    return [{k: v for k, v in item.items() if k not in {"source", "row", "col", "path"}} for item in segments]


def _normalize_translation_map(translations) -> dict[str, str]:
    if isinstance(translations, dict):
        return {str(key): str(value) for key, value in translations.items()}
    mapping: dict[str, str] = {}
    for item in translations:
        if not isinstance(item, dict):
            continue
        seg_id = item.get("id")
        if not seg_id:
            continue
        value = item.get("translation")
        if value is None:
            value = item.get("text")
        if value is None:
            continue
        mapping[str(seg_id)] = str(value)
    return mapping


def _estimate_growth(source_text: str, translated_text: str) -> float:
    base = max(len((source_text or "").strip()), 1)
    return len((translated_text or "").strip()) / base


def assess_translation_risk(xlsx_path: str, translations) -> dict:
    parts = _load_parts(xlsx_path)
    meta = _collect_segments_meta(parts)
    mapping = _normalize_translation_map(translations)
    sheets: dict[str, dict] = {}
    for item in meta:
        translated = mapping.get(item["id"])
        if translated is None or item["kind"] != "cell":
            continue
        growth = _estimate_growth(item["text"], translated)
        bucket = sheets.setdefault(
            item["sheet_name"],
            {"sheet_name": item["sheet_name"], "max_growth_ratio": 1.0, "cells": [], "risk": "low"},
        )
        bucket["max_growth_ratio"] = max(bucket["max_growth_ratio"], growth)
        if growth >= 2.0:
            risk = "high"
        elif growth >= 1.35:
            risk = "medium"
        else:
            risk = "low"
        if risk == "high" or (risk == "medium" and bucket["risk"] == "low"):
            bucket["risk"] = risk
        if risk != "low":
            bucket["cells"].append(
                {
                    "id": item["id"],
                    "cell": item["cell"],
                    "table": item["table"],
                    "growth_ratio": round(growth, 3),
                    "risk": risk,
                }
            )
    values = list(sheets.values())
    return {"has_risk": any(item["risk"] != "low" for item in values), "sheets": values}


def _set_cell_inline_string(cell: ET.Element, text: str) -> None:
    cell.set("t", "inlineStr")
    for child in list(cell):
        cell.remove(child)
    is_el = ET.SubElement(cell, qn(NS_MAIN, "is"))
    t_el = ET.SubElement(is_el, qn(NS_MAIN, "t"))
    if text.startswith(" ") or text.endswith(" ") or "\n" in text:
        t_el.set(XML_SPACE, "preserve")
    t_el.text = text


def _get_cols_root(sheet_root: ET.Element) -> ET.Element:
    cols = sheet_root.find(qn(NS_MAIN, "cols"))
    if cols is not None:
        return cols
    cols = ET.Element(qn(NS_MAIN, "cols"))
    children = list(sheet_root)
    insert_at = 0
    for idx, child in enumerate(children):
        if child.tag == qn(NS_MAIN, "sheetData"):
            insert_at = idx
            break
    sheet_root.insert(insert_at, cols)
    return cols


def _set_column_width(sheet_root: ET.Element, col_idx: int, width: float) -> None:
    cols = _get_cols_root(sheet_root)
    target_width = f"{width:.2f}"
    new_elements = []
    updated = False
    for col_el in list(cols):
        current_min = int(col_el.get("min", "1"))
        current_max = int(col_el.get("max", str(current_min)))
        if not (current_min <= col_idx <= current_max):
            continue
        cols.remove(col_el)
        if current_min < col_idx:
            before = ET.Element(qn(NS_MAIN, "col"), dict(col_el.attrib))
            before.set("max", str(col_idx - 1))
            new_elements.append(before)
        target = ET.Element(qn(NS_MAIN, "col"), dict(col_el.attrib))
        target.set("min", str(col_idx))
        target.set("max", str(col_idx))
        target.set("width", target_width)
        target.set("customWidth", "1")
        new_elements.append(target)
        if col_idx < current_max:
            after = ET.Element(qn(NS_MAIN, "col"), dict(col_el.attrib))
            after.set("min", str(col_idx + 1))
            new_elements.append(after)
        updated = True
        break
    if not updated:
        new_elements.append(
            ET.Element(
                qn(NS_MAIN, "col"),
                {"min": str(col_idx), "max": str(col_idx), "width": target_width, "customWidth": "1"},
            )
        )
    for item in new_elements:
        cols.append(item)


def _set_row_height(sheet_root: ET.Element, row_idx: int, height: float) -> None:
    for row_el in sheet_root.iterfind(f".//{qn(NS_MAIN, 'row')}"):
        if int(row_el.get("r", "0")) == row_idx:
            row_el.set("ht", f"{height:.2f}")
            row_el.set("customHeight", "1")
            break


def _normalize_output_path(input_path: str, output_path: str | None, target_lang: str | None) -> str:
    if output_path:
        return output_path
    src = Path(input_path)
    suffix = f".{target_lang}" if target_lang else ".translated"
    return str(src.with_name(f"{src.stem}{suffix}{src.suffix}"))


def apply_translation_map(
    xlsx_path: str,
    translations,
    *,
    output_path: str | None = None,
    target_lang: str | None = None,
    auto_format_tables: bool = True,
) -> str:
    parts = _load_parts(xlsx_path)
    meta = _collect_segments_meta(parts)
    mapping = _normalize_translation_map(translations)
    workbook_root = ET.fromstring(parts["xl/workbook.xml"])
    sheet_name_nodes = workbook_root.findall(f".//{qn(NS_MAIN, 'sheet')}")

    sheet_roots: dict[str, ET.Element] = {}
    for item in meta:
        if item["kind"] == "sheet_name":
            translated = mapping.get(item["id"])
            if translated is not None:
                sheet_name_nodes[item["sheet_index"] - 1].set("name", translated[:31])
    parts["xl/workbook.xml"] = ET.tostring(workbook_root, encoding="utf-8", xml_declaration=True)

    growth_by_sheet_col: dict[str, dict[int, float]] = defaultdict(dict)
    growth_by_sheet_row: dict[str, dict[int, float]] = defaultdict(dict)

    cell_items = [item for item in meta if item["kind"] == "cell"]
    grouped_cells: dict[str, list[dict]] = defaultdict(list)
    for item in cell_items:
        grouped_cells[item["path"]].append(item)

    for sheet_path, items in grouped_cells.items():
        root = ET.fromstring(parts[sheet_path])
        sheet_roots[sheet_path] = root
        cell_lookup = {cell.get("r"): cell for cell in root.iterfind(f".//{qn(NS_MAIN, 'c')}") if cell.get("r")}
        for item in items:
            translated = mapping.get(item["id"])
            if translated is None:
                continue
            cell = cell_lookup.get(item["cell"])
            if cell is None:
                continue
            _set_cell_inline_string(cell, translated)
            growth = _estimate_growth(item["text"], translated)
            current_col_growth = growth_by_sheet_col[sheet_path].get(item["col"], 1.0)
            current_row_growth = growth_by_sheet_row[sheet_path].get(item["row"], 1.0)
            growth_by_sheet_col[sheet_path][item["col"]] = max(current_col_growth, growth)
            growth_by_sheet_row[sheet_path][item["row"]] = max(current_row_growth, growth)

        if auto_format_tables:
            for col_idx, growth in growth_by_sheet_col[sheet_path].items():
                if growth < 1.25:
                    continue
                width = min(max(12 + math.ceil((growth - 1.0) * 10), 12), 42)
                _set_column_width(root, col_idx, width)
            for row_idx, growth in growth_by_sheet_row[sheet_path].items():
                if growth < 1.6:
                    continue
                height = min(max(15 * min(math.ceil(growth), 4), 20), 72)
                _set_row_height(root, row_idx, height)
        parts[sheet_path] = ET.tostring(root, encoding="utf-8", xml_declaration=True)

    chart_groups: dict[str, list[dict]] = defaultdict(list)
    for item in meta:
        if item["kind"] == "chart_text":
            chart_groups[item["path"]].append(item)
    for path, items in chart_groups.items():
        root = ET.fromstring(parts[path])
        nodes = [node for node in root.iterfind(f".//{qn(NS_DRAW, 't')}")]
        for item in items:
            translated = mapping.get(item["id"])
            if translated is None:
                continue
            index = int(item["id"].rsplit(":", 1)[1])
            if index < len(nodes):
                nodes[index].text = translated
        parts[path] = ET.tostring(root, encoding="utf-8", xml_declaration=True)

    resolved_output = _normalize_output_path(xlsx_path, output_path, target_lang)
    Path(resolved_output).parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(resolved_output, "w", zipfile.ZIP_DEFLATED) as zf:
        for name, data in parts.items():
            zf.writestr(name, data)
    return resolved_output
