"""PRO-EXCEL v1.3.0 — 数据清洗模块 (Batch 1: data_cleaner)

合并自 excel_master 的批量清洗能力:
- 缺失值处理 (drop/fill)
- 重复行清洗
- 异常值检测 (IQR/z-score)
- 空格/格式清理
- 列名规范化
- 类型推断与转换
- 一站式 workbook 清洗入口
"""

from __future__ import annotations

import re
from collections import Counter
from copy import deepcopy
from pathlib import Path
from typing import Any

import numpy as np
from openpyxl import load_workbook
from openpyxl.utils import get_column_letter

from .delivery_kernel import import_tabular_data


# ── 缺失值处理 ──────────────────────────────────────────────

def clean_missing(
    rows: list[dict],
    *,
    strategy: str = "drop",
    fill_value: Any = None,
    columns: list[str] | None = None,
) -> list[dict]:
    """缺失值处理。

    Args:
        rows: 行数据列表
        strategy: drop | fill_mean | fill_median | fill_mode | fill_value | fill_forward
        fill_value: strategy=fill_value 时的填充值
        columns: 限定处理的列，None 为全部

    Returns:
        处理后的行列表
    """
    if not rows:
        return rows

    cols = columns or list(rows[0].keys())
    result = deepcopy(rows)

    if strategy == "drop":
        return [row for row in result if not any(
            row.get(col) is None or (isinstance(row.get(col), str) and row.get(col).strip() == "")
            for col in cols
        )]

    if strategy == "drop_any":
        # 只要任一列为空就删
        return [row for row in result if not any(
            row.get(col) is None or (isinstance(row.get(col), str) and row.get(col).strip() == "")
            for col in cols
        )]

    # 计算填充值
    for col in cols:
        values = [row.get(col) for row in result if row.get(col) is not None
                  and not (isinstance(row.get(col), str) and row.get(col).strip() == "")]
        numeric_vals = [v for v in values if isinstance(v, (int, float))]

        if strategy == "fill_mean" and numeric_vals:
            fill = float(np.mean(numeric_vals))
        elif strategy == "fill_median" and numeric_vals:
            fill = float(np.median(numeric_vals))
        elif strategy == "fill_mode" and values:
            fill = Counter(values).most_common(1)[0][0]
        elif strategy == "fill_value":
            fill = fill_value
        elif strategy == "fill_forward":
            last = None
            for row in result:
                val = row.get(col)
                if val is not None and not (isinstance(val, str) and val.strip() == ""):
                    last = val
                elif last is not None:
                    row[col] = last
            continue
        else:
            continue

        for row in result:
            val = row.get(col)
            if val is None or (isinstance(val, str) and val.strip() == ""):
                row[col] = fill

    return result


# ── 重复行清洗 ──────────────────────────────────────────────

def clean_duplicates(
    rows: list[dict],
    *,
    subset: list[str] | None = None,
    keep: str = "first",
) -> list[dict]:
    """去重。

    Args:
        rows: 行数据列表
        subset: 用于判断重复的列，None 为全部列
        keep: first | last | False (全删)

    Returns:
        去重后的行列表
    """
    seen = set()
    result = []

    for row in rows:
        if subset:
            key = tuple(str(row.get(col)) for col in subset)
        else:
            key = tuple(str(v) for v in row.values())

        if key in seen:
            if keep == "last":
                # 替换之前保留的
                for i, r in enumerate(result):
                    sk = tuple(str(r.get(col)) for col in subset) if subset else tuple(str(v) for v in r.values())
                    if sk == key:
                        result[i] = row
                        break
            # keep == False 或 first: 跳过
            continue
        seen.add(key)
        result.append(row)

    return result


# ── 异常值检测 ──────────────────────────────────────────────

def clean_outliers(
    rows: list[dict],
    *,
    method: str = "iqr",
    columns: list[str] | None = None,
    threshold: float = 1.5,
    action: str = "clip",
) -> list[dict]:
    """异常值检测与处理。

    Args:
        rows: 行数据列表
        method: iqr | zscore
        columns: 限定处理的数值列，None 为自动检测
        threshold: IQR 倍数 (method=iqr) 或 z-score 阈值 (method=zscore)
        action: clip (截断) | drop (删除) | flag (标记)

    Returns:
        处理后的行列表
    """
    if not rows:
        return rows

    result = deepcopy(rows)
    cols = columns or _detect_numeric_columns(rows)

    for col in cols:
        values = [row.get(col) for row in result
                  if isinstance(row.get(col), (int, float)) and row.get(col) is not None]
        if len(values) < 4:
            continue

        arr = np.array(values, dtype=float)

        if method == "iqr":
            q1, q3 = np.percentile(arr, [25, 75])
            iqr = q3 - q1
            lower = q1 - threshold * iqr
            upper = q3 + threshold * iqr
        elif method == "zscore":
            mean, std = np.mean(arr), np.std(arr)
            lower = mean - threshold * std
            upper = mean + threshold * std
        else:
            continue

        for row in result:
            val = row.get(col)
            if not isinstance(val, (int, float)) or val is None:
                continue
            if val < lower or val > upper:
                if action == "clip":
                    row[col] = max(lower, min(val, upper))
                elif action == "flag":
                    row[f"{col}_outlier"] = True
                # drop: handled below

        if action == "drop":
            result = [row for row in result
                      if not (isinstance(row.get(col), (int, float))
                              and row.get(col) is not None
                              and (row[col] < lower or row[col] > upper))]

    return result


# ── 空格/格式清理 ───────────────────────────────────────────

def clean_whitespace(rows: list[dict]) -> list[dict]:
    """清理所有字符串列的首尾空格和多余空白。

    Returns:
        清理后的行列表
    """
    result = deepcopy(rows)
    for row in result:
        for key, val in row.items():
            if isinstance(val, str):
                cleaned = re.sub(r"\s+", " ", val).strip()
                row[key] = cleaned
    return result


# ── 列名规范化 ──────────────────────────────────────────────

def normalize_headers(rows: list[dict]) -> list[dict]:
    """列名规范化：去空格、统一命名。

    Returns:
        列名规范化后的行列表
    """
    if not rows:
        return rows

    result = deepcopy(rows)
    old_keys = list(result[0].keys())
    mapping = {}
    for key in old_keys:
        new_key = re.sub(r"\s+", "_", str(key).strip())
        new_key = re.sub(r"[^\w\u4e00-\u9fff]", "", new_key)
        if not new_key:
            new_key = f"col_{len(mapping)}"
        # 去重
        base = new_key
        counter = 1
        while new_key in mapping.values():
            new_key = f"{base}_{counter}"
            counter += 1
        mapping[key] = new_key

    new_rows = []
    for row in result:
        new_rows.append({mapping[k]: v for k, v in row.items()})
    return new_rows


# ── 类型推断与转换 ──────────────────────────────────────────

def type_infer_and_cast(rows: list[dict]) -> list[dict]:
    """自动推断每列类型并转换。

    - 全是数字 → int/float
    - 日期格式 → str (保留)
    - 布尔 → bool

    Returns:
        类型转换后的行列表
    """
    if not rows:
        return rows

    result = deepcopy(rows)
    cols = list(result[0].keys())

    for col in cols:
        values = [row.get(col) for row in result]
        non_empty = [v for v in values if v is not None and (not isinstance(v, str) or v.strip() != "")]

        if not non_empty:
            continue

        # 尝试 int
        if all(_try_int(v) is not None for v in non_empty):
            for row in result:
                if row.get(col) is not None:
                    row[col] = _try_int(row[col])
            continue

        # 尝试 float
        if all(_try_float(v) is not None for v in non_empty):
            for row in result:
                if row.get(col) is not None:
                    row[col] = _try_float(row[col])
            continue

    return result


# ── 一站式清洗入口 ──────────────────────────────────────────

def clean_workbook(
    source_path: str,
    output_path: str,
    *,
    sheet_name: str | None = None,
    clean_config: dict | None = None,
) -> dict:
    """一站式清洗已有 Excel 工作簿。

    Args:
        source_path: 源文件路径
        output_path: 输出路径
        sheet_name: 工作表名，None 为第一个
        clean_config: 清洗配置字典，支持:
            - missing: {strategy, fill_value, columns}
            - duplicates: {subset, keep}
            - outliers: {method, columns, threshold, action}
            - whitespace: bool
            - normalize_headers: bool
            - type_cast: bool

    Returns:
        {"output_path": ..., "report": {...}}
    """
    config = clean_config or {}
    rows = import_tabular_data(source_path, sheet_name=sheet_name)

    if not rows:
        raise ValueError("Source contains no rows")

    report = {"original_rows": len(rows), "steps": []}

    # 1. 空格清理
    if config.get("whitespace", True):
        rows = clean_whitespace(rows)
        report["steps"].append("whitespace")

    # 2. 列名规范化
    if config.get("normalize_headers", False):
        rows = normalize_headers(rows)
        report["steps"].append("normalize_headers")

    # 3. 缺失值
    if "missing" in config:
        mc = config["missing"]
        rows = clean_missing(rows, **mc)
        report["steps"].append(f"missing:{mc.get('strategy', 'drop')}")

    # 4. 去重
    if "duplicates" in config:
        dc = config["duplicates"]
        rows = clean_duplicates(rows, **dc)
        report["steps"].append("duplicates")

    # 5. 异常值
    if "outliers" in config:
        oc = config["outliers"]
        rows = clean_outliers(rows, **oc)
        report["steps"].append(f"outliers:{oc.get('method', 'iqr')}")

    # 6. 类型转换
    if config.get("type_cast", True):
        rows = type_infer_and_cast(rows)
        report["steps"].append("type_cast")

    report["final_rows"] = len(rows)
    report["removed_rows"] = report["original_rows"] - report["final_rows"]

    # 写入新文件
    from .excel_jsx import generate_from_rows
    from .normalizer import normalize_workbook_ui

    generate_from_rows(
        [{"name": sheet_name or "Cleaned", "headers": list(rows[0].keys()), "rows": rows}],
        output_path,
    )
    normalize_workbook_ui(output_path)

    return {"output_path": output_path, "report": report}


# ── 批量填充 ────────────────────────────────────────────────

def batch_fill(
    rows: list[dict],
    *,
    rules: list[dict],
) -> list[dict]:
    """批量填充（模拟 excel_master 的批量填充能力）。

    Args:
        rows: 行数据列表
        rules: 填充规则列表，每项为 {column, condition, fill_value}
            condition: 支持 {"eq": value} | {"contains": str} | {"is_empty": True} | {"regex": pattern}

    Returns:
        填充后的行列表
    """
    result = deepcopy(rows)
    for rule in rules:
        col = rule["column"]
        condition = rule["condition"]
        fill_val = rule["fill_value"]

        for row in result:
            current = row.get(col)

            match = False
            if "eq" in condition and current == condition["eq"]:
                match = True
            elif "contains" in condition and isinstance(current, str) and condition["contains"] in current:
                match = True
            elif "is_empty" in condition and (current is None or (isinstance(current, str) and current.strip() == "")):
                match = True
            elif "regex" in condition and isinstance(current, str) and re.search(condition["regex"], current):
                match = True

            if match:
                row[col] = fill_val

    return result


# ── 辅助函数 ────────────────────────────────────────────────

def _detect_numeric_columns(rows: list[dict]) -> list[str]:
    """自动检测数值列。"""
    if not rows:
        return []
    numeric = []
    for col in rows[0].keys():
        if any(isinstance(row.get(col), (int, float)) for row in rows):
            numeric.append(col)
    return numeric


def _try_int(value: Any) -> int | None:
    try:
        if isinstance(value, bool):
            return None
        if isinstance(value, str):
            return int(value.replace(",", "").replace(" ", ""))
        return int(value)
    except (ValueError, TypeError):
        return None


def _try_float(value: Any) -> float | None:
    try:
        if isinstance(value, bool):
            return None
        if isinstance(value, str):
            v = value.replace(",", "").replace(" ", "").replace("%", "")
            return float(v)
        return float(value)
    except (ValueError, TypeError):
        return None