"""PRO-EXCEL v1.3.1 — 机器学习模块 (Batch 3: ml_tools)

合并自 data-ml-analysis 的机器学习建模能力:
- 自动分类 (逻辑回归/随机森林)
- 自动回归 (线性回归/随机森林)
- 特征重要性
- 模型评估报告
"""

from __future__ import annotations

from collections import Counter
from math import ceil

import numpy as np
try:
    from sklearn.ensemble import RandomForestClassifier, RandomForestRegressor
    from sklearn.linear_model import LogisticRegression, LinearRegression
    from sklearn.metrics import (
        accuracy_score,
        classification_report,
        mean_absolute_error,
        mean_squared_error,
        r2_score,
    )
    from sklearn.model_selection import train_test_split
    from sklearn.preprocessing import LabelEncoder, StandardScaler
except ModuleNotFoundError as exc:  # pragma: no cover - exercised when ML deps are absent.
    RandomForestClassifier = RandomForestRegressor = None
    LogisticRegression = LinearRegression = None
    accuracy_score = classification_report = None
    mean_absolute_error = mean_squared_error = r2_score = None
    train_test_split = None
    LabelEncoder = StandardScaler = None
    _SKLEARN_IMPORT_ERROR: ModuleNotFoundError | None = exc
else:
    _SKLEARN_IMPORT_ERROR = None

def _require_sklearn() -> None:
    if _SKLEARN_IMPORT_ERROR is not None:
        raise ImportError(
            "scikit-learn is required for Pro Excel Gen ML helpers. "
            "Install dependencies with: pip install -r requirements.txt"
        ) from _SKLEARN_IMPORT_ERROR


def recommend_ml_task(rows: list[dict], *, target_column: str) -> dict:
    """Recommend classification or regression before training a model."""
    if not rows:
        return {"task": "unknown", "reason": "No rows supplied", "minimum_rows_ok": False}

    values = [row.get(target_column) for row in rows if row.get(target_column) not in (None, "")]
    if not values:
        return {
            "task": "unknown",
            "reason": f"Target column '{target_column}' has no usable values",
            "minimum_rows_ok": False,
        }

    unique_values = set(values)
    numeric_count = 0
    for value in values:
        try:
            float(str(value).replace(",", "").replace("%", ""))
            numeric_count += 1
        except (TypeError, ValueError):
            pass

    numeric_ratio = numeric_count / len(values)
    minimum_rows_ok = len(values) >= 10

    if numeric_ratio >= 0.9 and len(unique_values) > min(20, max(5, len(values) // 5)):
        task = "regression"
        reason = "Target is mostly numeric with many distinct values; predict a continuous number."
    else:
        task = "classification"
        reason = "Target has labels or a limited set of outcomes; predict a category/class."

    return {
        "task": task,
        "reason": reason,
        "minimum_rows_ok": minimum_rows_ok,
        "target_values": len(values),
        "unique_values": len(unique_values),
    }


# ── 自动分类 ────────────────────────────────────────────────

def auto_classify(
    rows: list[dict],
    *,
    target_column: str,
    feature_columns: list[str] | None = None,
    model_type: str = "auto",
    test_size: float = 0.2,
    random_state: int = 42,
) -> dict:
    """自动分类建模。

    Args:
        rows: 行数据列表
        target_column: 目标列（分类标签）
        feature_columns: 特征列，None 为自动选择数值列
        model_type: auto | logistic | random_forest
        test_size: 测试集比例
        random_state: 随机种子

    Returns:
        {model, accuracy, report, feature_importance, predictions}
    """
    _require_sklearn()
    if not rows:
        return {"error": "No data"}

    # 提取特征和目标
    X, y, feature_names = _prepare_features(rows, target_column, feature_columns)

    if len(X) < 10:
        return {"error": "Need at least 10 rows for classification"}
    if not feature_names or X.ndim != 2 or X.shape[1] == 0:
        return {"error": "Need at least 1 usable numeric feature column"}

    # 编码目标
    le = LabelEncoder()
    y_encoded = le.fit_transform(y)

    if len(np.unique(y_encoded)) < 2:
        return {"error": "Need at least 2 classes for classification"}
    class_counts = Counter(y_encoded)
    if min(class_counts.values()) < 2:
        return {"error": "Need at least 2 rows per class for stratified classification"}

    # 标准化
    scaler = StandardScaler()
    X_scaled = scaler.fit_transform(X)

    # 拆分
    n_classes = len(class_counts)
    test_count = max(ceil(len(X_scaled) * test_size), n_classes)
    if test_count >= len(X_scaled):
        return {"error": "Not enough rows to create a stratified test split"}
    X_train, X_test, y_train, y_test = train_test_split(
        X_scaled, y_encoded, test_size=test_count, random_state=random_state, stratify=y_encoded
    )

    # 选模型
    if model_type == "logistic":
        model = LogisticRegression(max_iter=1000, random_state=random_state)
    elif model_type == "random_forest":
        model = RandomForestClassifier(n_estimators=100, random_state=random_state)
    else:
        # auto: 尝试两种
        if len(np.unique(y_encoded)) == 2:
            model = LogisticRegression(max_iter=1000, random_state=random_state)
        else:
            model = RandomForestClassifier(n_estimators=100, random_state=random_state)

    model.fit(X_train, y_train)
    y_pred = model.predict(X_test)

    accuracy = accuracy_score(y_test, y_pred)
    labels = list(range(len(le.classes_)))
    report = classification_report(
        y_test,
        y_pred,
        labels=labels,
        target_names=[str(c) for c in le.classes_],
        output_dict=True,
        zero_division=0,
    )

    # 特征重要性
    importance = _get_feature_importance(model, feature_names)

    return {
        "model_type": type(model).__name__,
        "accuracy": round(accuracy, 4),
        "report": report,
        "feature_importance": importance,
        "classes": [str(c) for c in le.classes_],
        "test_samples": len(X_test),
    }


# ── 自动回归 ────────────────────────────────────────────────

def auto_regress(
    rows: list[dict],
    *,
    target_column: str,
    feature_columns: list[str] | None = None,
    model_type: str = "auto",
    test_size: float = 0.2,
    random_state: int = 42,
) -> dict:
    """自动回归建模。

    Args:
        rows: 行数据列表
        target_column: 目标列（数值）
        feature_columns: 特征列，None 为自动选择
        model_type: auto | linear | random_forest
        test_size: 测试集比例
        random_state: 随机种子

    Returns:
        {model, r2, mae, rmse, feature_importance}
    """
    _require_sklearn()
    if not rows:
        return {"error": "No data"}

    X, y, feature_names = _prepare_features(rows, target_column, feature_columns)

    if len(X) < 10:
        return {"error": "Need at least 10 rows for regression"}
    if not feature_names or X.ndim != 2 or X.shape[1] == 0:
        return {"error": "Need at least 1 usable numeric feature column"}

    # 标准化
    scaler = StandardScaler()
    X_scaled = scaler.fit_transform(X)

    # 拆分
    test_count = max(1, ceil(len(X_scaled) * test_size))
    if test_count >= len(X_scaled):
        return {"error": "Not enough rows to create a regression test split"}
    X_train, X_test, y_train, y_test = train_test_split(
        X_scaled, y, test_size=test_count, random_state=random_state
    )

    # 选模型
    if model_type == "linear":
        model = LinearRegression()
    elif model_type == "random_forest":
        model = RandomForestRegressor(n_estimators=100, random_state=random_state)
    else:
        model = RandomForestRegressor(n_estimators=100, random_state=random_state)

    model.fit(X_train, y_train)
    y_pred = model.predict(X_test)

    r2 = r2_score(y_test, y_pred)
    mae = mean_absolute_error(y_test, y_pred)
    rmse = np.sqrt(mean_squared_error(y_test, y_pred))

    importance = _get_feature_importance(model, feature_names)

    return {
        "model_type": type(model).__name__,
        "r2": round(r2, 4),
        "mae": round(mae, 4),
        "rmse": round(rmse, 4),
        "feature_importance": importance,
        "test_samples": len(X_test),
    }


# ── 特征重要性 ──────────────────────────────────────────────

def feature_importance(
    rows: list[dict],
    *,
    target_column: str,
    feature_columns: list[str] | None = None,
    task: str = "auto",
) -> list[dict]:
    """计算特征重要性。

    Args:
        rows: 行数据列表
        target_column: 目标列
        feature_columns: 特征列
        task: classification | regression | auto

    Returns:
        [{feature, importance}, ...]
    """
    _require_sklearn()
    if not rows:
        return []

    X, y, feature_names = _prepare_features(rows, target_column, feature_columns)
    if len(X) < 2 or not feature_names or X.ndim != 2 or X.shape[1] == 0:
        return []

    if task == "auto":
        # 判断分类还是回归
        unique_ratio = len(np.unique(y)) / len(y)
        task = "classification" if unique_ratio < 0.1 and len(np.unique(y)) <= 20 else "regression"

    if task == "classification":
        le = LabelEncoder()
        y = le.fit_transform(y)
        model = RandomForestClassifier(n_estimators=100, random_state=42)
    else:
        model = RandomForestRegressor(n_estimators=100, random_state=42)

    model.fit(X, y)
    return _get_feature_importance(model, feature_names)


# ── 模型评估报告 ────────────────────────────────────────────

def model_report(result: dict) -> str:
    """将模型结果格式化为可读报告。

    Args:
        result: auto_classify 或 auto_regress 的返回值

    Returns:
        Markdown 格式的评估报告
    """
    if "error" in result:
        return f"**错误**: {result['error']}"

    lines = [f"## 模型评估报告", f"", f"**模型类型**: {result['model_type']}"]

    if "accuracy" in result:
        lines.append(f"**准确率**: {result['accuracy']:.2%}")
        lines.append(f"**测试样本**: {result['test_samples']}")
        if "classes" in result:
            lines.append(f"**类别**: {', '.join(result['classes'])}")
    elif "r2" in result:
        lines.append(f"**R²**: {result['r2']:.4f}")
        lines.append(f"**MAE**: {result['mae']:.4f}")
        lines.append(f"**RMSE**: {result['rmse']:.4f}")
        lines.append(f"**测试样本**: {result['test_samples']}")

    if result.get("feature_importance"):
        lines.append("")
        lines.append("### 特征重要性")
        for fi in result["feature_importance"][:10]:
            lines.append(f"- **{fi['feature']}**: {fi['importance']:.4f}")

    return "\n".join(lines)


# ── 辅助函数 ────────────────────────────────────────────────

def _prepare_features(
    rows: list[dict],
    target_column: str,
    feature_columns: list[str] | None = None,
) -> tuple[np.ndarray, np.ndarray, list[str]]:
    """准备特征矩阵和目标向量。"""
    from .data_cleaner import _detect_numeric_columns

    if feature_columns is None:
        feature_columns = [c for c in _detect_numeric_columns(rows) if c != target_column]

    X_list = []
    y_list = []
    for row in rows:
        # 检查目标值
        target_val = row.get(target_column)
        if target_val is None:
            continue

        feats = []
        skip = False
        for col in feature_columns:
            val = row.get(col)
            if val is None:
                skip = True
                break
            # 处理字符串特征
            if isinstance(val, str):
                try:
                    val = float(val.replace(",", "").replace("%", ""))
                except (TypeError, ValueError):
                    skip = True
                    break
            feats.append(float(val))

        if not skip:
            X_list.append(feats)
            if isinstance(target_val, str):
                y_list.append(target_val)
            else:
                y_list.append(float(target_val))

    return np.array(X_list), np.array(y_list), feature_columns


def _get_feature_importance(model, feature_names: list[str]) -> list[dict]:
    """提取特征重要性。"""
    if hasattr(model, "feature_importances_"):
        importances = model.feature_importances_
    elif hasattr(model, "coef_"):
        importances = np.abs(model.coef_).flatten() if len(model.coef_.shape) > 1 else np.abs(model.coef_)
    else:
        return []

    result = [
        {"feature": feature_names[i], "importance": round(float(importances[i]), 4)}
        for i in range(len(feature_names))
    ]
    return sorted(result, key=lambda x: x["importance"], reverse=True)
