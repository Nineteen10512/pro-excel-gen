# ML Guide

Use this only when the user asks for prediction, classification, scoring, feature importance, or "machine learning" inside an Excel workflow.

## Decision Guide

- Choose `classification` when target values are labels or a small set of outcomes: win/loss, yes/no, churn segment, risk level, approval status.
- Choose `regression` when target values are continuous numbers: revenue, cost, delivery days, quantity, price, score, forecast amount.
- Call `recommend_ml_task(rows, target_column=...)` before training. Treat it as a guardrail, not a substitute for domain judgment.
- Require at least 10 usable rows for smoke-level training. For real business use, warn when rows are few, classes are imbalanced, or features are weak.

## Workflow

```python
from pro_excel_gen import recommend_ml_task, auto_classify, auto_regress, model_report

guide = recommend_ml_task(rows, target_column="outcome")
if guide["task"] == "classification":
    result = auto_classify(rows, target_column="outcome")
else:
    result = auto_regress(rows, target_column="outcome")

report = model_report(result)
```

## Guardrails

- Do not invent causal explanations. Feature importance is model evidence only.
- Keep source columns and target definition visible in the workbook.
- Report accuracy/R2/MAE/RMSE with sample count.
- If `scikit-learn` is missing, install `requirements.txt`; do not pretend ML ran.
