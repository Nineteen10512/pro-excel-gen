# pandas Integration

Use pandas when source data needs reshaping, joins, type cleanup, grouped analysis, or merging multiple CSV/XLSX inputs before Excel delivery.

## DataFrame to Workbook

```python
import pandas as pd
from pro_excel_gen import generate_from_rows, finalize_delivery_workbook

df = pd.read_csv("sales.csv")
df["gross_profit"] = df["revenue"] - df["cost"]
rows = df.to_dict("records")

generate_from_rows(
    [{"title": "Sales", "rows": rows, "table": {"name": "SalesTable"}}],
    "working.xlsx",
)
finalize_delivery_workbook("working.xlsx", "final.xlsx")
```

## Workbook Rules

- Use pandas for source preparation; use Excel formulas for workbook-facing derived metrics when possible.
- Preserve typed values: dates as dates, numbers as numbers, percentages as numeric ratios.
- After pandas cleanup, run delivery verification before handing off `.xlsx`.
- For user-editable workbooks, keep raw data and formula-backed summaries in separate sheets.
