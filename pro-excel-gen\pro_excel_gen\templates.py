from __future__ import annotations


TEMPLATES: dict[str, dict] = {
    "kpi_dashboard": {
        "name": "kpi_dashboard",
        "label": "KPI Dashboard",
        "description": "Executive KPI dashboard with scorecard and trend chart.",
        "keywords": ["dashboard", "kpi", "metric", "board", "monthly review"],
    },
    "finance_model": {
        "name": "finance_model",
        "label": "Finance Model",
        "description": "Finance and planning model with assumptions, drivers, and outputs.",
        "keywords": ["finance", "fp&a", "model", "budget", "forecast"],
    },
    "sales_tracker": {
        "name": "sales_tracker",
        "label": "Sales Tracker",
        "description": "Pipeline and revenue tracker with owner and stage fields.",
        "keywords": ["sales", "pipeline", "crm", "deal", "revenue tracker"],
    },
    "inventory_tracker": {
        "name": "inventory_tracker",
        "label": "Inventory Tracker",
        "description": "Stock ledger with quantities, reorder status, and aging support.",
        "keywords": ["inventory", "warehouse", "stock", "sku"],
    },
    "project_plan": {
        "name": "project_plan",
        "label": "Project Plan",
        "description": "Timeline and owner tracker for project execution.",
        "keywords": ["project", "timeline", "owner", "plan", "pm"],
    },
}


def list_templates() -> list[dict]:
    return [dict(item) for item in TEMPLATES.values()]


def get_template(name: str) -> dict | None:
    item = TEMPLATES.get(name)
    return dict(item) if item else None


def recommend_template(task: str) -> dict:
    lowered = (task or "").strip().lower()
    if not lowered:
        return dict(TEMPLATES["kpi_dashboard"])
    best_name = "kpi_dashboard"
    best_score = -1
    for name, item in TEMPLATES.items():
        score = 0
        for token in item["keywords"]:
            if token in lowered:
                score += len(token)
        if score > best_score:
            best_name = name
            best_score = score
    return dict(TEMPLATES[best_name])
