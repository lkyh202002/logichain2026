from __future__ import annotations

import json
from pathlib import Path

from openpyxl import load_workbook

ROOT = Path(__file__).resolve().parent.parent
MODEL_PATH = ROOT / "models" / "LogiChain_2025_Base_Case_Financial_Model.xlsx"
DASHBOARD_DIR = ROOT / "dashboard"
DATA_FILE = DASHBOARD_DIR / "dashboard-data.js"

FALLBACK = {
    "title": "LogiChain 2025 - Base-Case Financial Dashboard",
    "status": "BASE CASE",
    "source": "models/LogiChain_2025_Base_Case_Financial_Model.xlsx",
    "kpis": [
        {"label": "COGS per set", "value": 182.35, "unit": "USD/set"},
        {"label": "Monthly output", "value": 20400, "unit": "sets"},
        {"label": "Annual output", "value": 244800, "unit": "sets"},
        {"label": "Annual COGS", "value": 44691600, "unit": "USD/year"},
    ],
    "components": [
        {"label": "Front bracket", "share": 21.0, "value": 38.2},
        {"label": "Rotor hub", "share": 19.0, "value": 34.2},
        {"label": "Caliper body", "share": 15.0, "value": 27.6},
        {"label": "Brake pads", "share": 12.0, "value": 21.1},
        {"label": "Fasteners", "share": 10.0, "value": 18.2},
    ],
    "materials": [
        {"label": "Steel", "share": 29.0, "value": 52.5},
        {"label": "Aluminum", "share": 24.0, "value": 43.3},
        {"label": "Rubber", "share": 17.0, "value": 31.4},
        {"label": "Cast iron", "share": 14.0, "value": 25.9},
        {"label": "Adhesive", "share": 8.0, "value": 14.1},
    ],
    "suppliers": [
        {"label": "Supplier A", "share": 36.0, "value": 65.7},
        {"label": "Supplier B", "share": 28.0, "value": 50.2},
        {"label": "Supplier C", "share": 19.0, "value": 35.4},
        {"label": "Supplier D", "share": 10.0, "value": 18.5},
    ],
    "capacity": [
        {"line": "L0", "coverage": 1.0, "status": "OK"},
        {"line": "L1", "coverage": 0.94, "status": "Shortfall"},
        {"line": "L2", "coverage": 0.98, "status": "OK"},
    ],
    "sensitivity": {
        "labels": ["5%", "10%", "15%", "20%"],
        "values": [125000, 260000, 395000, 530000],
    },
}


def fmt_money(v):
    try:
        return float(v)
    except Exception:
        return 0.0


def cell(ws, ref):
    try:
        return ws[ref].value
    except Exception:
        return None


def try_float(value):
    if value is None:
        return 0.0
    if isinstance(value, str):
        cleaned = value.strip().replace("$", "").replace(",", "").replace("%", "")
        if not cleaned:
            return 0.0
        if cleaned.startswith("="):
            return 0.0
        try:
            return float(cleaned)
        except ValueError:
            return 0.0
    try:
        return float(value)
    except Exception:
        return 0.0


def extract_rows(ws, heading_text: str, value_col: str, label_col: str = "B", share_col: str = "F"):
    target_rows = []
    for row in range(1, ws.max_row + 1):
        val = cell(ws, f"B{row}")
        if isinstance(val, str) and heading_text.lower() in val.lower():
            start = row + 2
            for r in range(start, ws.max_row + 1):
                name = cell(ws, f"{label_col}{r}")
                if name is None:
                    continue
                if isinstance(name, str) and "Total" in name:
                    break
                value = try_float(cell(ws, f"{value_col}{r}"))
                share = try_float(cell(ws, f"{share_col}{r}"))
                if name != "" and (value or share):
                    target_rows.append({
                        "label": str(name),
                        "value": value,
                        "share": share * 100 if share <= 1 else share,
                    })
            break
    return target_rows[:6]


def extract_capacity(ws):
    rows = []
    for r in range(5, min(ws.max_row, 25)):
        line = cell(ws, f"A{r}")
        if line is None:
            continue
        if isinstance(line, str) and line.strip().startswith("Line"):
            continue
        coverage = try_float(cell(ws, f"J{r}"))
        status = "OK" if coverage >= 1 else "Shortfall"
        rows.append({"line": str(line), "coverage": coverage, "status": status})
    return rows[:6]


def build_from_workbook():
    data = FALLBACK.copy()
    if not MODEL_PATH.exists():
        return data

    try:
        wb = load_workbook(filename=MODEL_PATH, data_only=True)
    except Exception:
        return data

    summary = wb["Summary"] if "Summary" in wb.sheetnames else None
    if summary is not None:
        data["status"] = str(cell(summary, "B3") or "BASE CASE")
        data["kpis"] = [
            {"label": "COGS per set", "value": try_float(cell(summary, "C7")), "unit": "USD/set"},
            {"label": "Monthly output", "value": try_float(cell(summary, "C8")), "unit": "sets"},
            {"label": "Annual output", "value": try_float(cell(summary, "C9")), "unit": "sets"},
            {"label": "Annual COGS", "value": try_float(cell(summary, "C11")), "unit": "USD/year"},
        ]
        data["components"] = extract_rows(summary, "2. COGS Structure by Component", "E", "C")
        data["materials"] = extract_rows(summary, "3. COGS Structure by Raw Material", "E", "C")
        data["suppliers"] = extract_rows(summary, "4. COGS by Supplier", "E", "B")

    capacity = wb["Capacity Check"] if "Capacity Check" in wb.sheetnames else None
    if capacity is not None:
        data["capacity"] = extract_capacity(capacity)

    return data


def write_dashboard_data(data: dict):
    DASHBOARD_DIR.mkdir(parents=True, exist_ok=True)
    json_value = json.dumps(data, ensure_ascii=False, indent=2)
    DATA_FILE.write_text(f"window.dashboardData = {json_value};\n", encoding="utf-8")


def main():
    data = build_from_workbook()
    write_dashboard_data(data)
    print(f"Generated dashboard data at {DATA_FILE}")


if __name__ == "__main__":
    main()
