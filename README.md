# logichain2026

Dashboard for the LogiChain 2025 base-case financial model.

## Repository contents

- `models/LogiChain_2025_Base_Case_Financial_Model.xlsx` — generated Excel model
- `scripts/build_base_case_model.py` — workbook generator for the base-case financial model
- `scripts/build_dashboard.py` — dashboard generator based on the model output
- `dashboard/index.html` — generated dashboard page

## Build the dashboard

```bash
python scripts/build_dashboard.py
```

This script reads the financial model workbook and generates a browser-ready dashboard in `dashboard/index.html`.
