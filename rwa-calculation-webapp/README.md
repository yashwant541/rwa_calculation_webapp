# RWA Calculation webapp

Everything the **RWA Calculation** Dataiku webapp needs, in one self-contained folder.
This is **separate** from the *Category Formula Studio* / *Formula Lens* project — you can
move or rename this whole directory and nothing breaks.

```
rwa-calculation-webapp/
  frontend/
    body.html            → webapp HTML pane
    style.css             → webapp CSS pane
    app.js               → webapp JS pane
  backend/
    backend.py            → the full Python pane: your orchestrator + the additions
    backend_additions.py  → just the appended block, for reference
  python-lib/             → add this whole folder to the project's Python libraries
    webapp_core/          → shared formula engine (AST + helpers + auth)
    excel_deep_trace.py   → ExcelDeepTraceEngine (deep formula tracer)
  tests/                  → pytest suite for the engine (16 tests) — not deployed
  dev/serve.py            → local preview: mock API + real trace engine — not deployed
  INTEGRATION.md          → step-by-step deploy guide  ← start here
```

## What it is

A multipage webapp (Categories → Data → Select & validate → Run → Outputs → Compare →
**Trace** → Admin) that keeps your backend as the **orchestrator** — Admin only maps each
category to its managed folders and its recipe. The **Trace** pages open a produced
workbook and let you walk any row's calculations node by node: which `IF` branch fired,
what each sub-expression evaluated to, which `VLOOKUP` matched, and the precedents back to
the raw inputs.

## Quick start

```bash
cd rwa-calculation-webapp
pip install flask openpyxl pandas numpy pytest
python -m pytest -q          # 16 tests
python dev/serve.py          # http://127.0.0.1:5178  — clickable preview
```

Then follow **INTEGRATION.md** to deploy into Dataiku.

## Note on `python-lib/webapp_core/`

This is a **copy** of the shared formula engine so this project stands alone. It is the
same code as `webapp/python-lib/webapp_core/` in the Category Formula Studio project. If
that engine is ever changed there (`formula_translate.py` / `formula_trace.py` /
`compute.py` / `sample_parser.py`), re-copy those files here to stay in sync — or keep
them frozen; the RWA webapp only reads from the engine, never writes to it.
