# Deploying the RWA Calculation webapp

The backend stays your orchestrator. This adds a thin admin layer (folders + recipe
name/id per category) and the deep formula-trace pages, all in the *Category Formula
Studio* design system. Multipage, hash-routed:
**Categories → Category workflow (Data → Select & validate → Run → Outputs → Compare →
Trace) → Trace page → Admin**.

---

## 1. Config managed folder

Create one **empty** managed folder (e.g. `rwa_webapp_config`). Add its id to the
**project variables**:

```json
{ "RWA_WEBAPP_CONFIG_FOLDER": "aB3dK9Zx" }
```

The backend writes `admin.json` + `category_overrides.json` there itself — you never
hand-author anything in that folder.

---

## 2. Python libraries

**Project → Libraries → Python** — add the contents of this folder's `python-lib/`:

```
python-lib/
  webapp_core/          (the whole package)
  excel_deep_trace.py
```

(so `import webapp_core` and `import excel_deep_trace` resolve).

> **If an older `excel_deep_trace.py` is already on the libraries, replace it with this
> one.** The Trace page reports which build is loaded — its summary line carries
> `engine_build: rwa-deep-trace 2026.09`. If you see *"An old `excel_deep_trace.py` is on
> the project libraries"* or `__init__() got an unexpected keyword argument`, the stale
> module is still being imported — replace the file and **restart the backend**.

---

## 3. Backend

Paste **`backend/backend.py`** into the webapp's Python pane. It is your full backend
(the orchestrator — folder management, `recipe.run()`, validate, compare) **plus** the RWA
webapp additions already appended at the end. Nothing else to merge.

> `backend/backend_additions.py` is just that appended block on its own, kept for
> reference / for the case where your live backend has moved on and you want to re-apply
> only the delta.

The appended block adds:

* admin auth — `POST /api/login`, `POST /api/admin/change-credentials` (seeded
  `admin` / `changeme`, hashed via `webapp_core.auth`);
* config overrides — `GET /api/admin/config`, `POST /api/admin/category/<key>`; the edits
  are stored in `category_overrides.json` and merged into `CATEGORY_CONFIG` **in memory**
  on start-up and on each save, so `get_category_config` / `get_folder` keep working
  unchanged;
* `GET /api/whoami` — app title + current user for the top bar;
* `POST /excel-trace/open-output` — open a produced workbook from a managed folder into a
  trace session (no re-upload);
* `POST /excel-trace/columns`, `POST /excel-trace/mismatch-rows`,
  `POST /excel-trace/expression-tree`;
* **redefines `compare_all_category_outputs` and `get_category_output_schema`** so the
  comparison runs on the workbooks in the **`template` folder** (the ones that carry live
  Excel formulas), reading them with `data_only=True` so it sees the cached numbers. It
  falls back to the `output` folder if the template folder has no usable workbooks. Each
  comparison row also carries `SOURCE_ROW_NUMBER` so you can jump straight from a mismatch
  to that row in Trace. If a compared column has **no cached values** (the recipe wrote
  formulas without computing them) the response includes a `warnings[]` entry.

Your existing `/excel-trace/analyze` / `explain-cell` / `row-journey` / `compare-cells` /
`findings` / `formulas` routes are untouched and also work — they need
`ExcelDeepTraceEngine`, which now resolves from `python-lib/`.

Your existing `from excel_deep_trace import ExcelDeepTraceEngine` at the top of the file
now has a real module behind it.

---

## 4. Webapp panes

| pane | file |
|---|---|
| HTML | `frontend/body.html` |
| CSS | `frontend/style.css` |
| JS | `frontend/app.js` |
| Python | `backend/backend.py` (full orchestrator + additions) |

**Settings → "This web app has a Python backend"** must be enabled.

Restart the backend, open the webapp. First admin sign-in is `admin` / `changeme` (you're
forced to change it).

---

## 5. The pages

| Route | Page |
|---|---|
| `#/` | Categories grid — status per category |
| `#/c/<key>` | Workflow stepper (see below) |
| `#/c/<key>/trace?file=<path>` | Deep formula trace of a produced workbook |
| `#/admin` / `#/admin/<key>` | Admin — folders + recipe name/id per category; Settings tab for credentials |

**Workflow steps** (all against your existing `/api/*`):

1. **Data files** — drag-drop upload, delete, download; single mapping file with replace.
2. **Select & validate** — tick the files to process → `/api/validate-inputs`.
3. **Run** — `/api/run` (triggers the category's Dataiku recipe).
4. **Outputs** — output + template lists; 🔬 on each `.xlsx` output jumps to the Trace page.
5. **Compare** — column pairs → `/api/compare-all-outputs`, run over the **template**
   folder workbooks (live formulas). Shows `Compared from: template folder` and any
   *"no cached values"* warnings. "Trace mismatched rows →" opens the template workbook.
6. **Trace** — pick a `.xlsx` → the Trace page (defaults to the **template** folder).

**Trace page** — opens the workbook, then:

* pick a sheet + row → **row journey**: every formula cell in that row as a card
  (formula, value, plain-English narrative), each with "Open full trace →";
* **Mismatch rows** panel — after a numeric comparison, scan the workbook for rows where
  the two columns disagree and jump to a row's journey;
* **Findings** panel — workbook lint; click a cell to open its trace;
* **trace drawer** — the collapsible expression tree (taken `IF` branch highlighted, every
  sub-value, `VLOOKUP` match tables), Value flow, Explain, Precedents graph, Source. Cell
  references in the tree are links → re-root the trace on that cell (breadcrumb
  backtracking).

---

## 6. Admin — what it writes

`category_overrides.json`:

```jsonc
{
  "corporate": {
    "display_name": "Corporate",
    "recipe_name": "Corporate RWA",
    "recipe_id": "compute_QvDsUspv",
    "data_folder_id": "…", "mapping_folder_id": "…",
    "output_folder_id": "…", "template_folder_id": "…"
  }
}
```

Merged over your hard-coded `CATEGORY_CONFIG` (only the keys above). Comparison rules,
calculation handlers and tolerances stay in `CATEGORY_CONFIG` — not editable from the UI,
by design.

---

## 7. Requirement — output workbooks must carry live formulas

The Trace feature reads **Excel formulas** from the produced workbook. If your
`dataframe_to_excel_bytes` writes plain values, there is nothing to trace in those files.
Have the category recipe (or a template workbook) emit the RWA formula in each computed
column — e.g. `=ROUND(A2*B2*(1-D2),2)` per row and `=VLOOKUP(...)` for the risk weight —
the way `webapp_core.compute.build_workbook` does in Category Formula Studio. Value-only
workbooks open fine but the Trace page says *"no live Excel formulas — nothing to trace"*.

The output / template folders can hold **any number of files** (multiple runs,
`*_COMPARISON_*.xlsx`, `_webapp_metadata/`). The Trace picker lists every `.xlsx` output
except the `*_COMPARISON_*` workbooks, and `open-output` opens exactly the one file you
pick — it never scans the whole folder.

---

## 8. Local preview

```bash
pip install flask openpyxl pandas numpy
python dev/serve.py      # http://127.0.0.1:5178
```

Mocks every `/api/*` endpoint with sample data and drives the Trace pages with the **real**
`ExcelDeepTraceEngine` over a synthetic RWA workbook, so the whole UI is clickable.
