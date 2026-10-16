"""Row-level formula tracing for the **Formula Lens** webapp.

`compute.evaluate` runs each computed column as one vectorised `eval` over the whole
DataFrame - fast, but opaque.  This module does the opposite: it walks a formula AST
(`formula_translate.parse_ast`) for **one row**, recording at every node its inputs, its
value and - for `IF` / `IFS` / `AND` / `OR` / `IFERROR` - *which branch was taken*.  The
result is a JSON tree the front end renders as a collapsible "why did this cell get this
value" explorer, with click-through backtracking into upstream computed columns.

Public API
----------
* ``trace_cell(recipe, values_df, row_pos, column)``  -> the trace tree for one cell
* ``dependency_graph(recipe)``                        -> which columns/lookups feed which
* ``branch_report(recipe, values_df, column)``        -> IF-branch hit counts over all rows
* ``column_health(recipe, values_df)``                -> empty / error cell counts per column
* ``narrate(node)``                                   -> plain-English walk of the taken path

Pure module: no ``dataiku``.  Reuses the scalar-capable helpers in ``compute``.
"""

from __future__ import annotations

import math
import operator as _op
from typing import Any, Dict, List, Optional, Tuple

import numpy as np
import pandas as pd

from . import compute
from .compute import Lookups, ToggleValue
from .sample_parser import col_index_to_letter
from .formula_translate import (
    parse_ast, Node, Num, Bool, Nan, Str, Col, FixedCell, SheetCell, SheetRange,
    Param, Paren, Unary, Bin, Cmp, ConcatOp, Call, If, Vlookup, Manual,
)

_HELPERS = {
    name: getattr(compute, name)
    for name in [
        "IF", "IFERROR", "AND", "OR", "NOT", "MIN", "MAX", "SUM", "AVERAGE", "COUNT",
        "ROUND", "ROUNDUP", "ROUNDDOWN", "ABS", "INT", "MOD", "SQRT", "POWER",
        "CONCAT", "LEFT", "RIGHT", "MID", "LEN", "UPPER", "LOWER", "TRIM", "TEXT",
        "VALUE", "YEAR", "MONTH", "DAY", "TODAY", "NOW", "ISBLANK", "ISNUMBER", "ISERROR",
    ]
}

_CMP_FN = {"==": _op.eq, "!=": _op.ne, "<": _op.lt, ">": _op.gt, "<=": _op.le, ">=": _op.ge}


# --------------------------------------------------------------------------- value utils
def _is_nan(v: Any) -> bool:
    return isinstance(v, float) and math.isnan(v)


def _is_empty(v: Any) -> bool:
    return v is None or _is_nan(v) or (isinstance(v, str) and v == "")


def _is_errval(v: Any) -> bool:
    return v is None or (isinstance(v, float) and (math.isnan(v) or math.isinf(v)))


def _num(v: Any) -> Any:
    """Coerce a numeric-looking value to a number for arithmetic; leave others alone."""
    if isinstance(v, bool) or isinstance(v, (int, float)):
        return v
    if isinstance(v, str):
        s = v.strip().replace(",", "")
        try:
            return int(s)
        except ValueError:
            try:
                return float(s)
            except ValueError:
                return v
    return v


def _truthy(v: Any) -> bool:
    if isinstance(v, ToggleValue):
        return bool(v)
    if v is None:
        return False
    if _is_nan(v):
        return True  # numpy np.where treats NaN as true - stay faithful to compute
    if isinstance(v, str):
        return v.strip().lower() not in ("", "false", "no", "n", "0")
    try:
        return bool(v)
    except Exception:  # noqa: BLE001
        return True


def _json(v: Any) -> Any:
    """JSON-safe rendering of a cell value."""
    if v is None or _is_nan(v):
        return None
    if isinstance(v, (bool, int, str)):
        return v
    if isinstance(v, float):
        return None if math.isinf(v) else v
    if isinstance(v, (np.integer,)):
        return int(v)
    if isinstance(v, (np.floating,)):
        f = float(v)
        return None if (math.isnan(f) or math.isinf(f)) else f
    if isinstance(v, (np.bool_,)):
        return bool(v)
    if isinstance(v, (pd.Timestamp,)):
        return v.isoformat()
    if isinstance(v, pd.DataFrame):
        return f"<lookup table: {len(v)} rows>"
    return str(v)


def _repr(v: Any) -> str:
    if v is None or _is_nan(v):
        return "∅ empty"
    if isinstance(v, float):
        if math.isinf(v):
            return "∞" if v > 0 else "-∞"
        return repr(round(v, 10)).rstrip("0").rstrip(".") if v != int(v) else str(int(v))
    if isinstance(v, bool) or isinstance(v, np.bool_):
        return "TRUE" if v else "FALSE"
    if isinstance(v, (np.integer, int)):
        return str(int(v))
    if isinstance(v, (np.floating,)):
        return _repr(float(v))
    if isinstance(v, str):
        return f'"{v}"'
    if isinstance(v, pd.Timestamp):
        return v.date().isoformat()
    if isinstance(v, pd.DataFrame):
        return f"table[{len(v)}×{v.shape[1]}]"
    return str(v)


def _dtype(v: Any) -> str:
    if v is None or _is_nan(v):
        return "empty"
    if isinstance(v, (bool, np.bool_)):
        return "bool"
    if isinstance(v, (int, np.integer)):
        return "int"
    if isinstance(v, float):
        return "inf" if math.isinf(v) else "float"
    if isinstance(v, (np.floating,)):
        return _dtype(float(v))
    if isinstance(v, str):
        return "text"
    if isinstance(v, pd.Timestamp):
        return "date"
    if isinstance(v, pd.DataFrame):
        return "table"
    return type(v).__name__


def _values_close(a: Any, b: Any, tol: float = 1e-9) -> bool:
    if _is_empty(a) and _is_empty(b):
        return True
    if _is_empty(a) != _is_empty(b):
        return False
    an, bn = _num(a), _num(b)
    if isinstance(an, (int, float)) and isinstance(bn, (int, float)):
        try:
            return abs(float(an) - float(bn)) <= tol + 1e-6 * abs(float(bn))
        except (ValueError, OverflowError):
            return False
    return str(a).strip() == str(b).strip()


def _norm_expr(s: str) -> str:
    return "".join((s or "").split())


# --------------------------------------------------------------- Excel-ish rendering
_BIN_XL = {"+": "+", "-": "-", "*": "×", "/": "÷", "**": "^"}
_CMP_XL = {"==": "=", "!=": "≠", "<": "<", ">": ">", "<=": "≤", ">=": "≥"}


def _excel(n: Node) -> str:
    """A short, readable Excel-ish rendering of a node (uses column *names*)."""
    t = type(n).__name__
    if isinstance(n, Num):
        return n.text
    if isinstance(n, Bool):
        return n.py.upper()
    if isinstance(n, Nan):
        return "(empty)"
    if isinstance(n, Str):
        return f'"{n.value}"'
    if isinstance(n, Col):
        return n.name
    if isinstance(n, FixedCell):
        return f"${n.a1}"
    if isinstance(n, SheetCell):
        return f"{n.sheet}!{n.a1}"
    if isinstance(n, SheetRange):
        return f"{n.sheet}!"
    if isinstance(n, Param):
        return n.name
    if isinstance(n, Paren):
        return f"({_excel(n.inner)})"
    if isinstance(n, Unary):
        return f"{n.op}{_excel(n.operand)}"
    if isinstance(n, Bin):
        return f"{_excel(n.left)} {_BIN_XL.get(n.op, n.op)} {_excel(n.right)}"
    if isinstance(n, Cmp):
        return f"{_excel(n.left)} {_CMP_XL.get(n.op, n.op)} {_excel(n.right)}"
    if isinstance(n, ConcatOp):
        return f"{_excel(n.left)} & {_excel(n.right)}"
    if isinstance(n, If):
        parts = [_excel(n.cond), _excel(n.then)] + ([_excel(n.els)] if n.els is not None else [])
        return f"IF({', '.join(parts)})"
    if isinstance(n, Vlookup):
        return f"VLOOKUP({_excel(n.key)}, {n.sheet}!, {_excel(n.col_index)})"
    if isinstance(n, Manual):
        return f"{n.fname}(…)"
    if isinstance(n, Call):
        return f"{n.name}({', '.join(_excel(a) for a in n.args)})"
    return t


def _literal(n: Node) -> Any:
    """Best-effort constant value of a simple node (VLOOKUP col-index / approx flag)."""
    if isinstance(n, Num):
        try:
            return int(n.text)
        except ValueError:
            return float(n.text)
    if isinstance(n, Bool):
        return n.py == "True"
    if isinstance(n, Str):
        return n.value
    return None


# =========================================================================== tracer
class _Tracer:
    def __init__(self, recipe: Dict[str, Any], values_df: pd.DataFrame, lookups: Lookups):
        self.recipe = recipe
        self.df = values_df
        self.lookups = lookups
        self.row: Optional[pd.Series] = None
        self.notes: List[str] = []
        self._id = 0
        self.if_ids: Dict[int, int] = {}
        self.computed = {c["name"] for c in recipe.get("computed_columns", [])}
        self.toggles = {
            t["name"]: compute._toggle_value(t.get("value"))
            for t in recipe.get("toggles", [])
        }
        self.tmeta: Dict[str, Dict[str, str]] = {}
        for t in recipe.get("toggles", []):
            if t.get("sheet"):
                src = f"toggle sheet ‘{t['sheet']}’ cell {t.get('cell', '?')}"
            elif t.get("cell"):
                src = f"fixed data-sheet cell {t['cell']}"
            else:
                src = "a bare word used directly in the formula"
            self.tmeta[t["name"]] = {"value": str(t.get("value")), "source": src}

    def _nid(self) -> int:
        self._id += 1
        return self._id

    def _node(self, kind: str, n: Node, value: Any, role: str,
              children: Optional[List[dict]] = None, **detail) -> dict:
        return {
            "id": self._nid(),
            "kind": kind,
            "role": role,
            "excel": _excel(n),
            "py": n.to_python(),
            "value": _json(value),
            "value_repr": _repr(value),
            "dtype": _dtype(value),
            "evaluated": True,
            "children": children or [],
            "detail": {k: v for k, v in detail.items() if v is not None},
        }

    def _err(self, n: Node, role: str, msg: str, children=None) -> Tuple[dict, Any]:
        d = self._node("error", n, None, role, children or [], error=msg)
        return d, None

    # -- structural (un-evaluated) rendering of a branch that was not taken
    def _shallow(self, n: Node, role: str) -> dict:
        return {
            "id": self._nid(),
            "kind": "skipped",
            "role": role,
            "excel": _excel(n),
            "py": n.to_python(),
            "value": None,
            "value_repr": "not evaluated",
            "dtype": "skipped",
            "evaluated": False,
            "children": [self._shallow(c, "arg") for c in n.kids()],
            "detail": {"branch_id": self.if_ids[id(n)]} if id(n) in self.if_ids else {},
        }

    # -- main dispatch -----------------------------------------------------------
    def ev(self, n: Node, role: str = "root") -> Tuple[dict, Any]:
        if isinstance(n, Paren):
            return self.ev(n.inner, role)

        if isinstance(n, Num):
            return self._node("literal", n, _literal(n), role), _literal(n)
        if isinstance(n, Bool):
            return self._node("literal", n, n.py == "True", role), n.py == "True"
        if isinstance(n, Nan):
            return self._node("literal", n, float("nan"), role), float("nan")
        if isinstance(n, Str):
            return self._node("literal", n, n.value, role), n.value

        if isinstance(n, Col):
            return self._col(n, role)
        if isinstance(n, Param):
            return self._param(n, role)
        if isinstance(n, FixedCell):
            d = self._node("fixedcell", n, float("nan"), role,
                           note=f"fixed cell ${n.a1} is not carried in the input - treated as empty")
            return d, float("nan")
        if isinstance(n, SheetCell):
            try:
                v = self.lookups.cell(n.sheet, n.a1)
            except Exception as e:  # noqa: BLE001
                return self._err(n, role, f"cell lookup failed: {e}")
            return self._node("cell", n, v, role, sheet=n.sheet, a1=n.a1), v
        if isinstance(n, SheetRange):
            return self._node("table", n, f"<{n.sheet}>", role, sheet=n.sheet), None

        if isinstance(n, Unary):
            cd, cv = self.ev(n.operand, "operand")
            try:
                v = -_num(cv) if n.op == "-" else +_num(cv)
            except Exception as e:  # noqa: BLE001
                return self._err(n, role, f"unary {n.op}: {e}", [cd])
            return self._node("unary", n, v, role, [cd], operator=n.op), v

        if isinstance(n, Bin):
            return self._bin(n, role)
        if isinstance(n, Cmp):
            return self._cmp(n, role)
        if isinstance(n, ConcatOp):
            ld, lv = self.ev(n.left, "left")
            rd, rv = self.ev(n.right, "right")
            v = compute.CONCAT(lv, rv)
            return self._node("concat", n, v, role, [ld, rd]), v

        if isinstance(n, If):
            return self._if(n, role)
        if isinstance(n, Vlookup):
            return self._vlookup(n, role)
        if isinstance(n, Manual):
            return self._err(n, role,
                             f"{n.fname}() needs a hand-written pandas expression - "
                             "per-node trace is not available for it")
        if isinstance(n, Call):
            return self._call(n, role)

        return self._err(n, role, f"cannot trace node {type(n).__name__}")

    # -- leaves ----------------------------------------------------------------
    def _col(self, n: Col, role: str) -> Tuple[dict, Any]:
        if not n.resolved:
            return self._err(n, role, f"column letter {n.letter} is not in the schema")
        if self.row is None or n.name not in self.row.index:
            return self._err(n, role, f"column '{n.name}' is not available")
        v = self.row[n.name]
        d = self._node("col", n, v, role, column=n.name,
                       computed=(n.name in self.computed))
        return d, v

    def _param(self, n: Param, role: str) -> Tuple[dict, Any]:
        v = self.toggles.get(n.name)
        meta = self.tmeta.get(n.name, {})
        if v is None:
            v = float("nan")
            self.notes.append(f"toggle '{n.name}' has no value set")
        d = self._node("param", n, v, role, toggle=n.name,
                       source=meta.get("source", "(unknown source)"),
                       toggle_value=meta.get("value", _repr(v)))
        return d, v

    # -- operators -----------------------------------------------------------
    def _bin(self, n: Bin, role: str) -> Tuple[dict, Any]:
        ld, lv = self.ev(n.left, "left")
        rd, rv = self.ev(n.right, "right")
        a, b = _num(lv), _num(rv)
        try:
            if n.op == "+":
                v = a + b
            elif n.op == "-":
                v = a - b
            elif n.op == "*":
                v = a * b
            elif n.op == "/":
                with np.errstate(divide="ignore", invalid="ignore"):
                    v = float(np.float64(a) / np.float64(b))
            elif n.op == "**":
                v = a ** b
            else:  # pragma: no cover
                raise ValueError(n.op)
        except ZeroDivisionError:
            v = float("inf")
        except Exception as e:  # noqa: BLE001
            return self._err(n, role, f"'{n.op}' on {_repr(lv)} and {_repr(rv)}: {e}", [ld, rd])
        d = self._node("binop", n, v, role, [ld, rd], operator=n.op)
        return d, v

    def _cmp(self, n: Cmp, role: str) -> Tuple[dict, Any]:
        ld, lv = self.ev(n.left, "left")
        rd, rv = self.ev(n.right, "right")
        a, b = lv, rv
        an, bn = _num(lv), _num(rv)
        if isinstance(an, (int, float)) and isinstance(bn, (int, float)) and not isinstance(an, bool):
            a, b = an, bn
        try:
            v = bool(_CMP_FN[n.op](a, b))
        except TypeError:
            v = bool(_CMP_FN[n.op](str(a), str(b)))
        d = self._node("compare", n, v, role, [ld, rd], operator=n.op)
        return d, v

    # -- branches -----------------------------------------------------------
    def _if(self, n: If, role: str) -> Tuple[dict, Any]:
        cd, cv = self.ev(n.cond, "cond")
        taken_then = _truthy(cv)
        kids = [cd]
        if taken_then:
            td, val = self.ev(n.then, "then")
            td["taken"] = True
            kids.append(td)
            if n.els is not None:
                sd = self._shallow(n.els, "else")
                sd["taken"] = False
                kids.append(sd)
        else:
            sd = self._shallow(n.then, "then")
            sd["taken"] = False
            kids.append(sd)
            if n.els is not None:
                ed, val = self.ev(n.els, "else")
                ed["taken"] = True
                kids.append(ed)
            else:
                val = False  # Excel IF with no ELSE returns FALSE
        d = self._node("if", n, val, role, kids,
                       branch=("then" if taken_then else "else"),
                       branch_id=self.if_ids.get(id(n)),
                       cond_empty=_is_empty(cv) or None)
        return d, val

    def _call(self, n: Call, role: str) -> Tuple[dict, Any]:
        name = n.name
        if name in ("AND", "OR"):
            want = name == "OR"
            kids: List[dict] = []
            result = not want
            sc = None
            for k, a in enumerate(n.args):
                ad, av = self.ev(a, "arg")
                t = _truthy(av)
                ad["truthy"] = t
                kids.append(ad)
                if t == want:
                    result = want
                    sc = k
                    for rest in n.args[k + 1:]:
                        s = self._shallow(rest, "arg")
                        s["skipped"] = True
                        kids.append(s)
                    break
            d = self._node("logical", n, result, role, kids, operator=name,
                           short_circuit_at=sc)
            return d, result
        if name == "NOT":
            ad, av = self.ev(n.args[0], "arg") if n.args else (self._node("literal", Nan(), None, "arg"), None)
            v = not _truthy(av)
            return self._node("logical", n, v, role, [ad], operator="NOT"), v
        if name == "IFERROR":
            ad, av = self.ev(n.args[0], "primary")
            errored = ad.get("kind") == "error" or _is_errval(av)
            if errored:
                fd, fv = self.ev(n.args[1], "fallback")
                ad["taken"] = False
                fd["taken"] = True
                d = self._node("iferror", n, fv, role, [ad, fd],
                               error_caught=(ad.get("detail", {}).get("error") or _repr(av)))
                return d, fv
            fd = self._shallow(n.args[1], "fallback")
            ad["taken"] = True
            fd["taken"] = False
            return self._node("iferror", n, av, role, [ad, fd]), av

        kids = []
        vals = []
        for a in n.args:
            ad, av = self.ev(a, "arg")
            kids.append(ad)
            vals.append(av)
        fn = _HELPERS.get(name)
        if fn is None:
            return self._err(n, role, f"no helper named {name}()", kids)
        try:
            v = fn(*vals)
        except Exception as e:  # noqa: BLE001
            return self._err(n, role, f"{name}(): {type(e).__name__}: {e}", kids)
        return self._node("func", n, v, role, kids, func=name), v

    # -- vlookup ------------------------------------------------------------
    def _vlookup(self, n: Vlookup, role: str) -> Tuple[dict, Any]:
        kd, kv = self.ev(n.key, "key")
        ci = _literal(n.col_index)
        ci = int(ci) if isinstance(ci, (int, float)) else 2
        approx = bool(_truthy(_literal(n.approx)))
        detail: Dict[str, Any] = {"sheet": n.sheet, "key": _repr(kv),
                                  "col_index": ci, "approx": approx}
        try:
            frame = self.lookups.df(n.sheet)
            headers = [str(c) for c in frame.columns]
            detail["headers"] = headers
            detail["returned_column"] = headers[ci - 1] if 0 < ci <= len(headers) else None
            keys = list(frame.iloc[:, 0])
            matched = None
            for i, kk in enumerate(keys):
                if _values_close(kk, kv) or str(kk).strip().lower() == str(kv).strip().lower():
                    matched = i
                    break
            rows = []
            for i, (_, r) in enumerate(frame.iterrows()):
                if i >= 60:
                    break
                rows.append({"cells": [_json(x) for x in r.tolist()], "matched": i == matched})
            detail["rows"] = rows
            detail["matched_row"] = matched
        except Exception as e:  # noqa: BLE001
            detail["lookup_error"] = str(e)
        try:
            v = self.lookups.vlookup(kv, n.sheet, ci, approx)
        except Exception as e:  # noqa: BLE001
            return self._err(n, role, f"VLOOKUP: {e}", [kd])
        d = self._node("vlookup", n, v, role, [kd], **detail)
        return d, v


# =========================================================================== helpers
def _headers_by_letter(recipe: Dict[str, Any]) -> Dict[str, str]:
    """Map output-column letters -> names, matching how ``build_workbook`` lays the
    output out (which is the frame the live Excel formula's letters refer to)."""
    return {col_index_to_letter(i): name
            for i, name in enumerate(compute.ordered_columns(recipe))}


def _index_ifs(root: Node) -> Dict[int, int]:
    return {id(n): i for i, n in enumerate(x for x in root.walk() if isinstance(x, If))}


def _collect_kind(node: dict, kind: str, out: List[dict]) -> None:
    if node.get("kind") == kind:
        out.append(node)
    for c in node.get("children", []):
        _collect_kind(c, kind, out)


# --------------------------------------------------------------------------- narrative
def _narr(n: dict, out: List[str]) -> None:
    k = n.get("kind")
    if k == "if":
        cond = n["children"][0]
        br = n["detail"].get("branch", "?")
        empty = " (empty → treated as true)" if n["detail"].get("cond_empty") else ""
        out.append(f"Since {cond['excel']} is {cond['value_repr']}{empty}, "
                   f"the {br.upper()} branch is used.")
        for c in n["children"][1:]:
            if c.get("taken"):
                _narr(c, out)
    elif k == "iferror":
        if n["detail"].get("error_caught"):
            out.append(f"The primary expression failed ({n['detail']['error_caught']}), "
                       f"so the fallback is used.")
        for c in n["children"]:
            if c.get("taken"):
                _narr(c, out)
    elif k == "logical":
        sc = n["detail"].get("short_circuit_at")
        op = n["detail"].get("operator")
        if sc is not None:
            out.append(f"{op} short-circuited at argument {sc + 1} → {n['value_repr']}.")
        for c in n["children"]:
            _narr(c, out)
    elif k == "vlookup":
        d = n["detail"]
        if d.get("matched_row") is not None:
            out.append(f"VLOOKUP({d.get('key')}) matched a row in ‘{d.get('sheet')}’ "
                       f"and returned {n['value_repr']}.")
        else:
            out.append(f"VLOOKUP({d.get('key')}) found no match in ‘{d.get('sheet')}’ "
                       f"→ {n['value_repr']}.")
    elif k == "error":
        out.append(f"⚠ {n['detail'].get('error')}")
    else:
        for c in n.get("children", []):
            _narr(c, out)


def narrate(node: dict) -> List[str]:
    out: List[str] = []
    _narr(node, out)
    return out


# =========================================================================== public
def trace_cell(recipe: Dict[str, Any], values_df: pd.DataFrame, row_pos: int,
               column: str, lookups: Optional[Lookups] = None) -> Dict[str, Any]:
    """Trace one computed cell.  Returns a JSON-ready dict (see module docstring)."""
    lookups = lookups or Lookups(recipe.get("lookup_tables") or {})
    cc = {c["name"]: c for c in recipe.get("computed_columns", [])}
    if column not in cc:
        return {"ok": False, "error": f"'{column}' is not a computed column"}
    if row_pos < 0 or row_pos >= len(values_df):
        return {"ok": False, "error": f"row {row_pos} is out of range (0–{len(values_df) - 1})"}

    col = cc[column]
    tr = _Tracer(recipe, values_df, lookups)
    tr.row = values_df.iloc[row_pos]
    formula = col.get("excel_formula") or ""
    stored = (col.get("pandas_expr") or "").strip()
    notes: List[str] = []
    source = "excel"

    try:
        root_ast, pnotes = parse_ast(formula, _headers_by_letter(recipe))
        notes += pnotes
        tr.if_ids = _index_ifs(root_ast)
        if stored and _norm_expr(root_ast.to_python()) != _norm_expr(stored):
            notes.append("the saved pandas expression was hand-edited and no longer matches "
                         "the Excel formula - this trace follows the Excel formula")
        root, value = tr.ev(root_ast, "root")
    except Exception as exc:  # noqa: BLE001 - fall back to the whole-expression value
        source = "python"
        vec = tr.row[column] if column in tr.row.index else None
        root = {
            "id": 1, "kind": "opaque", "role": "root",
            "excel": formula or stored, "py": stored or "np.nan",
            "value": _json(vec), "value_repr": _repr(vec), "dtype": _dtype(vec),
            "evaluated": True, "children": [], "detail": {},
        }
        value = vec
        notes.append(f"this column's Excel formula could not be parsed ({exc}); showing the "
                     "computed value only. Per-node trace needs a translatable Excel formula.")

    vec_val = tr.row[column] if column in tr.row.index else None
    mismatch = source == "excel" and not _values_close(value, vec_val)
    narrative = [f"{column} = {_repr(value)}."] + narrate(root)

    return {
        "ok": True,
        "column": column,
        "row": int(row_pos),
        "source": source,
        "formula": formula,
        "pandas_expr": stored,
        "value": _json(value),
        "value_repr": _repr(value),
        "vector_value": _json(vec_val),
        "vector_value_repr": _repr(vec_val),
        "mismatch": bool(mismatch),
        "narrative": narrative,
        "notes": notes + tr.notes,
        "root": root,
        "row_data": {k: _json(v) for k, v in tr.row.items()},
        "computed_columns": [c["name"] for c in recipe.get("computed_columns", [])],
    }


def dependency_graph(recipe: Dict[str, Any]) -> Dict[str, Any]:
    """Nodes = input columns / computed columns / lookups / toggles; edges = 'X feeds Y'."""
    canonical = list(recipe.get("canonical_schema", []))
    computed = sorted(recipe.get("computed_columns", []), key=lambda c: c.get("position", 1e9))
    comp_names = [c["name"] for c in computed]
    known = set(canonical) | set(comp_names)
    hbl = _headers_by_letter(recipe)

    nodes: List[dict] = [{"id": n, "kind": "input"} for n in canonical]
    edges: List[dict] = []
    seen_edges = set()

    def edge(src: str, dst: str, via: str) -> None:
        key = (src, dst, via)
        if key not in seen_edges and src != dst:
            seen_edges.add(key)
            edges.append({"source": src, "target": dst, "via": via})

    lookup_nodes: set = set()
    toggle_nodes: set = set()

    for c in computed:
        nodes.append({"id": c["name"], "kind": "computed",
                      "position": c.get("position"),
                      "formula": c.get("excel_formula", ""),
                      "expr": c.get("pandas_expr", ""),
                      "notes": c.get("notes", [])})
        try:
            root, _ = parse_ast(c.get("excel_formula", ""), hbl)
        except Exception:  # noqa: BLE001
            continue
        for nd in root.walk():
            if isinstance(nd, Col) and nd.resolved and nd.name in known:
                edge(nd.name, c["name"], "column")
            elif isinstance(nd, (SheetRange, SheetCell)):
                lid = f"≡ {nd.sheet}"
                lookup_nodes.add(lid)
                edge(lid, c["name"], "lookup")
            elif isinstance(nd, Vlookup):
                lid = f"≡ {nd.sheet}"
                lookup_nodes.add(lid)
                edge(lid, c["name"], "lookup")
            elif isinstance(nd, Param):
                tid = f"⚙ {nd.name}"
                toggle_nodes.add(tid)
                edge(tid, c["name"], "toggle")

    for lid in sorted(lookup_nodes):
        nodes.append({"id": lid, "kind": "lookup"})
    for tid in sorted(toggle_nodes):
        nodes.append({"id": tid, "kind": "toggle"})

    # layer assignment (longest-path from a source) for a tidy left-to-right layout
    depth: Dict[str, int] = {}

    def resolve(name: str, stack: tuple = ()) -> int:
        if name in depth:
            return depth[name]
        if name in stack:
            return 0
        ins = [e["source"] for e in edges if e["target"] == name]
        d = 0 if not ins else 1 + max(resolve(s, stack + (name,)) for s in ins)
        depth[name] = d
        return d

    for nd in nodes:
        nd["layer"] = resolve(nd["id"])

    return {"nodes": nodes, "edges": edges,
            "computed_order": comp_names, "n_layers": (max(depth.values()) + 1 if depth else 1)}


def branch_report(recipe: Dict[str, Any], values_df: pd.DataFrame, column: str,
                  lookups: Optional[Lookups] = None, cap: int = 4000) -> Dict[str, Any]:
    """For each ``IF`` in ``column``'s formula, how many rows take each branch."""
    lookups = lookups or Lookups(recipe.get("lookup_tables") or {})
    cc = {c["name"]: c for c in recipe.get("computed_columns", [])}
    if column not in cc:
        return {"ok": False, "error": f"'{column}' is not a computed column"}
    n_rows = min(len(values_df), cap)
    stats: Dict[int, dict] = {}
    for i in range(n_rows):
        res = trace_cell(recipe, values_df, i, column, lookups)
        if not res.get("ok"):
            continue
        ifs: List[dict] = []
        _collect_kind(res["root"], "if", ifs)
        for ifn in ifs:
            bid = ifn["detail"].get("branch_id")
            if bid is None:
                continue
            s = stats.setdefault(bid, {"branch_id": bid, "cond": ifn["children"][0]["excel"],
                                       "then": 0, "else": 0, "then_row": None, "else_row": None,
                                       "cond_empty": 0})
            br = ifn["detail"].get("branch")
            s[br] += 1
            if ifn["detail"].get("cond_empty"):
                s["cond_empty"] += 1
            if s[f"{br}_row"] is None:
                s[f"{br}_row"] = i
    return {"ok": True, "column": column, "rows_scanned": n_rows,
            "capped": len(values_df) > cap,
            "branches": [stats[k] for k in sorted(stats)]}


def column_health(recipe: Dict[str, Any], values_df: pd.DataFrame) -> Dict[str, Any]:
    """Per computed column: how many cells are empty / non-empty, and a rollup flag."""
    out = []
    for c in sorted(recipe.get("computed_columns", []), key=lambda c: c.get("position", 1e9)):
        name = c["name"]
        if name not in values_df.columns:
            out.append({"column": name, "present": False})
            continue
        s = values_df[name]
        empty = int(s.map(_is_empty).sum())
        total = int(len(s))
        out.append({
            "column": name, "present": True, "total": total,
            "empty": empty, "filled": total - empty,
            "all_empty": empty == total and total > 0,
            "some_empty": 0 < empty < total,
        })
    return {"ok": True, "columns": out,
            "errors": list(values_df.attrs.get("compute_errors", []))}
