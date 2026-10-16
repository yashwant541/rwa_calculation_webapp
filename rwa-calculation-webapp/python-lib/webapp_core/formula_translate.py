"""Best-effort Excel-formula -> Python/pandas expression translator **plus a formula AST**.

The translator output is a *suggestion*: the Admin UI shows it in an editable textarea
next to the original formula, and the admin fixes anything the translator could not
handle. Every generated expression runs in the restricted namespace built by
``compute.evaluate`` which provides: ``df`` (canonical + already-computed columns),
``np``, ``pd``, the helpers ``IF AND OR NOT ROUND ROUNDUP ROUNDDOWN ABS INT MOD SQRT
POWER MIN MAX SUM AVERAGE COUNT CONCAT LEFT RIGHT MID LEN UPPER LOWER TRIM TEXT VALUE
YEAR MONTH DAY TODAY NOW ISBLANK ISNUMBER ISERROR IFERROR VLOOKUP CELL`` and ``PARAM``
(toggle/named-value accessor).

Formulas are expected in the *generalized* form produced by ``sample_parser`` where the
data row number has been replaced by the ``{r}`` placeholder (e.g. ``=C{r}*D{r}``).

Two entry points, one grammar:

* ``translate(formula, headers) -> (python_expr, notes)`` - the string the Admin UI edits.
* ``parse_ast(formula, headers) -> (Node, notes)`` - the same parse as a tree the
  *tracer* (``formula_trace``) walks node-by-node.  ``Node.to_python()`` renders exactly
  the string ``translate`` returns, so the two never drift (locked by
  ``tests/test_formula_ast.py``).
"""

from __future__ import annotations

import re
from typing import Dict, List, Optional, Tuple

_TOKEN_SPEC = [
    ("WS", r"\s+"),
    ("STRING", r'"(?:[^"]|"")*"'),
    ("SHEETRANGE", r"(?:'[^']+'|[A-Za-z_][A-Za-z0-9_ ]*)!\$?[A-Za-z]{1,3}\$?\d+:\$?[A-Za-z]{1,3}\$?\d+"),
    ("SHEETCELL", r"(?:'[^']+'|[A-Za-z_][A-Za-z0-9_ ]*)!\$?[A-Za-z]{1,3}\$?\d+"),
    ("TOKREF", r"\{tok:[^}]*\}"),
    ("RELCELL", r"\$?[A-Za-z]{1,3}\{r\}"),
    ("ABSCELL", r"\$?[A-Za-z]{1,3}\$?\d+"),
    ("NUMBER", r"\d+(?:\.\d+)?(?:[eE][+-]?\d+)?"),
    ("FUNC", r"[A-Za-z_][A-Za-z0-9_.]*(?=\s*\()"),
    ("IDENT", r"[A-Za-z_][A-Za-z0-9_.]*"),
    ("OP", r"<>|<=|>=|[-+*/^&=<>]"),
    ("LPAREN", r"\("),
    ("RPAREN", r"\)"),
    ("COMMA", r","),
]
_MASTER_RE = re.compile("|".join(f"(?P<{n}>{p})" for n, p in _TOKEN_SPEC))

_FUNC_MAP = {
    "IF": "IF", "IFERROR": "IFERROR", "AND": "AND", "OR": "OR", "NOT": "NOT",
    "ROUND": "ROUND", "ROUNDUP": "ROUNDUP", "ROUNDDOWN": "ROUNDDOWN",
    "ABS": "ABS", "INT": "INT", "MOD": "MOD", "SQRT": "SQRT", "POWER": "POWER",
    "MIN": "MIN", "MAX": "MAX", "SUM": "SUM", "AVERAGE": "AVERAGE", "COUNT": "COUNT",
    "CONCAT": "CONCAT", "CONCATENATE": "CONCAT",
    "LEFT": "LEFT", "RIGHT": "RIGHT", "MID": "MID", "LEN": "LEN",
    "UPPER": "UPPER", "LOWER": "LOWER", "TRIM": "TRIM", "TEXT": "TEXT", "VALUE": "VALUE",
    "YEAR": "YEAR", "MONTH": "MONTH", "DAY": "DAY", "TODAY": "TODAY", "NOW": "NOW",
    "ISBLANK": "ISBLANK", "ISNUMBER": "ISNUMBER", "ISERROR": "ISERROR",
    "COALESCE": "IFERROR",
}

# Functions the tracer treats specially (short-circuit / branch / error semantics).
LOGICAL_FUNCS = {"AND", "OR", "NOT"}


class _Tok:
    __slots__ = ("kind", "val")

    def __init__(self, kind: str, val: str):
        self.kind, self.val = kind, val

    def __repr__(self):  # pragma: no cover - debug only
        return f"{self.kind}:{self.val}"


def _tokenize(s: str) -> List[_Tok]:
    out: List[_Tok] = []
    pos = 0
    while pos < len(s):
        m = _MASTER_RE.match(s, pos)
        if not m:
            raise ValueError(f"Cannot tokenize near: {s[pos:pos + 20]!r}")
        pos = m.end()
        kind = m.lastgroup
        if kind == "WS":
            continue
        out.append(_Tok(kind, m.group()))
    return out


def _sheet_name(ref: str) -> str:
    name = ref.split("!", 1)[0]
    if name.startswith("'") and name.endswith("'"):
        name = name[1:-1]
    return name


# =========================================================================== AST
class Node:
    """Base class for formula AST nodes.

    ``to_python()`` renders the exact expression string the string-only translator
    produced.  ``kids()`` yields child nodes (for the tracer's tree walk).
    """

    def to_python(self) -> str:  # pragma: no cover - always overridden
        raise NotImplementedError

    def kids(self) -> List["Node"]:
        return []

    def walk(self):
        yield self
        for c in self.kids():
            yield from c.walk()


class Num(Node):
    def __init__(self, text: str):
        self.text = text

    def to_python(self) -> str:
        return self.text


class Bool(Node):
    def __init__(self, py: str):  # "True" / "False"
        self.py = py

    def to_python(self) -> str:
        return self.py


class Nan(Node):
    def to_python(self) -> str:
        return "np.nan"


class Str(Node):
    def __init__(self, value: str):
        self.value = value

    def to_python(self) -> str:
        return repr(self.value)


class Col(Node):
    """A data / already-computed column reference (was ``A{r}``)."""

    def __init__(self, name: str, resolved: bool = True, letter: Optional[str] = None):
        self.name, self.resolved, self.letter = name, resolved, letter

    def to_python(self) -> str:
        return f"df[{self.name!r}]"


class FixedCell(Node):
    """An absolute same-sheet cell (``$B$2``) that cannot be resolved at compute time."""

    def __init__(self, a1: str):
        self.a1 = a1

    def to_python(self) -> str:
        return "np.nan"


class SheetCell(Node):
    def __init__(self, sheet: str, a1: str):
        self.sheet, self.a1 = sheet, a1

    def to_python(self) -> str:
        return f"CELL({self.sheet!r}, {self.a1!r})"


class SheetRange(Node):
    def __init__(self, sheet: str):
        self.sheet = sheet

    def to_python(self) -> str:
        return f"LOOKUP_TABLE({self.sheet!r})"


class Param(Node):
    """A toggle / named value - ``PARAM('Name')`` (from ``{tok:Name}`` or a bare word)."""

    def __init__(self, name: str):
        self.name = name

    def to_python(self) -> str:
        return f"PARAM({self.name!r})"


class Paren(Node):
    def __init__(self, inner: Node):
        self.inner = inner

    def kids(self):
        return [self.inner]

    def to_python(self) -> str:
        return f"({self.inner.to_python()})"


class Unary(Node):
    def __init__(self, op: str, operand: Node):
        self.op, self.operand = op, operand

    def kids(self):
        return [self.operand]

    def to_python(self) -> str:
        return f"({self.op}{self.operand.to_python()})"


class Bin(Node):
    """Arithmetic: op in ``+ - * / **``."""

    def __init__(self, op: str, left: Node, right: Node):
        self.op, self.left, self.right = op, left, right

    def kids(self):
        return [self.left, self.right]

    def to_python(self) -> str:
        return f"({self.left.to_python()} {self.op} {self.right.to_python()})"


class Cmp(Node):
    """Comparison: op already in Python form (``== != < > <= >=``)."""

    def __init__(self, op: str, left: Node, right: Node):
        self.op, self.left, self.right = op, left, right

    def kids(self):
        return [self.left, self.right]

    def to_python(self) -> str:
        return f"({self.left.to_python()} {self.op} {self.right.to_python()})"


class ConcatOp(Node):
    """The Excel ``&`` operator (left-associative), rendered as nested ``CONCAT(...)``."""

    def __init__(self, left: Node, right: Node):
        self.left, self.right = left, right

    def kids(self):
        return [self.left, self.right]

    def to_python(self) -> str:
        return f"CONCAT({self.left.to_python()}, {self.right.to_python()})"


class Call(Node):
    """A translated function call - ``name`` is the helper name in the eval namespace."""

    def __init__(self, name: str, args: List[Node]):
        self.name, self.args = name, args

    def kids(self):
        return list(self.args)

    def to_python(self) -> str:
        return f"{self.name}({', '.join(a.to_python() for a in self.args)})"


class If(Node):
    """``IF(cond, then[, else])`` - the tracer evaluates only the branch that is taken."""

    def __init__(self, cond: Node, then: Node, els: Optional[Node] = None):
        self.cond, self.then, self.els = cond, then, els

    def kids(self):
        return [self.cond, self.then] + ([self.els] if self.els is not None else [])

    def to_python(self) -> str:
        parts = [self.cond.to_python(), self.then.to_python()]
        if self.els is not None:
            parts.append(self.els.to_python())
        return f"IF({', '.join(parts)})"


class Vlookup(Node):
    def __init__(self, key: Node, sheet: str, col_index: Node, approx: Node):
        self.key, self.sheet, self.col_index, self.approx = key, sheet, col_index, approx

    def kids(self):
        return [self.key, self.col_index, self.approx]

    def to_python(self) -> str:
        return (f"VLOOKUP({self.key.to_python()}, {self.sheet!r}, "
                f"{self.col_index.to_python()}, {self.approx.to_python()})")


class Manual(Node):
    """A function that needs a hand-written pandas expression (INDEX/MATCH/...)."""

    def __init__(self, fname: str):
        self.fname = fname

    def to_python(self) -> str:
        return f"MANUAL({self.fname!r})"


# =========================================================================== parser
class _Parser:
    def __init__(self, tokens: List[_Tok], header_by_letter: Dict[str, str]):
        self.toks = tokens
        self.i = 0
        self.headers = header_by_letter
        self.notes: List[str] = []

    # -- token helpers
    def _peek(self):
        return self.toks[self.i] if self.i < len(self.toks) else None

    def _next(self):
        t = self.toks[self.i]
        self.i += 1
        return t

    def _eat(self, kind):
        t = self._peek()
        if not t or t.kind != kind:
            raise ValueError(f"Expected {kind}, got {t}")
        return self._next()

    def _op_is(self, *vals):
        t = self._peek()
        return bool(t and t.kind == "OP" and t.val in vals)

    # -- grammar
    def parse(self) -> Node:
        return self._concat()

    def _concat(self) -> Node:
        left = self._compare()
        while self._op_is("&"):
            self._next()
            left = ConcatOp(left, self._compare())
        return left

    def _compare(self) -> Node:
        left = self._addsub()
        while self._op_is("=", "<>", "<", ">", "<=", ">="):
            op = self._next().val
            py = {"=": "==", "<>": "!="}.get(op, op)
            left = Cmp(py, left, self._addsub())
        return left

    def _addsub(self) -> Node:
        left = self._muldiv()
        while self._op_is("+", "-"):
            op = self._next().val
            left = Bin(op, left, self._muldiv())
        return left

    def _muldiv(self) -> Node:
        left = self._power()
        while self._op_is("*", "/"):
            op = self._next().val
            left = Bin(op, left, self._power())
        return left

    def _power(self) -> Node:
        left = self._unary()
        while self._op_is("^"):
            self._next()
            left = Bin("**", left, self._unary())
        return left

    def _unary(self) -> Node:
        t = self._peek()
        if t and t.kind == "OP" and t.val in ("-", "+"):
            self._next()
            return Unary(t.val, self._unary())
        return self._primary()

    def _primary(self) -> Node:
        t = self._peek()
        if t is None:
            raise ValueError("Unexpected end of formula")
        if t.kind == "LPAREN":
            self._next()
            inner = self._concat()
            self._eat("RPAREN")
            return Paren(inner)
        if t.kind == "NUMBER":
            return Num(self._next().val)
        if t.kind == "STRING":
            raw = self._next().val[1:-1].replace('""', '"')
            return Str(raw)
        if t.kind == "FUNC":
            return self._funccall()
        if t.kind == "TOKREF":
            raw = self._next().val
            return Param(raw[len("{tok:"): -1])
        if t.kind == "RELCELL":
            letter = self._next().val.split("{")[0].replace("$", "").upper()
            col = self.headers.get(letter)
            if col is None:
                self.notes.append(f"column letter {letter} not in the data schema")
                return Col(letter, resolved=False, letter=letter)
            return Col(col, resolved=True, letter=letter)
        if t.kind == "ABSCELL":
            self._next()
            self.notes.append(
                f"fixed cell {t.val} on the data sheet cannot be resolved at compute time"
            )
            return FixedCell(t.val.replace("$", ""))
        if t.kind == "SHEETCELL":
            self._next()
            sheet = _sheet_name(t.val)
            a1 = t.val.split("!", 1)[1].replace("$", "")
            return SheetCell(sheet, a1)
        if t.kind == "SHEETRANGE":
            self._next()
            return SheetRange(_sheet_name(t.val))
        if t.kind == "IDENT":
            name = self._next().val
            if name.upper() in ("TRUE", "FALSE"):
                return Bool(name.capitalize())
            return Param(name)
        raise ValueError(f"Unexpected token {t}")

    def _funccall(self) -> Node:
        fname = self._next().val.upper()
        self._eat("LPAREN")
        args: List[Node] = []
        if self._peek() and self._peek().kind != "RPAREN":
            while True:
                args.append(self._concat())
                if self._peek() and self._peek().kind == "COMMA":
                    self._next()
                    continue
                break
        self._eat("RPAREN")

        if fname == "VLOOKUP":
            sheet = None
            if len(args) > 1:
                for nd in args[1].walk():
                    if isinstance(nd, (SheetRange, SheetCell)):
                        sheet = nd.sheet
                        break
            key = args[0] if args else Nan()
            col_index = args[2] if len(args) > 2 else Num("2")
            approx = args[3] if len(args) > 3 else Bool("False")
            if sheet is None:
                self.notes.append("VLOOKUP table is not a Sheet!range reference")
                sheet = "UNKNOWN"
            return Vlookup(key, sheet, col_index, approx)

        if fname in ("HLOOKUP", "INDEX", "MATCH", "XLOOKUP", "LOOKUP"):
            self.notes.append(f"{fname} needs manual translation")
            return Manual(fname)

        if fname in ("TRUE", "FALSE"):
            return Bool(fname.capitalize())

        if fname == "IFS":
            # IFS(c1,v1,c2,v2,...) -> nested IF
            self.notes.append("IFS converted to nested IF")
            node: Node = Nan()
            for k in range(len(args) - 2, -1, -2):
                node = If(args[k], args[k + 1], node)
            return node

        if fname == "IF":
            if len(args) >= 3:
                return If(args[0], args[1], args[2])
            if len(args) == 2:
                return If(args[0], args[1], None)
            return Call("IF", args)  # 0/1 args: keep whatever the old path produced

        mapped = _FUNC_MAP.get(fname)
        if mapped is None:
            self.notes.append(f"unknown function {fname} kept as-is")
            mapped = fname
        return Call(mapped, args)


# =========================================================================== API
def _strip(excel_formula: str) -> str:
    s = excel_formula.strip()
    return s[1:] if s.startswith("=") else s


def parse_ast(excel_formula: str, header_by_letter: Dict[str, str]) -> Tuple[Node, List[str]]:
    """Parse ``excel_formula`` (generalized, ``{r}`` form) into an AST.

    Returns ``(root_node, notes)``.  Raises ``ValueError`` if the formula cannot be
    tokenized / parsed - callers that need the soft-failure string use ``translate``.
    """
    parser = _Parser(_tokenize(_strip(excel_formula)), header_by_letter)
    root = parser.parse()
    if parser.i != len(parser.toks):
        parser.notes.append("trailing tokens were ignored")
    return root, parser.notes


def translate(excel_formula: str, header_by_letter: Dict[str, str]) -> Tuple[str, List[str]]:
    """Return ``(python_expression, notes)``. Never raises - failures land in ``notes``."""
    s = _strip(excel_formula)
    if not s:
        return "np.nan", ["empty formula"]
    try:
        root, notes = parse_ast(excel_formula, header_by_letter)
        return root.to_python(), notes
    except Exception as exc:  # noqa: BLE001 - best effort, surface to admin
        return f"np.nan  # TODO translate: {s}", [f"could not parse: {exc}"]
