"""Pure Python Excel formula evaluator for subjective spreadsheet functions."""

import re
from datetime import date
from typing import Any

from app.cell_store import CellStore, _split_coord


class FormulaError(Exception):
    pass


class FormulaParser:
    def __init__(self, expr: str, sheet: str, store: CellStore):
        self.expr = expr.strip()
        self.sheet = sheet
        self.store = store
        self.pos = 0

    def parse(self) -> Any:
        result = self._parse_comparison()
        self._skip_ws()
        if self.pos < len(self.expr):
            raise FormulaError(f"Unexpected: {self.expr[self.pos:]}")
        return result

    def _peek(self, n: int = 0) -> str:
        if self.pos + n < len(self.expr):
            return self.expr[self.pos + n]
        return ""

    def _advance(self, n: int = 1) -> None:
        self.pos += n

    def _skip_ws(self) -> None:
        while self.pos < len(self.expr) and self.expr[self.pos].isspace():
            self.pos += 1

    def _parse_comparison(self) -> Any:
        left = self._parse_concat()
        self._skip_ws()
        ops = [("<=", 2), (">=", 2), ("<>", 2), ("=", 1), ("<", 1), (">", 1)]
        for op, ln in ops:
            if self.expr[self.pos : self.pos + ln] == op:
                self._advance(ln)
                right = self._parse_concat()
                return self._compare(left, op, right)
        return left

    def _compare(self, left: Any, op: str, right: Any) -> bool:
        ln, rn = _to_num(left), _to_num(right)
        if op == "<=":
            return ln <= rn
        if op == ">=":
            return ln >= rn
        if op == "<":
            return ln < rn
        if op == ">":
            return ln > rn
        if op == "=":
            return str(left) == str(right)
        if op == "<>":
            return str(left) != str(right)
        return False

    def _parse_concat(self) -> Any:
        parts = [self._parse_add()]
        self._skip_ws()
        while self._peek() == "&":
            self._advance()
            parts.append(self._parse_add())
        if len(parts) == 1:
            return parts[0]
        return "".join(format_display_value(p) for p in parts)

    def _parse_add(self) -> Any:
        left = self._parse_mul()
        while True:
            self._skip_ws()
            if self._peek() in ("+", "-"):
                op = self._peek()
                self._advance()
                right = self._parse_mul()
                if op == "+":
                    left = _to_num(left) + _to_num(right)
                else:
                    left = _to_num(left) - _to_num(right)
            else:
                break
        return left

    def _parse_mul(self) -> Any:
        left = self._parse_unary()
        while True:
            self._skip_ws()
            if self._peek() == "/":
                self._advance()
                right = self._parse_unary()
                rn = _to_num(right)
                left = _to_num(left) / rn if rn != 0 else 0
            elif self._peek() == "*":
                self._advance()
                right = self._parse_unary()
                left = _to_num(left) * _to_num(right)
            else:
                break
        return left

    def _parse_unary(self) -> Any:
        self._skip_ws()
        if self._peek() == "+":
            self._advance()
            return self._parse_unary()
        if self._peek() == "-":
            self._advance()
            return -_to_num(self._parse_unary())
        return self._parse_primary()

    def _parse_primary(self) -> Any:
        self._skip_ws()
        ch = self._peek()
        if ch == '"':
            return self._parse_string()
        if ch.isdigit() or (ch == "." and self._peek(1).isdigit()):
            return self._parse_number()
        if ch == "$":
            m = re.match(r"\$?([A-Z]+)\$?(\d+)", self.expr[self.pos:])
            if m:
                self.pos += len(m.group())
                return self.store.get(self.sheet, f"{m.group(1)}{m.group(2)}")
        if ch.isalpha() or ch == "'":
            return self._parse_ref_or_func()
        raise FormulaError(f"Bad token at {self.pos}: {self.expr[self.pos:]}")

    def _parse_string(self) -> str:
        self._advance()
        start = self.pos
        while self.pos < len(self.expr):
            if self.expr[self.pos] == '"':
                s = self.expr[start:self.pos]
                self._advance()
                return s
            self.pos += 1
        raise FormulaError("Unclosed string")

    def _parse_number(self) -> float:
        m = re.match(r"\d+(?:\.\d+)?", self.expr[self.pos:])
        if not m:
            raise FormulaError("Expected number")
        self.pos += len(m.group())
        return float(m.group())

    def _parse_ref_or_func(self) -> Any:
        if self._peek() == "'":
            return self._parse_sheet_ref()
        name = self._parse_name()
        self._skip_ws()
        if self._peek() == "(":
            return self._parse_function(name)
        if self._peek() == "!":
            self._advance()
            return self._parse_cell_or_range(name)
        if re.match(r"^[A-Z]+\d", name):
            return self.store.get(self.sheet, name)
        raise FormulaError(f"Unknown identifier {name}")

    def _parse_name(self) -> str:
        m = re.match(r"[A-Za-z0-9_ ]+", self.expr[self.pos:])
        if not m:
            raise FormulaError("Expected name")
        self.pos += len(m.group())
        return m.group().strip()

    def _parse_sheet_ref(self) -> Any:
        end = self.expr.index("'", self.pos + 1)
        sheet = self.expr[self.pos + 1:end]
        self.pos = end + 1
        if self._peek() != "!":
            raise FormulaError("Expected ! after sheet")
        self._advance()
        return self._parse_cell_or_range(sheet)

    def _parse_cell_or_range(self, sheet: str) -> Any:
        self._skip_ws()
        start = self._parse_cell_ref()
        self._skip_ws()
        if self._peek() == ":":
            self._advance()
            end = self._parse_cell_ref()
            _, sr = _split_coord(start)
            _, er = _split_coord(end)
            sc, _ = _split_coord(start)
            ec, _ = _split_coord(end)
            if er > sr:
                return self.store.range_values_2d(sheet, start, end)
            return self.store.range_values(sheet, start, end)
        return self.store.get(sheet, start)

    def _parse_cell_ref(self) -> str:
        m = re.match(r"\$?([A-Z]+)\$?(\d+)", self.expr[self.pos:])
        if not m:
            raise FormulaError("Expected cell ref")
        self.pos += len(m.group())
        return f"{m.group(1)}{m.group(2)}"

    def _parse_function(self, name: str) -> Any:
        self._advance()
        args: list[Any] = []
        self._skip_ws()
        if self._peek() == ")":
            self._advance()
            return self._call_func(name.upper(), args)
        while True:
            args.append(self._parse_comparison())
            self._skip_ws()
            if self._peek() == ",":
                self._advance()
                self._skip_ws()
                continue
            if self._peek() == ")":
                self._advance()
                break
            raise FormulaError("Expected , or ) in function")
        return self._call_func(name.upper(), args)

    def _call_func(self, name: str, args: list[Any]) -> Any:
        if name == "IF":
            cond, a, b = args[0], args[1], args[2]
            return a if _is_true(cond) else b
        if name == "COUNTIF":
            return _countif(_flatten_values(args[0]), str(args[1]))
        if name == "COUNTA":
            vals = _flatten_values(args[0])
            return sum(1 for v in vals if v is not None and str(v) != "")
        if name == "MAX":
            nums = [_to_num(v) for v in _flatten_values(args[0])]
            return max(nums) if nums else 0
        if name == "AVERAGE":
            nums = [_to_num(v) for v in _flatten_values(args[0])]
            return sum(nums) / len(nums) if nums else 0
        if name == "INDEX":
            arr, row, col = args[0], int(_to_num(args[1])), int(_to_num(args[2]))
            return _index_2d(arr, row, col)
        if name == "MATCH":
            lookup, arr, _mode = args[0], _flatten_values(args[1]), args[2]
            lookup_s = str(lookup)
            for i, v in enumerate(arr):
                if str(v) == lookup_s:
                    return i + 1
            return 0
        if name == "TEXT":
            val, fmt = args[0], str(args[1]).strip('"')
            if fmt.upper() == "MMDDYY":
                d = date.today()
                return f"{d.month:02d}{d.day:02d}{str(d.year)[-2:]}"
            return str(val)
        if name == "TODAY":
            return date.today()
        raise FormulaError(f"Unsupported function {name}")


def _flatten_values(val: Any) -> list[Any]:
    out: list[Any] = []
    if isinstance(val, list):
        for item in val:
            if isinstance(item, list):
                out.extend(item)
            else:
                out.append(item)
    else:
        out.append(val)
    return out


def format_display_value(val: Any) -> str:
    if val is None:
        return ""
    if isinstance(val, float) and val == int(val):
        return str(int(val))
    if isinstance(val, date):
        return val.isoformat()
    return str(val)


def _as_list(val: Any) -> list[Any]:
    if isinstance(val, list):
        return val
    return [val]


def _to_num(val: Any) -> float:
    if val is None or val == "":
        return 0.0
    if isinstance(val, bool):
        return float(val)
    if isinstance(val, (int, float)):
        return float(val)
    try:
        return float(str(val))
    except ValueError:
        return 0.0


def _is_true(val: Any) -> bool:
    if isinstance(val, bool):
        return val
    if isinstance(val, (int, float)):
        return val != 0
    if isinstance(val, str):
        return val.lower() in ("true", "yes")
    return bool(val)


def _countif(values: list[Any], pattern: str) -> int:
    pattern = pattern.strip('"')
    if pattern == "*!*":
        return sum(1 for v in values if isinstance(v, str) and "!" in v)
    if pattern.startswith("*") and pattern.endswith("*"):
        mid = pattern[1:-1]
        return sum(1 for v in values if mid in str(v))
    return sum(1 for v in values if str(v) == pattern)


def _index_2d(arr: Any, row: int, col: int) -> Any:
    if not isinstance(arr, list):
        return arr
    flat = arr
    if flat and isinstance(flat[0], list):
        rows = flat
        if row - 1 < len(rows) and col - 1 < len(rows[row - 1]):
            return rows[row - 1][col - 1]
        return 0
    idx = (row - 1) * len(flat) + (col - 1) if col else row - 1
    if 0 <= idx < len(flat):
        return flat[idx]
    return 0


def recalculate(store: CellStore, max_passes: int = 30) -> None:
    formulas = list(store.formulas.items())
    for _ in range(max_passes):
        changed = False
        for (sheet, coord), formula in formulas:
            try:
                parser = FormulaParser(formula[1:], sheet, store)
                new_val = parser.parse()
                new_disp = format_display_value(new_val)
                old_disp = format_display_value(store.values.get((sheet, coord), ""))
                if new_disp != old_disp:
                    store.set_computed(sheet, coord, new_disp)
                    changed = True
            except Exception:
                continue
        if not changed:
            break
