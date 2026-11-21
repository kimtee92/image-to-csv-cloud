"""Rewrite Excel 365-only functions into formulas older Excel can evaluate.

The client workbooks (DPAM's own timekeeping templates) use LET, SWITCH and
XLOOKUP. The originals open fine everywhere because Excel saved their computed
results in the file. Our export has to drop those cached results and force a full
recalculation (the times changed), and Excel 2016/2019 cannot evaluate those
functions, so every affected cell showed #NAME? (first reported by DPAM on the
24 Chicken ND column, Nov 2026).

So before export every formula is rewritten into functions available since Excel
2010, keeping the result the same:

  LET(n1, e1, ..., body)     -> body with each name replaced by its expression
  SWITCH(x, v1, r1, ..., d)  -> nested IF (IF(x=v1, r1, ...); x=TRUE() -> IF(v1, r1, ...))
  XLOOKUP(L, A, R, nf, 0)    -> IF(ISNA(MATCH(L, INDEX(A,0), 0)), nf, INDEX(R, MATCH(...)))
  XLOOKUP(L, A, R, nf, -1/1) -> IFERROR(INDEX(R, MATCH(AGGREGATE(14|15, 6, ...), ...)), nf)

The rewrites have to work in a plain (non-array) cell, because most of these
cells are plain formulas. Legacy Excel would reduce a range inside a scalar
function, such as TIMEVALUE(LEFT(range, ...)), to a single value. INDEX(expr, 0)
and AGGREGATE's array form both evaluate their argument as an array without
needing an array formula, so every lookup array goes through one of them.

Anything this module cannot rewrite safely is left unchanged and reported, so
the preflight check in deploy.sh can catch it.
"""
from __future__ import annotations

import re

# Functions that need Excel 2021/365. AGGREGATE (2010) also carries the _xlfn.
# prefix in the file format, but every Excel from 2010 on understands it.
UNSUPPORTED = ("LET", "SWITCH", "XLOOKUP")
_LEGACY_OK_XLFN = {"AGGREGATE"}
MAX_FORMULA_LEN = 8192          # Excel's hard limit on formula text


class _Unsupported(Exception):
    pass


def _scan(f: str):
    """Yield (index, char, in_string) over a formula, tracking "..." and '...' quoting."""
    in_dq = in_sq = False
    i, n = 0, len(f)
    while i < n:
        ch = f[i]
        if in_dq:
            if ch == '"':
                if i + 1 < n and f[i + 1] == '"':
                    yield i, ch, True
                    i += 1
                    yield i, f[i], True
                    i += 1
                    continue
                in_dq = False
            yield i, ch, True
        elif in_sq:
            if ch == "'":
                if i + 1 < n and f[i + 1] == "'":
                    yield i, ch, True
                    i += 1
                    yield i, f[i], True
                    i += 1
                    continue
                in_sq = False
            yield i, ch, True
        else:
            if ch == '"':
                in_dq = True
                yield i, ch, True
            elif ch == "'":
                in_sq = True
                yield i, ch, True
            else:
                yield i, ch, False
        i += 1


def _split_call(f: str, open_idx: int) -> tuple[list[str], int]:
    """Split the arguments of the call whose '(' is at open_idx.

    Returns (args, index of the matching ')').
    """
    depth = 0
    args: list[str] = []
    start = open_idx + 1
    for i, ch, quoted in _scan(f):
        if i < open_idx:
            continue
        if quoted:
            continue
        if ch in "({":
            depth += 1
        elif ch in ")}":
            depth -= 1
            if depth == 0:
                args.append(f[start:i])
                return args, i
        elif ch == "," and depth == 1:
            args.append(f[start:i])
            start = i + 1
    raise _Unsupported("unbalanced parentheses")


def _find_call(f: str, name: str) -> int:
    """Index of the first unquoted `_xlfn.NAME(`, or -1."""
    token = f"_xlfn.{name}("
    quoted_at = {i for i, _, q in _scan(f) if q}
    pos = f.find(token)
    while pos != -1:
        if pos not in quoted_at and (pos == 0 or not (f[pos - 1].isalnum() or f[pos - 1] in "_.")):
            return pos
        pos = f.find(token, pos + 1)
    return -1


def _replace_names(expr: str, env: dict[str, str]) -> str:
    """Replace _xlpm.<name> tokens outside quotes with their (parenthesised) values."""
    if not env:
        return expr
    out: list[str] = []
    quoted = [q for _, _, q in _scan(expr)]
    pat = re.compile(r"_xlpm\.([A-Za-z_][A-Za-z0-9_.]*)")
    last = 0
    for m in pat.finditer(expr):
        if quoted[m.start()]:
            continue
        name = m.group(1)
        if name not in env:
            continue
        out.append(expr[last:m.start()])
        out.append(env[name])
        last = m.end()
    out.append(expr[last:])
    return "".join(out)


def _strip(s: str) -> str:
    return s.strip()


def _rewrite_let(args: list[str]) -> str:
    if len(args) < 3 or len(args) % 2 == 0:
        raise _Unsupported("LET needs name/value pairs and a body")
    env: dict[str, str] = {}
    for k in range(0, len(args) - 1, 2):
        name = _strip(args[k])
        if not name.startswith("_xlpm."):
            raise _Unsupported(f"LET name {name!r} is not an _xlpm. name")
        env[name[len("_xlpm."):]] = "(" + _strip(_replace_names(args[k + 1], env)) + ")"
    return "(" + _strip(_replace_names(args[-1], env)) + ")"


def _rewrite_switch(args: list[str]) -> str:
    if len(args) < 3:
        raise _Unsupported("SWITCH needs a value and at least one pair")
    subject = _strip(args[0])
    rest = [_strip(a) for a in args[1:]]
    default = rest.pop() if len(rest) % 2 == 1 else "NA()"
    pairs = list(zip(rest[0::2], rest[1::2]))
    is_true = subject.upper() in ("TRUE()", "TRUE")
    out = default
    for value, result in reversed(pairs):
        cond = value if is_true else f"({subject})=({value})"
        out = f"IF({cond}, {result}, {out})"
    return out


def _rewrite_xlookup(args: list[str]) -> str:
    if not 3 <= len(args) <= 6:
        raise _Unsupported("XLOOKUP argument count")
    a = [_strip(x) for x in args] + [""] * (6 - len(args))
    look, arr, ret, nf, mode, search = a
    nf = nf or "NA()"
    mode = mode or "0"
    if search not in ("", "1"):
        raise _Unsupported(f"XLOOKUP search_mode {search}")
    lookup_arr = f"INDEX({arr},0)"
    if mode == "0":
        pos = f"MATCH({look}, {lookup_arr}, 0)"
        return f"IF(ISNA({pos}), {nf}, INDEX({ret}, {pos}))"
    if mode in ("-1", "1"):
        # Largest value <= L (or smallest >= L). AGGREGATE ignores the #DIV/0!
        # from entries that fail the test (option 6) and evaluates the array
        # without needing an array formula. None found -> #NUM! -> not-found.
        fn, op = ("14", "<=") if mode == "-1" else ("15", ">=")
        best = f"_xlfn.AGGREGATE({fn}, 6, {lookup_arr}/({lookup_arr}{op}({look})), 1)"
        return f"IFERROR(INDEX({ret}, MATCH({best}, {lookup_arr}, 0)), {nf})"
    raise _Unsupported(f"XLOOKUP match_mode {mode}")


_REWRITERS = {"LET": _rewrite_let, "SWITCH": _rewrite_switch, "XLOOKUP": _rewrite_xlookup}


def downgrade(formula: str) -> str:
    """Return `formula` with LET/SWITCH/XLOOKUP rewritten, innermost call first.

    Raises _Unsupported (caught by downgrade_safe) when a call can't be rewritten.
    """
    f = formula
    for _ in range(500):                     # each pass removes one call
        call = _innermost(f)
        if call is None:
            return f
        idx, close, name, args = call
        f = f[:idx] + _REWRITERS[name](args) + f[close + 1:]
    raise _Unsupported("too many nested calls")


def _innermost(f: str):
    calls = []
    for name in _REWRITERS:
        start = 0
        while True:
            rel = _find_call(f[start:], name)
            if rel == -1:
                break
            idx = start + rel
            args, close = _split_call(f, idx + len(f"_xlfn.{name}"))
            calls.append((idx, close, name, args))
            start = idx + 1
    if not calls:
        return None
    # innermost = a call that contains no other call
    for c in calls:
        if not any(o is not c and c[0] < o[0] and o[1] <= c[1] for o in calls):
            return c
    return None


def downgrade_safe(formula: str) -> tuple[str, str | None]:
    """(new formula, problem). On any problem the original is returned unchanged."""
    if "_xlfn." not in formula and "_xlpm." not in formula:
        return formula, None
    try:
        out = downgrade(formula)
    except _Unsupported as exc:
        return formula, str(exc)
    if len(out) > MAX_FORMULA_LEN:
        return formula, f"rewritten formula is {len(out)} chars (Excel max {MAX_FORMULA_LEN})"
    left = unsupported_functions(out)
    if left:
        return formula, f"still uses {', '.join(sorted(left))}"
    return out, None


def unsupported_functions(formula: str) -> set[str]:
    """New-Excel functions (or LET names) a formula still uses."""
    found = {m for m in re.findall(r"_xlfn\.([A-Z0-9_.]+)\(", formula)} - _LEGACY_OK_XLFN
    if "_xlpm." in formula:
        found.add("LET names (_xlpm.)")
    return found
