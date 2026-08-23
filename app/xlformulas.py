"""Small helpers for building Excel formula strings, so every schedule
builder constructs formulas the same way instead of hand-formatting strings
inline (a real source of typo bugs when the string never gets evaluated at
write time - see tests/test_formulas.py, which actually evaluates every
formula this module can produce via the `formulas` library rather than
trusting the string looks right).

sumifs_exact() deliberately uses SUMPRODUCT rather than native SUMIFS: found
via testing that the `formulas` verification library mishandles SUMIFS text
criteria that look like a leading-zero number (e.g. account code "0010")
- reproduced in isolation, `formulas` returns #VALUE! for SUMIFS but the
mathematically identical SUMPRODUCT((range=criteria)*sumrange) evaluates
correctly. SUMPRODUCT is standard Excel-2007-era syntax either way, so this
isn't a compromise made to dodge the checker - it's the safer choice with
no known downside, used everywhere instead of SUMIFS for exact-match sums.

Deliberately restricted to functions that evaluate reliably outside modern
Excel too: SUMPRODUCT/IF/IFERROR/ABS/AND/MAX/MIN need no prefix and are
supported everywhere. Never emit XLOOKUP/SORT/FILTER/UNIQUE/SEQUENCE -
they're spilling array functions with no spill metadata in an
openpyxl-written file, so only the top-left cell would ever get a value.
"""


def cell_ref(sheet: str, cell: str) -> str:
    """Quotes the sheet name only if it needs it (contains a space or a
    character that isn't a letter/digit/underscore)."""
    needs_quote = not sheet.replace("_", "").isalnum() or " " in sheet
    sheet_part = f"'{sheet}'" if needs_quote else sheet
    return f"{sheet_part}!{cell}"


def range_ref(sheet: str, col: str, first_row: int, last_row: int) -> str:
    return cell_ref(sheet, f"${col}${first_row}:${col}${last_row}")


def sumifs_exact(sum_range: str, *criteria_pairs: tuple[str, str]) -> str:
    """criteria_pairs: (criteria_range, criteria_expr) tuples, ANDed
    together, each an exact-match test. criteria_expr is inserted as-is, so
    pass a cell reference (A2) or a quoted literal (quote("Sales"))."""
    conditions = "*".join(f"({rng}={expr})" for rng, expr in criteria_pairs)
    return f"=SUMPRODUCT(({conditions})*{sum_range})"


def sum_of_values(sum_range: str, criteria_range: str, values: list[str]) -> str:
    """Sum where criteria_range matches ANY of several discrete values (e.g.
    every account code belonging to one fixed-asset category) - a single
    exact-match test can't OR across values, so this ORs a list of
    equality tests inside one SUMPRODUCT instead."""
    if not values:
        return "=0"
    ors = "+".join(f"({criteria_range}={v})" for v in values)
    return f"=SUMPRODUCT((({ors})>0)*{sum_range})"


def quote(value: str) -> str:
    return '"' + str(value).replace('"', '""') + '"'


def literal(value) -> str:
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, (int, float)):
        return repr(value)
    return quote(value)
