"""HTML helpers for importing pandoc-converted markdown.

When pandoc converts a Word document with complex (merged-cell) tables to gfm,
those tables come out as `<table>` blocks with `<tr>` / `<td colspan=… rowspan=…>`
elements rather than markdown pipe tables. `html_table_to_grid` parses such a
block and flattens rowspan/colspan into a rectangular grid (spanned cells
repeat their text), which can then be handed to `add_three_line_table`.

Requires `lxml`, already a hard dependency of the skill.
"""
from __future__ import annotations

import lxml.html


def html_table_to_grid(html_str: str) -> list[list[str]]:
    """Parse an HTML `<table>` (any well-formed `<table>` substring works) and
    return its content as a list of rows. Each row is a list of strings, one
    per column. Whitespace inside cells is collapsed to single spaces.

    Rowspan / colspan expansion: a cell with `colspan=N` is duplicated N times
    in its row; a cell with `rowspan=M` carries its text into the same column
    in the following M−1 rows. This converts merged tables into rectangular
    grids that the Tsinghua 三线表 style can display (Tsinghua's 三线表
    convention has no vertical merges).

    Returns an empty list if the HTML contains no rows.
    """
    root = lxml.html.fromstring(html_str)
    grid: list[list[str]] = []
    rowspans: dict[int, list] = {}  # col_index -> [text, rows_remaining]
    for tr in root.xpath(".//tr"):
        row: dict[int, str] = {}
        # carry pending rowspans into this row
        for col in list(rowspans):
            text, rem = rowspans[col]
            row[col] = text
            rem -= 1
            if rem <= 0:
                del rowspans[col]
            else:
                rowspans[col][1] = rem
        c = 0
        for cell in tr.xpath("./th|./td"):
            while c in row:
                c += 1
            text = " ".join(cell.text_content().split())
            cs = int(cell.get("colspan", 1) or 1)
            rs = int(cell.get("rowspan", 1) or 1)
            for k in range(cs):
                row[c + k] = text
                if rs > 1:
                    rowspans[c + k] = [text, rs - 1]
            c += cs
        width = (max(row) + 1) if row else 0
        grid.append([row.get(i, "") for i in range(width)])
    w = max((len(r) for r in grid), default=0)
    return [r + [""] * (w - len(r)) for r in grid]
