"""Public API for driving the 清华综合论文训练 Word template via python-docx.

Quick start:

    from tsinghua_thesis import helpers as h
    doc = h.open_template()
    h.set_cover_info(doc, title_cn="...", author="...", department="...",
                     major="...", advisor="...", date="二○二六年六月")
    h.set_abstract(doc, cn_text="...", cn_keywords=["k1","k2"],
                   en_text="...", en_keywords=["k1","k2"])
    h.add_chapter(doc, "引言")
    h.add_body(doc, "第一段正文。")
    h.save(doc, "thesis.docx")

After saving, open the .docx in Word/WPS and press F9 to refresh 插图清单/附表清单
TOC fields.
"""
from __future__ import annotations
import os
from pathlib import Path
from typing import Literal, Optional, Sequence

from docx import Document as _DocumentCtor
from docx.document import Document
from docx.shared import Cm
from docx.table import Table
from docx.text.paragraph import Paragraph

from . import styles as S
from ._xml import (
    find_paragraph_by_text,
    iter_paragraphs_after,
    replace_paragraph_text,
    insert_paragraph_after,
    remove_paragraph,
    set_table_cell_text,
    append_omml,
)


# ---------------------------------------------------------------------------
# Template resolution
# ---------------------------------------------------------------------------

_TEMPLATE_PATH_FILE = Path(__file__).resolve().parent.parent / "template" / "TEMPLATE_PATH.txt"
_DOWNLOAD_HINT = (
    "Please obtain the official 清华综合论文训练论文模板.docx from "
    "the 清华大学教务处 download page, then configure its path by either:\n"
    "  1. passing template_path=Path(...) explicitly,\n"
    "  2. setting env var TSINGHUA_THESIS_TEMPLATE=/path/to/template.docx, or\n"
    f"  3. writing the path into {_TEMPLATE_PATH_FILE}"
)


def _resolve_template_path(template_path: Optional[Path]) -> Path:
    if template_path is not None:
        p = Path(template_path).expanduser()
        if not p.exists():
            raise FileNotFoundError(f"Template not found at explicit path: {p}")
        return p
    env = os.environ.get("TSINGHUA_THESIS_TEMPLATE")
    if env:
        p = Path(env).expanduser()
        if not p.exists():
            raise FileNotFoundError(f"TSINGHUA_THESIS_TEMPLATE points to missing file: {p}")
        return p
    if _TEMPLATE_PATH_FILE.exists():
        first_line = _TEMPLATE_PATH_FILE.read_text(encoding="utf-8").splitlines()
        # skip empty + comment lines
        for line in first_line:
            line = line.strip()
            if line and not line.startswith("#"):
                p = Path(line).expanduser()
                if not p.exists():
                    raise FileNotFoundError(
                        f"{_TEMPLATE_PATH_FILE} points to missing file: {p}"
                    )
                return p
    raise FileNotFoundError(_DOWNLOAD_HINT)


# ---------------------------------------------------------------------------
# I/O
# ---------------------------------------------------------------------------

def open_template(template_path: Optional[Path] = None) -> Document:
    """Open the official template .docx as the starting Document."""
    return _DocumentCtor(str(_resolve_template_path(template_path)))


def save(doc: Document, out_path: Path | str) -> Path:
    """Save document. Refuses to overwrite the template itself."""
    out = Path(out_path).expanduser().resolve()
    template = _resolve_template_path(None) if _can_resolve_template() else None
    if template is not None and out == template.resolve():
        raise ValueError(
            f"Refusing to overwrite the template at {template}. "
            "Pass a different out_path."
        )
    out.parent.mkdir(parents=True, exist_ok=True)
    doc.save(str(out))
    return out


def _can_resolve_template() -> bool:
    try:
        _resolve_template_path(None)
        return True
    except FileNotFoundError:
        return False


# ---------------------------------------------------------------------------
# Cover (Table 0 + 封面* paragraphs)
# ---------------------------------------------------------------------------

# Cover table row → label (must match what's in the template)
_COVER_TABLE_ROWS = {
    "department": 0,   # 系别
    "major":      1,   # 专业
    "author":     2,   # 姓名
    "advisor":    3,   # 指导教师
}
_COVER_TABLE_TEXT_COL = 2  # the column holding the value


def set_cover_info(
    doc: Document,
    *,
    title_cn: str,
    author: str,
    department: str,
    major: str,
    advisor: str,
    date: str,
) -> None:
    """Fill the cover page: title paragraph + 4x3 info table + date paragraph.

    Args:
        title_cn: Chinese thesis title (replaces template placeholder text in
                  the 封面论文题目 paragraph).
        author: 姓名.
        department: 系别 e.g. "自动化系".
        major: 专业 e.g. "自动化".
        advisor: e.g. "张三 教授".
        date: e.g. "二○二六年六月" (Chinese reading is conventional on cover).
    """
    # 1. Replace title paragraph (style: 封面论文题目). It's the second 封面论文题目
    #    paragraph (first one is intentionally blank in the template).
    title_paragraphs = [p for p in doc.paragraphs if p.style.name == S.COVER_TITLE]
    if not title_paragraphs:
        raise ValueError(f"Template missing any '{S.COVER_TITLE}' paragraph.")
    # Find the one with actual placeholder text; fall back to the last one
    target = next((p for p in title_paragraphs if p.text.strip()), title_paragraphs[-1])
    replace_paragraph_text(target, title_cn)

    # 2. Fill the 4x3 cover table
    if not doc.tables:
        raise ValueError("Template missing cover info table.")
    cover_table = doc.tables[0]
    field_values = {
        "department": department,
        "major": major,
        "author": author,
        "advisor": advisor,
    }
    for field, row_idx in _COVER_TABLE_ROWS.items():
        cell = cover_table.rows[row_idx].cells[_COVER_TABLE_TEXT_COL]
        set_table_cell_text(cell, field_values[field])

    # 3. Fill the date paragraph (style 封面作者信息, contains '年' in the placeholder)
    for p in doc.paragraphs:
        if p.style.name == S.COVER_AUTHOR_INFO and "年" in p.text:
            replace_paragraph_text(p, date)
            break


# ---------------------------------------------------------------------------
# Abstract / Keywords
# ---------------------------------------------------------------------------

def set_abstract(
    doc: Document,
    *,
    cn_text: str,
    cn_keywords: Sequence[str],
    en_text: str,
    en_keywords: Sequence[str],
) -> None:
    """Replace the 摘要 and Abstract sections (text + keyword line) with user content.

    `cn_text` and `en_text` may contain newlines; each line becomes a paragraph
    styled `段落` (matching the template's existing abstract paragraphs).
    """
    _replace_abstract_section(
        doc,
        heading_text="摘    要",
        new_body=cn_text,
        new_keywords_line="关键词：" + "；".join(cn_keywords),
        keyword_marker="关键词：",
    )
    _replace_abstract_section(
        doc,
        heading_text="Abstract",
        new_body=en_text,
        new_keywords_line="Keywords: " + "; ".join(en_keywords),
        keyword_marker="Keywords:",
    )


def _replace_abstract_section(
    doc: Document,
    *,
    heading_text: str,
    new_body: str,
    new_keywords_line: str,
    keyword_marker: str,
) -> None:
    heading = find_paragraph_by_text(doc, heading_text)
    if heading is None or heading.style.name != S.CHAPTER_TITLE:
        # Re-search ignoring style if the strict match failed
        heading = find_paragraph_by_text(doc, heading_text.strip())
    if heading is None:
        raise ValueError(f"Could not find abstract heading: {heading_text!r}")

    # Collect existing body paragraphs between heading and keyword line.
    body_paragraphs = []
    keyword_paragraph = None
    for p in iter_paragraphs_after(doc, heading):
        if keyword_marker in p.text:
            keyword_paragraph = p
            break
        if p.style.name == S.CHAPTER_TITLE:
            # Reached the next section without finding keywords; abort defensively
            break
        body_paragraphs.append(p)

    # Strategy: rewrite the first body paragraph to first line of new content;
    # for additional lines, insert siblings before the keyword paragraph.
    lines = [ln for ln in new_body.split("\n") if ln.strip()]
    if not lines:
        lines = [""]

    if body_paragraphs:
        # Rewrite the first body paragraph; insert remaining lines after it;
        # delete any leftover original-template body paragraphs entirely.
        replace_paragraph_text(body_paragraphs[0], lines[0])
        anchor = body_paragraphs[0]
        for line in lines[1:]:
            anchor = insert_paragraph_after(anchor, line, style=S.BODY_LEGACY)
        for extra in body_paragraphs[1:]:
            remove_paragraph(extra)

    if keyword_paragraph is not None:
        replace_paragraph_text(keyword_paragraph, new_keywords_line)


# ---------------------------------------------------------------------------
# Body structure
# ---------------------------------------------------------------------------

def add_chapter(doc: Document, title: str) -> Paragraph:
    """Append a chapter title using `章标题-无级别`. Use for 引言, 第 N 章, 结论."""
    return doc.add_paragraph(title, style=S.CHAPTER_TITLE)


def add_section(doc: Document, title: str, level: Literal[1, 2, 3, 4] = 1) -> Paragraph:
    """Append a section heading. level 1..4 → Heading 1..4."""
    style_name = {1: S.HEADING_1, 2: S.HEADING_2, 3: S.HEADING_3, 4: S.HEADING_4}[level]
    return doc.add_paragraph(title, style=style_name)


def add_body(doc: Document, text: str) -> Paragraph:
    """Append a body paragraph styled `论文正文段落`."""
    return doc.add_paragraph(text, style=S.BODY)


# ---------------------------------------------------------------------------
# Floats: Figures
# ---------------------------------------------------------------------------

def add_figure(
    doc: Document,
    image_path: Path | str,
    caption: str,
    *,
    width_cm: float = 12.0,
    label: Optional[str] = None,
) -> Paragraph:
    """Insert a figure and its caption.

    The figure occupies a `图片`-styled paragraph; the caption appears below
    in a `Caption`-styled paragraph reading e.g. "图 X.Y label  caption".
    """
    img_p = doc.add_paragraph(style=S.FIGURE_PARAGRAPH)
    run = img_p.add_run()
    run.add_picture(str(image_path), width=Cm(width_cm))

    text = f"{label}  {caption}" if label else caption
    cap_p = doc.add_paragraph(text, style=S.FIG_CAPTION)
    return cap_p


# ---------------------------------------------------------------------------
# Floats: Three-line table
# ---------------------------------------------------------------------------

def add_three_line_table(
    doc: Document,
    header: Sequence[str],
    rows: Sequence[Sequence[str]],
    caption: str,
    *,
    label: Optional[str] = None,
) -> Table:
    """Build a 三线表-styled table with its caption above (表-题注)."""
    cap_text = f"{label}  {caption}" if label else caption
    doc.add_paragraph(cap_text, style=S.TABLE_CAPTION)

    ncols = len(header)
    table = doc.add_table(rows=1 + len(rows), cols=ncols)
    table.style = S.THREE_LINE_TABLE
    # Header row
    for c, h in enumerate(header):
        table.rows[0].cells[c].text = h
    # Body rows
    for r, row in enumerate(rows, start=1):
        if len(row) != ncols:
            raise ValueError(
                f"row {r-1} has {len(row)} cells but header has {ncols} columns"
            )
        for c, val in enumerate(row):
            table.rows[r].cells[c].text = str(val)
    return table


# ---------------------------------------------------------------------------
# Equations (pandoc-backed OMML)
# ---------------------------------------------------------------------------

def add_equation(
    doc: Document,
    latex: str,
    *,
    label: Optional[str] = None,
) -> Paragraph:
    """Insert a display equation as native Word OMML, styled `公式`.

    `label` (e.g. "(2-3)") is appended as trailing text in the same paragraph
    so it shows on the right; you can use Word's tab stops to right-align it.

    Requires pandoc on PATH. Raises RuntimeError with install hint if absent.
    """
    from ._equation import latex_to_omml

    p = doc.add_paragraph(style=S.EQUATION)
    for omath in latex_to_omml(latex, display=True):
        append_omml(p, omath)
    if label:
        # tab + label for right-alignment if user has set a right tab stop
        p.add_run("\t" + label)
    return p


def add_inline_equation(paragraph: Paragraph, latex: str) -> None:
    """Embed an inline equation inside an existing paragraph (OMML, not text)."""
    from ._equation import latex_to_omml

    for omath in latex_to_omml(latex, display=False):
        append_omml(paragraph, omath)


# ---------------------------------------------------------------------------
# Code
# ---------------------------------------------------------------------------

def add_code_block(
    doc: Document,
    code: str,
    language: Optional[str] = None,
) -> Paragraph:
    """Append code styled `行间代码`. Each input line → one paragraph.

    The `language` argument is currently recorded as a leading comment in the
    first paragraph for context; styling is otherwise uniform.
    """
    lines = code.split("\n")
    first_p = None
    if language:
        first_p = doc.add_paragraph(f"# language: {language}", style=S.CODE_BLOCK)
    for line in lines:
        p = doc.add_paragraph(line if line else " ", style=S.CODE_BLOCK)
        if first_p is None:
            first_p = p
    return first_p  # type: ignore[return-value]


def add_inline_code(paragraph: Paragraph, code: str) -> None:
    """Append a run with `行内代码` character style to an existing paragraph."""
    run = paragraph.add_run(code)
    # Apply character style
    char_style = paragraph.part.document.styles[S.INLINE_CODE]
    run.style = char_style


# ---------------------------------------------------------------------------
# References & Appendix
# ---------------------------------------------------------------------------

def add_reference(doc: Document, entry: str) -> Paragraph:
    """Append one reference entry styled `参考文献`.

    Entries should already be formatted per GB/T 7714—2015; this skill does
    not parse BibTeX (see references/known-limitations.md).
    """
    return doc.add_paragraph(entry, style=S.REFERENCE)


def add_appendix_heading(
    doc: Document,
    title: str,
    level: Literal[0, 1, 2, 3] = 0,
) -> Paragraph:
    """Append `附录标题` (level=0) or `附录标题 1/2/3`."""
    style_name = {
        0: S.APPENDIX_HEADING_0,
        1: S.APPENDIX_HEADING_1,
        2: S.APPENDIX_HEADING_2,
        3: S.APPENDIX_HEADING_3,
    }[level]
    return doc.add_paragraph(title, style=style_name)


def add_symbols_table(
    doc: Document,
    rows: Sequence[tuple[str, str]],
) -> Paragraph:
    """Append entries to the 符号和缩略语说明表.

    Each row is one paragraph styled `符号和缩略语说明表` with a tab between
    symbol and description (the template's existing rows use this layout).
    """
    last = None
    for sym, desc in rows:
        text = f"{sym}\t{desc}"
        last = doc.add_paragraph(text, style=S.SYMBOLS_TABLE)
    if last is None:
        raise ValueError("add_symbols_table: empty rows")
    return last


# ---------------------------------------------------------------------------
# TOC of figures / tables — placeholder helper
# ---------------------------------------------------------------------------

def insert_toc_placeholder(
    doc: Document,
    kind: Literal["figures", "tables"],
) -> Paragraph:
    """Insert a chapter-titled placeholder for 插图清单 or 附表清单.

    The actual TOC field already exists in the template; this helper is for
    cases where you've built a doc from scratch and want a labeled spot for
    the user to insert the field manually in Word (Insert > Table of Figures).
    """
    title = "插图清单" if kind == "figures" else "附表清单"
    heading = doc.add_paragraph(title, style=S.CHAPTER_TITLE)
    doc.add_paragraph(
        f"[在此插入 Word > 引用 > 插入表目录，类型选择 '{title.replace('清单','')}']",
        style=S.NORMAL,
    )
    return heading
