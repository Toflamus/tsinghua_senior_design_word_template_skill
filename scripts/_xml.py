"""Low-level OOXML helpers used internally by the public API.

Avoid using these from user code — the public API is in helpers.py.
"""
from __future__ import annotations
from typing import Iterable, Optional
from copy import deepcopy

from docx.document import Document
from docx.oxml import OxmlElement
from docx.text.paragraph import Paragraph
from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
M_NS = "http://schemas.openxmlformats.org/officeDocument/2006/math"
W = lambda tag: f"{{{W_NS}}}{tag}"
M = lambda tag: f"{{{M_NS}}}{tag}"


def find_paragraph_by_text(doc: Document, needle: str) -> Optional[Paragraph]:
    """Return the first paragraph whose stripped text equals or contains `needle`."""
    for p in doc.paragraphs:
        if needle in p.text:
            return p
    return None


def iter_paragraphs_after(doc: Document, anchor: Paragraph) -> Iterable[Paragraph]:
    """Yield paragraphs strictly after `anchor`, in document order.

    Compares the underlying XML element (anchor._p) rather than Python object
    identity — `doc.paragraphs` returns fresh Paragraph wrappers each call.
    """
    anchor_el = anchor._p
    seen = False
    for p in doc.paragraphs:
        if seen:
            yield p
        if p._p is anchor_el:
            seen = True


def clear_paragraph(p: Paragraph) -> None:
    """Remove all runs from a paragraph but keep its style + paragraph properties."""
    # Remove every <w:r> child while preserving <w:pPr>
    for child in list(p._p):
        if child.tag == W("r"):
            p._p.remove(child)


def replace_paragraph_text(p: Paragraph, text: str) -> None:
    """Clear all runs and write a single run with `text`. Preserves the style."""
    clear_paragraph(p)
    p.add_run(text)


def append_omml(paragraph: Paragraph, omath_el: etree._Element) -> None:
    """Append a deep-copied <m:oMath> element inside `paragraph`'s body XML.

    Word renders OMML inside a paragraph as a sibling of <w:r>, not nested in one.
    """
    paragraph._p.append(deepcopy(omath_el))


def insert_paragraph_after(paragraph: Paragraph, text: str = "", style: Optional[str] = None) -> Paragraph:
    """Insert a new paragraph immediately after `paragraph`. Mirrors a missing python-docx API.

    Uses python-docx's OxmlElement so the new <w:p> is a CT_P that supports
    style assignment.
    """
    new_p = OxmlElement("w:p")
    paragraph._p.addnext(new_p)
    from docx.text.paragraph import Paragraph as P
    np = P(new_p, paragraph._parent)
    if style:
        np.style = paragraph.part.document.styles[style]
    if text:
        np.add_run(text)
    return np


def remove_paragraph(paragraph: Paragraph) -> None:
    """Detach a paragraph from its parent (removes it entirely from the document)."""
    parent = paragraph._p.getparent()
    if parent is not None:
        parent.remove(paragraph._p)


def set_table_cell_text(cell, text: str, *, keep_style: bool = True) -> None:
    """Replace cell content with a single paragraph containing `text`.

    keep_style preserves the first paragraph's style and the cell's properties.
    """
    if not cell.paragraphs:
        cell.add_paragraph(text)
        return
    first = cell.paragraphs[0]
    # Clear extra paragraphs
    for p in cell.paragraphs[1:]:
        p._p.getparent().remove(p._p)
    replace_paragraph_text(first, text)
