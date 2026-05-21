"""LaTeX → OMML conversion via pandoc subprocess.

Pandoc is a system dependency. If missing, raise with an install hint.
"""
from __future__ import annotations
import shutil
import subprocess
import tempfile
from pathlib import Path
from typing import List

from lxml import etree

from ._xml import W, M


_PANDOC_INSTALL_HINT = (
    "pandoc not found on PATH. Install it:\n"
    "  Debian/Ubuntu/WSL: sudo apt install pandoc\n"
    "  macOS:             brew install pandoc\n"
    "  Or download from https://pandoc.org/installing.html"
)


def _check_pandoc() -> str:
    path = shutil.which("pandoc")
    if not path:
        raise RuntimeError(_PANDOC_INSTALL_HINT)
    return path


def latex_to_omml(latex: str, *, display: bool = True) -> List[etree._Element]:
    """Convert a LaTeX math expression to OMML <m:oMath> elements.

    Returns a list (usually length 1) of deep-copyable lxml elements ready to
    splice into a Word paragraph.

    Args:
        latex: the raw LaTeX (no $ wrappers; we add them).
        display: True for display math ($$ ... $$), False for inline ($ ... $).

    Raises:
        RuntimeError if pandoc is missing.
        ValueError on empty input or pandoc failure.
    """
    if not latex or not latex.strip():
        raise ValueError("latex_to_omml: empty input")

    _check_pandoc()
    wrapped = f"$$\n{latex}\n$$\n" if display else f"${latex}$\n"

    with tempfile.TemporaryDirectory() as td:
        td = Path(td)
        src = td / "in.md"
        out = td / "out.docx"
        src.write_text(wrapped, encoding="utf-8")

        # Pandoc reads markdown (so the $...$ math escapes work) and writes docx.
        result = subprocess.run(
            ["pandoc", "-f", "markdown", "-t", "docx", "-o", str(out), str(src)],
            capture_output=True, text=True,
        )
        if result.returncode != 0:
            raise ValueError(f"pandoc failed (exit {result.returncode}):\n{result.stderr}")

        # Extract <m:oMath> elements from the resulting docx
        import zipfile
        with zipfile.ZipFile(out) as zf:
            with zf.open("word/document.xml") as f:
                tree = etree.parse(f)
        root = tree.getroot()
        # Find any oMath or oMathPara descendants
        omath_elements = root.findall(f".//{M('oMath')}")
        if not omath_elements:
            raise ValueError(
                f"pandoc produced no <m:oMath> for input: {latex!r}\n"
                f"Output XML (first 500 chars): {etree.tostring(root, pretty_print=True).decode()[:500]}"
            )
        return omath_elements
