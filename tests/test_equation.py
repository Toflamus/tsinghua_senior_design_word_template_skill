"""Round-trip LaTeX → OMML tests. Requires pandoc."""
import shutil
import pytest
import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

pytestmark = pytest.mark.skipif(
    not shutil.which("pandoc"),
    reason="pandoc not installed; see references/known-limitations.md",
)

from scripts._equation import latex_to_omml  # noqa: E402
from lxml import etree  # noqa: E402


@pytest.mark.parametrize("latex", [
    r"E = mc^2",
    r"\frac{a}{b}",
    r"\sum_{i=1}^n x_i",
    r"\int_0^\infty e^{-x^2}\,dx = \frac{\sqrt{\pi}}{2}",
    r"x \in \mathbb{R}^n, \;\; A x \le b",
])
def test_display_equation(latex):
    elements = latex_to_omml(latex, display=True)
    assert len(elements) >= 1
    for el in elements:
        # Verify it's an m:oMath element
        assert el.tag.endswith("}oMath"), f"Expected m:oMath, got {el.tag}"
        # Should have at least one child
        assert len(el) > 0


def test_inline_equation():
    elements = latex_to_omml(r"a + b", display=False)
    assert len(elements) >= 1
    assert all(el.tag.endswith("}oMath") for el in elements)


def test_empty_input_raises():
    with pytest.raises(ValueError):
        latex_to_omml("", display=True)
    with pytest.raises(ValueError):
        latex_to_omml("   \n   ", display=True)
