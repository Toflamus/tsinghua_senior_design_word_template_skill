"""tsinghua-thesis-template skill — Python helpers package.

Public API is in ``helpers`` (also re-exported here for convenience).
Style-name constants live in ``styles``.
"""
from . import helpers, styles  # noqa: F401
from .helpers import (  # noqa: F401
    open_template,
    save,
    set_cover_info,
    set_abstract,
    add_chapter,
    add_section,
    add_body,
    add_figure,
    add_three_line_table,
    add_equation,
    add_inline_equation,
    add_code_block,
    add_inline_code,
    add_reference,
    add_appendix_heading,
    add_symbols_table,
    insert_toc_placeholder,
)

__all__ = [
    "open_template", "save",
    "set_cover_info", "set_abstract",
    "add_chapter", "add_section", "add_body",
    "add_figure", "add_three_line_table",
    "add_equation", "add_inline_equation",
    "add_code_block", "add_inline_code",
    "add_reference", "add_appendix_heading",
    "add_symbols_table", "insert_toc_placeholder",
    "helpers", "styles",
]
