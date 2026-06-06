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
    add_body_chapter,
    add_section,
    add_body,
    add_rich_body,
    add_figure,
    add_stacked_figure,
    add_three_line_table,
    add_equation,
    add_inline_equation,
    add_code_block,
    add_inline_code,
    add_reference,
    add_appendix_heading,
    add_symbols_table,
    insert_toc_placeholder,
    AnchorInserter,
    find_chapter_anchor,
    clear_example_body,
    clear_template_instruction_textboxes,
    renumber_caption_fields,
    render_inline_markup_in,
    strip_inline_html,
)
from ._image import safe_image  # noqa: F401
from ._html import html_table_to_grid  # noqa: F401

__all__ = [
    "open_template", "save",
    "set_cover_info", "set_abstract",
    "add_chapter", "add_body_chapter", "add_section", "add_body", "add_rich_body",
    "add_figure", "add_stacked_figure", "add_three_line_table",
    "add_equation", "add_inline_equation",
    "add_code_block", "add_inline_code",
    "add_reference", "add_appendix_heading",
    "add_symbols_table", "insert_toc_placeholder",
    "AnchorInserter", "find_chapter_anchor",
    "clear_example_body", "clear_template_instruction_textboxes",
    "renumber_caption_fields", "render_inline_markup_in",
    "strip_inline_html", "safe_image", "html_table_to_grid",
    "helpers", "styles",
]
