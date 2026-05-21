"""Smallest end-to-end example: cover + abstract + 1 chapter + 1 reference.

Run from the skill directory (or with the skill on sys.path):

    python examples/minimal_thesis.py

Outputs `minimal_thesis.docx` in the current working directory.

Prerequisites:
- Template path configured (see SKILL.md > Configure template path).
- python-docx and lxml installed.
"""
import sys
from pathlib import Path

# Allow running directly from a checkout without `pip install -e`
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from scripts import helpers as h  # noqa: E402


def main() -> Path:
    doc = h.open_template()

    h.set_cover_info(
        doc,
        title_cn="示例：清华综合论文训练模板 Skill 验证文档",
        author="某 某 某",
        department="自动化系",
        major="自动化",
        advisor="××× 教授",
        date="二○二六年六月",
    )

    h.set_abstract(
        doc,
        cn_text=(
            "本文展示了 tsinghua-thesis-template skill 的端到端用法。"
            "通过 python-docx 驱动学校 Word 模板，按命名样式插入封面、摘要、"
            "正文和参考文献。\n"
            "演示了批量自动化能力，但 TOC、公式、签名表等仍需在 Word 中收尾。"
        ),
        cn_keywords=["综合论文训练", "Word 模板", "python-docx", "自动化"],
        en_text=(
            "This document demonstrates end-to-end use of the "
            "tsinghua-thesis-template skill. It populates the official Tsinghua "
            "thesis template via python-docx using template-correct named styles.\n"
            "Bulk automation is shown; TOC fields, equations, and signature "
            "tables still require finishing in Word."
        ),
        en_keywords=["thesis training", "Word template", "python-docx", "automation"],
    )

    h.add_chapter(doc, "引言")
    h.add_body(doc, "这是引言的第一段。论文应当包含研究目的、背景和方法概述。")
    h.add_body(doc, "这是引言的第二段。可以引用相关工作，参见参考文献章节。")

    h.add_reference(
        doc,
        "竺可桢. 物理学[M]. 北京: 科学出版社, 1973: 56-60.",
    )

    out = h.save(doc, "minimal_thesis.docx")
    print(f"Saved: {out}")
    return out


if __name__ == "__main__":
    main()
