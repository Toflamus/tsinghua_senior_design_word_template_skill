"""Full skeleton: every template section with placeholder content.

Use as a starting point for a fresh thesis — fork and fill in.

Run from the skill directory:

    python examples/full_skeleton.py
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from scripts import helpers as h  # noqa: E402


def main() -> Path:
    doc = h.open_template()

    # ------- Cover -------
    h.set_cover_info(
        doc,
        title_cn="<填入中文论文标题>",
        author="<姓名>",
        department="<系别>",
        major="<专业>",
        advisor="<指导教师 职称>",
        date="二○二六年六月",
    )

    # ------- Abstract / Keywords -------
    h.set_abstract(
        doc,
        cn_text="<中文摘要：研究目的、研究方法、研究结果、研究结论>",
        cn_keywords=["<关键词1>", "<关键词2>", "<关键词3>"],
        en_text="<English abstract: purpose, methods, results, conclusions>",
        en_keywords=["<keyword1>", "<keyword2>", "<keyword3>"],
    )

    # ------- Symbols & abbreviations (front matter, optional) -------
    h.add_symbols_table(doc, rows=[
        ("MILP", "Mixed-Integer Linear Programming 混合整数线性规划"),
        ("GDP", "Generalized Disjunctive Programming 广义析取规划"),
    ])

    # ------- Body chapters -------
    # 引言 (unnumbered, uses 章标题-无级别)
    h.add_chapter(doc, "引言")
    h.add_body(doc, "<引言段落 1>")
    h.add_body(doc, "<引言段落 2>")

    # 第 1 章 文献综述
    h.add_section(doc, "第 1 章  文献综述", level=1)
    h.add_section(doc, "1.1 国内外研究现状", level=2)
    h.add_body(doc, "<本节正文>")
    h.add_section(doc, "1.1.1 子主题", level=3)
    h.add_body(doc, "<子节正文>")

    # 第 2 章 模型与方法
    h.add_section(doc, "第 2 章  模型与方法", level=1)
    h.add_section(doc, "2.1 数学模型", level=2)
    h.add_body(doc, "<模型描述>")
    try:
        h.add_equation(doc, r"\min_{x \ge 0}\; c^T x \;\;\text{s.t.}\;\; Ax \le b",
                       label="(2-1)")
    except RuntimeError as e:
        print(f"[warn] equation skipped (pandoc not installed): {e}")
        h.add_body(doc, "<公式 (2-1) — 待 pandoc 安装后自动插入>")

    # 第 3 章 实验
    h.add_section(doc, "第 3 章  实验", level=1)
    h.add_body(doc, "<实验设计与结果>")
    h.add_three_line_table(
        doc,
        header=["实验", "指标 A", "指标 B"],
        rows=[
            ["基线", "0.00", "0.00"],
            ["方案 1", "0.00", "0.00"],
        ],
        caption="实验结果汇总。",
        label="表 3-1",
    )

    # 第 4 章 结论
    h.add_section(doc, "第 4 章  结论与展望", level=1)
    h.add_body(doc, "<结论 1>")
    h.add_body(doc, "<结论 2>")

    # ------- References -------
    # Heading already exists in template under Title style. Just append entries.
    for entry in [
        "<参考文献条目 1，按 GB/T 7714—2015 格式>",
        "<参考文献条目 2>",
    ]:
        h.add_reference(doc, entry)

    # ------- Appendix -------
    h.add_appendix_heading(doc, "附录 A  <附录标题>", level=0)
    h.add_body(doc, "<附录正文>")

    # 致谢 / 声明 / 在学期间研究成果 / 训练记录表:
    #   保留模板原文，在 Word 中编辑文字与签名。

    out = h.save(doc, "thesis_skeleton.docx")
    print(f"Saved: {out}")
    print("Next: open in Word, F9 to refresh TOC fields, fill 致谢/声明/训练记录表 by hand.")
    return out


if __name__ == "__main__":
    main()
