"""Demo: a chapter that uses add_figure, add_three_line_table, add_equation,
add_inline_equation, add_inline_code, add_code_block.

Run from the skill directory:

    python examples/chapter_with_figure_table.py

Outputs `chapter_demo.docx`.

Notes:
- A small placeholder PNG is generated on the fly so the script is self-contained.
- The equation steps require pandoc on PATH. If missing, those calls are skipped
  with a console warning (the rest of the doc still builds).
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from scripts import helpers as h  # noqa: E402


def _make_placeholder_png(path: Path) -> Path:
    """Create a valid 1x1 red PNG so add_figure has something to insert."""
    import struct, zlib
    sig = b"\x89PNG\r\n\x1a\n"
    # IHDR: width=1, height=1, depth=8, color_type=2 (RGB), no interlace
    ihdr_data = struct.pack(">IIBBBBB", 1, 1, 8, 2, 0, 0, 0)
    ihdr = (struct.pack(">I", len(ihdr_data)) + b"IHDR" + ihdr_data
            + struct.pack(">I", zlib.crc32(b"IHDR" + ihdr_data) & 0xffffffff))
    # IDAT: filter byte 0x00 + 1 red pixel (RGB)
    raw = b"\x00\xff\x00\x00"
    compressed = zlib.compress(raw)
    idat = (struct.pack(">I", len(compressed)) + b"IDAT" + compressed
            + struct.pack(">I", zlib.crc32(b"IDAT" + compressed) & 0xffffffff))
    iend = struct.pack(">I", 0) + b"IEND" + struct.pack(">I", zlib.crc32(b"IEND") & 0xffffffff)
    path.write_bytes(sig + ihdr + idat + iend)
    return path


def main() -> Path:
    doc = h.open_template()

    # Minimal cover so the doc is openable
    h.set_cover_info(
        doc,
        title_cn="图、表、公式与代码插入演示",
        author="某 某 某",
        department="自动化系",
        major="自动化",
        advisor="××× 教授",
        date="二○二六年六月",
    )

    h.add_chapter(doc, "第 1 章  示例")
    h.add_section(doc, "图与表", level=2)
    h.add_body(doc, "下文演示 figure / three-line table 的插入方式。")

    img_path = _make_placeholder_png(Path("_placeholder.png"))
    h.add_figure(
        doc,
        img_path,
        caption="一个占位图，用于演示图与图题的样式。",
        label="图 1-1",
        width_cm=4.0,
    )

    h.add_three_line_table(
        doc,
        header=["求解策略", "运行时间 (s)", "最优 gap (%)"],
        rows=[
            ["MILP",      "1234.5",  "0.01"],
            ["GDP Hull",  " 876.2",  "0.05"],
            ["GDP BigM",  "1502.8",  "0.20"],
        ],
        caption="不同求解策略的性能对比。",
        label="表 1-1",
    )

    h.add_section(doc, "公式与代码", level=2)
    inline_p = h.add_body(
        doc, "如下式所示，目标函数为线性目标 c "
    )
    # Append an inline equation right inside this paragraph
    try:
        h.add_inline_equation(inline_p, r"c^T x")
        inline_p.add_run("，约束为 ")
        h.add_inline_equation(inline_p, r"Ax \le b")
        inline_p.add_run("。")
    except RuntimeError as e:
        print(f"[warn] inline equation skipped: {e}")
        inline_p.add_run(" (公式略，需 pandoc)")

    try:
        h.add_equation(
            doc,
            r"\min_{x \in X}\; c^T x \quad \text{s.t.} \quad A x \le b",
            label="(1-1)",
        )
    except RuntimeError as e:
        print(f"[warn] display equation skipped: {e}")
        h.add_body(doc, "公式 (1-1) 略 — 需安装 pandoc。")

    code_p = h.add_body(doc, "示例：使用 ")
    h.add_inline_code(code_p, "Pyomo")
    code_p.add_run(" 构造模型并交给 ")
    h.add_inline_code(code_p, "Gurobi")
    code_p.add_run(" 求解。")

    h.add_code_block(
        doc,
        code=(
            "from pyomo.environ import *\n"
            "m = ConcreteModel()\n"
            "m.x = Var(within=NonNegativeReals)\n"
            "m.obj = Objective(expr=m.x)\n"
            "SolverFactory('gurobi').solve(m)"
        ),
        language="python",
    )

    h.add_reference(
        doc,
        "HU C, OU T, CHANG H, et al. Deep GRU Neural Network Prediction "
        "and Feedforward Compensation[J]. IEEE Transactions on ..., 2024.",
    )

    out = h.save(doc, "chapter_demo.docx")
    print(f"Saved: {out}")
    return out


if __name__ == "__main__":
    main()
