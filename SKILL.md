---
name: tsinghua-thesis-template
description: Author a 清华大学综合论文训练 (Tsinghua undergrad thesis) by driving the official Word template via python-docx helpers. Applies template-correct Chinese-named styles (章标题-无级别, 论文正文段落, 三线表, 公式, 参考文献, etc.), fills cover info, inserts figures/tables/code, and converts LaTeX equations to native Word OMML via pandoc.
when_to_use: User is writing or editing a 清华综合论文训练 thesis, needs to produce or modify the official .docx, mentions 综合论文训练 / 毕业论文模板 / 章标题-无级别 / 论文正文段落, or asks to insert chapters/figures/tables/equations into the template programmatically.
allowed-tools: Read, Edit, Write, Bash
---

# Tsinghua Thesis Template Skill

Drive the 清华大学综合论文训练 Word template (`01 综合论文训练论文模板.docx`) from Python. Saves you from hand-applying 194 named styles in Word.

## Overview

The official Tsinghua undergraduate thesis (综合论文训练) is delivered as a `.docx` template with 194 pre-defined paragraph/character/table styles and a fixed section ordering (封面 → 摘要 → Abstract → 插图清单 → 附表清单 → 符号缩略语 → 引言 + 主体 → 参考文献 → 附录 → 致谢 → 声明 → 在学期间研究成果 → 综合论文训练记录表).

This skill provides reusable Python helpers (built on `python-docx`) to programmatically populate the template while keeping all styles template-correct.

Authoritative reference repo: <https://github.com/Toflamus/tsinghua_senior_design_word_template_skill>

## Prerequisites

1. **Python deps**: `pip install python-docx>=1.1 lxml`
2. **Equations (optional)**: `pandoc` on PATH. Debian/Ubuntu/WSL: `sudo apt install pandoc`. macOS: `brew install pandoc`.
3. **The template `.docx`** — obtain from the 教务处 download page; this skill does **not** redistribute it (copyright).

## Configure template path

`open_template()` resolves the template path in this order:

1. Explicit `template_path=Path(...)` argument
2. `$TSINGHUA_THESIS_TEMPLATE` environment variable
3. `<skill_dir>/template/TEMPLATE_PATH.txt` (first non-comment line is the path)

Typical setup (option 3):

```bash
cp ~/.claude/skills/tsinghua_senior_design_word_template_skill/template/TEMPLATE_PATH.txt.example \
   ~/.claude/skills/tsinghua_senior_design_word_template_skill/template/TEMPLATE_PATH.txt
# then edit TEMPLATE_PATH.txt to point at your local .docx
```

## Quick start

```python
import sys
sys.path.insert(0, "/path/to/skill_dir")  # or pip install -e
from scripts import helpers as h

doc = h.open_template()
h.set_cover_info(
    doc,
    title_cn="多产品聚烯烃生产计划的混合整数规划研究",
    author="王朝龙",
    department="化工系",
    major="化学工程与工业生物工程",
    advisor="××× 教授",
    date="二○二六年六月",
)
h.set_abstract(
    doc,
    cn_text="本文研究了……\n第二段……",
    cn_keywords=["生产调度", "混合整数规划", "Pyomo"],
    en_text="This paper studies...\nSecond paragraph...",
    en_keywords=["production scheduling", "MILP", "Pyomo"],
)
h.add_chapter(doc, "引言")
h.add_section(doc, "研究背景", level=2)
h.add_body(doc, "聚烯烃生产是石化行业重要的中间产品环节……")
h.add_equation(doc, r"\min_{x \in X} c^T x", label="(1-1)")
h.add_reference(doc, "竺可桢. 物理学[M]. 北京: 科学出版社, 1973: 56-60.")
h.save(doc, "thesis_draft.docx")
```

After saving, **open the `.docx` in Word/WPS and press F9** to refresh 插图清单/附表清单 fields.

## Style cheat sheet

| Content | Helper | Underlying style |
|---|---|---|
| Chapter title (引言/摘要/第N章/参考文献) | `add_chapter` | `章标题-无级别` |
| Section heading (1.1) | `add_section(..., level=2)` | `Heading 2` |
| Body paragraph | `add_body` | `论文正文段落` |
| Figure | `add_figure` | `图片` + `Caption` |
| 三线表 with caption | `add_three_line_table` | `表-题注` + `三线表` |
| Display equation | `add_equation` | `公式` (with native OMML) |
| Inline equation | `add_inline_equation` | OMML in current paragraph |
| Code block | `add_code_block` | `行间代码` |
| Inline code | `add_inline_code` | `行内代码` (character style) |
| Reference entry | `add_reference` | `参考文献` |
| Appendix heading | `add_appendix_heading` | `附录标题` / `附录标题 1-3` |
| 符号缩略语 row | `add_symbols_table` | `符号和缩略语说明表` |
| Cover info (4 fields + title + date) | `set_cover_info` | cover table + `封面*` |
| Abstract + keywords (zh/en) | `set_abstract` | `章标题-无级别` + `段落` |

Full 194-style reference: `references/styles.md`.

## Helper index

All in `scripts/helpers.py`:

- `open_template(template_path=None) -> Document`
- `save(doc, out_path) -> Path`
- `set_cover_info(doc, *, title_cn, author, department, major, advisor, date)`
- `set_abstract(doc, *, cn_text, cn_keywords, en_text, en_keywords)`
- `add_chapter(doc, title) -> Paragraph`
- `add_section(doc, title, level=1) -> Paragraph`
- `add_body(doc, text) -> Paragraph`
- `add_figure(doc, image_path, caption, *, width_cm=12.0, label=None) -> Paragraph`
- `add_three_line_table(doc, header, rows, caption, *, label=None) -> Table`
- `add_equation(doc, latex, *, label=None) -> Paragraph` (display, OMML)
- `add_inline_equation(paragraph, latex) -> None`
- `add_code_block(doc, code, language=None) -> Paragraph`
- `add_inline_code(paragraph, code) -> None`
- `add_reference(doc, entry) -> Paragraph`
- `add_appendix_heading(doc, title, level=0) -> Paragraph`
- `add_symbols_table(doc, rows: list[(symbol, desc)])`
- `insert_toc_placeholder(doc, kind: "figures"|"tables")`

## Hard limitations

- **TOC fields**: `插图清单` / `附表清单` / `目录` are Word fields. Press F9 in Word to refresh after generation.
- **Cross-references**: "图 3.1" is currently plain text. Word's REF field-based cross-refs are out of scope (could be added later).
- **BibTeX**: not supported. `add_reference` accepts pre-formatted GB/T 7714—2015 strings.
- **Merged-cell tables**: skill builds simple rectangular tables. The template's cover table is edited in place (don't rebuild it).
- **Fonts**: Linux/WSL users may need `apt install fonts-wqy-zenhei fonts-noto-cjk` for Chinese rendering when previewing.

Full detail: `references/known-limitations.md`.

## Examples

- `examples/minimal_thesis.py` — smallest end-to-end (cover + abstract + 1 chapter + 1 reference)
- `examples/chapter_with_figure_table.py` — figures/tables/equations/code demo
- `examples/full_skeleton.py` — every template section with placeholder content

## Extending

When adding a helper:

1. Update `scripts/styles.py` first if a new style name is involved.
2. Add the function to `scripts/helpers.py` with type hints + 1-line docstring.
3. Export it from `scripts/__init__.py`.
4. Add a row to the style cheat sheet above and a one-line entry in `## Helper index`.
5. Add a smoke-test assertion in `tests/test_smoke.py`.

Style invariant: do **not** hard-code style names inside `helpers.py` — always reference `styles.S.XXX` constants.
