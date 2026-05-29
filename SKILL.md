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
h.clear_example_body(doc)                                # wipe template's sample body
anchor = h.find_chapter_anchor(doc, "参考文献", style="Title")
ins = h.AnchorInserter(doc, anchor)                      # insert before 参考文献

ins(h.add_body_chapter, "引言")                          # Heading 1; Word adds "第 1 章"
ins(h.add_section, "研究背景", level=2)                  # Heading 2; Word adds "1.1"
ins(h.add_body, "聚烯烃生产是石化行业重要的中间产品环节……")
ins(h.add_equation, r"\min_{x \in X} c^T x", label="(1-1)")
h.save(doc, "thesis_draft.docx")
```

After saving, **open the `.docx` in Word/WPS and press F9** to refresh 目录/插图清单/附表清单 fields.

## Critical: heading style ≠ chapter title style

The Tsinghua template uses **three different styles** for chapter-level headings — picking the wrong one breaks Word's auto-numbering and produces duplicated numbers like "第 2 章 2. 建模分析":

| Section | Style | Helper | Auto-numbered? |
|---|---|---|---|
| 摘要 / Abstract / 插图清单 / 附表清单 / 符号缩略语 / 综合论文训练记录表 | `章标题-无级别` | `add_chapter` | No |
| 引言 / 第 N 章（正文章节）| `Heading 1` | `add_body_chapter` | **Yes** — Word prefixes "第 N 章" |
| 参考文献 / 致谢 / 声明 / 在学期间研究成果 | `Title` | (template default — don't append) | No |

**Never** prefix the heading text with the chapter number yourself. Word will add it. Write:

```python
ins(h.add_body_chapter, "建模分析")          # ✅ renders as "第 2 章 建模分析"
ins(h.add_body_chapter, "第 2 章 建模分析")  # ❌ renders as "第 2 章 第 2 章 建模分析"

ins(h.add_section, "聚烯烃排产问题描述", level=2)   # ✅ renders as "2.1 聚烯烃..."
ins(h.add_section, "2.1 聚烯烃排产问题描述", level=2)  # ❌ duplicated number prefix
```

The same rule applies to `add_section(level=2/3/4)` — Word auto-numbers as N.M / N.M.K / N.M.K.L.

**Appendix headings have the same trap.** `附录标题` auto-numbers as "附录 X", `附录标题 2/3` as "X.N" / "X.N.M". Only `附录标题 1` is *not* auto-numbered. Pass only the title text on the auto-numbered levels:

```python
ins(h.add_appendix_heading, "序列式 MILP 基准模型", level=0)         # ✅ "附录 A 序列式..."
ins(h.add_appendix_heading, "附录 A 序列式 MILP 基准模型", level=0)  # ❌ "附录 A 附录 A 序列式..."

ins(h.add_appendix_heading, "初始约束", level=2)        # ✅ "A.1 初始约束"
ins(h.add_appendix_heading, "1. 初始约束", level=2)     # ❌ "A.1 1. 初始约束"
```

**Reference entries have the same trap too.** The `参考文献` style auto-numbers as "[N]" in document order. Pass only the citation body; Word numbers it. Append entries in citation order so the in-text "[N]" citations align with the reference list:

```python
ins(h.add_reference, "葛凌生. 产业链的多元化和高端化推动...")        # ✅ "[1] 葛凌生..."
ins(h.add_reference, "[1] 葛凌生. 产业链的多元化和高端化推动...")   # ❌ "[1] [1] 葛凌生..."
```

## Anchored insertion (avoid appending past the back matter)

The template's back matter (参考文献 → 附录 → 致谢 → 声明 → 在学期间研究成果 → 训练记录表) lives at the *end* of the doc. Plain `doc.add_paragraph` appends past the back matter — your body chapters end up after 综合论文训练记录表, which is wrong.

Use `AnchorInserter` to insert before the references section instead, after first calling `clear_example_body` to remove the template's sample 引言/图表示例 chapters:

```python
h.clear_example_body(doc)
anchor = h.find_chapter_anchor(doc, "参考文献", style="Title")
ins = h.AnchorInserter(doc, anchor)
ins(h.add_body_chapter, "引言")
ins(h.add_body_chapter, "建模分析")
ins(h.add_section, "聚烯烃排产问题描述", level=2)
ins(h.add_figure, "fig.png", caption="...", label="图 2-1")  # 2 paragraphs, moved together
ins(h.add_equation, r"E = mc^2", label="(2-1)")
```

Each `ins(fn, …)` call invokes `fn(doc, …)` (the helper appends as usual) and then moves the just-added element(s) to immediately before the anchor — preserving order across calls and grouping multi-paragraph adds (figure = image + caption) together.

## Style cheat sheet

| Content | Helper | Underlying style |
|---|---|---|
| **Body chapter** (引言, 第 N 章) — text is **just** the chapter name | `add_body_chapter` | `Heading 1` |
| **Front-matter chapter** (摘要, Abstract, 插图清单, 附表清单, 符号缩略语, 综合论文训练记录表) | `add_chapter` | `章标题-无级别` |
| **Back-matter chapter** (参考文献, 致谢, 声明, 在学期间研究成果) — *use template defaults* | (do not append) | `Title` |
| Section heading §N.M — text is **just** the section name | `add_section(..., level=2)` | `Heading 2` |
| Subsection §N.M.K | `add_section(..., level=3)` | `Heading 3` |
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
- `add_chapter(doc, title) -> Paragraph` — front-matter chapter (`章标题-无级别`)
- `add_body_chapter(doc, title) -> Paragraph` — body chapter (`Heading 1`, auto-numbered)
- `add_section(doc, title, level=1) -> Paragraph`
- `add_body(doc, text) -> Paragraph`
- `add_rich_body(doc, text) -> Paragraph` — body paragraph rendering inline `$math$`/`\(math\)`/`**bold**`/`` `code` `` markup
- `add_figure(doc, image_path, caption, *, width_cm=12.0, label=None) -> Paragraph`
- `add_three_line_table(doc, header, rows, caption, *, label=None) -> Table`
- `add_equation(doc, latex, *, label=None, on_error="raise"|"text") -> Paragraph` (display, OMML; rewrites `\atop`→`\substack`; `on_error="text"` falls back to raw text instead of raising)
- `add_inline_equation(paragraph, latex) -> None`
- `add_code_block(doc, code, language=None) -> Paragraph`
- `add_inline_code(paragraph, code) -> None`
- `add_reference(doc, entry) -> Paragraph`
- `add_appendix_heading(doc, title, level=0) -> Paragraph`
- `add_symbols_table(doc, rows: list[(symbol, desc)])`
- `insert_toc_placeholder(doc, kind: "figures"|"tables")`
- `AnchorInserter(doc, anchor)` — class; `ins(fn, *args, **kw)` inserts before `anchor`
- `find_chapter_anchor(doc, contains, *, style=None) -> Paragraph`
- `clear_example_body(doc) -> None` — wipe template's sample 引言/图表示例 chapters
- `renumber_caption_fields(doc, *, appendix_heading_style=None) -> None` — bake correct "图 N.M"/"表 N.M" into caption field caches so they display right without F9

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
