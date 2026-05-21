# Style Decision Tree

Quick lookup: "I have content of type X — which helper / style do I use?"

## I have a heading

```
Is it the chapter title of an unnumbered front/back section (摘要, Abstract, 引言, 参考文献, 致谢)?
├─ YES  → add_chapter(doc, "...")              [style: 章标题-无级别]
└─ NO   → numbered chapter "第 N 章 标题"?
          ├─ YES  → add_section(doc, "...", level=1)   [Heading 1]
          ├─ Section like "1.1"   → add_section(level=2)
          ├─ Subsection "1.1.1"   → add_section(level=3)
          └─ Sub-subsection "1.1.1.1" → add_section(level=4)
```

## I have body text

```
Is it inside an abstract section?
├─ YES → set_abstract(...) replaces it       [style: 段落 — template's abstract style]
└─ NO  → add_body(doc, text)                  [style: 论文正文段落]
```

## I have a figure

```
add_figure(doc, image_path, caption="说明文字", label="图 X-Y")
                              ↑                    ↑
                              caption text         numbering label
[Styles: 图片 for image paragraph, Caption for caption paragraph]
```

## I have a table

```
Is it a 三线表 (3-row-line table, standard academic)?
├─ YES → add_three_line_table(doc, header, rows, caption, label="表 X-Y")
│         [Styles: 三线表 + 表-题注]
└─ NO  → use python-docx Table API directly + apply doc.styles["Normal Table"]
         or another table style from styles.md
```

## I have a math expression

```
LaTeX source available?
├─ YES + standalone display → add_equation(doc, latex, label="(X-Y)")
│         [pandoc converts to OMML, paragraph style: 公式]
├─ YES + inline in a paragraph → add_inline_equation(paragraph, latex)
└─ NO  (only know how to type into Word's equation editor)
         → leave a placeholder paragraph with add_body(...) and finalize in Word
```

## I have code

```
Block of code?
├─ YES → add_code_block(doc, code, language="python")   [style: 行间代码]
└─ NO inline literal like `np.array`?
       → in an existing paragraph: add_inline_code(paragraph, code)
         [character style: 行内代码]
```

## I have a reference entry

```
add_reference(doc, "竺可桢. 物理学[M]. 北京: 科学出版社, 1973: 56-60.")
[style: 参考文献]
# Format the string yourself per GB/T 7714—2015
```

## I have an appendix

```
add_appendix_heading(doc, "附录 A  XX", level=0)   [附录标题]
   add_appendix_heading(doc, "A.1 ...", level=1)   [附录标题 1]
   add_body(doc, "...")
```

## I have a symbol / abbreviation entry

```
add_symbols_table(doc, [
    ("MILP", "Mixed-Integer Linear Programming"),
    ("γ", "比热比=1.4"),
])
[style: 符号和缩略语说明表]
```

## I'm editing the cover

```
set_cover_info(
    doc,
    title_cn="<paper title>",
    author="<name>",
    department="<dept>",
    major="<major>",
    advisor="<advisor>",
    date="二○二六年六月",
)
# This edits in place: cover title paragraph + 4×3 info table + date paragraph.
# Do NOT rebuild the cover table or any 封面* paragraph manually.
```

## I'm editing the abstract

```
set_abstract(
    doc,
    cn_text="第一段...\n第二段...",
    cn_keywords=["k1", "k2", "k3"],
    en_text="...",
    en_keywords=["k1", "k2"],
)
# Replaces the placeholder paragraphs between 摘要/Abstract heading and the
# keyword line. The keyword line is rewritten as "关键词：xx；yy" / "Keywords: ...".
```

## I can't find a helper for what I need

1. Check `references/styles.md` for the right style name.
2. Use python-docx's `doc.add_paragraph(text, style=<name>)` directly.
3. If you find yourself repeating it, add a new helper to `scripts/helpers.py` (see "Extending" in SKILL.md).
