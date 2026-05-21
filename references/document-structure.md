# Document Structure

The 清华综合论文训练 template enforces a fixed section sequence. Helpers below assume you build sections in this order.

## Section sequence

| # | Section | Style of heading | Helper to populate | Notes |
|---|---|---|---|---|
| 1 | 封面 (Cover) | `封面*` (4 styles) + 4×3 table | `set_cover_info` | Title, dept, major, author, advisor, date. **No English title on cover.** |
| 2 | 关于论文使用授权的说明 | `使用授权说明标题` | edit in Word | Static legal text + signature table (Table 1). |
| 3 | 摘要 | `章标题-无级别` | `set_abstract(cn_text=..., cn_keywords=...)` | Body paragraphs use `段落` style; keyword line uses `Normal`. |
| 4 | Abstract | `章标题-无级别` | `set_abstract(en_text=..., en_keywords=...)` | Same handling as CN. |
| 5 | 插图清单 | `章标题-无级别` | (auto in Word, press F9) | TOC field already in template; do not regenerate. |
| 6 | 附表清单 | `章标题-无级别` | (auto in Word, press F9) | Same. |
| 7 | 符号和缩略语说明 | `章标题-无级别` | `add_symbols_table(rows=[(sym, desc), ...])` | Each row uses `符号和缩略语说明表` style. |
| 8 | 引言 / 第 N 章 | `Heading 1` (numbered) **or** `章标题-无级别` (unnumbered like 引言/结论) | `add_chapter` (unnumbered) or `add_section(level=1)` (numbered) | The template uses `Heading 1` for "引言" — match the existing pattern. |
| 8a | 1.1 二级标题 | `Heading 2` | `add_section(level=2)` | |
| 8b | 1.1.1 三级标题 | `Heading 3` | `add_section(level=3)` | |
| 8c | Body | `论文正文段落` | `add_body` | |
| 8d | Figure | `图片` + `Caption` | `add_figure` | Caption text format: `图 X.Y  说明`. |
| 8e | Table | `表-题注` + `三线表` | `add_three_line_table` | Caption above table. Three-line styled. |
| 8f | Equation (display) | `公式` (OMML inserted) | `add_equation(latex, label="(X-Y)")` | Pandoc converts LaTeX → OMML. |
| 8g | Equation (inline) | OMML inserted in existing paragraph | `add_inline_equation(paragraph, latex)` | |
| 8h | Code block | `行间代码` | `add_code_block` | One line per paragraph. |
| 8i | Inline code | `行内代码` (char style) | `add_inline_code(paragraph, code)` | |
| 9 | 参考文献 | `Title` (template uses Title here) | `add_chapter(...)` or just `add_reference` after a heading | Each entry: `add_reference(entry)` using `参考文献` style. Format per GB/T 7714—2015. |
| 10 | 附录 X | `附录标题` (top), `附录标题 1/2/3` | `add_appendix_heading(level=0..3)` | Then `add_body` etc. inside. |
| 11 | 致谢 | `Title` | edit in Word | Static-ish; you can use `add_chapter` + `add_body`. |
| 12 | 声明 | `Title` + Table 1 (signature) | edit in Word | Signature table requires manual filling in Word. |
| 13 | 在学期间参加课题的研究成果 | `Title` | edit in Word | Add entries with `add_body` or paste. |
| 14 | 综合论文训练记录表 | `章标题-无级别` + Table 4 (7×6) | edit in Word | Final grade table — fill in Word by hand. |

## What this skill does not generate

- **TOC field codes** (`目录` / `插图清单` / `附表清单`) — already present in template. Press F9 in Word to refresh.
- **Cover authorization signatures** — Table 1's signature cells (作者签名/导师签名) are for handwriting; leave blank or fill in Word.
- **训练记录表** — Final grades table, instructor-signed; fill in Word.

## "Static" sections — edit in Word, not via this skill

These sections have boilerplate text the template publishes verbatim; modifying them programmatically risks losing the school's standard wording. Edit text in Word directly:

- 关于论文使用授权的说明
- 声明 (the legal statement, not the signature table)

## Recommended build order

If generating a draft from outline, follow:

```python
doc = h.open_template()
h.set_cover_info(doc, ...)
h.set_abstract(doc, ...)
h.add_symbols_table(doc, [...])     # 符号缩略语 (optional)
# 主体: 引言 + 各章
h.add_chapter(doc, "引言")           # uses 章标题-无级别; alternative: add_section(level=1, ...)
h.add_body(doc, "...")
# ... 章节内容 ...
# 参考文献 (heading already in template via Title style — see §9 above)
h.add_reference(doc, "...")
# 附录 / 致谢 — typically left static or edited in Word
h.save(doc, "out.docx")
```

After save: open in Word, F9 to refresh TOC, manually finish signature/grade tables, eyeball cover.
