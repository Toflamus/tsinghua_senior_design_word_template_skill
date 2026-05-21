# Known Limitations

## Hard limits of python-docx

python-docx does not natively handle several Word features. These require manual work in Word/WPS after generation:

### 1. TOC fields (目录, 插图清单, 附表清单)

The template embeds Word `TOC` field codes for these three lists. python-docx cannot evaluate them. **You must press F9 in Word** (or right-click → Update Field) after opening the saved `.docx`.

If you skipped using the template's existing fields (e.g. you built a doc from scratch), this skill provides `insert_toc_placeholder(doc, kind)` to add a labeled spot where you can later: `Word > 引用 > 插入表目录` → choose 图 or 表.

### 2. Cross-references (e.g. "如图 3.1 所示")

Currently rendered as **plain text**. You must keep these synchronized manually when figure/table/equation numbers shift. Word's `REF` field-based cross-references are a possible future enhancement but not implemented.

Recommended workflow:

- Use stable label prefixes like `图 X-Y` where X = chapter, Y = sequence.
- Search-and-replace in Word if numbers change.

### 3. Equation labels (e.g. "(2-3)")

`add_equation(latex, label="(2-3)")` inserts the label as **trailing text** in the equation paragraph (with a tab so right-alignment works if you set a right tab stop). The label is not a Word field — cross-referencing it requires Word field codes set up manually.

### 4. Merged-cell tables

`add_three_line_table` builds rectangular tables. The template's existing complex tables — **cover info table (Table 0, 4×3)**, **signature table (Table 1)**, and **训练记录表 (Table 4, 7×6 with merged cells)** — are handled differently:

- Cover table: this skill *edits cells in place* via `set_cover_info`, never rebuilds it.
- Signature / 训练记录表: left for the user to fill in Word.

If you need to build a table with merged cells, use `python-docx` cell merging APIs directly on the result of `add_three_line_table` (or write raw XML via `scripts/_xml.py`).

### 5. BibTeX → 参考文献

Not implemented. `add_reference(entry)` accepts pre-formatted GB/T 7714—2015 strings:

```python
h.add_reference(doc, "竺可桢. 物理学[M]. 北京: 科学出版社, 1973: 56-60.")
```

A future enhancement could integrate `pybtex` or `bibtexparser` to format from `.bib`.

## Equation conversion (pandoc dependency)

`add_equation` / `add_inline_equation` shell out to **pandoc** to convert LaTeX → OMML. Pandoc must be on PATH:

- Debian/Ubuntu/WSL: `sudo apt install pandoc`
- macOS: `brew install pandoc`
- Windows: `winget install --id JohnMacFarlane.Pandoc`
- Or download from <https://pandoc.org/installing.html>

If pandoc is missing, `RuntimeError` is raised with the install hint.

Coverage: pandoc handles most standard LaTeX math (`\frac`, `\sum`, `\int`, subscripts, superscripts, Greek, etc.). Unusual macros from custom `.sty` files won't work.

If pandoc produces formatting that doesn't match the rest of the document (font size, italic conventions), open the equation in Word's equation editor and apply the template's `公式` paragraph style — the OMML content is editable post-insertion.

## Fonts on Linux / WSL

When previewing the `.docx` (e.g. via LibreOffice or unoconv), Chinese characters may render as missing glyphs unless CJK fonts are installed:

```bash
sudo apt install fonts-wqy-zenhei fonts-noto-cjk
fc-cache -f -v
```

Word/WPS on Windows/macOS bundle Chinese fonts and don't need this.

## Template version drift

The skill is designed against the specific style names in the version of `01 综合论文训练论文模板.docx` we inspected. If 教务处 releases a new template that renames styles, helpers will raise `KeyError` on `doc.styles[...]`. The fix:

1. Re-inspect the new template: extract all style names.
2. Update constants in `scripts/styles.py`.
3. Re-run `tests/test_smoke.py`.

`template/style_fingerprint.json` (planned) records expected style names so the skill can detect drift on `open_template()` and warn loudly.

## Out-of-scope (won't be added)

- LaTeX export / dual-source thesis (Word ↔ LaTeX bidirectional sync).
- Authority signatures / digital seals.
- Auto thesis grading.
- Non-清华 templates (a future skill can fork).
