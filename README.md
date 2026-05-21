# tsinghua-thesis-template

A Claude Code **skill** that drives the 清华大学综合论文训练 (Tsinghua undergraduate thesis) Word template programmatically. Insert chapters, figures, tables, equations, and references with the right Chinese-named styles — no manual `Heading 2 / 论文正文段落` selection in Word.

## What's a Claude Code skill?

A folder under `~/.claude/skills/` containing a `SKILL.md` that Claude Code auto-loads at startup. When the model sees a thesis-related ask, it consults this skill's docs and helpers instead of guessing.

See <https://docs.claude.com/en/docs/claude-code/skills> for the format spec.

## Install

```bash
git clone https://github.com/Toflamus/tsinghua_senior_design_word_template_skill \
  ~/.claude/skills/tsinghua_senior_design_word_template_skill

# Python deps (or use your conda env):
pip install python-docx>=1.1 lxml

# For equations (optional but recommended):
sudo apt install pandoc      # Debian/Ubuntu/WSL
# brew install pandoc        # macOS

# Restart Claude Code so the new skill folder is discovered.
```

## Configure the template path

The template `.docx` is **not bundled** (copyright). Point the skill at your local copy via one of:

1. Pass `template_path=Path(...)` to `open_template`.
2. Set `$TSINGHUA_THESIS_TEMPLATE=/path/to/template.docx`.
3. Copy and edit the path file:
   ```bash
   cp template/TEMPLATE_PATH.txt.example template/TEMPLATE_PATH.txt
   $EDITOR template/TEMPLATE_PATH.txt   # paste the absolute path on a line by itself
   ```

The template (`01 综合论文训练论文模板.docx`) is published by 清华大学教务处 — obtain it from the official download page.

## Quick start

```python
from scripts import helpers as h

doc = h.open_template()
h.set_cover_info(
    doc,
    title_cn="<paper title>",
    author="<name>",
    department="<dept>",
    major="<major>",
    advisor="<advisor>",
    date="二○二六年六月",
)
h.set_abstract(doc, cn_text="...", cn_keywords=["..."],
                    en_text="...", en_keywords=["..."])
h.add_chapter(doc, "引言")
h.add_body(doc, "...")
h.add_equation(doc, r"E = mc^2", label="(1-1)")     # native OMML via pandoc
h.add_reference(doc, "竺可桢. 物理学[M]. 北京: 科学出版社, 1973.")
h.save(doc, "draft.docx")
```

Open the saved `.docx` in Word/WPS and **press F9** to refresh TOC / 插图清单 / 附表清单 fields.

## What's in here

```
tsinghua_senior_design_word_template_skill/
├── SKILL.md                      # entry point Claude Code reads
├── README.md                     # this file
├── LICENSE                       # MIT
├── pyproject.toml
├── scripts/
│   ├── helpers.py                # public API
│   ├── styles.py                 # style-name constants
│   ├── _xml.py                   # low-level OOXML helpers
│   └── _equation.py              # pandoc LaTeX → OMML
├── references/
│   ├── styles.md                 # full 194-style table
│   ├── document-structure.md     # locked section ordering + helper mapping
│   ├── known-limitations.md      # what doesn't work; pandoc deps; font caveats
│   └── style-decision-tree.md    # "I have X content → use helper Y"
├── examples/
│   ├── minimal_thesis.py
│   ├── chapter_with_figure_table.py
│   └── full_skeleton.py
├── tests/
│   ├── test_smoke.py             # end-to-end smoke test (skips without template)
│   └── test_equation.py          # OMML round-trip (skips without pandoc)
└── template/
    ├── TEMPLATE_PATH.txt.example # copy → TEMPLATE_PATH.txt with your local path
    └── style_fingerprint.json    # expected style names for version-drift detection
```

## Tests

```bash
cd ~/.claude/skills/tsinghua_senior_design_word_template_skill
pip install pytest
pytest tests/
```

Tests skip cleanly if the template path or pandoc is missing.

## Limitations

See [references/known-limitations.md](references/known-limitations.md). Highlights:

- TOC fields require `F9` refresh in Word.
- Cross-references (`图 3.1` text) are plain text, not Word REF fields.
- BibTeX → 参考文献 not supported (pre-formatted GB/T 7714—2015 strings only).
- Cover signature + 训练记录表 are filled in Word by hand.

## License

MIT — see [LICENSE](LICENSE). The license covers this skill's code and documentation only; the template `.docx` itself is **not** distributed here and is governed by its own copyright (清华大学).

## Contributing

PRs welcome. Convention:

1. Update `scripts/styles.py` first if a new template style is introduced.
2. Reference style constants from `helpers.py`; never hard-code names.
3. Add a smoke-test assertion to `tests/test_smoke.py`.
4. Update the cheat sheet in `SKILL.md` and the decision tree in `references/`.
