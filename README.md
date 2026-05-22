# 🎓 tsinghua-thesis-template

一个 Claude Code **skill**，用来程序化地驱动清华大学综合论文训练（本科毕设）官方 Word 模板。✍️ 让 Claude Code 帮你直接往 `.docx` 里插入章节、图、表、公式和参考文献，并自动套用模板里那些中文样式名（`章标题-无级别`、`论文正文段落`、`三线表`、`公式`、`参考文献` 等等），不用再在 Word 里手动一段段挑样式。

## 🤔 什么是 Claude Code skill？

`~/.claude/skills/` 下的一个文件夹，里面放一份 `SKILL.md`。Claude Code 启动时会自动扫描并加载它，之后当你跟 Claude 聊到毕设相关的事情，它会**自动**翻这里的文档和 helper，而不是凭空猜样式名。✨

格式规范见 👉 <https://docs.claude.com/en/docs/claude-code/skills>。

## 📦 安装

```bash
git clone https://github.com/Toflamus/tsinghua_senior_design_word_template_skill \
  ~/.claude/skills/tsinghua_senior_design_word_template_skill

# 🐍 Python 依赖（推荐装到你的 conda 环境里）：
pip install python-docx>=1.1 lxml

# 🧮 公式渲染需要（可选，但强烈推荐）：
sudo apt install pandoc      # Debian / Ubuntu / WSL
# brew install pandoc        # macOS

# 🔄 装完后重启 Claude Code，让它扫到这个新 skill。
```

## 🗂️ 配置模板路径

⚠️ 由于版权原因，模板 `.docx` **不打包在仓库里**。你需要从下面三种方式里挑一种，告诉 skill 你本地模板的位置：

1. 调用 `open_template(template_path=Path(...))` 时显式传入；
2. 设置环境变量 `$TSINGHUA_THESIS_TEMPLATE=/绝对路径/template.docx`；
3. 复制并填写路径文件：
   ```bash
   cp template/TEMPLATE_PATH.txt.example template/TEMPLATE_PATH.txt
   $EDITOR template/TEMPLATE_PATH.txt   # 把绝对路径单独写一行
   ```

📥 模板文件（`01 综合论文训练论文模板.docx`）由清华大学教务处发布，从官方下载页获取即可。

## 🚀 快速上手

```python
from scripts import helpers as h

doc = h.open_template()
h.set_cover_info(
    doc,
    title_cn="<论文题目>",
    author="<姓名>",
    department="<系名>",
    major="<专业>",
    advisor="<导师>",
    date="二○二六年六月",
)
h.set_abstract(doc, cn_text="...", cn_keywords=["..."],
                    en_text="...", en_keywords=["..."])
h.add_chapter(doc, "引言")
h.add_body(doc, "...")
h.add_equation(doc, r"E = mc^2", label="(1-1)")     # 🧮 走 pandoc 转成原生 OMML
h.add_reference(doc, "竺可桢. 物理学[M]. 北京: 科学出版社, 1973.")
h.save(doc, "draft.docx")
```

💡 打开保存好的 `.docx`（Word 或 WPS 都行），**按 F9** 刷新目录 / 插图清单 / 附表清单 字段。

## 📁 项目结构

```
tsinghua_senior_design_word_template_skill/
├── 📘 SKILL.md                   # Claude Code 加载时读的入口
├── 📄 README.md                  # 本文档
├── ⚖️  LICENSE                    # MIT
├── 🔧 pyproject.toml
├── 🐍 scripts/
│   ├── helpers.py                # 对外公开的 API
│   ├── styles.py                 # 模板里所有样式名常量（唯一来源）
│   ├── _xml.py                   # 底层 OOXML helper
│   └── _equation.py              # pandoc LaTeX → OMML 转换
├── 📚 references/
│   ├── styles.md                 # 194 个样式全表，按用途分组
│   ├── document-structure.md     # 论文章节固定顺序 + 每节对应的 helper
│   ├── known-limitations.md      # 做不到的事情、pandoc 依赖、字体注意点
│   └── style-decision-tree.md    # "我要插 X 内容 → 用 helper Y" 的决策树
├── 💡 examples/
│   ├── minimal_thesis.py         # 最短可用示例
│   ├── chapter_with_figure_table.py
│   └── full_skeleton.py          # 把模板的每个章节占位都生成一遍
├── ✅ tests/
│   ├── test_smoke.py             # 端到端冒烟测试（没有模板时自动 skip）
│   └── test_equation.py          # OMML 往返测试（没有 pandoc 时自动 skip）
└── 📋 template/
    ├── TEMPLATE_PATH.txt.example # 复制成 TEMPLATE_PATH.txt 并填本地路径
    └── style_fingerprint.json    # 记录模板里所有样式名，便于检测模板版本漂移
```

## 🧪 跑测试

```bash
cd ~/.claude/skills/tsinghua_senior_design_word_template_skill
pip install pytest
pytest tests/
```

🟢 没配置模板路径或没装 pandoc 时，对应测试会自动 skip，不会让 CI 飘红。

## ⚠️ 已知限制

详见 [references/known-limitations.md](references/known-limitations.md)。简要说：

- 📑 目录、插图清单、附表清单这些字段需要在 Word 里按 `F9` 手动刷新；
- 🔗 正文里 "图 3.1" 这种**交叉引用是纯文本**，不会自动跟着图序变。Word 的 REF 字段没实现；
- 📚 参考文献需要你提供按 GB/T 7714—2015 格式预排好的字符串，不支持 BibTeX 自动生成；
- ✍️ 封面手签字、训练记录表这种必须人填的部分，仍然要在 Word 里手填。

## ⚖️ 许可证

MIT — 详见 [LICENSE](LICENSE)。该许可只覆盖本 skill 自己的代码和文档；模板 `.docx` 文件本身**不在本仓库内分发**，其版权归清华大学所有，请遵守其原始版权要求。

## 🤝 贡献

欢迎 PR。约定：

1. 引入新的模板样式时，**先**改 `scripts/styles.py`；
2. `helpers.py` 里只引用 `styles.py` 的常量，绝对不要在 helper 里硬写中文样式名；
3. 加新 helper 时，往 `tests/test_smoke.py` 里追一条断言；
4. 同步更新 `SKILL.md` 里的样式速查表和 `references/` 里的决策树。
