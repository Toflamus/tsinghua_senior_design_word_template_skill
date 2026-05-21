"""Style-name constants for the 清华综合论文训练 Word template.

Single source of truth. If the template's style names change (new template
version), update only this file and helpers will pick it up.
"""

# --- Headings ---
CHAPTER_TITLE = "章标题-无级别"      # 章标题 (also used for 摘要/Abstract/参考文献/插图清单 etc.)
HEADING_1 = "Heading 1"            # 引言 / 第 X 章 (when numbered)
HEADING_2 = "Heading 2"            # 1.1
HEADING_3 = "Heading 3"            # 1.1.1
HEADING_4 = "Heading 4"            # 1.1.1.1
TITLE = "Title"                    # 致谢/声明/在学期间研究成果

# --- Body ---
BODY = "论文正文段落"              # 推荐的正文样式
BODY_LEGACY = "段落"               # 模板里摘要部分使用的样式

# --- Floats ---
FIGURE_PARAGRAPH = "图片"          # 图片所在段
FIG_CAPTION = "Caption"            # 图题
TABLE_CAPTION = "表-题注"          # 表题（base = Caption）
THREE_LINE_TABLE = "三线表"        # 表格样式

# --- Equations ---
EQUATION = "公式"

# --- Code ---
INLINE_CODE = "行内代码"           # 字符样式
CODE_BLOCK = "行间代码"            # 段落样式

# --- References ---
REFERENCE = "参考文献"

# --- Appendix ---
APPENDIX_HEADING_0 = "附录标题"
APPENDIX_HEADING_1 = "附录标题 1"
APPENDIX_HEADING_2 = "附录标题 2"
APPENDIX_HEADING_3 = "附录标题 3"

# --- Front matter ---
SYMBOLS_TABLE = "符号和缩略语说明表"
LISTING_HEADING = "标题-插图表格清单&符号缩略语说明"
TOF_ENTRY = "table of figures"     # also used by 附表清单

# --- Cover ---
COVER_TSINGHUA = "封面清华大学"
COVER_BANNER = "封面综合论文训练"
COVER_TITLE = "封面论文题目"
COVER_AUTHOR_INFO = "封面作者信息"

# --- TOC ---
TOC_HEADING = "TOC Heading"
TOC_1 = "toc 1"
TOC_2 = "toc 2"
TOC_3 = "toc 3"

# --- Misc / normal ---
NORMAL = "Normal"
LIST_PARAGRAPH = "List Paragraph"

# --- Critical style names that must exist in any compatible template ---
REQUIRED_STYLES = (
    CHAPTER_TITLE,
    HEADING_1, HEADING_2, HEADING_3,
    BODY, BODY_LEGACY,
    FIG_CAPTION, TABLE_CAPTION, THREE_LINE_TABLE,
    EQUATION,
    INLINE_CODE, CODE_BLOCK,
    REFERENCE,
    APPENDIX_HEADING_0,
    SYMBOLS_TABLE,
    COVER_TITLE, COVER_AUTHOR_INFO,
)
