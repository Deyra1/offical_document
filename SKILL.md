---
name: official-document
description: 按照 GB/T 9704—2012《党政机关公文格式》国家标准生成规范公文格式的文档。用于撰写对外正式文档材料（如版号申请材料、报告、函件、通知等），确保格式统一规范。触发词：公文、公文格式、正式文档、对外材料、版号申请、官方文件、规范格式文档、写公文、拟公文。
---

# 公文格式文档生成

## Overview

按照 GB/T 9704—2012《党政机关公文格式》国家标准，生成格式规范、排版统一的对外正式文档。适用于版号申请材料、正式报告、函件、通知等场景。

## 技术实现约束（必须遵守）

### 生成工具：python-docx（唯一选择）

**禁止使用 docx-js (Node.js)**。原因：中文引号（`""`）与 JavaScript 字符串分隔符冲突，会导致大量语法错误和调试浪费。

使用前检查依赖：
```bash
python3 -c "import docx" 2>/dev/null || pip3 install python-docx
```

### 字体安装（首次使用必须执行）

本 skill 自带标准公文字体，位于 `assets/fonts/` 目录：
- `仿宋_GB2312.ttf` — 正文用
- `方正小标宋简.TTF` — 文档大标题用

**每次执行前，必须先检查并安装字体：**

```bash
# 检查字体是否已安装
ls ~/Library/Fonts/仿宋_GB2312.ttf ~/Library/Fonts/方正小标宋简.TTF 2>/dev/null

# 如果缺失，从 skill 目录安装
SKILL_FONTS="$HOME/.config/piebox/skills/official-document/assets/fonts"
cp "$SKILL_FONTS/仿宋_GB2312.ttf" ~/Library/Fonts/ 2>/dev/null
cp "$SKILL_FONTS/方正小标宋简.TTF" ~/Library/Fonts/ 2>/dev/null
```

安装后无需重启，python-docx 直接引用字体名即可（Word 打开文档时会使用已安装的字体渲染）。

**字体注册名称（在 python-docx 中使用）：**

| 公文要素 | 字体文件 | python-docx 中的字体名 |
|---------|---------|----------------------|
| 正文 | 仿宋_GB2312.ttf | `FangSong_GB2312` |
| 文档大标题 | 方正小标宋简.TTF | `FZXiaoBiaoSong-B05S` |
| 一级标题（黑体） | 系统自带 | `STHeiti` |
| 页码（宋体） | 系统自带 | `Songti SC` |

**注意**：macOS 没有 `fc-list` 命令，不要使用。检查字体是否安装直接 `ls ~/Library/Fonts/` 即可。

### python-docx 固定参数速查表

```python
from docx import Document
from docx.shared import Pt, Mm, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH as WDA  # 注意：是 JUSTIFY 不是 JUSTIFIED
from docx.oxml.ns import qn

# 页面设置
PAGE_WIDTH = Mm(210)
PAGE_HEIGHT = Mm(297)
MARGIN_TOP = Mm(37)
MARGIN_BOTTOM = Mm(35)
MARGIN_LEFT = Mm(28)
MARGIN_RIGHT = Mm(26)

# 字号
PT_COVER_TITLE = Pt(24)   # 小一号 - 封面标题
PT_DOC_TITLE = Pt(22)     # 二号 - 正文大标题
PT_BODY = Pt(16)           # 三号 - 正文/各级小标题
PT_NOTE = Pt(14)           # 四号 - 页码/版记

# 行距与缩进
LINE_SPACING = Pt(28.95)  # 固定行距（22行/页）
LINE_SPACING_RULE = 4     # WD_LINE_SPACING.EXACTLY 的数值
FIRST_LINE_INDENT = Cm(0.85)  # 首行缩进2字符（三号字宽）

# 对齐方式
# WDA.JUSTIFY = 两端对齐（正文）
# WDA.CENTER = 居中（标题）
# WDA.LEFT = 左对齐
```

### 设置东亚字体的固定写法

```python
run = paragraph.add_run(text)
run.font.name = 'FangSong_GB2312'
run._element.rPr.rFonts.set(qn('w:eastAsia'), 'FangSong_GB2312')
run.font.size = Pt(16)
```

**必须同时设置 `font.name` 和 `w:eastAsia`**，否则中文字符不会应用指定字体。

### 中文引号处理

在 Python 字符串中，中文引号使用 Unicode 转义：
- 左双引号 `"` → `\u201c`
- 右双引号 `"` → `\u201d`
- 左单引号 `'` → `\u2018`
- 右单引号 `'` → `\u2019`

示例：`'主题项目\u201c疯狂派对\u201d'`

如果源文档中的引号是 ASCII `"`，转换时要替换为全角 `\u201c \u201d`。

## 操作流程

1. **确认依赖可用**：`python3 -c "import docx"` 
2. **读取源内容**：用 python-docx 或 textutil 读取用户提供的文件
3. **确定封面**：版号材料/产品说明 → 需要封面；一般公文 → 不需要
4. **生成文档**：用 python-docx 编写脚本，一次性执行
5. **验证输出**：检查文件存在且内容正确

## 公文格式规范（GB/T 9704—2012）

### 版面

- A4 (210mm × 297mm)
- 页边距：上37mm 下35mm 左28mm 右26mm
- 版心：156mm × 225mm
- 每面22行，每行28字

### 字体字号

| 要素 | 字体 | python-docx 字体名 | 字号 |
|------|------|-------------------|------|
| 正文 | 仿宋 | `FangSong_GB2312` | 三号(16pt) |
| 文档标题 | 方正小标宋 | `FZXiaoBiaoSong-B05S` | 二号(22pt) |
| 一级标题（一、） | 黑体 | `STHeiti` | 三号(16pt) |
| 二级标题（1.1） | 楷体/仿宋 | `FangSong_GB2312` | 三号(16pt) |
| 三/四级标题 | 仿宋 | `FangSong_GB2312` | 三号(16pt) |
| 页码 | 宋体 | `Songti SC` | 四号(14pt) |

### 正文编排

- 首行缩进2字符，回行顶格
- 行距：固定值28.95磅
- 对齐：两端对齐
- 文字颜色：黑色

### 结构层次序数

- 第一层：`一、` `二、` `三、`（黑体）
- 第二层：`（一）` `（二）`（楷体）或 `1.1` `1.2`（产品说明常用）
- 第三层：`1.` `2.` `3.`（仿宋）
- 第四层：`（1）` `（2）` `（3）`（仿宋）

### 页码

- 四号宋体阿拉伯数字
- 格式：`— 1 —`（数字左右各一条一字线）
- 封面不计页码，正文从1开始

### 成文日期

- 阿拉伯数字标全，不编虚位
- 示例：`2024年3月5日`（不写 `03月05日`）

## 封面规则

### 何时需要封面

- 版号申请材料/产品说明 → **需要封面**
- 一般公文（通知、函、报告） → **不需要封面**

### 封面布局（python-docx 实现）

```python
# 封面：8个空行 → 标题(居中) → 12个空行 → 日期(居中)
# 然后 doc.add_page_break()

# 封面标题
for _ in range(8): 
    add_empty_para(doc)  # 空行撑到 1/3 位置

add_para(doc, '产品名称说明', font='FZXiaoBiaoSong-B05S', size=Pt(24), 
         bold=True, align=WDA.CENTER, indent=False)

for _ in range(12): 
    add_empty_para(doc)  # 填充到底部

add_para(doc, '2026年5月', font='FangSong_GB2312', size=Pt(16),
         align=WDA.CENTER, indent=False)

doc.add_page_break()
```

### 封面标题命名（不加书名号）

- 游戏产品 → `xxx游戏产品说明`
- 软件产品 → `xxx软件说明`
- 日期：当前年月（如 `2026年5月`）

## 内容转换指引

当用户提供现有文本/文档进行格式转换时：

1. **读取源内容**：
   - `.docx` → `python-docx` 直接读取
   - `.doc` → `textutil -convert txt -stdout file.doc`
   - `.txt/.md` → 直接读取
2. **识别结构**：自动识别标题、正文段落、层级序号、表格
3. **格式映射**：
   - 源标题 → 二号黑体居中
   - 源正文 → 三号仿宋(Songti SC)，首行缩进
   - 源序号列表 → 按层次序数规范重排
   - 源表格 → 保留内容，字体统一为仿宋
   - 源文中 ASCII `"` 引号 → 替换为 `\u201c \u201d` 全角引号
4. **输出新文档**：`原文件名_公文格式.docx`，不修改源文件

## python-docx 代码模板

每次生成文档时，使用以下骨架代码（复制后填入具体内容）：

```python
from docx import Document
from docx.shared import Pt, Mm, Cm
from docx.enum.text import WD_ALIGN_PARAGRAPH as WDA
from docx.oxml.ns import qn

FONT_BODY = 'FangSong_GB2312'
FONT_HEITI = 'STHeiti'
FONT_TITLE = 'FZXiaoBiaoSong-B05S'
FONT_SONG = 'Songti SC'
PT_BODY = Pt(16)
PT_TITLE = Pt(22)
PT_COVER = Pt(24)
LINE_SP = Pt(28.95)
INDENT = Cm(0.85)


def ap(doc, text, fn=FONT_BODY, sz=PT_BODY, b=False, al=WDA.JUSTIFY, ind=True):
    """添加一个段落"""
    p = doc.add_paragraph()
    p.alignment = al
    pf = p.paragraph_format
    pf.line_spacing = LINE_SP
    pf.line_spacing_rule = 4  # EXACTLY
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    if ind:
        pf.first_line_indent = INDENT
    run = p.add_run(text)
    run.font.name = fn
    run._element.rPr.rFonts.set(qn('w:eastAsia'), fn)
    run.font.size = sz
    run.font.bold = b
    return p


def h1(d, t): return ap(d, t, fn=FONT_HEITI, b=True, ind=False)
def h2(d, t): return ap(d, t, ind=True)
def h3(d, t): return ap(d, t, ind=True)
def bd(d, t, b=False): return ap(d, t, b=b, ind=True)
def em(d): return ap(d, '', ind=False)


def sc(cell, text, fn=FONT_BODY, sz=PT_BODY, b=False, al=WDA.LEFT):
    """设置表格单元格"""
    cell.text = ''
    p = cell.paragraphs[0]
    p.alignment = al
    run = p.add_run(text)
    run.font.name = fn
    run._element.rPr.rFonts.set(qn('w:eastAsia'), fn)
    run.font.size = sz
    run.font.bold = b


doc = Document()
for s in doc.sections:
    s.page_width = Mm(210)
    s.page_height = Mm(297)
    s.top_margin = Mm(37)
    s.bottom_margin = Mm(35)
    s.left_margin = Mm(28)
    s.right_margin = Mm(26)

if doc.paragraphs:
    doc.paragraphs[0]._element.getparent().remove(doc.paragraphs[0]._element)

# === 封面（版号材料需要）===
for _ in range(8): em(doc)
ap(doc, '文档标题', fn=FONT_TITLE, sz=PT_COVER, b=True, al=WDA.CENTER, ind=False)
for _ in range(12): em(doc)
ap(doc, '2026年5月', al=WDA.CENTER, ind=False)
doc.add_page_break()

# === 正文 ===
ap(doc, '文档标题', fn=FONT_TITLE, sz=PT_TITLE, b=True, al=WDA.CENTER, ind=False)
em(doc)

h1(doc, '一、章节标题')
h2(doc, '1.1  小节标题')
bd(doc, '正文内容...')

# === 保存 ===
doc.save('/path/to/output.docx')
```

## 简化应用说明

版号申请材料等对外文档核心格式清单：

- A4 纸张，页边距 37/35/28/26 mm
- 封面页：产品名称（不加书名号）+ 年月
- 标题：二号方正小标宋(FZXiaoBiaoSong-B05S)居中
- 正文：三号仿宋(FangSong_GB2312)，首行缩进，固定行距28.95磅
- 结构层次：一、→ 1.1 → 1.1.1（产品说明常用数字序号）
- 页码：`— 1 —` 格式，封面不计
