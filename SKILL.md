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
| 一级标题（黑体） | 系统/Windows 自带 | `SimHei` |
| 二级标题（楷体） | 系统/Windows 自带 | `KaiTi` |
| 页码（宋体） | 系统自带 | `SimSun` |

**注意**：macOS 没有 `fc-list` 命令，不要使用。检查字体是否安装直接 `ls ~/Library/Fonts/` 即可。

**关于字体名选择**：python-docx 中写入的字体名是 Word 文档的声明，不要求本机安装该字体。应使用 Windows 标准字体名（`SimHei`、`KaiTi`、`SimSun`），而非 macOS 特有字体名（如 `STHeiti`、`STKaiti`、`Songti SC`），因为最终文档是交给对方在 Windows 环境查看。

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

### 中文字符串禁止用 Unicode 转义

**直接写中文字符**，不要将中文内容用 `\uXXXX` 形式编码。原因：容易写错码点（如"疯"写成"疌"），且不可读、不可审查。

```python
# ✅ 正确
title = '疯狂派对游戏产品说明'

# ❌ 错误 - 容易写错且不可读
title = '\u75af\u72c2\u6d3e\u5bf9...'
```

仅引号等特殊字符允许使用 Unicode 转义（因为它们与 Python 字符串语法冲突）。

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
| 一级标题（一、） | 黑体 | `SimHei` | 三号(16pt) |
| 二级标题（1.1） | 楷体 | `KaiTi` | 三号(16pt) |
| 三/四级标题 | 仿宋 | `FangSong_GB2312` | 三号(16pt) |
| 页码 | 宋体 | `SimSun` | 四号(14pt) |

### 正文编排

- 首行缩进2字符，回行顶格
- 行距：固定值28.95磅
- 对齐：两端对齐
- 文字颜色：黑色

### 结构层次序数

- 第一层：`一、` `二、` `三、`（黑体 SimHei）
- 第二层：`（一）` `（二）`（楷体 KaiTi）或 `1.1` `1.2`（产品说明常用）
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
3. **提取图片**：从源 .docx 中提取所有图片，在新文档中按原位置插入
4. **格式映射**：
   - 源标题 → 按层级设置对应格式（见下方标题层级规则）
   - 源正文 → 三号仿宋(FangSong_GB2312)，首行缩进
   - 源序号列表 → 按层次序数规范重排
   - 源表格 → 保留内容，按公文表格规范重排
   - 源文中 ASCII `"` 引号 → 替换为 `\u201c \u201d` 全角引号
   - 源图片 → 保持原始尺寸，居中排布，图片前后各保留一个空段落
5. **输出新文档**：`原文件名_公文格式.docx`，不修改源文件

## 标题层级规则（必须严格遵守）

各级标题必须有明确的视觉层级区分：

| 层级 | 格式示例 | 字体 | 字号 | 加粗 | 缩进 | Word 样式 |
|------|---------|------|------|------|------|----------|
| 文档大标题 | 疯狂派对游戏产品说明 | FZXiaoBiaoSong-B05S | 二号(22pt) | 是 | 居中，无缩进 | 正文/无 |
| 一级标题 | 一、背景与设定 | SimHei | 三号(16pt) | 是 | 顶格（无缩进） | Heading 1 |
| 二级标题 | 1.1  游戏背景 | KaiTi | 三号(16pt) | 是 | 首行缩进2字符 | Heading 2 |
| 三级标题 | 1.1.1  好人阵营 | FangSong_GB2312 | 三号(16pt) | 是 | 首行缩进2字符 | Heading 3 |
| 正文 | 普通段落文字 | FangSong_GB2312 | 三号(16pt) | 否 | 首行缩进2字符 | 正文 |

### 必须设置 Word 内置 Heading 样式

**标题段落必须使用 `doc.add_paragraph(style='Heading N')` 创建**，而不是普通段落。原因：只有使用 Word 内置 Heading 样式，才能在 Word 中通过"引用 → 目录"自动生成目录。

设置 Heading 样式后，仍然需要手动覆盖字体、字号、行距、缩进、**颜色**等属性（因为 Word 默认的 Heading 样式是蓝色字体，不符合公文规范）。

**必须显式设置 `run.font.color.rgb = RGBColor(0, 0, 0)`**，否则标题会继承 Heading 样式的蓝色。`run.font.color.rgb = None` 表示"不设置"（继承样式颜色），不等于黑色。

**关键区别**：
- 一级标题用 **黑体(SimHei) + 加粗**，顶格
- 二级标题用 **楷体(KaiTi) + 加粗**，首行缩进
- 三级标题用 **仿宋加粗**，与正文区分仅靠加粗
- 正文用 **仿宋不加粗**

## 图片处理规则

转换文档时，**必须保留原文档中的所有图片**：

1. **提取图片**：从源 .docx 的 `word/media/` 中提取所有图片文件
2. **插入位置**：按原文档中图片出现的段落位置，在对应文字之后插入
3. **注意空表格中的图片**：原文档可能使用空表格来并排展示图片，解析时必须检查表格元素中的 `a:blip` 引用，将其中的图片也提取并按顺序插入
4. **图片格式**：
   - 居中对齐
   - 保持原始宽高比
   - 宽度不超过版心宽度（156mm），超出则等比缩小
   - 图片段落无首行缩进
4. **python-docx 插入图片代码**：

```python
from docx.shared import Mm
from docx.enum.text import WD_ALIGN_PARAGRAPH as WDA

def add_image(doc, image_path, max_width_mm=150):
    """插入图片，居中，限制最大宽度"""
    from PIL import Image
    img = Image.open(image_path)
    w, h = img.size
    # 按最大宽度等比缩放
    width = Mm(max_width_mm)
    
    p = doc.add_paragraph()
    p.alignment = WDA.CENTER
    p.paragraph_format.first_line_indent = None
    p.paragraph_format.space_before = Pt(6)
    p.paragraph_format.space_after = Pt(6)
    run = p.add_run()
    run.add_picture(image_path, width=width)
    return p
```

5. **如果无法安装 PIL/Pillow**，直接设置固定宽度 `Mm(150)` 即可

## 表格格式规范

公文中的表格样式要求：

1. **边框**：单线边框（`single`），线宽 4-6（约 0.5pt），颜色黑色
2. **表头**：
   - 字体：仿宋(FangSong_GB2312)，加粗
   - 对齐：水平居中 + 垂直居中
   - 无底色（公文表格不加灰色底纹）
3. **正文单元格**：
   - 字体：仿宋(FangSong_GB2312)，不加粗
   - 对齐：水平居中 + 垂直居中
4. **表格宽度**：撑满版心宽度
5. **行距**：单倍行距（不同于正文的固定行距28.95磅）
6. **单元格内边距**：上下 60 twips，左右 108 twips

**python-docx 表格代码模板**：

```python
from docx.oxml.ns import qn, nsdecls
from docx.oxml import parse_xml
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT

def set_table_border(table):
    """设置表格为公文标准单线边框"""
    tbl = table._tbl
    tblPr = tbl.find(qn('w:tblPr'))
    if tblPr is None:
        tblPr = parse_xml(f'<w:tblPr {nsdecls("w")}></w:tblPr>')
        tbl.insert(0, tblPr)
    borders = parse_xml(
        f'<w:tblBorders {nsdecls("w")}>'
        '<w:top w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:left w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:bottom w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:right w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:insideH w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '<w:insideV w:val="single" w:sz="4" w:space="0" w:color="000000"/>'
        '</w:tblBorders>'
    )
    tblPr.append(borders)


def sc(cell, text, fn=FONT_BODY, sz=PT_BODY, b=False):
    """设置表格单元格：水平居中 + 垂直居中 + 单倍行距"""
    from docx.oxml import OxmlElement
    # 必须清空并重建段落，否则 run 属性可能不生效
    for p in cell.paragraphs:
        p._element.getparent().remove(p._element)
    new_p = OxmlElement('w:p')
    cell._tc.append(new_p)
    from docx.text.paragraph import Paragraph
    p = Paragraph(new_p, cell)
    
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    p.alignment = WDA.CENTER
    p.paragraph_format.line_spacing = None  # 单倍行距
    p.paragraph_format.first_line_indent = None
    p.paragraph_format.space_before = Pt(0)
    p.paragraph_format.space_after = Pt(0)
    
    run = p.add_run(text)
    run.font.name = fn
    rPr = run._element.get_or_add_rPr()
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:eastAsia'), fn)
    rFonts.set(qn('w:ascii'), fn)
    rFonts.set(qn('w:hAnsi'), fn)
    rPr.append(rFonts)
    run.font.size = sz
    run.font.bold = b


# 使用示例
table = doc.add_table(rows=6, cols=3)
set_table_border(table)

# 设置表头（加粗）
for i, text in enumerate(['角色名称', '主动技能', '角色设定与胜利条件']):
    sc(table.rows[0].cells[i], text, b=True)

# 设置数据行（不加粗）
sc(table.rows[1].cells[0], '魔术师')
sc(table.rows[1].cells[2], '描述文字...')
```

**表格禁止事项**：
- ❌ 不要加灰色底纹/背景色（ShadingType）
- ❌ 不要用粗边框（sz 不超过 6）
- ❌ 不要用彩色边框
- ❌ 不要用固定行距28.95磅（表格内用单倍行距）

## python-docx 代码模板

每次生成文档时，使用以下骨架代码（复制后填入具体内容）：

```python
from docx import Document
from docx.shared import Pt, Mm, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH as WDA
from docx.oxml.ns import qn

FONT_BODY = 'FangSong_GB2312'
FONT_HEITI = 'SimHei'
FONT_KAITI = 'KaiTi'
FONT_TITLE = 'FZXiaoBiaoSong-B05S'
FONT_SONG = 'SimSun'
PT_BODY = Pt(16)
PT_TITLE = Pt(22)
PT_COVER = Pt(24)
LINE_SP = Pt(28.95)
INDENT = Cm(0.85)


def ap(doc, text, fn=FONT_BODY, sz=PT_BODY, b=False, al=WDA.JUSTIFY, ind=True, style=None):
    """添加一个段落"""
    p = doc.add_paragraph(style=style) if style else doc.add_paragraph()
    p.alignment = al
    pf = p.paragraph_format
    pf.line_spacing = LINE_SP
    pf.line_spacing_rule = 4  # EXACTLY
    pf.space_before = Pt(0)
    pf.space_after = Pt(0)
    if ind:
        pf.first_line_indent = INDENT
    else:
        pf.first_line_indent = None
    run = p.add_run(text)
    run.font.name = fn
    run._element.rPr.rFonts.set(qn('w:eastAsia'), fn)
    run.font.size = sz
    run.font.bold = b
    run.font.color.rgb = RGBColor(0, 0, 0)  # Heading 样式默认蓝色，必须显式设为黑色
    return p


def h1(d, t): return ap(d, t, fn=FONT_HEITI, b=True, ind=False, style='Heading 1')   # 一、顶格黑体加粗
def h2(d, t): return ap(d, t, fn=FONT_KAITI, b=True, ind=True, style='Heading 2')   # 1.1 缩进楷体加粗
def h3(d, t): return ap(d, t, fn=FONT_BODY, b=True, ind=True, style='Heading 3')    # 1.1.1 缩进仿宋加粗
def bd(d, t, b=False): return ap(d, t, b=b, ind=True)                  # 正文
def em(d): return ap(d, '', ind=False)                                  # 空行


def sc(cell, text, fn=FONT_BODY, sz=PT_BODY, b=False, al=WDA.CENTER):
    """设置表格单元格（正确写法：重建段落确保 run 属性生效）"""
    from docx.oxml import OxmlElement
    from docx.text.paragraph import Paragraph
    for p in cell.paragraphs:
        p._element.getparent().remove(p._element)
    new_p = OxmlElement('w:p')
    cell._tc.append(new_p)
    p = Paragraph(new_p, cell)
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    p.alignment = al
    run = p.add_run(text)
    run.font.name = fn
    rPr = run._element.get_or_add_rPr()
    rFonts = OxmlElement('w:rFonts')
    rFonts.set(qn('w:eastAsia'), fn)
    rFonts.set(qn('w:ascii'), fn)
    rPr.append(rFonts)
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

## 转换时必须检查的事项

1. **图片是否保留** — 必须从源文档提取所有图片（包括空表格中的并排图片）并插入新文档
2. **标题层级是否区分** — 一级标题用 SimHei（黑体），二级标题用 KaiTi（楷体），三级标题用仿宋加粗，正文仿宋不加粗
3. **标题是否设置 Heading 样式** — 必须用 Heading 1/2/3 样式，否则无法自动生成目录
4. **表格是否规范** — 单线细边框、无底色、表头加粗居中、单倍行距
5. **字体名是否正确** — 使用 Windows 标准字体名（SimHei/KaiTi/SimSun），不要用 macOS 字体名（STHeiti/STKaiti/Songti SC）
6. **中文内容是否直接写** — 中文字符串直接写中文，不要用 `\uXXXX` 转义（仅引号等特殊字符例外）
