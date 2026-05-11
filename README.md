# 党政机关公文格式文档生成工具

基于 GB/T 9704-2012《党政机关公文格式》国家标准，自动生成规范格式的 Word 文档（.docx）。

## 标题层级格式

| 层级 | 编号格式 | 字体 | 示例 |
|------|---------|------|------|
| Heading 1 | 一、标题 | 3号黑体 | 一、总体要求 |
| Heading 2 | 1.1 标题 | 3号楷体 | 1.1 基本原则 |
| Heading 3 | 1.1.1 标题 | 3号仿宋体 | 1.1.1 具体措施 |
| Heading 4 | （1）标题 | 3号仿宋体 | （1）第一项 |

- Heading 2 的第一个数字跟随 Heading 1 序号（如第二章下为 2.1、2.2）
- 各级编号自动管理，上级标题递增时自动重置下级编号

## 安装

```bash
npm install
```

## 快速使用

### 方式一：快捷函数

```javascript
import { createDocument } from "./src/index.js";

const builder = createDocument({
  title: "关于做好XX工作的通知",
  content: [
    { type: "recipient", text: "各省、自治区、直辖市人民政府" },
    { type: "paragraph", text: "为贯彻落实..." },
    { type: "h1", text: "总体要求" },
    { type: "paragraph", text: "坚持以..." },
    { type: "h2", text: "基本原则" },
    { type: "h3", text: "具体措施" },
    { type: "h4", text: "第一项" },
  ],
});

await builder.save("output.docx");
```

### 方式二：Builder API

```javascript
import { OfficialDocumentBuilder } from "./src/index.js";

const doc = new OfficialDocumentBuilder();

doc
  .addTitle("关于印发《XX办法》的通知")
  .addRecipient("各有关单位")
  .addParagraph("正文内容...")
  .addHeading1("适用范围")
  .addParagraph("本办法适用于...")
  .addHeading2("具体规定")
  .addHeading3("细则")
  .addHeading4("条款");

await doc.save("output.docx");
```

## API

### `createDocument(config)`

快捷创建函数，支持的 content type：

| type | 说明 |
|------|------|
| `recipient` | 主送机关 |
| `h1` / `heading1` | 一级标题 |
| `h2` / `heading2` | 二级标题 |
| `h3` / `heading3` | 三级标题 |
| `h4` / `heading4` | 四级标题 |
| `paragraph` / `p` | 正文段落 |
| `blank` | 空行（可设 count） |
| `pageBreak` | 分页符 |

### `OfficialDocumentBuilder`

| 方法 | 说明 |
|------|------|
| `addTitle(text)` | 添加公文标题（2号小标宋体，居中） |
| `addRecipient(text)` | 添加主送机关（顶格） |
| `addHeading1(text)` | 一级标题（自动编号） |
| `addHeading2(text)` | 二级标题（自动编号） |
| `addHeading3(text)` | 三级标题（自动编号） |
| `addHeading4(text)` | 四级标题（自动编号） |
| `addParagraph(text)` | 正文段落 |
| `addBlankLine(n)` | 添加空行 |
| `addPageBreak()` | 分页 |
| `build()` | 返回 Document 对象 |
| `save(path)` | 保存为 .docx 文件 |

## 运行示例

```bash
npm run example
```

生成的文档在 `output/` 目录下。

## 测试

```bash
npm test
```

## 格式规范依据

- 页面：A4 (210mm × 297mm)
- 天头（上边距）：37mm
- 订口（左边距）：28mm
- 版心：156mm × 225mm
- 正文：3号仿宋体（16pt），每面22行，每行28字
- 行距：固定值 29pt

## License

MIT
