/**
 * 党政机关公文格式文档生成工具
 * 遵循 GB/T 9704-2012《党政机关公文格式》国家标准
 *
 * 标题层级格式：
 *   Heading 1: 一、标题内容 （黑体）
 *   Heading 2: 1.1 标题内容 （楷体）
 *   Heading 3: 1.1.1 标题内容 （仿宋体）
 *   Heading 4: （1）标题内容 （仿宋体）
 */

import {
  Document,
  Packer,
  Paragraph,
  TextRun,
  AlignmentType,
  HeadingLevel,
  Header,
  Footer,
  PageNumber,
  PageBreak,
  Table,
  TableRow,
  TableCell,
  BorderStyle,
  WidthType,
  LevelFormat,
  TabStopType,
  TabStopPosition,
} from "docx";
import fs from "fs";

// ============================================================
// 常量定义 - GB/T 9704-2012 标准参数
// ============================================================

/** 中文数字映射 */
const CHINESE_NUMBERS = [
  "一", "二", "三", "四", "五", "六", "七", "八", "九", "十",
  "十一", "十二", "十三", "十四", "十五", "十六", "十七", "十八", "十九", "二十",
];

/** 页面尺寸 - A4 (DXA单位, 1mm ≈ 56.7 DXA) */
const PAGE = {
  width: 11906,   // 210mm
  height: 16838,  // 297mm
};

/**
 * 页边距 - GB/T 9704-2012 规定
 * 天头(上): 37mm, 版心高225mm, 下边距: 297-37-225 = 35mm
 * 订口(左): 28mm, 版心宽156mm, 右边距: 210-28-156 = 26mm
 */
const MARGIN = {
  top: 2098,    // 37mm
  bottom: 1984, // 35mm
  left: 1588,   // 28mm
  right: 1474,  // 26mm
};

/** 字号对照表 (单位: half-point, 即 Word 中的字号×2) */
const FONT_SIZE = {
  "小标宋_2号": 44,  // 2号 = 22pt → 44 half-points
  "黑体_3号": 32,    // 3号 = 16pt → 32 half-points
  "楷体_3号": 32,
  "仿宋_3号": 32,
  "仿宋_4号": 28,    // 4号 = 14pt → 28 half-points
  "宋体_4号半角": 28,
};

/** 字体名称 */
const FONT = {
  fangsong: "仿宋",
  heiti: "黑体",
  kaiti: "楷体",
  xiaobiaosong: "方正小标宋简体",
  songti: "宋体",
};

/** 行距 - 每面22行，版心高225mm，行距约 225/22 ≈ 10.23mm ≈ 29pt (固定值) */
const LINE_SPACING = {
  fixed: 580, // 29pt × 20 twips = 580 (DXA中行距单位)
};

// ============================================================
// 编号管理器 - 自动管理各级标题编号
// ============================================================

class NumberingManager {
  constructor() {
    this.counters = {
      h1: 0,  // 一、二、三…
      h2: 0,  // x.1, x.2…
      h3: 0,  // x.x.1, x.x.2…
      h4: 0,  // （1）（2）…
    };
  }

  /** 重置所有计数器 */
  reset() {
    this.counters = { h1: 0, h2: 0, h3: 0, h4: 0 };
  }

  /** 获取 Heading1 编号: 一、二、三… */
  nextH1(title) {
    this.counters.h1++;
    this.counters.h2 = 0;
    this.counters.h3 = 0;
    this.counters.h4 = 0;
    const num = CHINESE_NUMBERS[this.counters.h1 - 1] || this.counters.h1.toString();
    return `${num}、${title}`;
  }

  /** 获取 Heading2 编号: 1.1, 1.2, 2.1… */
  nextH2(title) {
    this.counters.h2++;
    this.counters.h3 = 0;
    this.counters.h4 = 0;
    return `${this.counters.h1}.${this.counters.h2} ${title}`;
  }

  /** 获取 Heading3 编号: 1.1.1, 1.1.2… */
  nextH3(title) {
    this.counters.h3++;
    this.counters.h4 = 0;
    return `${this.counters.h1}.${this.counters.h2}.${this.counters.h3} ${title}`;
  }

  /** 获取 Heading4 编号: （1）（2）… */
  nextH4(title) {
    this.counters.h4++;
    return `（${this.counters.h4}）${title}`;
  }
}

// ============================================================
// 文档构建器
// ============================================================

class OfficialDocumentBuilder {
  constructor(options = {}) {
    this.numbering = new NumberingManager();
    this.children = [];
    this.options = {
      title: options.title || "",
      issuer: options.issuer || "",       // 发文机关
      docNumber: options.docNumber || "", // 发文字号
      recipient: options.recipient || "", // 主送机关
      showPageNumber: options.showPageNumber !== false,
      ...options,
    };
  }

  /**
   * 添加公文标题（2号小标宋体，居中）
   */
  addTitle(title) {
    this.options.title = title;
    this.children.push(
      new Paragraph({
        alignment: AlignmentType.CENTER,
        spacing: { before: 0, after: 0, line: LINE_SPACING.fixed, lineRule: "exact" },
        children: [
          new TextRun({
            text: title,
            font: { eastAsia: FONT.xiaobiaosong },
            size: FONT_SIZE["小标宋_2号"],
            bold: true,
          }),
        ],
      })
    );
    return this;
  }

  /**
   * 添加一级标题 - 格式: 一、标题内容（3号黑体）
   */
  addHeading1(title) {
    const text = this.numbering.nextH1(title);
    this.children.push(
      new Paragraph({
        spacing: { before: 0, after: 0, line: LINE_SPACING.fixed, lineRule: "exact" },
        indent: { firstLine: 640 }, // 首行缩进2字符
        children: [
          new TextRun({
            text: text,
            font: { eastAsia: FONT.heiti },
            size: FONT_SIZE["黑体_3号"],
            bold: true,
          }),
        ],
      })
    );
    return this;
  }

  /**
   * 添加二级标题 - 格式: 1.1 标题内容（3号楷体）
   */
  addHeading2(title) {
    const text = this.numbering.nextH2(title);
    this.children.push(
      new Paragraph({
        spacing: { before: 0, after: 0, line: LINE_SPACING.fixed, lineRule: "exact" },
        indent: { firstLine: 640 },
        children: [
          new TextRun({
            text: text,
            font: { eastAsia: FONT.kaiti },
            size: FONT_SIZE["楷体_3号"],
          }),
        ],
      })
    );
    return this;
  }

  /**
   * 添加三级标题 - 格式: 1.1.1 标题内容（3号仿宋体）
   */
  addHeading3(title) {
    const text = this.numbering.nextH3(title);
    this.children.push(
      new Paragraph({
        spacing: { before: 0, after: 0, line: LINE_SPACING.fixed, lineRule: "exact" },
        indent: { firstLine: 640 },
        children: [
          new TextRun({
            text: text,
            font: { eastAsia: FONT.fangsong },
            size: FONT_SIZE["仿宋_3号"],
          }),
        ],
      })
    );
    return this;
  }

  /**
   * 添加四级标题 - 格式: （1）标题内容（3号仿宋体）
   */
  addHeading4(title) {
    const text = this.numbering.nextH4(title);
    this.children.push(
      new Paragraph({
        spacing: { before: 0, after: 0, line: LINE_SPACING.fixed, lineRule: "exact" },
        indent: { firstLine: 640 },
        children: [
          new TextRun({
            text: text,
            font: { eastAsia: FONT.fangsong },
            size: FONT_SIZE["仿宋_3号"],
          }),
        ],
      })
    );
    return this;
  }

  /**
   * 添加正文段落（3号仿宋体，首行缩进2字符）
   */
  addParagraph(text) {
    this.children.push(
      new Paragraph({
        spacing: { before: 0, after: 0, line: LINE_SPACING.fixed, lineRule: "exact" },
        indent: { firstLine: 640 }, // 2字符 = 2×16pt×20 = 640 twips
        children: [
          new TextRun({
            text: text,
            font: { eastAsia: FONT.fangsong },
            size: FONT_SIZE["仿宋_3号"],
          }),
        ],
      })
    );
    return this;
  }

  /**
   * 添加主送机关（顶格，3号仿宋体）
   */
  addRecipient(recipient) {
    this.options.recipient = recipient;
    this.children.push(
      new Paragraph({
        spacing: { before: 0, after: 0, line: LINE_SPACING.fixed, lineRule: "exact" },
        children: [
          new TextRun({
            text: recipient + "：",
            font: { eastAsia: FONT.fangsong },
            size: FONT_SIZE["仿宋_3号"],
          }),
        ],
      })
    );
    return this;
  }

  /**
   * 添加空行
   */
  addBlankLine(count = 1) {
    for (let i = 0; i < count; i++) {
      this.children.push(
        new Paragraph({
          spacing: { before: 0, after: 0, line: LINE_SPACING.fixed, lineRule: "exact" },
          children: [],
        })
      );
    }
    return this;
  }

  /**
   * 添加分页符
   */
  addPageBreak() {
    this.children.push(new Paragraph({ children: [new PageBreak()] }));
    return this;
  }

  /**
   * 构建文档对象
   */
  build() {
    const doc = new Document({
      styles: {
        default: {
          document: {
            run: {
              font: { eastAsia: FONT.fangsong },
              size: FONT_SIZE["仿宋_3号"],
            },
          },
        },
      },
      sections: [
        {
          properties: {
            page: {
              size: {
                width: PAGE.width,
                height: PAGE.height,
              },
              margin: {
                top: MARGIN.top,
                bottom: MARGIN.bottom,
                left: MARGIN.left,
                right: MARGIN.right,
              },
            },
          },
          headers: this.options.showPageNumber
            ? undefined
            : undefined,
          footers: this.options.showPageNumber
            ? {
                default: new Footer({
                  children: [
                    new Paragraph({
                      alignment: AlignmentType.CENTER,
                      children: [
                        new TextRun({
                          children: [PageNumber.CURRENT],
                          font: { eastAsia: FONT.songti },
                          size: FONT_SIZE["宋体_4号半角"],
                        }),
                      ],
                    }),
                  ],
                }),
              }
            : undefined,
          children: this.children,
        },
      ],
    });
    return doc;
  }

  /**
   * 构建并保存为文件
   */
  async save(filePath) {
    const doc = this.build();
    const buffer = await Packer.toBuffer(doc);
    fs.writeFileSync(filePath, buffer);
    return filePath;
  }
}

// ============================================================
// 快捷函数
// ============================================================

/**
 * 创建公文文档
 * @param {Object} config - 文档配置
 * @param {string} config.title - 公文标题
 * @param {Array} config.content - 内容数组，每项格式为 { type, text }
 * @param {Object} config.options - 其他选项
 * @returns {OfficialDocumentBuilder}
 *
 * @example
 * const doc = createDocument({
 *   title: "关于做好XX工作的通知",
 *   content: [
 *     { type: "recipient", text: "各省、自治区、直辖市人民政府" },
 *     { type: "paragraph", text: "为贯彻落实..." },
 *     { type: "h1", text: "总体要求" },
 *     { type: "paragraph", text: "坚持以..." },
 *     { type: "h2", text: "基本原则" },
 *     { type: "h3", text: "具体措施" },
 *     { type: "h4", text: "第一项措施" },
 *   ]
 * });
 */
function createDocument(config = {}) {
  const builder = new OfficialDocumentBuilder(config.options || {});

  if (config.title) {
    builder.addTitle(config.title);
    builder.addBlankLine(1);
  }

  if (config.content && Array.isArray(config.content)) {
    for (const item of config.content) {
      switch (item.type) {
        case "recipient":
          builder.addRecipient(item.text);
          break;
        case "h1":
        case "heading1":
          builder.addHeading1(item.text);
          break;
        case "h2":
        case "heading2":
          builder.addHeading2(item.text);
          break;
        case "h3":
        case "heading3":
          builder.addHeading3(item.text);
          break;
        case "h4":
        case "heading4":
          builder.addHeading4(item.text);
          break;
        case "paragraph":
        case "p":
          builder.addParagraph(item.text);
          break;
        case "blank":
          builder.addBlankLine(item.count || 1);
          break;
        case "pageBreak":
          builder.addPageBreak();
          break;
        default:
          builder.addParagraph(item.text || "");
      }
    }
  }

  return builder;
}

export { createDocument, OfficialDocumentBuilder, NumberingManager };
