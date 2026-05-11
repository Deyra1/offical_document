/**
 * 示例：生成一份公文格式文档
 * 运行方式: node examples/generate-example.js
 */

import { createDocument, OfficialDocumentBuilder } from "../src/index.js";

// ============================================================
// 方式一：使用快捷函数 createDocument
// ============================================================
async function example1() {
  const builder = createDocument({
    title: "关于进一步加强公文规范化管理工作的通知",
    content: [
      { type: "recipient", text: "各省、自治区、直辖市人民政府，国务院各部委、各直属机构" },
      { type: "paragraph", text: "为深入贯彻落实党中央、国务院关于推进公文规范化管理的有关要求，进一步提高公文质量和效率，现就有关事项通知如下。" },

      { type: "h1", text: "总体要求" },
      { type: "paragraph", text: "坚持以习近平新时代中国特色社会主义思想为指导，全面贯彻党的二十大精神，认真落实公文处理工作各项制度规定，不断提升公文规范化水平。" },

      { type: "h1", text: "主要任务" },
      { type: "h2", text: "规范公文格式" },
      { type: "paragraph", text: "各级各部门应当严格按照GB/T 9704-2012《党政机关公文格式》国家标准制发公文，确保格式统一规范。" },
      { type: "h3", text: "标题格式要求" },
      { type: "paragraph", text: "公文标题应当准确简要概括公文的主要内容并标明公文种类，一般用2号小标宋体字，编排于红色分隔线下空二行位置。" },
      { type: "h3", text: "正文格式要求" },
      { type: "paragraph", text: "正文一般用3号仿宋体字，编排于主送机关名称下一行，每个自然段左空二字，回行顶格。" },
      { type: "h4", text: "一级结构层次用黑体字" },
      { type: "paragraph", text: "文中一级结构层次序数用中文数字加顿号标注，采用黑体字排印。" },
      { type: "h4", text: "二级结构层次用楷体字" },
      { type: "paragraph", text: "文中二级结构层次序数用阿拉伯数字加点标注，采用楷体字排印。" },

      { type: "h2", text: "加强审核把关" },
      { type: "paragraph", text: "健全公文审核机制，严格执行公文处理程序，确保公文内容准确、格式规范、文字精炼。" },

      { type: "h1", text: "组织保障" },
      { type: "h2", text: "加强组织领导" },
      { type: "paragraph", text: "各地区各部门要高度重视公文规范化管理工作，切实加强组织领导，明确责任分工。" },
      { type: "h2", text: "强化培训指导" },
      { type: "paragraph", text: "定期开展公文处理业务培训，提高工作人员公文写作和处理能力。" },
    ],
  });

  await builder.save("output/example-quickstart.docx");
  console.log("✓ 方式一完成: output/example-quickstart.docx");
}

// ============================================================
// 方式二：使用 Builder API 逐步构建
// ============================================================
async function example2() {
  const doc = new OfficialDocumentBuilder();

  doc
    .addTitle("关于印发《XX管理办法》的通知")
    .addBlankLine(1)
    .addRecipient("各有关单位")
    .addParagraph("《XX管理办法》已经研究通过，现印发给你们，请认真贯彻执行。")
    .addHeading1("适用范围")
    .addParagraph("本办法适用于全国范围内的相关管理活动。")
    .addHeading1("管理原则")
    .addHeading2("依法依规")
    .addParagraph("坚持依法行政，严格按照法律法规和规章制度办事。")
    .addHeading2("公开透明")
    .addParagraph("坚持信息公开，保障公民知情权和参与权。")
    .addHeading3("信息公开方式")
    .addParagraph("通过政府网站、新闻发布会等渠道及时公开相关信息。")
    .addHeading3("公开时限要求")
    .addParagraph("相关信息应当自形成或变更之日起20个工作日内予以公开。")
    .addHeading4("主动公开事项")
    .addParagraph("涉及公共利益的重大决策信息应当主动公开。")
    .addHeading4("依申请公开事项")
    .addParagraph("公民、法人或其他组织可依法申请获取相关信息。")
    .addHeading1("附则")
    .addParagraph("本办法自发布之日起施行。");

  await doc.save("output/example-builder.docx");
  console.log("✓ 方式二完成: output/example-builder.docx");
}

// ============================================================
// 运行示例
// ============================================================
import { mkdirSync } from "fs";
mkdirSync("output", { recursive: true });

await example1();
await example2();

console.log("\n所有示例文档已生成到 output/ 目录");
