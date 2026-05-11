/**
 * 测试：验证公文文档生成工具的核心功能
 */

import { createDocument, OfficialDocumentBuilder, NumberingManager } from "../src/index.js";
import fs from "fs";
import assert from "assert";
import http from "http";
import net from "net";
import path from "path";
import { fileURLToPath } from "url";
import { spawn } from "child_process";

// ============================================================
// 测试 NumberingManager
// ============================================================
function testNumberingManager() {
  const nm = new NumberingManager();

  // H1 编号: 一、二、三…
  assert.strictEqual(nm.nextH1("测试"), "一、测试");
  assert.strictEqual(nm.nextH1("第二章"), "二、第二章");

  // H2 编号: 跟随 H1 序号
  assert.strictEqual(nm.nextH2("节一"), "2.1 节一");
  assert.strictEqual(nm.nextH2("节二"), "2.2 节二");

  // H3 编号
  assert.strictEqual(nm.nextH3("小节"), "2.2.1 小节");
  assert.strictEqual(nm.nextH3("小节二"), "2.2.2 小节二");

  // H4 编号
  assert.strictEqual(nm.nextH4("条款"), "（1）条款");
  assert.strictEqual(nm.nextH4("条款二"), "（2）条款二");

  // 新 H1 重置下级编号
  assert.strictEqual(nm.nextH1("第三章"), "三、第三章");
  assert.strictEqual(nm.nextH2("新节"), "3.1 新节");
  assert.strictEqual(nm.nextH3("新小节"), "3.1.1 新小节");
  assert.strictEqual(nm.nextH4("新条款"), "（1）新条款");

  // 新 H2 重置 H3 和 H4
  assert.strictEqual(nm.nextH2("另一节"), "3.2 另一节");
  assert.strictEqual(nm.nextH3("又一小节"), "3.2.1 又一小节");
  assert.strictEqual(nm.nextH4("又一条款"), "（1）又一条款");

  // reset
  nm.reset();
  assert.strictEqual(nm.nextH1("重来"), "一、重来");

  console.log("✓ NumberingManager 测试通过");
}

// ============================================================
// 测试 OfficialDocumentBuilder
// ============================================================
async function testBuilder() {
  const builder = new OfficialDocumentBuilder();
  builder
    .addTitle("测试文档")
    .addRecipient("测试单位")
    .addParagraph("正文内容。")
    .addHeading1("第一章")
    .addHeading2("第一节")
    .addHeading3("小节")
    .addHeading4("条款");

  const outputPath = "/tmp/test-official-doc.docx";
  await builder.save(outputPath);

  // 验证文件存在且非空
  const stats = fs.statSync(outputPath);
  assert(stats.size > 0, "生成的文件不应为空");
  assert(stats.size > 1000, "生成的 docx 文件大小应大于 1KB");

  // 验证是有效的 ZIP (docx)
  const buffer = fs.readFileSync(outputPath);
  assert.strictEqual(buffer[0], 0x50, "文件应以 PK 开头 (ZIP格式)");
  assert.strictEqual(buffer[1], 0x4B, "文件应以 PK 开头 (ZIP格式)");

  fs.unlinkSync(outputPath);
  console.log("✓ OfficialDocumentBuilder 测试通过");
}

// ============================================================
// 测试 createDocument 快捷函数
// ============================================================
async function testCreateDocument() {
  const builder = createDocument({
    title: "快捷函数测试",
    content: [
      { type: "recipient", text: "测试对象" },
      { type: "p", text: "段落一" },
      { type: "h1", text: "章" },
      { type: "h2", text: "节" },
      { type: "h3", text: "小节" },
      { type: "h4", text: "条" },
      { type: "heading1", text: "第二章" },
      { type: "heading2", text: "第二节" },
      { type: "blank", count: 2 },
      { type: "paragraph", text: "结尾" },
    ],
  });

  assert(builder instanceof OfficialDocumentBuilder);

  const outputPath = "/tmp/test-create-doc.docx";
  await builder.save(outputPath);

  const stats = fs.statSync(outputPath);
  assert(stats.size > 1000);

  fs.unlinkSync(outputPath);
  console.log("✓ createDocument 快捷函数测试通过");
}

async function getFreePort() {
  return new Promise((resolve, reject) => {
    const tester = net.createServer();
    tester.listen(0, "127.0.0.1", () => {
      const address = tester.address();
      tester.close(() => resolve(address.port));
    });
    tester.on("error", reject);
  });
}

async function request({ port, method, path: requestPath, body, headers = {} }) {
  return new Promise((resolve, reject) => {
    const req = http.request(
      {
        host: "127.0.0.1",
        port,
        method,
        path: requestPath,
        headers,
      },
      (res) => {
        const chunks = [];
        res.on("data", (chunk) => chunks.push(chunk));
        res.on("end", () => {
          resolve({
            statusCode: res.statusCode,
            headers: res.headers,
            body: Buffer.concat(chunks),
          });
        });
      }
    );

    req.on("error", reject);
    if (body) {
      req.write(body);
    }
    req.end();
  });
}

async function waitForServer(port, retries = 50) {
  for (let i = 0; i < retries; i++) {
    try {
      const response = await request({ port, method: "GET", path: "/health" });
      if (response.statusCode === 200 && response.body.toString() === "ok") {
        return;
      }
    } catch {
      // 服务器尚未就绪，继续重试
    }

    await new Promise((resolve) => setTimeout(resolve, 100));
  }

  throw new Error("预览服务器未能在预期时间内启动");
}

async function testPreviewServer() {
  const __filename = fileURLToPath(import.meta.url);
  const __dirname = path.dirname(__filename);
  const projectRoot = path.resolve(__dirname, "..");
  const port = await getFreePort();

  const server = spawn(process.execPath, [path.join(projectRoot, "preview-server.js")], {
    cwd: projectRoot,
    env: { ...process.env, PORT: String(port) },
    stdio: ["ignore", "pipe", "pipe"],
  });

  let stderr = "";
  server.stderr.on("data", (chunk) => {
    stderr += chunk.toString();
  });

  try {
    await waitForServer(port);

    const page = await request({ port, method: "GET", path: "/" });
    assert.strictEqual(page.statusCode, 200);
    assert.match(page.body.toString("utf8"), /<form method="POST" action="\/generate">/);

    const formBody = new URLSearchParams({
      title: "预览测试",
      recipient: "测试单位",
      body: "第一段\n第二段",
    }).toString();

    const generated = await request({
      port,
      method: "POST",
      path: "/generate",
      body: formBody,
      headers: {
        "Content-Type": "application/x-www-form-urlencoded",
        "Content-Length": Buffer.byteLength(formBody),
      },
    });

    assert.strictEqual(generated.statusCode, 200, stderr || "预览接口应返回 200");
    assert.match(
      generated.headers["content-type"],
      /application\/vnd\.openxmlformats-officedocument\.wordprocessingml\.document/
    );
    assert.strictEqual(generated.body[0], 0x50, "下载文件应以 PK 开头");
    assert.strictEqual(generated.body[1], 0x4B, "下载文件应以 PK 开头");

    console.log("✓ Preview 服务器测试通过");
  } finally {
    server.kill("SIGTERM");
    await new Promise((resolve) => server.once("exit", resolve));
  }
}

// ============================================================
// 运行所有测试
// ============================================================
try {
  testNumberingManager();
  await testBuilder();
  await testCreateDocument();
  await testPreviewServer();
  console.log("\n✅ 所有测试通过！");
} catch (err) {
  console.error("\n❌ 测试失败:", err.message);
  process.exit(1);
}
