import http from "http";
import { Packer } from "docx";
import { createDocument } from "./src/index.js";

const host = "127.0.0.1";
const port = Number(process.env.PORT ?? "3000");

function escapeHtml(value = "") {
  return value
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;");
}

function renderPage(values = {}) {
  const title = escapeHtml(values.title || "关于进一步加强公文规范化管理工作的通知");
  const recipient = escapeHtml(values.recipient || "各有关单位");
  const body = escapeHtml(values.body || "请在此输入正文内容。\n\n支持换行，将按段落写入文档。");

  return `<!DOCTYPE html>
<html lang="zh-CN">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>PieBox Preview</title>
    <style>
      body { font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif; margin: 40px auto; max-width: 760px; padding: 0 16px; line-height: 1.5; }
      h1 { margin-bottom: 8px; }
      p { color: #555; }
      label { display: block; margin: 16px 0 6px; font-weight: 600; }
      input, textarea, button { width: 100%; box-sizing: border-box; font: inherit; padding: 10px 12px; }
      textarea { min-height: 240px; resize: vertical; }
      button { margin-top: 16px; cursor: pointer; }
      .hint { margin-top: 16px; font-size: 14px; color: #666; }
    </style>
  </head>
  <body>
    <h1>公文生成预览</h1>
    <p>该项目本身是 Node.js 文档生成库。为适配 PieBox 预览，这里提供一个最小 HTTP 页面，用于在线生成并下载 .docx 文件。</p>
    <form method="POST" action="/generate">
      <label for="title">标题</label>
      <input id="title" name="title" value="${title}" />

      <label for="recipient">主送机关</label>
      <input id="recipient" name="recipient" value="${recipient}" />

      <label for="body">正文</label>
      <textarea id="body" name="body">${body}</textarea>

      <button type="submit">生成并下载 DOCX</button>
    </form>
    <div class="hint">健康检查：<code>/health</code></div>
  </body>
</html>`;
}

async function readBody(request) {
  const chunks = [];
  for await (const chunk of request) {
    chunks.push(chunk);
  }
  return Buffer.concat(chunks).toString("utf8");
}

async function generateDocument(formData) {
  const title = (formData.get("title") || "未命名公文").toString().trim();
  const recipient = (formData.get("recipient") || "").toString().trim();
  const body = (formData.get("body") || "").toString();

  const content = [];
  if (recipient) {
    content.push({ type: "recipient", text: recipient });
  }

  const paragraphs = body
    .split(/\r?\n+/)
    .map((item) => item.trim())
    .filter(Boolean);

  if (paragraphs.length === 0) {
    content.push({ type: "paragraph", text: "" });
  } else {
    for (const paragraph of paragraphs) {
      content.push({ type: "paragraph", text: paragraph });
    }
  }

  const builder = createDocument({ title, content });
  return Packer.toBuffer(builder.build());
}

const server = http.createServer(async (request, response) => {
  try {
    if (request.method === "GET" && request.url === "/health") {
      response.writeHead(200, { "Content-Type": "text/plain; charset=utf-8" });
      response.end("ok");
      return;
    }

    if (request.method === "GET" && request.url === "/") {
      response.writeHead(200, { "Content-Type": "text/html; charset=utf-8" });
      response.end(renderPage());
      return;
    }

    if (request.method === "POST" && request.url === "/generate") {
      const body = await readBody(request);
      const formData = new URLSearchParams(body);
      const buffer = await generateDocument(formData);

      response.writeHead(200, {
        "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        "Content-Disposition": 'attachment; filename="preview.docx"',
        "Content-Length": buffer.length,
      });
      response.end(buffer);
      return;
    }

    response.writeHead(404, { "Content-Type": "text/plain; charset=utf-8" });
    response.end("Not Found");
  } catch (error) {
    response.writeHead(500, { "Content-Type": "text/plain; charset=utf-8" });
    response.end(`生成失败: ${error.message}`);
  }
});

server.listen(port, host, () => {
  const address = server.address();
  const actualPort = typeof address === "object" && address ? address.port : port;
  console.log(`Preview server running at http://${host}:${actualPort}`);
});

function shutdown() {
  server.close(() => process.exit(0));
}

process.on("SIGTERM", shutdown);
process.on("SIGINT", shutdown);
