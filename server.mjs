import { createServer } from "node:http";
import { readFile } from "node:fs/promises";
import { existsSync } from "node:fs";
import { extname, join, normalize } from "node:path";
import { fileURLToPath } from "node:url";

const ROOT = fileURLToPath(new URL(".", import.meta.url));
const PUBLIC_DIR = join(ROOT, "public");

// Keep the project dependency-free while still supporting the usual .env workflow.
if (existsSync(join(ROOT, ".env"))) {
  const envText = await readFile(join(ROOT, ".env"), "utf8");
  for (const line of envText.split(/\r?\n/)) {
    const match = line.match(/^\s*([A-Za-z_][A-Za-z0-9_]*)\s*=\s*(.*?)\s*$/);
    if (match && process.env[match[1]] === undefined) process.env[match[1]] = match[2].replace(/^['"]|['"]$/g, "");
  }
}
// veFaaS exposes port 8000 by default; use PORT=3000 for a local dev server.
const PORT = Number(process.env.PORT || 8000);

let tokenCache = { value: "", expiresAt: 0 };

const SAMPLE_MEETING = `会议主题：Q4 新用户 onboarding 优化\n日期：2026-09-03\n参会人：产品-林晓，设计-周宁，研发-陈默，运营-王悦\n\n王悦：最近新用户完成首次配置的比例偏低，客服反馈主要卡在权限配置和不知道下一步做什么。\n林晓：Q4 先解决首次配置流程，目标是把完成率从当前基线提升 15%。\n陈默：研发可以在 9 月 18 日前增加权限配置向导，并支持保存草稿。\n周宁：设计在 9 月 11 日前给出两套向导方案，重点验证首屏信息是否足够清楚。\n王悦：运营负责找 10 位新用户做可用性测试，9 月 16 日前回收反馈。\n林晓：决定采用“权限向导 + 下一步清单”的方案，灰度一周后看数据。\n\n待确认：灰度的具体用户比例和数据看板负责人尚未确定。`;

const SYSTEM_PROMPT = `你是会议执行助理。请把会议纪要转为可执行的“决策”和“任务”。
只使用纪要中明确出现的信息，不要臆造负责人、日期或指标；不确定的内容放入 open_questions。
输出必须是 JSON，不要 Markdown，不要解释。
JSON 结构：
{
  "meeting_title": string,
  "summary": string,
  "decisions": [{"id": string, "content": string, "owner": string, "deadline": string, "confidence": "high"|"medium"|"low", "source": string}],
  "tasks": [{"id": string, "title": string, "description": string, "owner": string, "due_date": string, "priority": "P0"|"P1"|"P2", "status": "待开始", "decision_id": string, "source": string}],
  "open_questions": [string]
}
日期统一使用 YYYY-MM-DD；纪要没有给出日期时返回空字符串。每个任务都必须能由一个人执行，并尽量关联一个 decision_id。`;

function json(res, status, body) {
  const data = JSON.stringify(body);
  res.writeHead(status, {
    "Content-Type": "application/json; charset=utf-8",
    "Content-Length": Buffer.byteLength(data),
    "Access-Control-Allow-Origin": "*"
  });
  res.end(data);
}

function text(res, status, body, contentType = "text/plain; charset=utf-8") {
  res.writeHead(status, { "Content-Type": contentType });
  res.end(body);
}

async function body(req) {
  let raw = "";
  for await (const chunk of req) raw += chunk;
  if (!raw) return {};
  try { return JSON.parse(raw); } catch { throw new Error("请求体不是有效 JSON"); }
}

function documentIdFromUrl(value) {
  if (!value) return "";
  const input = String(value).trim();
  // Supports docx links and a directly supplied document token.
  const match = input.match(/\/docx\/([A-Za-z0-9_-]+)/i) || input.match(/^[A-Za-z0-9_-]{10,}$/);
  return match ? (match[1] || match[0]) : "";
}

function wikiTokenFromUrl(value) {
  const match = String(value || "").match(/\/wiki\/([A-Za-z0-9_-]+)/i);
  return match ? match[1] : "";
}

async function feishuToken() {
  if (tokenCache.value && tokenCache.expiresAt > Date.now() + 30_000) return tokenCache.value;
  if (!process.env.FEISHU_APP_ID || !process.env.FEISHU_APP_SECRET) {
    throw new Error("未配置 FEISHU_APP_ID / FEISHU_APP_SECRET");
  }
  const response = await fetch("https://open.feishu.cn/open-apis/auth/v3/tenant_access_token/internal", {
    method: "POST",
    headers: { "Content-Type": "application/json; charset=utf-8" },
    body: JSON.stringify({ app_id: process.env.FEISHU_APP_ID, app_secret: process.env.FEISHU_APP_SECRET })
  });
  const result = await response.json();
  if (!response.ok || result.code !== 0 || !result.tenant_access_token) {
    throw new Error(`获取飞书 Token 失败：${result.msg || response.status}`);
  }
  tokenCache = { value: result.tenant_access_token, expiresAt: Date.now() + (Number(result.expire) || 7200) * 1000 };
  return tokenCache.value;
}

async function readFeishuDoc(documentId) {
  const token = await feishuToken();
  const response = await fetch(`https://open.feishu.cn/open-apis/docx/v1/documents/${encodeURIComponent(documentId)}/raw_content`, {
    headers: { Authorization: `Bearer ${token}` }
  });
  const result = await response.json();
  if (!response.ok || result.code !== 0) {
    throw new Error(`读取飞书文档失败：${result.msg || response.status}`);
  }
  return result.data?.content || "";
}

async function resolveWikiDocumentId(wikiToken) {
  const token = await feishuToken();
  const response = await fetch(`https://open.feishu.cn/open-apis/wiki/v2/spaces/get_node?token=${encodeURIComponent(wikiToken)}`, {
    headers: { Authorization: `Bearer ${token}` }
  });
  const result = await response.json();
  if (!response.ok || result.code !== 0) throw new Error(`读取飞书 Wiki 节点失败：${result.msg || response.status}`);
  const node = result.data?.node || {};
  if (node.obj_type && node.obj_type !== "docx") throw new Error(`当前 Wiki 页面类型为 ${node.obj_type}，请提供飞书文档（docx）链接`);
  return node.obj_token || node.objToken || "";
}

function extractTextFromChat(result) {
  const content = result?.choices?.[0]?.message?.content;
  if (Array.isArray(content)) return content.map((item) => item.text || "").join("");
  return String(content || "");
}

function normalizeResult(value) {
  const empty = { meeting_title: "未命名会议", summary: "", decisions: [], tasks: [], open_questions: [] };
  let parsed = value;
  if (typeof value === "string") {
    const cleaned = value.replace(/^```json\s*/i, "").replace(/```\s*$/i, "").trim();
    parsed = JSON.parse(cleaned);
  }
  const result = { ...empty, ...parsed };
  result.decisions = Array.isArray(result.decisions) ? result.decisions : [];
  result.tasks = Array.isArray(result.tasks) ? result.tasks : [];
  result.open_questions = Array.isArray(result.open_questions) ? result.open_questions : [];
  result.decisions = result.decisions.map((d, i) => ({
    id: d.id || `D${i + 1}`,
    content: d.content || "",
    owner: d.owner || "",
    deadline: d.deadline || "",
    confidence: d.confidence || "medium",
    source: d.source || ""
  }));
  result.tasks = result.tasks.map((t, i) => ({
    id: t.id || `T${i + 1}`,
    title: t.title || "",
    description: t.description || "",
    owner: t.owner || "",
    due_date: t.due_date || "",
    priority: t.priority || "P1",
    status: t.status || "待开始",
    decision_id: t.decision_id || "",
    source: t.source || ""
  }));
  return result;
}

async function analyzeWithAI(meetingText) {
  if (!process.env.OPENAI_API_KEY) {
    if (String(process.env.ALLOW_SAMPLE_MODE).toLowerCase() === "false") {
      throw new Error("未配置 OPENAI_API_KEY");
    }
    if (meetingText.trim() !== SAMPLE_MEETING.trim()) {
      throw new Error("当前处于示例模式：请点击“填入示例”体验，或配置 OPENAI_API_KEY 后分析真实纪要。");
    }
    return normalizeResult({
      meeting_title: "Q4 新用户 onboarding 优化",
      summary: "围绕新用户首次配置完成率偏低的问题，团队决定采用权限向导和下一步清单，并通过灰度观察效果。",
      decisions: [{ id: "D1", content: "采用“权限向导 + 下一步清单”方案，灰度一周后评估数据。", owner: "林晓", deadline: "", confidence: "high", source: "林晓：决定采用“权限向导 + 下一步清单”的方案，灰度一周后看数据。" }],
      tasks: [
        { id: "T1", title: "开发权限配置向导和保存草稿能力", description: "完成 onboarding 权限配置向导，支持用户中途保存。", owner: "陈默", due_date: "2026-09-18", priority: "P0", status: "待开始", decision_id: "D1", source: "陈默：研发可以在 9 月 18 日前增加权限配置向导，并支持保存草稿。" },
        { id: "T2", title: "输出两套向导设计方案", description: "重点验证首屏信息是否足够清楚。", owner: "周宁", due_date: "2026-09-11", priority: "P1", status: "待开始", decision_id: "D1", source: "周宁：设计在 9 月 11 日前给出两套向导方案。" },
        { id: "T3", title: "组织 10 位新用户可用性测试", description: "招募用户、执行测试并回收反馈。", owner: "王悦", due_date: "2026-09-16", priority: "P1", status: "待开始", decision_id: "D1", source: "王悦：运营负责找 10 位新用户做可用性测试。" }
      ],
      open_questions: ["灰度用户比例尚未确定。", "数据看板负责人尚未确定。"]
    });
  }
  const base = (process.env.OPENAI_BASE_URL || "https://api.openai.com/v1").replace(/\/$/, "");
  const response = await fetch(`${base}/chat/completions`, {
    method: "POST",
    headers: { "Content-Type": "application/json", Authorization: `Bearer ${process.env.OPENAI_API_KEY}` },
    body: JSON.stringify({
      model: process.env.OPENAI_MODEL || "gpt-4o-mini",
      temperature: 0.1,
      response_format: { type: "json_object" },
      messages: [
        { role: "system", content: SYSTEM_PROMPT },
        { role: "user", content: `请分析以下会议纪要：\n\n${meetingText}` }
      ]
    })
  });
  const result = await response.json();
  if (!response.ok || result.error) throw new Error(`AI 分析失败：${result.error?.message || response.status}`);
  try { return normalizeResult(extractTextFromChat(result)); }
  catch { throw new Error("AI 返回内容不是可解析的 JSON，请重试或调整模型配置"); }
}

async function route(req, res) {
  const url = new URL(req.url, `http://${req.headers.host || "localhost"}`);
  if (req.method === "GET" && url.pathname === "/api/config") {
    return json(res, 200, { feishu: Boolean(process.env.FEISHU_APP_ID && process.env.FEISHU_APP_SECRET), ai: Boolean(process.env.OPENAI_API_KEY), sample_mode: String(process.env.ALLOW_SAMPLE_MODE).toLowerCase() !== "false" });
  }
  if (req.method === "POST" && url.pathname === "/api/analyze") {
    const input = await body(req);
    let meetingText = String(input.meeting_text || "").trim();
    let documentUrl = String(input.document_url || "").trim();
    // Be forgiving when a user pastes a Feishu URL into the large meeting-text box.
    if (!documentUrl && /^https?:\/\/[^\s]+\/(?:wiki|docx)\/[A-Za-z0-9_-]+/i.test(meetingText)) {
      documentUrl = meetingText;
      meetingText = "";
    }
    let documentId = input.document_id || documentIdFromUrl(documentUrl);
    if (!documentId) {
      const wikiToken = wikiTokenFromUrl(documentUrl);
      if (wikiToken) documentId = await resolveWikiDocumentId(wikiToken);
    }
    if (!meetingText && documentId) meetingText = await readFeishuDoc(documentId);
    if (!meetingText && String(process.env.ALLOW_SAMPLE_MODE).toLowerCase() !== "false") meetingText = SAMPLE_MEETING;
    if (!meetingText) throw new Error("请提供飞书文档链接或会议纪要文本");
    const result = await analyzeWithAI(meetingText);
    return json(res, 200, { source: documentId ? "feishu_doc" : "text", document_id: documentId, result });
  }
  if (req.method === "POST" && url.pathname === "/api/confirm") {
    const input = await body(req);
    const result = normalizeResult(input.result || {});
    if (!result.tasks.length) throw new Error("没有可确认的任务");
    return json(res, 200, { ok: true, task_count: result.tasks.length, confirmed_at: new Date().toISOString() });
  }
  if (req.method === "GET") {
    const requested = normalize(url.pathname === "/" ? "/index.html" : url.pathname);
    const file = join(PUBLIC_DIR, requested);
    if (!file.startsWith(PUBLIC_DIR) || !existsSync(file)) return text(res, 404, "Not found");
    const mime = { ".html": "text/html; charset=utf-8", ".css": "text/css; charset=utf-8", ".js": "text/javascript; charset=utf-8" }[extname(file)] || "application/octet-stream";
    return text(res, 200, await readFile(file), mime);
  }
  return json(res, 404, { error: "Not found" });
}

const server = createServer(async (req, res) => {
  try { await route(req, res); }
  catch (error) { json(res, 400, { error: error.message || "未知错误" }); }
});

server.listen(PORT, () => console.log(`Meeting-to-Action running at http://localhost:${PORT}`));
