/**
 * toiletFill.js — Scriptable 脚本（厕所巡查记录自动填报）
 *
 * 由 iOS 快捷指令通过 URL 调用：
 *   scriptable:///run?scriptName=toiletFill&date=10.31&time=9:36&sheet=二楼中庭男厕
 *
 * 功能：连金山文档 MCP → 定位当天时段行 → 填5个单元格 → 上传照片 → 嵌DISPIMG → 验证
 * 首次运行无 Token 时自动走 WPS 授权流程（弹链接→轮询拿Token→存Keychain）
 *
 * 配置见下方 CONFIG。
 */

// ===================== CONFIG =====================
const CONFIG = {
  // writeTarget: "copy" = 厕所巡查记录表[副本]（测试）; "official" = 厕所巡查记录表（正式）
  writeTarget: "copy",
  defaultName: "王凯",
  defaultStatus: "无",
  defaultCleared: "∖",
  // 快捷指令保存照片的路径（相对 Scriptable iCloud 容器根）
  imagePath: "ToiletUpload/photo.jpeg",
  fileKeyword: "厕所巡查记录表",
};
// ===================================================

const EP = "https://mcp-center.wps.cn/skill_hub/mcp";
const EXCHANGE = "https://api.wps.cn/office/v5/ai/skill_hub/wps_auth/exchange";
const DATA_START = 3;
const PIC_COL = 8;

function fmt(msg) { console.log(msg); }

// ---------- HTTP / MCP ----------
function headers(sid) {
  const h = {
    "Authorization": "Bearer " + getToken(),
    "Content-Type": "application/json",
    "Accept": "application/json, text/event-stream",
  };
  if (sid) h["Mcp-Session-Id"] = sid;
  return h;
}

async function post(payload, sid) {
  const req = new Request(EP, {
    method: "POST", headers: headers(sid), body: JSON.stringify(payload),
  });
  try { req.timeoutInterval = 25; } catch (e) {}
  let text = "";
  try { text = await req.loadString(); }
  catch (e) { throw new Error("网络请求失败（25秒超时或连接失败）: " + e); }
  const respHeaders = req.response ? req.response.headers : {};
  const newSid = respHeaders["Mcp-Session-Id"] || respHeaders["mcp-session-id"] || sid || null;
  // 兼容 JSON 与 SSE 两种响应
  let json = null;
  try { json = JSON.parse(text); }
  catch (e) {
    for (const line of text.split("\n")) {
      const t = line.trim();
      if (t.startsWith("data:")) {
        const p = t.slice(5).trim();
        if (p && p !== "[DONE]") { try { json = JSON.parse(p); } catch (e2) {} }
      }
    }
  }
  return { json, sid: newSid };
}

let mcpSid = null;
let rpcId = 0;

async function rpc(method, params, notify) {
  rpcId += 1;
  const payload = { jsonrpc: "2.0", method: method };
  if (params !== undefined && params !== null) payload.params = params;
  if (!notify) payload.id = rpcId;
  const { json, sid } = await post(payload, mcpSid);
  if (sid) mcpSid = sid;
  if (notify) return null;
  if (!json) throw new Error("RPC 无响应: " + method);
  if (json.error) throw new Error("RPC错误: " + JSON.stringify(json.error).slice(0, 200));
  const result = json.result;
  if (!result || !result.content || !result.content[0]) {
    // initialize 等调用直接返回 result
    return result || {};
  }
  const inner = JSON.parse(result.content[0].text);
  if (inner.code !== undefined && inner.code !== 0) {
    throw new Error(inner.message || ("code=" + inner.code));
  }
  return inner;
}

async function callTool(name, args) {
  const inner = await rpc("tools/call", { name: name, arguments: args });
  return inner.data !== undefined ? inner.data : inner;
}

// ---------- Token ----------
function hasToken() {
  try {
    return Keychain.contains("kdocs_token") && (Keychain.get("kdocs_token") || "").length > 0;
  } catch (e) {
    return false;
  }
}

function getToken() {
  if (!hasToken()) throw new Error("未配置 Token，请先单独运行一次本脚本完成授权");
  return Keychain.get("kdocs_token");
}

function uuid() {
  return "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, c => {
    const r = Math.random() * 16 | 0;
    return (c === "x" ? r : (r & 0x3 | 0x8)).toString(16);
  });
}

async function authFlow() {
  // 先给最快路径：从剪贴板导入（电脑上复制好的 Token 发到手机后复制）
  let clip = "";
  try { clip = Pasteboard.paste() || ""; } catch (e) { clip = ""; }
  const clipTok = String(clip).replace(/^bearer\s+/i, "").trim();
  if (clipTok.length > 50) {
    const imp = new Alert();
    imp.title = "检测到剪贴板疑似 Token";
    imp.message = "长度 " + clipTok.length + " 字符。是否导入为金山文档 Token？";
    imp.addAction("✅ 导入");
    imp.addAction("不用，走在线授权");
    const pick = await imp.present();
    if (pick === 0) {
      Keychain.set("kdocs_token", clipTok);
      const ok = new Alert();
      ok.title = "✅ Token 已保存（剪贴板导入）";
      ok.message = "试试运行快捷指令；若接口报鉴权失败(400006)再走在线授权。";
      await ok.present();
      return true;
    }
  }

  const code = uuid();
  const loginUrl = "https://account.wps.cn/login?cb=" +
    encodeURIComponent("https://api.wps.cn/office/v5/ai/skill_hub/users/callback?code=" + code);

  const alert = new Alert();
  alert.title = "WPS 在线授权";
  alert.message = "方式一：复制链接到电脑浏览器打开并登录授权\n方式二：直接在本机打开授权页\n\n授权后回到这里等待（5分钟内有效）\n\n" + loginUrl;
  alert.addAction("📋 复制链接");
  alert.addAction("📱 本机打开授权页");
  alert.addAction("取消");
  const idx = await alert.present();
  if (idx === 0) {
    Pasteboard.copy(loginUrl);
    const tip = new Alert();
    tip.title = "链接已复制";
    tip.message = "粘贴到电脑浏览器打开，登录 WPS 并确认授权。\n完成后回到这里点「好」开始等待。";
    tip.addAction("好，开始等待");
    await tip.present();
  } else if (idx === 1) {
    Safari.openInApp(loginUrl, false);
  } else {
    return false;
  }

  // 轮询 exchange（无论在哪台设备授权，服务端凭 code 发 Token）
  // 用 loadString + 全文正则搜 token：兼容任意嵌套结构，HTTP 错误也不再静默吞掉
  let lastRaw = "";
  for (let i = 0; i < 100; i++) {
    await sleep(3000);
    try {
      const req = new Request(EXCHANGE, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ code: code }),
      });
      try { req.timeoutInterval = 20; } catch (e) {}
      const text = await req.loadString();   // HTTP 4xx/5xx 也返回内容，不抛异常
      lastRaw = (text || "").slice(0, 300);
      const m = (text || "").match(/"token"\s*:\s*"([^"]+)"/);
      if (m && m[1]) {
        Keychain.set("kdocs_token", m[1]);
        const ok = new Alert();
        ok.title = "✅ Token 已保存";
        ok.message = "有效期约一年。下次直接运行快捷指令即可。";
        await ok.present();
        return true;
      }
    } catch (e) {
      lastRaw = "EXC: " + String(e).slice(0, 200);
    }
  }
  throw new Error("授权超时（5分钟）。最后响应: " + lastRaw);
}

function sleep(ms) {
  return new Promise(resolve => Timer.schedule(0.001 * ms, false, resolve));
}

// ---------- 业务 ----------
async function resolveFile() {
  if (CONFIG.writeTarget === "official") {
    const data = await callTool("search_files",
      { keyword: CONFIG.fileKeyword, type: "file_name", page_size: 20 });
    const items = (data.data && data.data.items) || [];
    for (const it of items) {
      const f = it.file || {};
      if (f.name === CONFIG.fileKeyword + ".xlsx") return f.id;
    }
    for (const it of items) {
      const f = it.file || {};
      if ((f.name || "").endsWith(".xlsx") &&
          (f.name || "").includes(CONFIG.fileKeyword) &&
          !(f.name || "").includes("[副本]")) return f.id;
    }
    throw new Error("未找到正式表《" + CONFIG.fileKeyword + ".xlsx》");
  }
  const data = await callTool("search_files",
    { keyword: CONFIG.fileKeyword, type: "file_name", page_size: 20 });
  const items = (data.data && data.data.items) || [];
  for (const it of items) {
    const f = it.file || {};
    if ((f.name || "").includes("[副本]")) return f.id;
  }
  if (items.length) return items[0].file.id;
  throw new Error("云盘中未找到厕所巡查记录表");
}

async function resolveSheet(fileId, sheetName) {
  const data = await callTool("sheet.get_sheets_info", { file_id: fileId });
  const sheets = data.detail.sheetsInfo;
  const want = sheetName.replace(/\s/g, "");
  for (const s of sheets) {
    if (s.sheetName.replace(/\s/g, "") === want) return s.sheetId;
  }
  throw new Error("未找到工作表「" + sheetName + "」");
}

async function readBlock(fileId, wid, r1, r2, c1, c2) {
  const data = await callTool("sheet.get_range_data", {
    file_id: fileId, worksheet_id: wid,
    range: { rowFrom: r1, rowTo: r2, colFrom: c1, colTo: c2 },
  });
  const cells = (data.detail && data.detail.rangeData) || [];
  const grid = {};
  for (const c of cells) grid[c.originRow + "," + c.originCol] = c.cellText || "";
  return grid;
}

function slotFromTime(timeStr) {
  const hour = parseInt(timeStr.split(":")[0], 10);
  if (hour < 12) return 1;
  if (hour < 19) return 2;
  return 3;
}

async function locateRow(fileId, wid, date, slot) {
  const grid = await readBlock(fileId, wid, DATA_START, 59, 0, 3);
  const rows = new Set();
  for (const k in grid) rows.add(parseInt(k.split(",")[0], 10));
  let emptyBase = null;
  const bases = [...rows].filter(r => (r - DATA_START) % 3 === 0).sort((a, b) => a - b);
  for (const base of bases) {
    const d = grid[base + ",1"] || "";
    if (d === date) {
      const row = base + slot - 1;
      const t = grid[row + ",2"] || "";
      if (t) throw new Error(date + " 第" + slot + "次已有记录（" + t + "），请人工确认");
      return row;
    }
    if (!d && emptyBase === null) emptyBase = base;
  }
  if (emptyBase !== null) {
    const row = emptyBase + slot - 1;
    if (grid[row + ",2"]) throw new Error("目标行已有记录，请人工确认");
    return row;
  }
  throw new Error("未找到可用日期块");
}

async function writeCells(fileId, wid, items) {
  for (let attempt = 0; attempt < 3; attempt++) {
    await rpc("tools/call", {
      name: "sheet.update_range_data",
      arguments: {
        file_id: fileId, sheetId: wid,
        rangeData: items.map(it => ({
          opType: "formula", rowFrom: it.row, rowTo: it.row,
          colFrom: it.col, colTo: it.col, formula: it.value,
        })),
      },
    });
    await sleep(2500);
    let allOk = true;
    for (const it of items) {
      const got = (await readBlock(fileId, wid, it.row, it.row, it.col, it.col))[it.row + "," + it.col] || "";
      if (got !== it.value) { allOk = false; break; }
    }
    if (allOk) return;
  }
  throw new Error("单元格写入多次验证未通过");
}

async function uploadImage(fileId, imageData, filename) {
  const b64 = imageData.toBase64String();
  const data = await callTool("upload_attachment", {
    file_id: fileId, filename: filename,
    content_base64: b64, content_type: "image/jpeg",
  });
  return data.object_id;
}

async function readCellRaw(fileId, wid, row, col) {
  const data = await callTool("sheet.get_range_data", {
    file_id: fileId, worksheet_id: wid,
    range: { rowFrom: row, rowTo: row, colFrom: col, colTo: col },
  });
  const cells = (data.detail && data.detail.rangeData) || [];
  return cells.find(c => c.originCol === col) || {};
}

async function insertPicture(fileId, wid, row, col, oid, width, height) {
  for (let attempt = 0; attempt < 3; attempt++) {
    await rpc("tools/call", {
      name: "sheet.range_data_batch_update",
      arguments: {
        file_id: fileId, worksheet_id: wid,
        range_data: [{
          row_from: row, row_to: row, col_from: col, col_to: col,
          op_type: "cell_operation_type_picture",
          cell_pic_info: { tag: "sheet_pic_type_attachment", pic_content: oid,
                           width: width, height: height },
        }],
      },
    });
    await sleep(2500);
    const c = await readCellRaw(fileId, wid, row, col);
    if (c.isCellPic && c.picData === oid) return c.fmlaText || "";
  }
  throw new Error("图片插入未持久化（重试3次）");
}

// ---------- 网络诊断 ----------
async function probe(url, expectCode) {
  const start = Date.now();
  const req = new Request(url);
  try { req.timeoutInterval = 15; } catch (e) {}
  try {
    await req.loadString();
    const code = req.response ? req.response.statusCode : "?";
    const ms = Date.now() - start;
    const mark = String(code) === String(expectCode) ? "✅" : "⚠️";
    return mark + " HTTP " + code + "（" + ms + "ms）";
  } catch (e) {
    return "❌ " + String(e).slice(0, 60) + "（" + (Date.now() - start) + "ms）";
  }
}

async function probePost(label, hdrs) {
  const start = Date.now();
  try {
    const req = new Request(EP, {
      method: "POST", headers: hdrs,
      body: JSON.stringify({ jsonrpc: "2.0", id: 1, method: "initialize",
        params: { protocolVersion: "2024-11-05", capabilities: {},
                  clientInfo: { name: "probe", version: "1.0" } } }),
    });
    try { req.timeoutInterval = 12; } catch (e) {}
    await req.loadString();
    const code = req.response ? req.response.statusCode : "?";
    return label + ": ✅ HTTP " + code + "（" + (Date.now() - start) + "ms）";
  } catch (e) {
    return label + ": ❌ " + String(e).slice(0, 50) + "（" + (Date.now() - start) + "ms）";
  }
}

async function diagnose() {
  const lines = [];
  lines.push("1) 基础网络 www.wps.cn: " + await probe("https://www.wps.cn/", 200));
  lines.push("2) 接口域名 GET(404属正常): " + await probe("https://mcp-center.wps.cn/", 404));
  // POST 探针矩阵：定位是哪个请求头触发挂起
  lines.push("2b) POST 仅Content-Type: " + await probePost(
    "", { "Content-Type": "application/json" }));
  lines.push("2c) POST +Authorization: " + await probePost(
    "", { "Content-Type": "application/json",
          "Authorization": "Bearer " + getToken() }));
  lines.push("2d) POST +Accept(SSE): " + await probePost(
    "", { "Content-Type": "application/json",
          "Authorization": "Bearer " + getToken(),
          "Accept": "application/json, text/event-stream" }));
  try {
    mcpSid = null;
    await rpcInit();
    lines.push("3) MCP initialize: ✅ 会话已建立" +
      (mcpSid ? "" : "（⚠️ 未取到会话ID，后续调用会失败！）"));
  } catch (e) {
    lines.push("3) MCP initialize: ❌ " + String(e && e.message ? e.message : e).slice(0, 120));
    return lines.join("\n") + "\n\n→ 第2步超时=手机网络到接口不通\n→ 第3步失败但第2步通=会话/协议问题";
  }
  try {
    await callTool("search_files",
      { keyword: CONFIG.fileKeyword, type: "file_name", page_size: 3 });
    lines.push("4) API调用(带Token): ✅");
  } catch (e) {
    lines.push("4) API调用(带Token): ❌ " + String(e && e.message ? e.message : e).slice(0, 120));
  }
  return lines.join("\n");
}

// ---------- 主流程 ----------
async function main() {
  let q;
  try {
    q = args.queryParameters;
  } catch (e) { q = {}; }

  // 无参数运行 → 设置菜单（诊断/测试/重配/清除 Token）
  if (!q || !q.date || !q.time || !q.sheet) {
    while (true) {
      const a = new Alert();
      a.title = "toiletFill 设置";
      a.message = "模式=" + CONFIG.writeTarget + "（" +
        (CONFIG.writeTarget === "official" ? "正式表" : "测试[副本]") + "）\n" +
        "Token: " + (hasToken() ? "已配置" : "未配置");
      a.addAction("🩺 网络诊断");
      if (hasToken()) a.addAction("🔍 测试 Token（连接接口）");
      a.addAction("🔑 重新配置 Token");
      if (hasToken()) a.addAction("🗑 清除 Token");
      a.addAction("退出");
      const pick = await a.present();
      const hasT = hasToken();
      if (pick === 0) {
        const report = await diagnose();
        const d = new Alert();
        d.title = "🩺 诊断报告";
        d.message = report;
        d.addAction("复制报告");
        d.addAction("好");
        const di = await d.present();
        if (di === 0) Pasteboard.copy(report);
      } else if (hasT && pick === 1) {
        try {
          await rpcInit();
          await callTool("search_files",
            { keyword: CONFIG.fileKeyword, type: "file_name", page_size: 3 });
          const ok = new Alert();
          ok.title = "✅ Token 有效";
          ok.message = "接口连接正常，可以跑快捷指令了。";
          await ok.present();
        } catch (e) {
          const bad = new Alert();
          bad.title = "❌ 连接失败";
          bad.message = String(e && e.message ? e.message : e).slice(0, 300) +
            "\n\n如果是鉴权问题，请「重新配置 Token」。\n或先选「🩺 网络诊断」定位。";
          bad.addAction("好");
          await bad.present();
        }
      } else if ((hasT && pick === 2) || (!hasT && pick === 1)) {
        await authFlow();   // 会先尝试剪贴板导入
      } else if (hasT && pick === 3) {
        Keychain.remove("kdocs_token");
        const c = new Alert();
        c.title = "已清除";
        c.message = "Token 已从 Keychain 移除。";
        c.addAction("好");
        await c.present();
      } else {
        return;
      }
    }
  }

  // 快捷指令调用：确保 Token 后连接口
  if (!hasToken()) {
    await authFlow();
    if (!hasToken()) return; // 用户取消授权
  }

  await rpcInit();

  const { date, time, sheet } = q;
  const slot = slotFromTime(time);

  // 读照片（Scriptable iCloud 容器内）
  const fm = FileManager.iCloud();
  const imgPath = CONFIG.imagePath;
  if (!fm.fileExists(imgPath)) {
    throw new Error("找不到照片 " + imgPath + "，请检查快捷指令「存储文件」路径");
  }
  const imgData = fm.readData(imgPath);
  const img = Image.fromData(imgData);
  const pxW = Math.round(img.size.width);
  const pxH = Math.round(img.size.height);

  const fileId = await resolveFile();
  const wid = await resolveSheet(fileId, sheet);
  const row = await locateRow(fileId, wid, date, slot);
  fmt("目标行: " + row);

  // 填 5 个单元格（次数列若为空也补）
  const items = [
    { row: row, col: 1, value: date },
    { row: row, col: 2, value: time },
    { row: row, col: 4, value: CONFIG.defaultStatus },
    { row: row, col: 5, value: CONFIG.defaultName },
    { row: row, col: 6, value: CONFIG.defaultCleared },
  ];
  const cnt = (await readBlock(fileId, wid, row, row, 0, 0))[row + ",0"] || "";
  if (!cnt) items.push({ row: row, col: 0, value: String(slot) });
  await writeCells(fileId, wid, items);

  // 上传 + 嵌图（尺寸=像素×11）
  const oid = await uploadImage(fileId, imgData, "toilet_" + date.replace(".", "") + "_" + time.replace(":", "") + ".jpeg");
  await insertPicture(fileId, wid, row, PIC_COL, oid, Math.round(pxW * 11), Math.round(pxH * 11));

  const done = new Alert();
  done.title = "✅ 填报成功";
  done.message = sheet + " " + date + " 第" + slot + "次 " + time +
    "\n巡查人: " + CONFIG.defaultName + "\n图片: 已嵌入（" + pxW + "×" + pxH + "）";
  done.addAction("好");
  await done.present();
}

async function rpcInit() {
  mcpSid = null;
  await rpc("initialize", {
    protocolVersion: "2024-11-05", capabilities: {},
    clientInfo: { name: "toiletFill-ios", version: "1.0" },
  });
  await rpc("notifications/initialized", {}, true);
}

// 错误兜底：弹窗显示，方便手机上看
main().catch(e => {
  const a = new Alert();
  a.title = "❌ 出错了";
  a.message = String(e && e.message ? e.message : e).slice(0, 400);
  a.addAction("好");
  a.present();
  console.error(String(e));
});
