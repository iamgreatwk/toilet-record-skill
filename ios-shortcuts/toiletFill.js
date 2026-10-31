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
  let text = "";
  try { text = await req.loadString(); }
  catch (e) { throw new Error("网络请求失败: " + e); }
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
function getToken() {
  const t = Keychain.get("kdocs_token");
  if (!t) throw new Error("未配置 Token，请先单独运行一次本脚本完成授权");
  return t;
}

function uuid() {
  return "xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx".replace(/[xy]/g, c => {
    const r = Math.random() * 16 | 0;
    return (c === "x" ? r : (r & 0x3 | 0x8)).toString(16);
  });
}

async function authFlow() {
  const code = uuid();
  const loginUrl = "https://account.wps.cn/login?cb=" +
    encodeURIComponent("https://api.wps.cn/office/v5/ai/skill_hub/users/callback?code=" + code);
  const alert = new Alert();
  alert.title = "WPS 授权";
  alert.message = "点击「打开授权页」后，在浏览器登录/确认，然后回到这里等待。\n（约一年有效）";
  alert.addAction("打开授权页");
  alert.addAction("取消");
  const idx = await alert.present();
  if (idx !== 0) return false;
  Safari.openInApp(loginUrl, false); // 用 App 内浏览器，返回方便
  // 轮询 exchange
  for (let i = 0; i < 100; i++) {
    await sleep(3000);
    try {
      const req = new Request(EXCHANGE, {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({ code: code }),
      });
      const resp = await req.loadJSON();
      if (resp && resp.code === 200 && resp.token) {
        Keychain.set("kdocs_token", resp.token);
        const ok = new Alert();
        ok.title = "✅ Token 已保存";
        ok.message = "有效期约一年。下次直接运行快捷指令即可。";
        await ok.present();
        return true;
      }
    } catch (e) { /* 继续轮询 */ }
  }
  throw new Error("授权超时（5分钟），请重试");
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

// ---------- 主流程 ----------
async function main() {
  let q;
  try {
    q = args.queryParameters;
  } catch (e) { q = {}; }

  // 无参数运行 → 只做授权
  if (!q || !q.date || !q.time || !q.sheet) {
    await rpcInit();
    const has = Keychain.contains("kdocs_token");
    if (!has) {
      await authFlow();
    } else {
      const a = new Alert();
      a.title = "toiletFill";
      a.message = "Token 已配置（" + CONFIG.writeTarget + " 模式）。请通过快捷指令调用。";
      a.addAction("好");
      await a.present();
    }
    return;
  }

  await rpcInit();
  if (!Keychain.contains("kdocs_token")) await authFlow();

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
