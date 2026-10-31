#!/usr/bin/env python3
"""金山文档 MCP 极简客户端（纯标准库，无依赖）

Token 获取优先级：
  1. 命令行 --token
  2. 环境变量 KDOCS_TOKEN
  3. 脚本同目录 .token 文件（第一行为 Token）
"""
import json
import os
import sys
import urllib.request

ENDPOINT = "https://mcp-center.wps.cn/skill_hub/mcp"


def load_token(explicit=None, script_dir=None):
    """返回标准化 Token（剥掉可能存在的 Bearer 前缀，发送时统一加）。
    兼容两种存量格式：裸 Token（手机端 exchange 接口返回）和
    "Bearer <token>"（PC 端 mcporter 配置迁移）。"""
    if explicit:
        tok = explicit.strip()
    else:
        env = os.environ.get("KDOCS_TOKEN")
        if env:
            tok = env.strip()
        else:
            if script_dir is None:
                script_dir = os.path.dirname(os.path.abspath(__file__))
            token_file = os.path.join(script_dir, ".token")
            if os.path.exists(token_file):
                with open(token_file, encoding="utf-8") as f:
                    tok = ""
                    for line in f:
                        line = line.strip()
                        if line and not line.startswith("#"):
                            tok = line
                            break
            else:
                print("错误：未找到 Token。请将 Token 写入脚本同目录 .token 文件，"
                      "或设置环境变量 KDOCS_TOKEN，或用 --token 传入。", file=sys.stderr)
                sys.exit(2)
    if tok.lower().startswith("bearer "):
        tok = tok[7:].strip()
    if not tok:
        print("错误：Token 为空。", file=sys.stderr)
        sys.exit(2)
    return tok


class KDocsMCP:
    def __init__(self, token):
        # 兼容裸 Token 与带 Bearer 前缀两种输入，统一剥前缀后发送时再加
        t = token.strip()
        if t.lower().startswith("bearer "):
            t = t[7:].strip()
        self.token = t
        self.sid = None
        self.next_id = 0

    def _post(self, payload):
        headers = {
            "Authorization": "Bearer " + self.token,
            "Content-Type": "application/json",
            "Accept": "application/json, text/event-stream",
        }
        if self.sid:
            headers["Mcp-Session-Id"] = self.sid
        req = urllib.request.Request(ENDPOINT, data=json.dumps(payload).encode(),
                                     headers=headers, method="POST")
        resp = urllib.request.urlopen(req, timeout=180)
        sid = resp.headers.get("Mcp-Session-Id")
        if sid:
            self.sid = sid
        ct = resp.headers.get("Content-Type", "")
        return ct, resp.read().decode("utf-8", errors="replace")

    @staticmethod
    def _parse(ct, body):
        if "text/event-stream" in ct:
            for line in body.splitlines():
                line = line.strip()
                if line.startswith("data:"):
                    p = line[5:].strip()
                    if p and p != "[DONE]":
                        try:
                            return json.loads(p)
                        except Exception:
                            pass
            return None
        try:
            return json.loads(body)
        except Exception:
            return None

    def rpc(self, method, params=None, notify=False):
        self.next_id += 1
        payload = {"jsonrpc": "2.0", "method": method}
        if params is not None:
            payload["params"] = params
        if not notify:
            payload["id"] = self.next_id
        ct, body = self._post(payload)
        if notify:
            return None
        return self._parse(ct, body)

    def init(self):
        self.rpc("initialize", {
            "protocolVersion": "2024-11-05",
            "capabilities": {},
            "clientInfo": {"name": "toilet-record", "version": "1.0"},
        })
        self.rpc("notifications/initialized", {}, notify=True)

    def call(self, name, args):
        """调用工具，自动解包 result.content[0].text 的内层 JSON"""
        r = self.rpc("tools/call", {"name": name, "arguments": args})
        if not r or "result" not in r:
            raise RuntimeError(f"工具调用失败: {json.dumps(r, ensure_ascii=False)[:300]}")
        inner = json.loads(r["result"]["content"][0]["text"])
        if inner.get("code") not in (0, None):
            raise RuntimeError(f"{name} 业务错误 {inner.get('code')}: {inner.get('message')}")
        return inner

    # ---- 常用封装 ----

    def read_cell(self, file_id, worksheet_id, row, col):
        """读取单个单元格，0-based。空格子返回 ''"""
        try:
            inner = self.call("sheet.get_range_data", {
                "file_id": file_id, "worksheet_id": worksheet_id,
                "range": {"rowFrom": row, "rowTo": row, "colFrom": col, "colTo": col},
            })
        except RuntimeError:
            return ""
        cells = ((inner.get("data", {}) or {}).get("detail", {}) or {}).get("rangeData") or []
        c = next((x for x in cells if x.get("originCol") == col), {})
        return c.get("cellText", "") or ""

    def read_block(self, file_id, worksheet_id, r1, r2, c1, c2):
        """读取矩形区域，返回 {(row,col): text}，0-based"""
        inner = self.call("sheet.get_range_data", {
            "file_id": file_id, "worksheet_id": worksheet_id,
            "range": {"rowFrom": r1, "rowTo": r2, "colFrom": c1, "colTo": c2},
        })
        cells = ((inner.get("data", {}) or {}).get("detail", {}) or {}).get("rangeData") or []
        grid = {}
        for c in cells:
            grid[(c["originRow"], c["originCol"])] = c.get("cellText", "") or ""
        return grid

    def write_cells(self, file_id, worksheet_id, items, retries=3, delay=3):
        """批量写单元格。items: [(row, col, value)]，0-based。带验证重试，返回成功列表"""
        import time
        done = []
        pending = list(items)
        for attempt in range(retries):
            if not pending:
                break
            self.call("sheet.update_range_data", {
                "file_id": file_id, "sheetId": worksheet_id,
                "rangeData": [{"opType": "formula", "rowFrom": r, "rowTo": r,
                               "colFrom": c, "colTo": c, "formula": v}
                              for r, c, v in pending],
            })
            time.sleep(delay)
            still = []
            for r, c, v in pending:
                if self.read_cell(file_id, worksheet_id, r, c) == v:
                    done.append((r, c, v))
                else:
                    still.append((r, c, v))
            pending = still
        if pending:
            raise RuntimeError(f"以下单元格写入失败: {pending}")
        return done

    def upload_image(self, file_id, image_path, filename=None):
        """上传图片附件，返回 object_id"""
        import base64
        if filename is None:
            filename = os.path.basename(image_path)
        with open(image_path, "rb") as f:
            b64 = base64.b64encode(f.read()).decode()
        inner = self.call("upload_attachment", {
            "file_id": file_id, "filename": filename,
            "content_base64": b64, "content_type": "image/jpeg",
        })
        return inner["data"]["object_id"]

    def insert_cell_picture(self, file_id, worksheet_id, row, col, object_id,
                            width, height, retries=3, delay=3):
        """在 (row,col) 插入单元格图片（DISPIMG）。width/height 单位≈像素×11"""
        import time
        item = {"row_from": row, "row_to": row, "col_from": col, "col_to": col,
                "op_type": "cell_operation_type_picture",
                "cell_pic_info": {"tag": "sheet_pic_type_attachment",
                                  "pic_content": object_id,
                                  "width": width, "height": height}}
        for attempt in range(retries):
            self.call("sheet.range_data_batch_update", {
                "file_id": file_id, "worksheet_id": worksheet_id,
                "range_data": [item],
            })
            time.sleep(delay)
            c = self.read_cell(file_id, worksheet_id, row, col)
            # isCellPic 需要读原始字段
            inner = self.call("sheet.get_range_data", {
                "file_id": file_id, "worksheet_id": worksheet_id,
                "range": {"rowFrom": row, "rowTo": row, "colFrom": col, "colTo": col},
            })
            cells = ((inner.get("data", {}) or {}).get("detail", {}) or {}).get("rangeData") or []
            cc = next((x for x in cells if x.get("originCol") == col), {})
            if cc.get("isCellPic") and cc.get("picData") == object_id:
                return cc.get("fmlaText", "")
            time.sleep(delay)
        raise RuntimeError(f"单元格图片插入未持久化 ({row},{col})")


def image_size(path):
    """纯标准库解析 JPEG/PNG 像素尺寸"""
    with open(path, "rb") as f:
        data = f.read(64 * 1024)
    if data[:8] == b"\x89PNG\r\n\x1a\n":
        w = int.from_bytes(data[16:20], "big")
        h = int.from_bytes(data[20:24], "big")
        return w, h
    if data[:2] == b"\xff\xd8":  # JPEG
        i = 2
        while i < len(data) - 9:
            if data[i] != 0xFF:
                i += 1
                continue
            marker = data[i + 1]
            if marker in (0xC0, 0xC1, 0xC2, 0xC3):
                h = int.from_bytes(data[i + 5:i + 7], "big")
                w = int.from_bytes(data[i + 7:i + 9], "big")
                return w, h
            if marker in (0xD8, 0x01) or 0xD0 <= marker <= 0xD7:
                i += 2
                continue
            seglen = int.from_bytes(data[i + 2:i + 4], "big")
            i += 2 + seglen
    raise ValueError(f"无法解析图片尺寸: {path}")
