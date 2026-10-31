#!/usr/bin/env python3
"""厕所巡查记录自动填报

用法：
  python3 fill_record.py --image 照片.jpeg --sheet "二楼中庭男厕" \
      --date 10.31 --time 9:36 [--name 王凯] [--dry-run]

流程：定位工作表 → 找日期+时段对应行 → 填日期/时间/巡查人/情况说明/是否已清除
     → 上传图片 → 插入 DISPIMG 单元格图片 → 逐项验证

时间→时段映射：9:00-9:30=第1次 / 15:00-15:30=第2次 / 20:00-20:30=第3次
（按小时判断：<12→1，<19→2，其余→3）
"""
import argparse
import sys
import os

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
from kdocs_mcp import KDocsMCP, load_token, image_size

DEFAULT_NAME = "王凯"
DEFAULT_STATUS = "无"
DEFAULT_CLEARED = "∖"
PIC_COL = 8          # 巡查图片列（0-based）
DATA_START = 3       # 数据区起始行（0-based）
FILE_KEYWORD = "厕所巡查记录表"


def slot_from_time(time_str):
    hour = int(time_str.split(":")[0])
    if hour < 12:
        return 1
    if hour < 19:
        return 2
    return 3


def resolve_file(m, file_id=None, file_name=None):
    """定位表格文件。
    - file_id 直接指定
    - file_name 指定名称（不含 .xlsx）：优先精确匹配，其次不含[副本]的同名文件
    - 都不传：默认优先 [副本]（测试用）"""
    if file_id:
        return file_id
    inner = m.call("search_files", {"keyword": FILE_KEYWORD, "type": "file_name",
                                    "page_size": 20})
    items = inner.get("data", {}).get("data", {}).get("items", [])
    cands = []
    for it in items:
        f = it.get("file", {})
        if f.get("name", "").endswith(".xlsx") and FILE_KEYWORD in f.get("name", ""):
            cands.append(f)
    if not cands:
        raise RuntimeError("云盘中未找到厕所巡查记录表，请确认文件名")
    if file_name:
        want = file_name.strip()
        # 1) 精确同名（含 .xlsx）
        for f in cands:
            if f["name"] == want + ".xlsx":
                return f["id"]
        # 2) 同名但不含[副本]
        for f in cands:
            if f["name"].replace(".xlsx", "") == want:
                return f["id"]
        # 3) 含关键词且不含[副本]
        for f in cands:
            if "[副本]" not in f["name"]:
                return f["id"]
        raise RuntimeError(f"未找到名为「{want}」的正式表格")
    # 默认：优先 [副本]（测试用）
    for f in cands:
        if "[副本]" in f["name"]:
            return f["id"]
    return cands[0]["id"]


def resolve_sheet(m, file_id, sheet_name):
    inner = m.call("sheet.get_sheets_info", {"file_id": file_id})
    sheets = inner["data"]["detail"]["sheetsInfo"]
    target = sheet_name.replace(" ", "")
    for s in sheets:
        if s["sheetName"].replace(" ", "") == target:
            return s["sheetId"], s["sheetName"]
    raise RuntimeError(f"未找到工作表「{sheet_name}」。现有：{[s['sheetName'] for s in sheets]}")


def locate_row(m, file_id, wid, date, slot, dry=False):
    """找日期+时段对应的数据行。
    结构：从 DATA_START 起每 3 行一天（次数1/2/3），次数1 行带位置清单。"""
    grid = m.read_block(file_id, wid, DATA_START, 59, 0, 3)
    rows = sorted({r for (r, c) in grid})
    # 按 3 行分块
    blocks = {}
    for r in rows:
        blocks.setdefault(DATA_START + (r - DATA_START) // 3 * 3, []).append(r)
    empty_base = None
    for base in sorted(blocks):
        date_in_base = grid.get((base, 1), "")
        if date_in_base == date:
            row = base + slot - 1
            t = grid.get((row, 2), "")
            if t:
                raise RuntimeError(f"{date} 第{slot}次已有时间记录（{t}），请人工确认")
            return row
        if not date_in_base and empty_base is None:
            empty_base = base
    if empty_base is not None:
        row = empty_base + slot - 1
        if grid.get((row, 2), ""):
            raise RuntimeError(f"目标行 {row} 已有记录，请人工确认")
        return row
    raise RuntimeError("未找到可用的日期块（表格数据区已满？）")


def main():
    ap = argparse.ArgumentParser(description="厕所巡查记录自动填报")
    ap.add_argument("--image", required=True, help="巡查照片路径")
    ap.add_argument("--sheet", required=True, help="厕所名，如：二楼中庭男厕")
    ap.add_argument("--date", required=True, help="日期，如：10.31")
    ap.add_argument("--time", required=True, help="时间，如：9:36")
    ap.add_argument("--name", default=DEFAULT_NAME, help="巡查人")
    ap.add_argument("--status", default=DEFAULT_STATUS, help="情况说明")
    ap.add_argument("--cleared", default=DEFAULT_CLEARED, help="是否已清除")
    ap.add_argument("--file-id", default=None, help="直接指定表格 file_id（跳过搜索）")
    ap.add_argument("--file-name", default=None,
                    help="表格文件名（不含.xlsx）。正式使用传：厕所巡查记录表；不传默认用[副本]")
    ap.add_argument("--token", default=None, help="金山文档 Token")
    ap.add_argument("--dry-run", action="store_true", help="只定位，不写入")
    args = ap.parse_args()

    m = KDocsMCP(load_token(args.token))
    m.init()

    file_id = resolve_file(m, args.file_id, args.file_name)
    wid, real_name = resolve_sheet(m, file_id, args.sheet)
    slot = slot_from_time(args.time)
    row = locate_row(m, file_id, wid, args.date, slot, args.dry_run)
    print(f"文件: {file_id}")
    print(f"工作表: {real_name} (worksheet_id={wid})")
    print(f"目标行: 0-based {row}（第{slot}次巡查）")

    if args.dry_run:
        print("[dry-run] 不写入")
        return

    # 填单元格：日期/时间/情况说明/巡查人/是否已清除（0-based 列 1/2/4/5/6）
    items = [(row, 1, args.date), (row, 2, args.time),
             (row, 4, args.status), (row, 5, args.name), (row, 6, args.cleared)]
    # 次数列为空则补上
    if not m.read_cell(file_id, wid, row, 0):
        items.append((row, 0, str(slot)))
    m.write_cells(file_id, wid, items)
    print("单元格填写: ✅")

    # 上传 + 插图
    oid = m.upload_image(file_id, args.image)
    print(f"图片上传: {oid}")
    w, h = image_size(args.image)
    fmla = m.insert_cell_picture(file_id, wid, row, PIC_COL, oid,
                                 width=round(w * 11), height=round(h * 11))
    print(f"图片嵌入: {fmla[:50]}")
    print(f"\n✅ 完成：{real_name} {args.date} 第{slot}次 {args.time} 巡查人={args.name}")


if __name__ == "__main__":
    main()
