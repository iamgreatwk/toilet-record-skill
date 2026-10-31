---
name: toilet-record
description: 厕所巡查记录表自动填报。当用户发来巡查照片并说明厕所位置时，自动读取照片水印时间，定位金山文档《厕所巡查记录表》对应工作表和数据行，填入日期/时间/巡查人/情况说明/是否已清除并嵌入 DISPIMG 单元格图片。触发词：巡查照片、厕所巡查、填巡查记录、巡查记录表。
---

# 厕所巡查记录自动填报 Skill

## 业务流程（与执行方式无关，先读懂这个）

用户提供巡查照片 + 厕所名（如"二楼中庭男厕"）时：

1. **读照片水印**：用视觉能力读取照片水印上的日期和时间（水印相机格式：`2026-10-31 09:36:15`）。
   - 微信传输的照片会丢 EXIF，**水印是唯一时间来源**
   - 推导：日期 `10.31`、时间 `9:36`（去秒、去前导零）、时段（按小时：<12→第1次，<19→第2次，其余→第3次；对应 9:00-9:30 / 15:00-15:30 / 20:00-20:30）
2. **执行填报**（按下方"执行方式"三级降级选择一种）
3. **汇报**：工作表名、日期、时段、时间、巡查人
4. **批量**：多张照片逐张处理，同一厕所连续处理，换厕所换表名

## 执行方式（按环境能力降级，能高不就低）

### 方式 A：能运行 Python（首选）

```bash
python3 <skill目录>/scripts/fill_record.py \
  --image "/path/to/照片.jpeg" --sheet "二楼中庭男厕" \
  --date "10.31" --time "9:36"
```

可选：`--name 王凯` `--status 无` `--cleared ∖` `--dry-run`（只定位）`--file-id` `--token`
`--file-name`：**正式使用必须传 `--file-name "厕所巡查记录表"`**（写入正式表）；不传默认写入 `[副本]`（测试用）。
脚本零依赖（纯标准库），内置：按名称搜文件、找表、定位日期+时段行、填 5 个单元格并逐个验证重试、上传图片、嵌 DISPIMG、验证落地。
安全检查：目标行已有时间记录会报错停下，交人工确认。

### 方式 B：客户端已挂金山文档 MCP 工具（如 kdocs.* 函数可直接调用）

按下列顺序调用（**参数名和取值严格照抄，全是实测踩坑结论**）：

1. 搜文件：`search_files {keyword:"厕所巡查记录表", type:"file_name"}` → 取名字含 `[副本]` 的 `.xlsx` 的 `file.id`
2. 列工作表：`sheet.get_sheets_info {file_id}` → 按 `sheetName` 匹配厕所名（可能有尾部空格，比对时去空格），取 `sheetId`
3. 读数据区定位行：`sheet.get_range_data {file_id, worksheet_id:<表id>, range:{rowFrom:3,rowTo:59,colFrom:0,colTo:3}}`（0-based！）
   - 行块结构：从 rowFrom=3 起每 3 行一天 = 次数1/2/3；次数1 行 col3 有位置清单
   - 找 `col1==日期` 的块 → 目标行 = 块首行 + (时段-1)；该行 col2 已有时间则停下问用户
4. 填单元格（**必须用 sheetId**，一次调用带 5 项，坐标 0-based）：
   `sheet.update_range_data {file_id, sheetId:<表id>, rangeData:[{opType:"formula",rowFrom:R,rowTo:R,colFrom:C,colTo:C,formula:"值"}...]}`
   - 项：col1=日期、col2=时间、col4="无"、col5="王凯"、col6="∖"；若 col0 为空补次数
   - **写完必须回读验证**（逐项读回对比），失败重试最多 3 次（该接口偶发静默丢弃）
5. 上传图片：`upload_attachment {file_id, filename:"xx.jpeg", content_base64:"<base64>", content_type:"image/jpeg"}` → `data.object_id`
   - ⚠️ base64 很大，若经命令行传参会超长；客户端原生 MCP 调用无此问题
6. 嵌图片（**唯一有效通道**，不是 update_range_data！坐标 0-based）：
   `sheet.range_data_batch_update {file_id, worksheet_id:<表id>, range_data:[{row_from:R,row_to:R,col_from:8,col_to:8,op_type:"cell_operation_type_picture",cell_pic_info:{tag:"sheet_pic_type_attachment",pic_content:"<object_id>",width:<像素宽×11>,height:<像素高×11>}}]}`
   - **width/height 必须带**，否则图片显示成小点；单位≈像素×11（1080×1440 → 11879×15839）
7. 验证：`sheet.get_range_data {file_id, worksheet_id, range:{rowFrom:R,rowTo:R,colFrom:8,colTo:8}}` → `isCellPic==true` 且 `picData==object_id`，且读回 col1/col2/col4/col5/col6 与写入一致

### 方式 C：只能执行 shell（curl）

用 curl 直连 MCP 端点 `https://mcp-center.wps.cn/skill_hub/mcp`，JSON-RPC 流程：
1. `initialize` → 从响应头 `Mcp-Session-Id` 取会话 ID → `notifications/initialized`
2. `tools/call {name:"<工具名>", arguments:{...}}`，参数与方式 B 完全相同
3. 大 base64 先写临时 JSON 文件再 `curl -d @file.json`，不要拼进命令行
4. 请求头：`Authorization: <Token>`、`Content-Type: application/json`、`Accept: application/json, text/event-stream`；后续请求带 `Mcp-Session-Id` 头
（有 Python 时强烈建议直接用方式 A，C 方式仅作兜底）

## Token 配置（首次必做）

来源优先级：`--token` 参数 > 环境变量 `KDOCS_TOKEN` > 脚本同目录 `.token` 文件（第一行纯 Token）。

获取：浏览器登录 https://www.kdocs.cn/latest → 右上角头像旁主菜单 → 「龙虾专属入口」→ 复制 Token。
过期特征：接口返回 `400006`，重新获取即可（约一年有效）。

## 必须遵守的接口要点（实测踩坑，违反必翻车）

1. `sheet.get_range_data` 读表格用 **`worksheet_id`**——传 `sheet_id` 会被**静默忽略，永远返回第一个表**！
2. `sheet.update_range_data` 写表格用 **`sheetId`**——传 `worksheet_id` 会被忽略！
3. 所有行列坐标 **0-based**。
4. `update_range_data` 的 `opType:"picture"` 是**假接口**（code=0 但永不落地），图片只能走 `range_data_batch_update` + `cell_operation_type_picture`。
5. `sheet.get_attachment_url` 拿不到 URL（url 恒为空），别走弯路。
6. 写入偶发静默丢弃 → 每次写完必须回读验证 + 重试。
7. 照片不要在对话里内联 base64，传本地路径给脚本。

## 目标表格结构

《厕所巡查记录表.xlsx》（正式表，实际使用传 `--file-name "厕所巡查记录表"`）或《厕所巡查记录表[副本].xlsx》（测试，默认）。脚本按名称自动搜索：
- 30 个工作表 = 5 层楼 × 东边/中庭/西边 × 男厕/女厕，表名如「二楼中庭男厕」
- 列序（0-based）：0=次数、1=日期、2=时间、3=位置、4=情况说明、5=巡查人、6=是否已清除、7=确认人、8=巡查图片、10=巡查时间段
- 数据从 0-based 第 3 行起，每天 3 行（次数 1/2/3 预填），第 1 次行预填位置勾选清单
- 固定填法：巡查人=王凯、情况说明=无、是否已清除=∖、确认人留空
- 日期格式 `10.31`、时间格式 `9:36`

## 故障排查

| 现象 | 处理 |
|---|---|
| 400006 | Token 过期，重新获取 |
| 找不到工作表 | 比对厕所名与表名（注意尾部空格） |
| "第N次已有时间记录" | 该行已被填过，与用户确认 |
| 图片 code=0 但读回 isCellPic=false | 自动重试 3 次仍失败则报用户，勿反复重写 |
| 图片显示成小点 | width/height 没带或单位错（应为像素×11） |
| 读到的永远是同一个表的数据 | 用错了 sheet_id，改 worksheet_id |
