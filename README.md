# 电子表格公式影响分析（Spreadsheet Formula Impact Analysis）

一个可自托管的 Web 应用：上传 Excel（`.xlsx`）文件后，系统解析工作表中的公式，
将**跨工作表引用**与**矩形区域引用**展开为单元格级别的依赖图，自动检测**循环引用**，
并可针对任意单元格列出其**所有直接 / 间接受影响单元格及完整影响路径**；此外提供
**依赖健康审计**，扫描孤立公式、空引用、悬空跨表引用、越界引用与高扇出节点。

- 后端：FastAPI + openpyxl
- 前端：原生 HTML / CSS / JavaScript 单页（无需构建，无需 Node）
- 部署：Dockerfile + docker-compose 一键启动

## 功能一览

- 📤 拖拽 / 点选上传 `.xlsx`（上限 25MB，内存中保留最近 5 个文件）
- 🔗 解析单元格引用、矩形区域（`A1:B10`）、整列（`B:B`）、整行（`2:5`）、
  绝对引用（`$A$1`）、跨表引用（`参数!B2`）、带空格的工作表名（`'月报 2026'!B2`）
- 🔁 Tarjan 强连通分量算法检测循环引用（含自环），页面顶部高亮展示
- 🩺 **依赖健康审计**（`GET /api/workbooks/{id}/audit`），扫描五类问题：

  | 级别 | 类型 | 含义 |
  |---|---|---|
  | 错误 error | `missing_sheet` | 跨表公式指向**不存在的工作表** |
  | 警告 warning | `out_of_sheet_range` | 跨表引用落在目标工作表**已用范围之外** |
  | 警告 warning | `empty_reference` | 单元格**被公式引用却为空**（无公式、无值） |
  | 警告 warning | `high_fanout` | 公式直接引用的单元格数**超过扇出阈值**（默认 50，可调） |
  | 提示 info | `orphan_formula` | 公式**未被任何其他公式引用**（孤立 / 疑似死计算） |

  每条问题给出级别、类型、所在工作表、单元格位置、中文说明与**证据边**（from → to），
  并汇总各级别 / 各类型计数；支持按**级别**和**工作表**筛选，一键导出审计 JSON。
- 🔍 选定单元格后展示：
  - **下游影响**：哪些单元格直接 / 间接引用了它（改了它会影响谁）
  - **上游依赖**：该单元格直接 / 间接依赖谁（它的计算需要谁）
  - 每个受影响单元格的**层级（深度）**和一条**最短影响路径**
- 🗂 可按工作表筛选、按单元格 / 公式文本搜索、切换是否只看公式单元格
- 📄 查看每个单元格的公式原文与 Excel 缓存计算值
- 💾 一键导出完整分析 JSON（工作表统计、公式清单、依赖边、循环环）

## 一键启动（Docker Compose）

```bash
docker compose up --build -d
```

然后浏览器打开：

> **http://localhost:8000**

上传工作簿后点击顶部「🩺 依赖健康审计」即可打开审计面板（支持级别 / 工作表筛选、
扇出阈值调整与审计 JSON 导出）。示例文件的 `审计示例` 工作表集中演示了全部五类问题。

停止：

```bash
docker compose down
```

查看日志 / 重启：

```bash
docker compose logs -f          # 跟踪日志
docker compose restart          # 重启服务（内存中的工作簿会清空，需重新上传）
```

如需修改端口，编辑 `docker-compose.yml` 中的 `"8000:8000"`（左为主机端口）。

## 本地开发（不使用 Docker）

需要 Python 3.11+：

```bash
pip install -r requirements.txt
python scripts/make_sample.py          # 生成示例文件（可选，仓库已附带）
uvicorn app.main:app --host 0.0.0.0 --port 8000 --reload
```

运行测试：

```bash
pip install pytest httpx
python -m pytest tests/ -v
```

## 端口与 API

服务监听 **8000** 端口，全部接口：

| 方法 | 路径 | 说明 |
|---|---|---|
| `GET` | `/` | 前端单页 |
| `POST` | `/api/upload` | 上传 xlsx（multipart 字段名 `file`），返回工作簿 id、统计与循环信息 |
| `GET` | `/api/workbooks/{id}` | 工作簿统计摘要 |
| `GET` | `/api/workbooks/{id}/cells?sheet=&q=&formulas_only=` | 单元格列表（支持工作表筛选 / 搜索） |
| `GET` | `/api/workbooks/{id}/impact?cell=参数!B2&direction=downstream` | 影响面分析；`direction` 取 `downstream`（默认，受影响）或 `upstream`（依赖链） |
| `GET` | `/api/workbooks/{id}/audit?level=&sheet=&fanout_threshold=50` | 依赖健康审计；可按级别（`error`/`warning`/`info`）与工作表筛选，`fanout_threshold` 为扇出阈值（1~100000，默认 50） |
| `GET` | `/api/workbooks/{id}/audit?download=true` | 以附件形式下载审计 JSON（同样支持上述筛选参数） |
| `GET` | `/api/workbooks/{id}/export` | 导出完整分析 JSON |
| `GET` | `/samples/sample_impact.xlsx` | 下载示例文件 |
| `GET` | `/api/health` | 健康检查 |

### 导出 JSON 结构

```json
{
  "filename": "sample_impact.xlsx",
  "summary": { "sheets": [...], "formula_count": 18, "edge_count": 43,
               "cycle_count": 2, "cycles": [["循环!A1","循环!B1"],["循环!C1"]] },
  "formulas": { "计算!B2": { "formula": "=B2*参数!B4", "value": null, "in_cycle": false } },
  "edges": [ { "from": "参数!B2", "to": "计算!B2" } ],
  "cycles": [ ["循环!A1", "循环!B1"], ["循环!C1"] ]
}
```

> 说明：`value` 取自 Excel 保存时的缓存计算结果。openpyxl 不执行公式计算，
> 由程序新建、从未在 Excel 中打开保存的文件其缓存值为 `null`，属正常现象。

### 接口示例（curl）

```bash
# 1) 上传，拿到 workbook_id
curl -F "file=@samples/sample_impact.xlsx" http://localhost:8000/api/upload
# -> {"workbook_id":"a1b2c3d4e5f6", ...}

# 2) 全量依赖健康审计（默认扇出阈值 50）
curl http://localhost:8000/api/workbooks/a1b2c3d4e5f6/audit

# 3) 只看错误级别、限定工作表、自定义扇出阈值
curl -G http://localhost:8000/api/workbooks/a1b2c3d4e5f6/audit \
  --data-urlencode "level=error" \
  --data-urlencode "sheet=审计示例" \
  --data-urlencode "fanout_threshold=20"

# 4) 下载审计 JSON（保留当前筛选）
curl -G -OJ http://localhost:8000/api/workbooks/a1b2c3d4e5f6/audit \
  --data-urlencode "download=true"
```

异常情况返回明确 4xx：

| 场景 | 状态码 | 响应示例 |
|---|---|---|
| 工作簿不存在 / 已被 LRU 淘汰 | `404` | `{"detail":"工作簿不存在或已被淘汰，请重新上传"}` |
| `fanout_threshold` 非整数或越界 | `400` | `{"detail":"fanout_threshold 必须是 1~100000 之间的整数"}` |
| `level` 取值非法 | `400` | `{"detail":"level 取值非法，应为：error, warning, info"}` |
| `sheet` 指定的工作表不存在 | `404` | `{"detail":"工作表不存在: Foo"}` |

### 审计 JSON 结构

```json
{
  "workbook_id": "a1b2c3d4e5f6",
  "filename": "sample_impact.xlsx",
  "fanout_threshold": 50,
  "filters": { "sheet": null, "level": null },
  "total": 20,
  "counts": { "error": 1, "warning": 5, "info": 14, "all": 20 },
  "counts_by_type": {
    "missing_sheet": 1, "out_of_sheet_range": 2, "empty_reference": 2,
    "high_fanout": 1, "orphan_formula": 14
  },
  "issues": [
    {
      "level": "error",
      "type": "missing_sheet",
      "type_label": "引用不存在的工作表",
      "sheet": "审计示例",
      "cell": "审计示例!B4",
      "cell_display": "审计示例!B4",
      "message": "公式中的跨表引用 不存在的表!A1 指向不存在的工作表「不存在的表」",
      "evidence": [
        { "from": "审计示例!B4", "to": "不存在的表!A1", "ref": "不存在的表!A1",
          "missing_sheet": "不存在的表" }
      ],
      "evidence_total": 1
    },
    {
      "level": "warning",
      "type": "out_of_sheet_range",
      "type_label": "跨表引用超出已用范围",
      "sheet": "审计示例",
      "cell": "审计示例!B6",
      "cell_display": "审计示例!B6",
      "message": "跨表引用 审计数据!B45:B60 超出工作表「审计数据」已用范围（51 行 × 3 列），其中 9 个单元格在已用范围之外",
      "evidence": [
        { "from": "审计示例!B6", "to": "审计数据!B45:B60", "ref": "审计数据!B45:B60",
          "target_sheet": "审计数据", "sheet_max_row": 51, "sheet_max_col": 3,
          "out_of_range_cells": 9 }
      ],
      "evidence_total": 1
    },
    {
      "level": "warning",
      "type": "empty_reference",
      "type_label": "引用空单元格",
      "sheet": "审计数据",
      "cell": "审计数据!C3",
      "cell_display": "审计数据!C3",
      "message": "单元格 审计数据!C3 为空（无公式、无值），却被 1 个公式引用",
      "evidence": [ { "from": "审计示例!B3", "to": "审计数据!C3" } ],
      "evidence_total": 1
    },
    {
      "level": "warning",
      "type": "high_fanout",
      "type_label": "扇出过大",
      "sheet": "审计示例",
      "cell": "审计示例!B7",
      "cell_display": "审计示例!B7",
      "message": "公式 审计示例!B7 直接引用 50 个单元格，超过扇出阈值 50",
      "fanout": 50,
      "threshold": 50,
      "evidence": [ { "from": "审计示例!B7", "to": "审计数据!B2" } ],
      "evidence_total": 50
    },
    {
      "level": "info",
      "type": "orphan_formula",
      "type_label": "孤立公式",
      "sheet": "审计示例",
      "cell": "审计示例!B1",
      "cell_display": "审计示例!B1",
      "message": "公式 审计示例!B1 未被任何其他公式引用，疑似孤立/废弃计算",
      "evidence": [],
      "evidence_total": 0
    }
  ]
}
```

> 说明：`counts` / `counts_by_type` 始终为**全量**计数，不受 `level` / `sheet`
> 筛选影响；`issues` 与 `total` 才是筛选后的结果。证据边超过 20 条时只返回前 20 条，
> 总数见 `evidence_total`。

## 示例文件

`samples/sample_impact.xlsx`（由 `scripts/make_sample.py` 生成，也可在首页直接下载）
包含 5 个工作表，覆盖所有分析场景：

| 工作表 | 内容 |
|---|---|
| `参数` | 纯输入项（单价、数量、折扣率、税率、月数） |
| `计算` | 跨表引用链：单价×数量 → 折扣 → 税额 → 月合计 → 年合计 |
| `汇总` | 跨表矩形区域 `SUM(计算!B2:B5)`、整列引用 `SUM(参数!B:B)` |
| `月报 2026` | 带空格的工作表名（公式里写作 `='月报 2026'!...`） |
| `循环` | 故意制造的循环引用：`A1↔B1` 互引、`C1=C1*2` 自环、`D1` 为环的下游 |
| `审计数据` | 50 行输入数据（用于扇出演示），`C3` 为范围内空单元格 |
| `审计示例` | 集中触发五类审计问题：孤立公式、空引用、悬空工作表引用、越界引用、50 格扇出 |

推荐体验路径：

1. 首页下载并上传 `sample_impact.xlsx`，顶部可看到红色循环引用提示；
2. 左侧选择 `参数!B2`（单价），右侧显示 14 个受影响单元格与逐层路径；
3. 切换到「上游依赖」，选择 `'月报 2026'!B3`，可看到完整回溯链直到 `参数!B2`；
4. 用工作表筛选定位 `循环!A1`，观察循环标记与影响面的截断行为；
5. 点击「🩺 依赖健康审计」打开审计面板：顶部按 错误/警告/提示 筛选，
   `审计示例` 工作表可看到全部五类问题；点击证据边中的单元格可直接跳回影响面视图；
   调整「扇出阈值」为 20 可观察更多高扇出告警；
6. 在审计面板点击「导出审计 JSON」或「导出分析 JSON」获得机器可读的完整结果。

## 实现要点

- `app/formula_parser.py`：词法式正则扫描公式，区分字符串字面量 / `#REF!` 错误 /
  函数名，避免把 `LOG10`、`"A1"` 误判为单元格；输出未展开的 `RawRef`。
- `app/graph.py`：
  - 加载两次工作簿（`data_only=False` 取公式、`data_only=True` 取缓存值）；
  - 区域展开为具体单元格，整行 / 整列按工作表已用范围截断，单个区域上限 5000 格
    （`MAX_RANGE_CELLS`，防 `SUM(A:XFD)` 之类打爆内存）；
  - 保留每个公式的**原始引用** `raw_refs`（含指向不存在工作表的引用，构图时跳过、审计时使用）；
  - 迭代版 Tarjan SCC 检测循环引用；
  - BFS 求影响面，记录最短路径、层级与环标记，默认最大深度 50 / 5000 节点。
- `app/audit.py`：在依赖图上做健康审计——原始引用检查缺失工作表 / 跨表越界，
  反向边聚合空引用（跨表越界不重复计数），无前向边公式判为孤立，
  出度超阈值判为高扇出；按 级别 → 工作表 → 行列 排序输出。
- `app/main.py`：FastAPI 路由与内存 LRU 式工作簿缓存（不落盘，重启即清空）。

## 目录结构

```
.
├── app/
│   ├── main.py              # FastAPI 入口与 API（含 /audit）
│   ├── formula_parser.py    # 公式引用解析
│   ├── graph.py             # 依赖图 / 环检测 / 影响面 BFS
│   ├── audit.py             # 依赖健康审计（五类问题）
│   └── static/index.html    # 前端单页（影响分析 + 审计面板）
├── samples/sample_impact.xlsx
├── scripts/make_sample.py   # 示例文件生成脚本
├── tests/                   # pytest 用例（42 个）
├── Dockerfile
├── docker-compose.yml
└── requirements.txt
```
