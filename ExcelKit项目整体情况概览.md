# ExcelKit精简版 整体情况概览

> 最后更新：2026-11-09 00:20
> 当前版本：v0.1.0（Git 版本管理）
> 当前文件：`W:\ExcelKit\ExcelKit项目整体情况概览.md`
> 项目根目录：`W:\ExcelKit\`
> Git：`W:\git\cmd\git.exe`；远端：`https://github.com/ksgujie/ExcelKit.git`（仅 `main`，`master` 跟踪 `origin/main`）

## 1. 项目目标与范围

- 目标：基于原 ExcelKit 精简的轻量级 Excel 读写库，日常办公高频功能优先。
- 核心功能：Workbook/Worksheet/Cell/Range、A1/0-based 地址（含 `$A$1`、XLSX/XLS 上限常量）、XLSX/XLS/CSV/TSV、样式、区域、公式、模板、图片（XLSX float/embed，PNG/JPEG）、合并、冻结窗格、自动筛选（XLSX 持久化）、页面设置。
- 明确不做：图表、批注、超链接、条件格式、数据验证、保护、行列插删、auto_fill、find/replace、Table、命名区域、richData 单元格内图片。XLS 图片写入抛 `UnsupportedFeatureError`。无兼容包袱。

## 2. 当前整体进度

| 模块 | 状态 | 说明 |
| --- | --- | --- |
| P1a/P1b/P2/P3a/P3c | 已完成 | `963c200` |
| ET.register_namespace 全局副作用 | 已完成（缓解） | 集中为 `_register_xlsx_namespaces()`，移除散落重复注册；stdlib `ET.register_namespace` 为进程级 API，无法完全作用域隔离，已注释说明 |
| dimension 纳入合并/图片锚点 | 已完成 | `writer/xlsx.py` 的 `dimension` 覆盖 `_max_row/_max_column` + `_merged` + 图片 anchor |
| 测试/文档/推送 | 已完成 | 60 tests OK；已提交推送 |

## 3. 已完成内容

- `963c200`：P1/P2/P3 修复，59 tests OK，已推送。
- 本轮：writer 命名空间注册集中化（减少重复全局调用）；`dimension` 覆盖合并范围与图片锚点；新增 `test_xlsx_dimension_covers_merge_and_image_anchor`。全量 **60 tests OK**。

## 4. 未完成/后续可选

- [ ] `ET.register_namespace` 的进程级副作用为 ET 固有限制；如需彻底隔离需自定义序列化（当前未引入）。
- [ ] GitHub Releases 旧发布如有需网页手动删除。

## 5. 关键文件与目录说明

| 路径 | 作用 |
| --- | --- |
| `W:\ExcelKit\excelkit\writer\xlsx.py` | XLSX 写出（集中命名空间注册、dimension 覆盖合并/图片） |
| `W:\ExcelKit\excelkit\reader\xlsx.py` | XLSX 读取 |
| `W:\ExcelKit\excelkit\core\cell.py` | Cell/CellValue（写回、严格 `as_int`） |
| `W:\ExcelKit\tests\test_file.py` | 文件往返/merge/freeze+filter/dimension 回归 |
| `W:\ExcelKit\tests\test_cell.py` | 单元格转换回归 |

## 6. 运行与验证命令

```powershell
cd W:\ExcelKit
python -m unittest tests.test_address tests.test_workbook tests.test_cell tests.test_append tests.test_range tests.test_formula tests.test_file tests.test_style tests.test_template tests.test_image -v
```

- 预期：全量 OK，退出码 0。
- 实际：**60 tests, OK**（2026-11-09 验证，退出码 0）。

## 7. 已知问题、风险与注意事项

- **双视图**：通过 `W:\eyes\Temp\opencode` 落盘并同步；以 git/测试结果为准。
- `ET.register_namespace` 为 stdlib 进程级 API，已集中注册并注释；无法在不改序列化方案前彻底消除全局影响。
- `#` 日期字面量赋值保持原字符串；`Cell.as_*` 写回型；`as_int` 严格整数；XLS 含图片抛 `UnsupportedFeatureError`。

## 8. 最近变更记录

### v0.1.0（2026-11-09）
- `963c200`：P1a/P1b/P2/P3a/P3c，59 tests OK。
- 本轮：writer `_register_xlsx_namespaces()` 集中化；`dimension` 覆盖合并/图片锚点；`test_xlsx_dimension_covers_merge_and_image_anchor`。全量 **60 tests OK**，已提交推送。

## 9. 给下一个 AI 账号的接续指令

- 读本文件；工作目录 `W:\ExcelKit`，Git `W:\git\cmd\git.exe`，双视图用 `W:\eyes\Temp\opencode` 落盘。
- 运行第 6 节确认全量通过；修改代码必须同步更新 `README.md`、`ExcelKit_Usage_Guide.md`、本概览并跑测试后再推送。
