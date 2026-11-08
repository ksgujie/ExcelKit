# ExcelKit精简版 整体情况概览

> 最后更新：2026-11-08 23:55
> 当前版本：v0.1.0（Git 版本管理）
> 当前文件：`W:\ExcelKit\ExcelKit项目整体情况概览.md`
> 项目根目录：`W:\ExcelKit\`
> Git：`W:\git\cmd\git.exe`；远端：`https://github.com/ksgujie/ExcelKit.git`（仅 `main`，当前 `master` 跟踪 `origin/main`）

## 1. 项目目标与范围

- 目标：基于原 ExcelKit 精简的轻量级 Excel 读写库，日常办公高频功能优先。
- 核心功能：Workbook/Worksheet/Cell/Range、A1/0-based 地址（含 `$A$1`、XLSX/XLS 上限常量）、XLSX/XLS/CSV/TSV、样式、区域、公式、模板、图片（XLSX float/embed，PNG/JPEG）、合并、冻结窗格、自动筛选（XLSX 持久化）、页面设置。
- 明确不做：图表、批注、超链接、条件格式、数据验证、保护、行列插删、auto_fill、find/replace、Table、命名区域、richData 单元格内图片。XLS 图片写入抛 `UnsupportedFeatureError`。库处于设计阶段，无兼容包袱。

## 2. 当前整体进度

| 模块 | 状态 | 说明 |
| --- | --- | --- |
| P1a Cell.as_* 写回 | 已完成 | `Cell.as_*` 转换后写回单元格并返回；`read().as_*` 只读 |
| P1b as_int 严格整数 | 已完成 | 非整数浮点（如 3.14）抛 `ValueError`，不截断；`3.0`/`"3.0"` 可转；失败不写回 |
| P2 freeze/autoFilter 持久化 | 已完成 | XLSX writer 输出 sheetViews/pane、autoFilter；reader 恢复 |
| P3a 清理无用代码 | 已完成 | 删除未使用 `parse_hash_datetime_literal` |
| P3c 加载器容错收窄 | 已完成 | 合并/图片解析捕获数据/文件异常，暴露 `NameError` 等代码错误 |
| 测试/文档 | 已完成 | 57 tests OK；README/Usage Guide/概览已同步 |
| 推送 GitHub | 待本次提交后执行 | `git push origin master:main` |

## 3. 已完成内容

- `670d825`（已在 `origin/main`）：地址常量/导出/`$A$1`、Image+add_image(float/embed)+XLSX 往返、XLS 图片异常、文档/测试。
- BUG-1：XLSX 合并读取丢失（`range_index` 未导入被宽泛异常吞掉）已修复，含 `test_xlsx_merge_roundtrip`。
- 本轮修复（待提交）：P1a/P1b/P2/P3a/P3c，新增 `test_xlsx_freeze_and_autofilter_roundtrip`，`test_cell` 更新写回/严格 int。
- 远端清理：仅 `origin/main`；旧 `simple`/旧标签已删；GitHub Releases 需网页手动处理。

## 4. 未完成/后续可选

- [ ] 本轮提交并 `git push origin master:main`。
- [ ] `ET.register_namespace` 全局副作用：已解释，属嵌入兼容性，本轮未改。
- [ ] `dimension` 未纳入合并/图片锚点范围（通常无害，按需优化）。

## 5. 关键文件与目录说明

| 路径 | 作用 |
| --- | --- |
| `W:\ExcelKit\excelkit\core\cell.py` | Cell/CellValue；`as_*` 写回、`as_int` 严格整数 |
| `W:\ExcelKit\excelkit\core\worksheet.py` | freeze_panes/auto_filter/image/merge 等 |
| `W:\ExcelKit\excelkit\writer\xlsx.py` | XLSX 写出（sheetViews/pane、autoFilter、merge、drawing） |
| `W:\ExcelKit\excelkit\reader\xlsx.py` | XLSX 读取（恢复 freeze/autoFilter/merge/image；收窄异常） |
| `W:\ExcelKit\excelkit\address.py` | 地址转换与 XLSX/XLS 上限常量 |
| `W:\ExcelKit\tests\test_cell.py` | 写回/严格 int 回归 |
| `W:\ExcelKit\tests\test_file.py` | XLSX roundtrip、merge、freeze/autoFilter |

## 6. 运行与验证命令

```powershell
cd W:\ExcelKit
python -m unittest tests.test_address tests.test_workbook tests.test_cell tests.test_append tests.test_range tests.test_formula tests.test_file tests.test_style tests.test_template tests.test_image -v
```

- 预期：全量 OK，退出码 0。
- 实际：**57 tests, OK**（2026-11-08 验证，退出码 0）。

## 7. 已知问题、风险与注意事项

- **双视图**：文件工具与 shell/Python 对 `W:\ExcelKit` 可能不一致；通过 `W:\eyes\Temp\opencode` 落盘并同步/校验。
- `#` 日期字面量赋值保持原字符串存储；`as_date/as_datetime` 可解析带/不带 `#`。
- `Cell.as_*` 为**写回型**（与 `read().as_*` 只读区分）；`as_int()` 对非整数小数抛 `ValueError`。
- XLSX 保存含图片到 `.xls` 抛 `UnsupportedFeatureError`。

## 8. 最近变更记录

### v0.1.0（2026-11-08）
- `670d825`：地址/图片/文档/56 tests 基线（含 merge 修复后）。
- 本轮：P1a/P1b 写回与严格 `as_int`；P2 `freeze_panes`/`auto_filter.range` XLSX 持久化；P3a 删除 `parse_hash_datetime_literal`；P3c reader 异常收窄。新增 freeze/autoFilter 回归测试。全量 **57 tests OK**。文档同步。

## 9. 给下一个 AI 账号的接续指令

- 先读本文件；工作目录 `W:\ExcelKit`，Git `W:\git\cmd\git.exe`，双视图用 `W:\eyes\Temp\opencode` 落盘。
- 运行第 6 节确认全量通过；确认 `origin/master:main` 状态；除非用户明确要求，不 force push。
- 修改代码必须同步更新 `README.md`、`ExcelKit_Usage_Guide.md`、本概览并跑测试。
