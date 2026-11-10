# ExcelKit精简版 整体情况概览

> 最后更新：2026-11-10 01:05
> 当前版本：v0.1.0（Git 版本管理）
> 当前文件：`W:\ExcelKit\ExcelKit项目整体情况概览.md`
> 项目根目录：`W:\ExcelKit\`
> Git：`W:\git\cmd\git.exe`；远端：`https://github.com/ksgujie/ExcelKit.git`（仅 `main`，`master` 跟踪 `origin/main`）

## 1. 项目目标与范围

- 目标：基于原 ExcelKit 精简的轻量级 Excel 读写库，日常办公高频功能优先。
- 核心功能：Workbook/Worksheet/Cell/Range、A1/0-based 地址、XLSX/XLS/CSV/TSV、样式、区域、公式、模板、图片（XLSX float/embed，PNG/JPEG）、合并、冻结窗格、自动筛选（XLSX 持久化）、页面设置。
- 明确不做：图表、批注、超链接、条件格式、数据验证、保护、行列插删、auto_fill、find/replace、Table、命名区域等。XLS 图片写入抛 `UnsupportedFeatureError`。

## 2. 当前整体进度

| 模块 | 状态 | 说明 |
| --- | --- | --- |
| styles.xml/Excel 兼容性修复 | 已完成 | 规范默认 styles.xml（fonts/fills/borders/xfs 属性完整）并补充 workbook bookViews，修复 Excel 打开单元格显示空值问题 |
| append_rows on_mismatch | 已完成 | `append_rows(..., on_mismatch="warn"|"error"|"ignore")` 默认 warn；error 抛 ValueError 且原子；warn/ignore 允许参差行并按需提醒 |
| auto_filter.enable() | 已完成 | 新增 `enable()`，`apply()` 作为兼容别名（启用筛选区域；当前不按条件隐藏行） |
| 测试/文档/推送 | 已完成 | 全量 **64 tests OK**，文档/概览同步，已提交并推送 |

## 3. 已完成内容

- `6126d19`：地址转换重命名（*_to_*）。
- `109f78b`：writer 命名空间集中化、dimension 覆盖合并/图片锚点。
- `963c200`：P1/P2/P3 修复。
- 本轮：修复 XLSX 默认样式/工作簿视图兼容性（解决 Excel 显示 A1/B1 为空）；`append_rows` 增加 `on_mismatch`（默认 warn，支持参差行/警告，error 原子抛 ValueError）；`auto_filter.enable()`（`apply()` 兼容别名）。全量 **64 tests OK**。

## 4. 未完成/后续可选

- [ ] `auto_filter` 按条件筛选并隐藏行（当前仅启用筛选区域/持久化，属后续增强）。
- [ ] 文档可继续补全为“每个 API 详细说明+范例”的完整手册（待统一安排）。
- [ ] GitHub Releases 旧发布如有需网页手动删除。

## 5. 关键文件与目录说明

| 路径 | 作用 |
| --- | --- |
| `excelkit/writer/xlsx.py` | XLSX 写出（规范 styles、bookViews、图片/冻结/筛选/合并等） |
| `excelkit/core/worksheet.py` | `append_rows` on_mismatch、`auto_filter.enable/apply`、add_image 等 |
| `tests/test_append.py` | 追加与 on_mismatch 回归测试 |
| `ExcelKit_Usage_Guide.md` | 使用文档 |
| `README.md` | 项目说明 |

## 6. 运行与验证命令

```powershell
cd W:\ExcelKit
python -m unittest tests.test_address tests.test_workbook tests.test_cell tests.test_append tests.test_range tests.test_formula tests.test_file tests.test_style tests.test_template tests.test_image -v
```

- 预期：全量 OK，退出码 0。
- 实际：**64 tests, OK**（2026-11-10 验证，退出码 0）。

## 7. 已知问题、风险与注意事项

- `ET.register_namespace` 为 stdlib 进程级 API，已集中注册并注释（无法完全作用域隔离）。
- `append_rows` 默认 `on_mismatch="warn"` 允许参差行并发出 `UserWarning`；需要严格表格时请用 `on_mismatch="error"`。
- `auto_filter.enable()/apply()` 仅启用筛选区域，不执行按条件隐藏行；XLSX 持久化筛选范围。

## 8. 最近变更记录

### v0.1.0（2026-11-10）
- `6126d19`：地址转换重命名（*_to_*）。
- `109f78b`：dimension/命名空间注册优化。
- 本轮：修复 styles.xml/bookViews Excel 兼容性；`append_rows` on_mismatch；`auto_filter.enable()`。全量 **64 tests OK**，已提交推送。

## 9. 给下一个 AI 账号的接续指令

- 读本文件；工作目录 `W:\ExcelKit`，Git `W:\git\cmd\git.exe`。
- 运行第 6 节确认全量通过；修改代码必须同步更新 `README.md`、`ExcelKit_Usage_Guide.md`、本概览并跑测试后再推送。
