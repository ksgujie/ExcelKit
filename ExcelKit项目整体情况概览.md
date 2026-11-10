# ExcelKit 精简版整体情况概览

> 最后更新：2026-11-10（API 重命名：`Workbook.sheet` → `get_sheet`，`Worksheet.export` → `export_csv`）
> 当前版本：v0.1.0（Git 版本管理）
> 当前文件：`W:\ExcelKit\ExcelKit项目整体情况概览.md`
> 项目根目录：`W:\ExcelKit\`
> Git：`W:\git\cmd\git.exe`；远端：`https://github.com/ksgujie/ExcelKit.git`（分支 `master` 跟踪 `origin/main`，远端仅保留 `main`）

## 1. 项目目标与范围

- 目标：轻量级 Excel 读写库，日常办公高频功能优先，代码精简、易维护。
- 核心功能：`Workbook/Worksheet/Cell/CellValue/Range`、A1/0-based 地址转换（含 `$A$1`、XLSX/XLS 上限常量）、XLSX/XLS/CSV/TSV 读写、单元格值与公式、基础样式、区域操作、模板渲染、基础公式计算、图片（XLSX float/embed，PNG/JPEG）、合并/冻结窗格/自动筛选（XLSX 持久化）、页面设置、基础表。
- 明确不做：图表、超链接、批注/备注、条件格式、数据验证、保护、行列插删、auto_fill、find/replace 增强、Table 扩展、命名区域、richData 单元格内图片等。`.xls` 图片写入抛 `UnsupportedFeatureError`。

## 2. 当前整体进度

| 模块 | 状态 | 说明 |
| --- | --- | --- |
| 地址与转换 API | 已完成 | `cell_to_index/index_to_cell/range_to_index/index_to_range/column_to_index/index_to_column`，XLSX/XLS 限制常量导出 |
| XLSX 读写与兼容性 | 已完成 | 纯标准库；规范默认 `styles.xml` 与 workbook `bookViews`；冻结/自动筛选/合并/图片等持久化与 roundtrip |
| 图片（XLSX） | 已完成 | PNG/JPEG，`float/embed`，读写回环；XLS 写图抛 `UnsupportedFeatureError`；CSV/TSV 忽略 |
| 追加/筛选 API | 已完成 | `append_rows(rows, *, on_mismatch="warn"|"error"|"ignore")`；`auto_filter.set/enable/apply(别名)/clear`（启用区域，不按条件隐藏行） |
| API 重命名（本轮） | 已完成 | `Workbook.sheet(...)` → `get_sheet(...)`；`Worksheet.export(...)` → `export_csv(..., delimiter=",")` |
| 文档与测试 | 已完成 | `README.md`、`ExcelKit_Usage_Guide.md`、本概览同步；全量 64 tests OK |

## 3. 关键 API 现状（节选）

- `Workbook`：`add_sheet/get_sheet/remove_sheet/move_sheet/copy_sheet`、`active/sheets`、`save/load/calculate/validate/render`
- `Worksheet`：`["A1"]`/`cell/range`、`append/append_rows`、`values/used_range/max_row/max_column`、`merge`、`freeze_panes`、`auto_filter(range/set/enable/apply/clear)`、`page`、`add_image/images/remove_image/clear_images`、`find/replace`、`export_csv(filename, *, delimiter=",")`、`tab_color`
- `Cell/CellValue`：`value/formula/style/address/index`，写回型 `as_*` 与只读 `read().as_*`，严格 `as_int` 等
- `Range`：`values/set_values/clear(values,styles,hyperlinks,notes)`（当前仅 values/styles 生效；hyperlinks/notes 为预留参数）、`copy_to/merge/unmerge/to_records`

## 4. 最近提交与验证

- 已推送：`d4d6478` `Fix XLSX styles/bookViews compatibility, append_rows on_mismatch, auto_filter.enable`
- 历史：`6126d19` 地址转换重命名、`109f78b` writer 优化、`963c200` P1/P2/P3 修复等（详见 git log）。
- 本轮待提交：`Workbook.get_sheet` / `Worksheet.export_csv` 重命名，文档/概览/测试同步，`.gitignore` 忽略 `build/`。
- 全量验证命令：
  ```powershell
  cd W:\ExcelKit
  python -m unittest tests.test_address tests.test_workbook tests.test_cell tests.test_append tests.test_range tests.test_formula tests.test_file tests.test_style tests.test_template tests.test_image -v
  ```
- 实际：**64 tests, OK**（退出码 0）。

## 5. 待办与注意事项

- [ ] `Range.clear` 的 `hyperlinks/notes`：当前库未实现超链接/批注，参数无实际清除效果；方案 C（实现超链接/批注并支持真实清除）**尚未获得 A–E 规格与明确授权**，不得直接编码。
- [ ] `auto_filter`：仅启用筛选区域/持久化，不按条件隐藏行（后续增强）。
- [ ] 文档可继续扩展为“每个 API 详细说明+范例”的完整手册。
- [ ] `ET.register_namespace` 为 stdlib 进程级副作用，已集中注册并注释。
- [ ] `append_rows` 默认 `on_mismatch="warn"` 允许参差行并 `UserWarning`；严格场景用 `on_mismatch="error"`。
- [ ] `.gitignore` 已包含 `build/`（本轮提交纳入）。

## 6. 给下一个 AI 账号的接续指令

- 先读本文件与 `README.md`、`ExcelKit_Usage_Guide.md`；工作目录 `W:\ExcelKit`，Git `W:\git\cmd\git.exe`。
- 任何代码修改前按流程更新本概览（阶段 A）并获得授权；修改后同步更新 `README.md`、`ExcelKit_Usage_Guide.md`、本概览，运行第 4 节全量测试通过后再提交推送 `origin master:main`。
- 继续方案 C 前必须先确认 A–E（批注类型、超链接类型/属性、API 命名与形状、XLS/CSV/TSV/copy_to 行为、分阶段策略）并获得明确授权。
