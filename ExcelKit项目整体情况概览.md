# ExcelKit精简版 整体情况概览

> 最后更新：2026-11-08 22:45
> 当前版本：v0.1.0（Git 提交保留历史）
> 当前文件：`W:\ExcelKit\ExcelKit项目整体情况概览.md`
> 项目根目录：`W:\ExcelKit\`
> Git：`W:\git\cmd\git.exe`；远端：`https://github.com/ksgujie/ExcelKit.git`

## 1. 项目目标与范围

- 目标：基于原 ExcelKit (ksgujie/ExcelKit) 精简出一个以实用为主的轻量级 Excel 读写库，只保留日常办公高频功能。
- 核心功能：Workbook/Worksheet/Cell/Range、A1 地址与 0-based 索引互转、XLSX/XLS/CSV/TSV 读写、单元格值与公式、基础样式、区域批量操作、模板渲染、基础公式计算、图片插入（XLSX：浮动/完全嵌入单元格）。
- 明确不做：图表、批注、超链接、条件格式、数据验证、工作簿/工作表保护、行列插入/删除、文档属性、auto_fill、find/replace、render_many/export_pages、Table、命名区域。Excel 365 richData/IMAGE 值类型不做。`.xls` 写入图片不支持（xlwt 限制，抛 `UnsupportedFeatureError`）。库处于设计阶段，无兼容包袱。

## 2. 当前整体进度

| 模块 | 状态 | 说明 |
| --- | --- | --- |
| 地址常量/函数 | 已完成 | `XLSX_MAX_ROWS/COLUMNS`、`XLS_MAX_ROWS/COLUMNS`（无 `MAX_ROW/MAX_COLUMN`）；支持 `$A$1`；顶层导出 6 个转换函数与 4 个常量 |
| 图片功能 (XLSX) | 已完成 | `ws.add_image`，`mode=float`/`embed`，PNG/JPEG，XLSX 读写往返；embed 可随合并区域扩展 |
| XLS 图片策略 | 已完成 | `.xls` 保存含图片抛 `UnsupportedFeatureError` |
| 文档同步 | 已完成 | README、ExcelKit_Usage_Guide.md、概览 |
| 测试 | 已完成 | 55 个单元测试全部通过（2026-11-08 22:45） |
| 推送 GitHub | 待执行 | 检查远端历史，分叉则推 `simple` 分支，不 force push master |

## 3. 已完成内容

- Workbook/Worksheet/Cell/Range、XLSX/XLS/CSV/TSV、样式、区域、公式、模板、冻结/合并/筛选/页面、`tab_color`。
- 已删除：命名区域、Table；XLS 依赖默认安装；`render` 默认 `strict=True`。
- 地址：`cell_index/cell_address/range_index/range_address/column_to_index/index_to_column` 顶层导出；支持 `$A$1`；`XLSX_MAX_ROWS=1048576`、`XLSX_MAX_COLUMNS=16384`、`XLS_MAX_ROWS=65536`、`XLS_MAX_COLUMNS=256`。
- 图片：`excelkit.core.image.Image`、`ws.add_image/remove_image/clear_images/images`；XLSX 生成 `xl/media`、`xl/drawings`、rels、Content_Types、`<drawing>`；修正 sheet XML 元素顺序（`mergeCells` 在 `sheetData` 后、`drawing` 靠后）；读取恢复图片。`#` 日期字面量赋值保持原字符串存储（功能唯一性）。
- 55 个单元测试通过（含 `test_image`）。

## 4. 未完成/待修改内容

- [ ] 提交本轮改动并推送 `https://github.com/ksgujie/ExcelKit.git`（先检查远端历史/分叉）。

## 5. 下一步具体修改点

1. **提交**：`W:\git\cmd\git.exe -C W:\ExcelKit add -A`，提交信息说明地址常量/导出/图片/文档/测试。
2. **远端检查**：`git remote add origin https://github.com/ksgujie/ExcelKit.git`（若无）；`git fetch origin`；若 `origin/master` 与本地历史分叉或无共同基础，执行 `git push -u origin HEAD:simple` 并报告；若可快进/正常，`git push -u origin master`。
3. **禁止**：未经用户明确确认不得 `--force`/`--force-with-lease` 覆盖远端 master。

## 6. 关键文件与目录说明

| 路径 | 作用 |
| --- | --- |
| `W:\ExcelKit\excelkit\address.py` | 地址/索引/列转换、`XLSX_MAX_*`/`XLS_MAX_*` |
| `W:\ExcelKit\excelkit\core\image.py` | 图片模型、PNG/JPEG 尺寸解析（标准库） |
| `W:\ExcelKit\excelkit\core\worksheet.py` | `add_image`/`_images` |
| `W:\ExcelKit\excelkit\writer\xlsx.py` | XLSX 写出（图片、元素顺序修正） |
| `W:\ExcelKit\excelkit\reader\xlsx.py` | XLSX 读取（图片） |
| `W:\ExcelKit\excelkit\writer\xls.py` | XLS 含图片抛 `UnsupportedFeatureError` |
| `W:\ExcelKit\excelkit\errors.py` | `UnsupportedFeatureError` |
| `W:\ExcelKit\tests\test_image.py` | 图片测试 |
| `W:\eyes\Temp\opencode\real\` | 真实源码工作副本（双视图落盘用） |

## 7. 运行与验证命令

```powershell
cd W:\ExcelKit
python -m unittest tests.test_address tests.test_workbook tests.test_cell tests.test_append tests.test_range tests.test_formula tests.test_file tests.test_style tests.test_template tests.test_image -v
```

- 预期结果：`Ran 55 tests ... OK`，退出码 0。
- 实际结果：55 tests OK（2026-11-08 22:45）。
- 退出码：0。

## 8. 已知问题、风险与注意事项

- **双视图**：文件工具与 shell/Python 对 `W:\ExcelKit` 可能不一致；改动用 `W:\eyes\Temp\opencode` 脚本落到真实磁盘并校验。
- XLSX 元素顺序已修正；若遇 Excel 修复提示需检查具体部件。
- `embed` 为 `twoCellAnchor editAs="twoCell"`（随单元格缩放），非 Excel 365 richData 置于单元格内值。
- 赋值时 `#` 开头的日期字面量当前按**原字符串**存储（用户确认，保持功能唯一性）；`as_date/as_datetime` 支持解析带/不带 `#` 的字符串。
- `.xls` 图片：保存含图片抛 `UnsupportedFeatureError`。

## 9. 最近变更记录

### v0.1.0（2026-11-08）

- 已有：`7313c2c` render strict 默认 True；`4c4011a` 删 Table、XLS 默认依赖；`1cbad03` tab_color。
- 本轮（已授权）：`XLSX_MAX_ROWS/COLUMNS`、`XLS_MAX_ROWS/COLUMNS`，无 `MAX_ROW/MAX_COLUMN`；地址函数顶层导出与 `$A$1`；图片 `add_image` float/embed、PNG/JPEG、XLSX 读写往返、`.xls` 含图片抛异常；修正 XLSX sheet XML 顺序；文档与测试同步（55 OK）；待推送 GitHub。

## 10. 给下一个 AI 账号的接续指令

- 先读取本文件，复述：55 测试通过，代码/文档已落地，**尚未推送**（第 4、5 节）。
- 工作目录 `W:\ExcelKit`，Git `W:\git\cmd\git.exe`。**双视图**：改动用 `W:\eyes\Temp\opencode` 脚本落盘。
- 按第 5 节提交并推送；若分叉推 `simple`，禁止未经确认 force push。
- 每次动代码前先更新本概览（阶段 A）。不要保留 `MAX_ROW/MAX_COLUMN`；不要添加图表/条件格式/保护等明确不做项。
