# ExcelKit精简版 整体情况概览

> 最后更新：2026-11-10（v0.1.2：XLS 自动筛选写入/读取，已发布）
> 当前版本：v0.1.2
> 当前文件：`W:\ExcelKit\ExcelKit项目整体情况概览.md`
> 项目根目录：`W:\ExcelKit\`
> Git：`W:\git\cmd\git.exe`，远程 `https://github.com/ksgujie/ExcelKit.git`，分支 `master` 对应 `origin/main`

## 1. 项目目标与范围

- 目标：基于原 ExcelKit (ksgujie/ExcelKit) 精简出一个以实用为主的轻量级 Excel 读写库，只保留日常办公高频功能。
- 核心功能：Workbook/Worksheet/Cell/Range 对象模型、A1 地址与 0-based 索引、XLSX/XLS/CSV/TSV 读写、单元格值与公式、基础样式、区域批量操作、模板渲染、基础公式计算。
- 明确不做：图表、超链接、批注/备注、条件格式、数据验证、工作簿/工作表保护、行列插入/删除、文档属性、auto_fill 序列填充、增强 find/replace、Table 扩展、命名区域、richData 单元格内图片、render_many/export_pages 等包装循环 API。

## 2. 当前整体进度

| 模块 | 状态 | 说明 |
| --- | --- | --- |
| 项目骨架与目录 | 已完成 | `W:\ExcelKit\excelkit\` |
| 核心对象模型 (Workbook/Worksheet/Cell/Range) | 已完成 | 全量 68 测试通过 |
| 文件读写 (XLSX/XLS/CSV/TSV) | 已完成 | XLSX 纯标准库，XLS 依赖 xlrd/xlwt |
| 样式系统 (Style/Font/Fill/Border/Alignment) | 已完成 | Border 保持现状 |
| 区域操作 (range/values/set_values/clear/copy_to/merge) | 已完成 | |
| 公式计算器与类型转换 | 已完成 | SUM/AVERAGE/IF/跨表等 |
| 模板渲染 (render) | 已完成 | 标量标签 + {loop} 循环行块 |
| 辅助功能 (冻结窗格/合并/自动筛选/页面设置/基础 Table) | 已完成 | auto_filter 只启用筛选区域，不按条件隐藏行 |
| XLSX 自动筛选持久化 | 已完成 | `<autoFilter ref="...">` |
| XLS 自动筛选写入/读取 | 已完成 | v0.1.2：AUTOFILTERINFO + `_FilterDatabase` |
| Wheel 打包 | 已完成 | `dist\excelkit-0.1.2-py3-none-any.whl` |

## 3. 已完成内容

- 精简版完整源码（excelkit 包）与 10 个测试模块，v0.1.2 时 68 个用例全部通过。
- 代码主目录：`W:\ExcelKit\excelkit\`；测试目录：`W:\ExcelKit\tests\`。
- 详细中文使用文档：`W:\ExcelKit\ExcelKit_Usage_Guide.md`。
- v0.1.1 已修复 XLSX inline string（`<is>`）导致字符串空白的问题，并移除 `auto_filter.apply()` 别名。
- v0.1.2 已实现 XLS 自动筛选：写出行程表 `AUTOFILTERINFO`(0x009D) + 全局 `SUPBOOK`(0x01AE)/`EXTERNSHEET`(0x0017)/`NAME`(0x0018, `_FilterDatabase`)，并回填 `BOUNDSHEET`(0x0085) 偏移；读取端用 xlrd `name_map["_filterdatabase"]` 还原范围。
- v0.1.2 已发布 Wheel：`W:\ExcelKit\dist\excelkit-0.1.2-py3-none-any.whl`。

## 4. 未完成/待修改内容

- [x] XLS 保存后自动筛选不生效 —— 已修复（BIFF8 记录注入）。
- [x] XLS 读取不还原自动筛选 —— 已修复（`_apply_auto_filters`）。
- [x] 文档未记录 XLS 自动筛选支持情况 —— 已同步 README 与使用文档。
- [x] 版本号与 Wheel —— 已升到 0.1.2 并重建 Wheel。
- [ ] 后续可选（未获授权，不在本轮范围）：XLS 的合并单元格、冻结窗格、列宽、图片写入。

## 5. 本轮具体修改点（已实施）

1. 文件：`W:\ExcelKit\excelkit\writer\xls.py`
   - `save_xls(workbook, filename)` 生成 BIFF 流后注入自动筛选记录：
     - 工作表子流：`DIMENSIONS`(0x0200) 前插入 `AUTOFILTERINFO`(0x009D, 列数)。
     - 全局流：`SST`(0x00FC) 前插入 `SUPBOOK`(0x01AE, 工作表数)、`EXTERNSHEET`(0x0017, 每表一条)、每个筛选表一条 `NAME`(0x0018, grbit=0x0021, 内置名 0x0D `_FilterDatabase`, `ptgArea3d` 0x3B)。
     - 全部插入完成后按最终流重新回填每条 `BOUNDSHEET`(0x0085) 的 `lbPlyPos`。
   - 用 `xlwt.CompoundDoc.XlsDoc().save()` 写出修补后的 BIFF 流。
   - 注入异常时 `warnings.warn` 并回退为不带筛选保存，绝不产出损坏文件；超出 XLS 上限（65536 行 / 256 列）跳过并警告。
2. 文件：`W:\ExcelKit\excelkit\reader\xls.py`
   - 新增 `_apply_auto_filters`：读取 `book.name_map["_filterdatabase"]`，用 `Name.scope` + `Name.result.value[0].coords` 还原 `_auto_filter_range`；单格区域输出 `"A1"`（与 XLSX 一致）。
3. 文件：`W:\ExcelKit\tests\test_file.py`
   - 新增 `test_xls_autofilter_roundtrip`、`test_xls_autofilter_writes_biff_records`。
4. 文件：`W:\ExcelKit\README.md`、`W:\ExcelKit\ExcelKit_Usage_Guide.md`、本概览
   - 记录 XLS 自动筛选已支持、XLS 仍不写图片/合并/冻结窗格/列宽；版本号 0.1.2。
5. 文件：`W:\ExcelKit\pyproject.toml`、`W:\ExcelKit\excelkit\__init__.py`
   - `version` / `__version__` 升到 `0.1.2`。

## 6. 关键文件与目录说明

| 路径 | 作用 | 备注 |
| --- | --- | --- |
| `W:\ExcelKit\excelkit\__init__.py` | 顶层公共 API 与 `__version__` | |
| `W:\ExcelKit\excelkit\core\workbook.py` | Workbook 类 | 工作簿生命周期 |
| `W:\ExcelKit\excelkit\core\worksheet.py` | Worksheet 类 | 含 `_auto_filter_range` |
| `W:\ExcelKit\excelkit\core\cell.py` | Cell、CellValue 类 | |
| `W:\ExcelKit\excelkit\core\range.py` | Range 类 | |
| `W:\ExcelKit\excelkit\core\page.py` | 页面设置（简化） | |
| `W:\ExcelKit\excelkit\core\table.py` | 基础 Table | |
| `W:\ExcelKit\excelkit\reader\xlsx.py` | XLSX 读取 | 标准库 zipfile/xml |
| `W:\ExcelKit\excelkit\reader\xls.py` | XLS 读取 | xlrd + `_apply_auto_filters` |
| `W:\ExcelKit\excelkit\reader\delimited.py` | CSV/TSV 读取 | 标准库 csv |
| `W:\ExcelKit\excelkit\writer\xlsx.py` | XLSX 写出 | 标准库 zipfile/xml |
| `W:\ExcelKit\excelkit\writer\xls.py` | XLS 写出 | xlwt + BIFF8 记录注入 |
| `W:\ExcelKit\excelkit\writer\delimited.py` | CSV/TSV 写出 | 标准库 csv |
| `W:\ExcelKit\excelkit\style.py` | 样式系统 | |
| `W:\ExcelKit\excelkit\address.py` | A1 地址解析 | `range_to_index` / `index_to_range` |
| `W:\ExcelKit\excelkit\calculation.py` | 公式计算器 | |
| `W:\ExcelKit\excelkit\template.py` | 模板渲染 | |
| `W:\ExcelKit\excelkit\errors.py` | 异常定义 | |
| `W:\ExcelKit\tests\test_*.py` | 10 个测试模块 | 本轮新增 XLS 筛选用例 |

## 7. 运行与验证命令

```powershell
cd W:\ExcelKit
python -m unittest tests.test_address tests.test_workbook tests.test_cell tests.test_append tests.test_range tests.test_formula tests.test_file tests.test_style tests.test_template tests.test_image -v
```

- 预期结果：Ran 68 tests, OK。
- 实际结果：Ran 68 tests in 0.210s，OK（2026-11-10）。
- 退出码：0。

Excel COM 验证（实际结果）：

| 文件 | 场景 | 结果 |
| --- | --- | --- |
| `xls_autofilter_check.xls` | `A1:B3` | `AutoFilterMode=True`，`Range=A1:B3` |
| `xls_autofilter_offset.xls` | `C3:D6`（数据更大） | `AutoFilterMode=True`，`Range=C3:D6` |
| `xls_autofilter_multi.xls` | 两个工作表 | S1 `A1:C6`、S2 `C3:D6`，均 `True` |
| `v012check\wheel_out.xls` | Wheel 安装后 `C3:D6` | `AutoFilterMode=True`，`Range=C3:D6` |

```powershell
$xl = New-Object -ComObject Excel.Application
$xl.Visible = $false; $xl.DisplayAlerts = $false
$wb = $xl.Workbooks.Open('W:\eyes\Temp\opencode\xls_autofilter_check.xls')
$ws = $wb.Worksheets.Item(1)
$ws.AutoFilterMode
$ws.AutoFilter.Range.Address($false, $false)
$wb.Close($false); $xl.Quit()
```

Wheel 全新虚拟环境验证（`W:\eyes\Temp\opencode\venv012`）：

- `excelkit.__version__` = `0.1.2`，`__file__` 位于 venv 的 `site-packages`。
- XLS 写入 `A1:B3` → 读回 `A1:B3`；BIFF 流含 1 个 `0x009D`、1 个 `0x0018`。
- `auto_filter` 无 `apply`、有 `enable`。

## 8. 已知问题、风险与注意事项

- XLS 读写依赖 xlrd/xlwt；XLSX/CSV/TSV 纯标准库。
- **W: 盘读写缓存风险**：`read`/`edit` 工具可能看到旧快照，所有项目文件读写请用 Python 脚本完成并校验字节数/中文数量。
- BIFF8 记录注入依赖 xlwt 内部记录顺序（BOF/EOF/BOUNDSHEET/DIMENSIONS/SST），xlwt 升级后需回归测试。
- 注入失败必须回退为"无筛选保存"，不能产出损坏文件。
- 自动筛选只启用区域/下拉，不实现筛选条件、不隐藏行。
- XLS 仍不写：图片、合并单元格、冻结窗格、列宽、公式缓存（现状，未在本轮范围）。
- 所有索引统一 0-based；A1 地址 1-based。
- `Range.clear` 方案 C 未获授权，不得编码。
- 运行本地脚本验证时勿用 `python script.py`（会把脚本目录置于 `sys.path[0]`，可能导入 `W:\eyes\Temp\opencode\excelkit` 旧副本），改用 `python -c "exec(open(..., encoding='utf-8').read())"` 且工作目录为 `W:\ExcelKit`。

## 9. 最近变更记录

### v0.1.2（2026-11-10）

- 新增：XLS 写出自动筛选（`AUTOFILTERINFO` + `_FilterDatabase` + `SUPBOOK`/`EXTERNSHEET` + `BOUNDSHEET` 偏移修正）。
- 新增：XLS 读取自动筛选（xlrd `_filterdatabase`），单格区域规范为 `"A1"`。
- 新增：XLS 自动筛选写入/读取测试（`tests/test_file.py`，共 9 个用例）。
- 修改：版本号 0.1.1 → 0.1.2；README、使用文档、本概览同步。
- 验证：68 tests OK；Excel COM 四个场景全部正确；Wheel 全新 venv 复验通过。
- 提交：`__COMMIT__`

### v0.1.1（2026-11-10）

- 修复：XLSX 字符串单元格 `<is>` 元素（旧版 `<inlineStr>` 会空白），读取端兼容旧格式。
- 修改：删除 `ws.auto_filter.apply()` 别名，仅保留 `set/enable/clear/range`。
- 新增：Wheel 打包与安装说明。
- 验证：66 tests OK；Excel COM 验证 `成绩.xlsx` 内容正常。

## 10. 给下一个 AI 账号的接续指令

- 先读取本文件，复述：当前 v0.1.2 已完成 XLS 自动筛选写入/读取，68 tests OK，Wheel 已发布。
- 下一轮若继续扩展 XLS 能力，建议顺序：合并单元格 → 冻结窗格 → 列宽；每次改动前仍执行阶段 A 落盘并等待授权。
- 注意：所有文件读写用 Python 脚本（避免 W: 盘工具缓存导致旧内容/乱码）。
- 验证：全量 unittest 退出码 0；Excel COM 检查 `AutoFilterMode` 与 `AutoFilter.Range`。
- 不要做：筛选条件/隐藏行、图表/超链接/批注/条件格式/数据验证/保护/auto_fill。
