# ExcelKit精简版 整体情况概览

> 最后更新：2026-11-08 21:10
> 当前版本：v0.1.0（已完成）
> 当前文件：`W:\ExcelKit\ExcelKit项目整体情况概览.md`
> 项目根目录：`W:\ExcelKit\`

## 1. 项目目标与范围

- 目标：基于原 ExcelKit (ksgujie/ExcelKit) 精简出一个以实用为主的轻量级 Excel 读写库，只保留日常办公高频功能，去掉图表/图片/条件格式/保护等装饰性功能。
- 核心功能：Workbook/Worksheet/Cell/Range 对象模型、A1 地址与 0-based 索引、XLSX/XLS/CSV/TSV 读写、单元格值与公式、基础样式、区域批量操作、模板渲染（标量标签 + 循环行块）、基础公式计算。
- 明确不做：图表、图片、批注、超链接、条件格式、数据验证、工作簿/工作表保护、行列插入/删除、文档属性、auto_fill 序列填充、find/replace、render_many/export_pages 等包装循环 API、数据表（Table）、命名区域。

## 2. 当前整体进度

| 模块 | 状态 | 说明 |
| --- | --- | --- |
| 项目骨架与目录 | 已完成 | W:\ExcelKit\excelkit\ |
| 核心对象模型 (Workbook/Worksheet/Cell/Range) | 已完成 | 46 测试全部通过 |
| 文件读写 (XLSX/XLS/CSV/TSV) | 已完成 | XLSX 纯标准库；XLS 依赖 xlrd/xlwt（默认安装） |
| 样式系统 (Style/Font/Fill/Border/Alignment) | 已完成 | |
| 区域操作 (range/values/set_values/clear/copy_to/merge) | 已完成 | |
| 公式计算器与类型转换 | 已完成 | SUM/AVERAGE/IF/跨表等 |
| 模板渲染 (render) | 已完成 | 标量标签 + {loop} 循环行块 |
| 辅助功能 (冻结窗格/合并/自动筛选/页面设置) | 已完成 | |
| 示例与测试 | 已完成 | 46 个单元测试全部通过 |

## 3. 已完成内容

- 分析原 GitHub 项目各版本 (v0.1.0 ~ v0.9.0) 功能，完成功能清单与精简版 API 设计。
- 下载 v0.9.0 源码至 W:\ExcelKit\demo\v0.9.0\ 并阅读 README、CHANGELOG、核心模块与示例。
- 编写精简版完整源码（excelkit 包）及 9 个测试模块共 46 个测试用例，全部通过。
- 代码主目录：`W:\ExcelKit\excelkit\`；测试目录：`W:\ExcelKit\tests\`。

## 4. 未完成/待修改内容

- [ ] 可选：增加更多示例文件（examples/）

## 5. 下一步具体修改点

- 无（核心功能已完成）。如需扩展，可按概览文件“明确不做”列表谨慎评估。

## 6. 关键文件与目录说明

| 路径 | 作用 | 备注 |
| --- | --- | --- |
| `W:\ExcelKit\excelkit\__init__.py` | 顶层公共 API（Workbook/Worksheet/Cell/CellValue/Range） | |
| `W:\ExcelKit\excelkit\core\workbook.py` | Workbook 类 | 工作簿生命周期 |
| `W:\ExcelKit\excelkit\core\worksheet.py` | Worksheet 类 | 工作表核心操作 |
| `W:\ExcelKit\excelkit\core\cell.py` | Cell、CellValue 类 | 单元格读写与转换 |
| `W:\ExcelKit\excelkit\core\range.py` | Range 类 | 区域批量操作 |
| `W:\ExcelKit\excelkit\core\page.py` | 页面设置（简化） | 方向、纸张、边距、打印区域 |
| `W:\ExcelKit\excelkit\reader\xlsx.py` | XLSX 读取 | 标准库 zipfile/xml |
| `W:\ExcelKit\excelkit\reader\xls.py` | XLS 读取 | xlrd |
| `W:\ExcelKit\excelkit\reader\delimited.py` | CSV/TSV 读取 | 标准库 csv |
| `W:\ExcelKit\excelkit\writer\xlsx.py` | XLSX 写出 | 标准库 zipfile/xml |
| `W:\ExcelKit\excelkit\writer\xls.py` | XLS 写出 | xlwt |
| `W:\ExcelKit\excelkit\writer\delimited.py` | CSV/TSV 写出 | 标准库 csv |
| `W:\ExcelKit\excelkit\style.py` | 样式系统 | Font/Fill/Border/Alignment/Style |
| `W:\ExcelKit\excelkit\address.py` | A1 地址解析 | cell_index/range_index 等 |
| `W:\ExcelKit\excelkit\calculation.py` | 公式计算器 | 基础函数 |
| `W:\ExcelKit\excelkit\template.py` | 模板渲染 | 标量标签 + {loop} |
| `W:\ExcelKit\excelkit\errors.py` | 异常定义 | |
| `W:\ExcelKit\tests\test_*.py` | 9 个测试模块，46 个用例 | 全部通过 |

## 7. 运行与验证命令

```powershell
# 运行全部测试（退出码须为 0）
python -m unittest tests.test_address tests.test_workbook tests.test_cell tests.test_append tests.test_range tests.test_formula tests.test_file tests.test_style tests.test_template -v
```
工作目录：`W:\ExcelKit`

- 预期结果：Ran 46 tests, OK。
- 实际结果：46 tests, OK（2026-11-08 21:10 验证）。
- 退出码：0。

## 8. 已知问题、风险与注意事项

- XLS 读写使用 xlrd/xlwt，已列为默认依赖，`pip install .` 时自动安装；XLSX/CSV/TSV 纯标准库。
- 公式计算器不是完整 Excel 引擎，仅支持基础算术、比较和常用函数（SUM/AVERAGE/MIN/MAX/COUNT/COUNTA/IF/AND/OR/NOT/ABS/INT/ROUND/CONCAT/LEN/LEFT/RIGHT/MID 及跨表引用）。
- 所有索引统一为 0-based，行列参数先行后列；A1 地址字符串遵循 Excel 原生表示（1-based）。
- 需在 Python 3.8+ 环境运行。
- XLSX 写出基于标准库 zipfile/xml 实现，无第三方依赖。

## 9. 最近变更记录

### v0.1.0（2026-11-08）

- 新增：基于原 ExcelKit v0.9.0 精简版，仅保留高频实用功能。
- 完成：Workbook/Worksheet/Cell/Range 对象模型、XLSX/XLS/CSV/TSV 读写、基础样式、区域批量操作、模板渲染、基础公式计算。
- 测试：46 个单元测试全部通过。
- 追加：`ws.color` 改为 `ws.tab_color` 并写入/读取 XLSX；删除未落盘的命名区域与数据表（Table）功能；XLS 依赖改为默认安装。
- 风险：无。

## 10. 给下一个 AI 账号的接续指令

- 先读取本文件，复述当前进展：精简版已完成，46 测试全部通过，主目录 `W:\ExcelKit\excelkit\`。
- 如需继续扩展，只在“明确不做”范围外谨慎添加；不要添加图表/图片/条件格式/保护/auto_fill。
- 每次修改代码必须同步更新文档：`README.md`、`ExcelKit_Usage_Guide.md`、本概览文件，并在提交前运行全部单元测试。
- 验证命令：在 `W:\ExcelKit\` 运行上述 unittest 命令，退出码须为 0。
