# ExcelKit精简版 整体情况概览

> 最后更新：2026-11-10（v0.1.1：XLSX inline string 修复、Wheel 打包）
> 当前版本：v0.1.1
> 当前文件：`W:\ExcelKit\ExcelKit项目整体情况概览.md`
> 项目根目录：`W:\ExcelKit\`
> Git：`W:\git\cmd\git.exe`，远程 `https://github.com/ksgujie/ExcelKit.git`，分支 `master` 对应 `origin/main`

## 1. 项目目标与范围

- 目标：基于原 ExcelKit (ksgujie/ExcelKit) 精简出一个以实用为主的轻量级 Excel 读写库，只保留日常办公高频功能，去掉图表/条件格式/保护等装饰性功能。
- 核心功能：Workbook/Worksheet/Cell/Range 对象模型、A1 地址与 0-based 索引、XLSX/XLS/CSV/TSV 读写、单元格值与公式、基础样式、区域批量操作、模板渲染（标量标签 + 循环行块）、基础公式计算。
- 明确不做：图表、超链接、批注/备注、条件格式、数据验证、工作簿/工作表保护、行列插入/删除、文档属性、auto_fill 序列填充、增强 find/replace、Table 扩展、命名区域、richData 单元格内图片、render_many/export_pages 等包装循环 API。

## 2. 当前整体进度

| 模块 | 状态 | 说明 |
| --- | --- | --- |
| 项目骨架与目录 | 已完成 | `W:\ExcelKit\excelkit\` |
| 核心对象模型 (Workbook/Worksheet/Cell/Range) | 已完成 | 全量 66 测试通过 |
| 文件读写 (XLSX/XLS/CSV/TSV) | 已完成 | XLSX 纯标准库，XLS 依赖 xlrd/xlwt |
| 样式系统 (Style/Font/Fill/Border/Alignment) | 已完成 | Border 保持现状 |
| 区域操作 (range/values/set_values/clear/copy_to/merge) | 已完成 | |
| 公式计算器与类型转换 | 已完成 | SUM/AVERAGE/IF/跨表等 |
| 模板渲染 (render) | 已完成 | 标量标签 + {loop} 循环行块 |
| 辅助功能 (冻结窗格/合并/自动筛选/页面设置/基础 Table) | 已完成 | auto_filter 只启用筛选区域，不按条件隐藏行；apply() 别名已删除 |
| 文档与测试 | 已完成 | `README.md`、`ExcelKit_Usage_Guide.md`、本概览同步；全量 66 tests OK |
| Wheel 打包 | 已完成 | `dist\excelkit-0.1.1-py3-none-any.whl` |

## 3. 已完成内容

- 分析原 GitHub 项目各版本 (v0.1.0 ~ v0.9.0) 功能，完成功能清单与精简版 API 设计。
- 编写精简版完整源码（excelkit 包）及 10 个测试模块共 66 个测试用例，全部通过。
- 代码主目录：`W:\ExcelKit\excelkit\`；测试目录：`W:\ExcelKit\tests\`。
- 详细中文使用文档：`W:\ExcelKit\ExcelKit_Usage_Guide.md`。

## 4. 最近提交与验证

- 已推送：`8e89991` `Fix XLSX inline string element causing Excel/WPS blank text; improve workbook compatibility`。
- 本轮：版本号升级到 `0.1.1`，删除 `auto_filter.apply()` 别名，新增 Wheel 构建产物与安装说明。
- 实测：全量 **66 tests, OK**（退出码 0）。
- Excel COM 实测修复后的 XLSX：`UsedRange=$A$1:$B$3`，`A1=姓名`、`B1=成绩`、`A2=张三`、`B2=95`。

## 5. 关键修复：XLSX 字符串单元格

- 标准写法：`<c r="A1" t="inlineStr"><is><t>姓名</t></is></c>`。
- 旧版误写：`<c r="A1" t="inlineStr"><inlineStr><t>姓名</t></inlineStr></c>`，Excel/WPS 会显示空白，数字仍正常。
- 写出端已改为 `<is>`；读取端同时兼容 `<is>` 与旧版 `<inlineStr>`。
- 回归测试：`test_xlsx_inline_string_uses_is_element`、`test_xlsx_reader_supports_legacy_inline_str_element`。

## 6. 关键文件与目录说明

| 路径 | 作用 | 备注 |
| --- | --- | --- |
| `W:\ExcelKit\excelkit\__init__.py` | 顶层公共 API 与 `__version__` | |
| `W:\ExcelKit\excelkit\core\workbook.py` | Workbook 类 | 工作簿生命周期 |
| `W:\ExcelKit\excelkit\core\worksheet.py` | Worksheet 类 | 工作表核心操作 |
| `W:\ExcelKit\excelkit\core\cell.py` | Cell、CellValue 类 | 单元格读写与转换 |
| `W:\ExcelKit\excelkit\core\range.py` | Range 类 | 区域批量操作 |
| `W:\ExcelKit\excelkit\core\page.py` | 页面设置（简化） | 方向、纸张、边距、打印区域 |
| `W:\ExcelKit\excelkit\core\table.py` | 基础 Table | 数据表定义 |
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
| `W:\ExcelKit\tests\test_*.py` | 10 个测试模块，66 个用例 | 全部通过 |

## 7. 运行与验证命令

```powershell
# 运行全部测试（退出码须为 0）
python -m unittest tests.test_address tests.test_workbook tests.test_cell tests.test_append tests.test_range tests.test_formula tests.test_file tests.test_style tests.test_template tests.test_image -v
```

工作目录：`W:\ExcelKit`

- 预期结果：Ran 66 tests, OK。
- 实际结果：66 tests, OK（2026-11-10 验证）。
- 退出码：0。

## 8. Wheel 构建与安装

```powershell
cd W:\ExcelKit
python -m pip wheel . --no-deps --no-build-isolation -w dist
pip install --force-reinstall dist\excelkit-0.1.1-py3-none-any.whl
```

安装后确认：

```powershell
python -c "import excelkit; print(excelkit.__version__, excelkit.__file__)"
```

## 9. 已知问题、风险与注意事项

- XLS 读写依赖 xlrd/xlwt，未安装时抛出清晰 ImportError；XLSX/CSV/TSV 纯标准库，不受影响。
- 公式计算器不是完整 Excel 引擎，仅支持基础算术、比较和常用函数（SUM/AVERAGE/MIN/MAX/COUNT/COUNTA/IF/AND/OR/NOT/ABS/INT/ROUND/CONCAT/LEN/LEFT/RIGHT/MID 及跨表引用）。
- 所有索引统一为 0-based，行列参数先行后列；A1 地址字符串遵循 Excel 原生表示（1-based）。
- 需在 Python 3.8+ 环境运行。
- **导入路径风险**：机器上可能同时存在多份 `excelkit`（源码、旧 site-packages、临时快照）。运行程序前先执行 `python -c "import excelkit; print(excelkit.__file__)"`，确认使用的是目标版本。
- `Range.clear` 方案 C 仍未获得 A–E 明确规格与授权，不得直接编码。
- `auto_filter` 仍只启用筛选区域，不按条件隐藏行；`apply()` 别名已删除，`find/replace` 不对称已确认不是问题。

## 10. 给下一个 AI 账号的接续指令

- 先读取本文件，复述当前进展：精简版已完成，66 测试全部通过，主目录 `W:\ExcelKit\excelkit\`，当前版本 v0.1.1。
- 如需继续扩展，只在“明确不做”范围外谨慎添加；不要添加图表/超链接/批注/条件格式/数据验证/保护/auto_fill 等。
- 验证命令：在 `W:\ExcelKit\` 运行上述 unittest 命令，退出码须为 0。
- 发布前重新构建 Wheel，并用全新虚拟环境安装验证。
