# ExcelKit 0.8.2

ExcelKit 是一个使用清晰对象模型读写 XLSX、XLS、CSV 与 TSV 文件的轻量级库。

逐项参数、返回值、异常及示例请参阅
[《ExcelKit 0.8.2 完整中文使用与 API 手册》](docs/API完整使用手册.md)。

## 安装

```bash
pip install excelkit-0.8.2-py3-none-any.whl
```

## 快速开始

```python
from excelkit import Workbook

workbook = Workbook()
worksheet = workbook.active

worksheet.append(["姓名", "成绩", "总分"])
worksheet.append(["张三", 95, None])
worksheet.append(["李四", 88, None])
worksheet["C2"].formula = "=SUM(B2:B2)"
worksheet["C3"].formula = "=SUM(B3:B3)"

workbook.save("成绩.xlsx")
```

## 坐标与索引规则

所有数字索引统一为 **0-based**，所有行列参数统一为 **先行、后列**：

```python
worksheet.cell(0, 0)       # A1
worksheet.cell(7, 2)       # C8

worksheet["C8"].row       # 7
worksheet["C8"].column    # 2
worksheet["C8"].index     # (7, 2)，顺序为先行后列
worksheet["C8"].address   # "C8"
```

A1 字符串仍遵循 Excel 原生表示，所以第一格写作 `A1`。`MAX_ROW` 和
`MAX_COLUMN` 表示可用数量，分别为 1048576 和 16384，并非最大索引。空工作表的
`max_row`、`max_column` 都是 `-1`。

## 核心 API

### Workbook

- `Workbook()`：创建空工作簿。
- `add_sheet(name)`：创建并返回工作表。
- `sheet(name_or_index)`：按不区分大小写的标签或非负 0-based 索引返回工作表。
- `remove_sheet(name_or_index)`：删除工作表并返回当前工作簿。
- `move_sheet(name_or_index, index)`：移动到指定 0-based 最终位置。
- `copy_sheet(name_or_index, new_name)`：完整复制内容、布局和打印设置。
- `add_named_range(name, area)`：创建工作簿级命名区域。
- `named_range(name)`、`named_ranges`、`remove_named_range(name)`：查询、枚举或删除
  命名区域。
- `sheets`：按创建顺序返回工作表 tuple。
- `active`：返回第一张工作表；空工作簿会创建 `Sheet1`。
- `Workbook.load(filename, *, encoding=None, delimiter=None, has_header=False)`：读取
  XLS、XLSX、XLSM、XLTX、CSV 或 TSV；文本文件可指定编码、分隔符和表头。
- `render(data=None, *, sheet_data=None, strict=False)`：使用公共或分工作表数据
  替换模板标签并展开循环行块。
- `calculate(*, strict=False)`：在 Python 中计算当前支持的公式并更新缓存结果。
- `save(filename, *, encoding=..., delimiter=..., formulas=False)`：保存 XLSX、XLS；
  单工作表还可保存 CSV/TSV。

### Worksheet

- `worksheet["A1"]`：日常 A1 单元格访问。
- `worksheet.cell("A1")`：显式 A1 单元格访问。
- `worksheet.cell(row, column)`：0-based 动态行列访问，先行后列。
- `worksheet.range("A1:C10")`：创建连续矩形区域。
- `append(values)`：追加一行；空表从索引 0、即 A1 开始。
- `append_rows(rows)`：原子校验并连续追加二维数据。
- `values`：返回从 A1 到已触及边界的全部普通值二维列表。
- `max_row`、`max_column`：已经触及的最大 0-based 索引；空表为 `-1`。
- `name`：读取或设置工作表名称；设置时同步 Workbook 名称索引。
- `color`：读取、设置或清除工作表标签颜色。
- `row(index)`、`column(index)`：设置 0-based 行高、列宽和隐藏状态。
- `merged_ranges`：返回全部合并区域的只读 tuple。
- `freeze_panes`：设置冻结后的第一个可滚动 A1 单元格，`None` 清除。
- `auto_filter.range`：设置或清除连续筛选区域；`auto_filter.set()` 设置条件并可用 `apply()` 隐藏不匹配行。
- `sort(address, *, keys, has_header=False)`：按 `SortKey` 对连续区域排序。
- `visibility`：使用 `Worksheet.VISIBLE`、`HIDDEN`、`VERY_HIDDEN` 控制可见状态。
- `add_chart(chart_type, anchor=...)`、`charts`：创建柱状、条形、折线或饼图。
- `add_image(filename, anchor=...)`、`images`：插入 PNG/JPEG 图片。
- `cell.note`：读取、设置或清除传统单元格批注。
- `find(query, ...)`、`replace(query, replacement, ...)`：在值或公式中查找、批量替换。
- `export(filename, ...)`：导出当前工作表为 CSV/TSV。
- `show_gridlines`：控制屏幕网格线。
- `page`：页面布局与打印设置唯一入口。
- `add_table(address, *, name, ...)`：在连续区域上创建基础 Excel 数据表；Table 支持
  `columns`、`resize()`、`append()`、`append_rows()`、`clear_data()` 和汇总行。
- `table(name)`、`tables`、`remove_table(name)`：查询、枚举或删除本表数据表定义。

`cell_at` 已彻底删除。数字坐标和 A1 地址统一由 `cell()` 处理。

工作表名称和标签颜色使用普通属性设置：

```python
worksheet.name = "销售明细"
worksheet.color = "4472C4"

assert workbook.sheet("销售明细") is worksheet
assert worksheet.color == "FF4472C4"
```

标签颜色可在 XLSX 中保存和读取；旧版 XLS 写出后端不支持标签颜色。

### Cell

- `row`、`column`：分别返回 0-based 行索引和列索引。
- `index`：以只读 `(row, column)` 元组一次返回 0-based 行列索引。
- `address`：对应的规范化 A1 地址。
- `value`：普通单元格返回已保存值，公式单元格返回当前有效计算结果；写入会清除
  同位置的公式。
- `formula`：公式；可包含或省略 `=`，读取时始终带 `=`；赋值 `None` 清除公式。
- `style`：完整不可变样式，支持字体、填充、边框、对齐和数字格式。
- `formula_status`：返回 `empty`、`pending`、`calculated` 或 `error`。
- `calculation_error`：返回最近一次 Python 公式计算错误说明。
- `copy_style(source)`：从另一个 Cell 复制完整样式，不复制值和公式。
- `set_value(value)`：写入值并返回当前 Cell，用于链式类型转换。
- `as_string()`、`as_int()`、`as_float()`、`as_bool()`、`as_date()`、
  `as_datetime()`：转换、写回并直接返回目标类型。
- `read()`：取得只读值快照，可读取 `.value` 或使用同一组 `as_*()` 而不写回。

`cell.value` 是统一读取入口，不需要先判断单元格是否含公式。公式不会在读取时自动
计算：结果可来自 Excel/WPS 已保存结果或 `workbook.calculate()`；尚无有效结果时为
`None`。修改任意输入值或公式会使工作簿内全部派生结果失效，随后可重新计算。

### Range

- `min_row`、`min_column`、`max_row`、`max_column`：0-based 边界索引。
- `values`：以二维 list 读取有效值；公式单元格显示当前结果，没有有效结果时为 `None`。
- `set_values(values)`：写入等形状二维数据并返回当前区域；普通值会覆盖原公式。
- `address`：规范化 A1 区域地址。
- `clear(values=..., styles=..., hyperlinks=..., notes=...)`：通过一个入口按需清除区域内容和附属信息。
- `copy_to(target, *, values=True, formulas=True, styles=True)`：复制到等尺寸区域，
  并按源目标偏移调整公式中的相对行列引用。
- `merge()`、`unmerge()`：合并或取消合并当前连续区域。

```python
worksheet.range("A1:B2").set_values([
    [1, 2],
    [3, 4],
])
```

## 地址工具

地址工具只从 `excelkit.address` 导入：

```python
from excelkit.address import (
    MAX_ROW,
    MAX_COLUMN,
    column_to_index,
    index_to_column,
    cell_index,
    range_index,
    range_address,
    cell_address,
)

assert column_to_index("A") == 0
assert column_to_index("AA") == 26
assert index_to_column(26) == "AA"
assert cell_index("C8") == (7, 2)
assert range_index("B3:D8") == (2, 1, 7, 3)
assert range_address(2, 1, 7, 3) == "B3:D8"
assert cell_address(7, 2) == "C8"
```

## 异常

异常只从 `excelkit.errors` 导入：

```python
from excelkit.errors import (
    ExcelKitError,
    FormulaCalculationError,
    InvalidAddressError,
    InvalidWorksheetNameError,
    InvalidFileError,
)
```

- `ExcelKitError`：ExcelKit 自定义异常的共同基类。
- `InvalidAddressError`：A1 地址或 0-based 行列索引无效。
- `InvalidWorksheetNameError`：工作表名称无效。
- `InvalidFileError`：读取的文件格式不支持或结构损坏。
- `FormulaCalculationError`：Python 端公式解析、依赖或计算失败。

## 日期与日期时间字面量

普通值入口会统一识别以下写法：

```python
from datetime import date, datetime

worksheet["A1"] = "#2026-8-1"
worksheet["A2"] = "#2026/8/1 12:33"
worksheet["A3"] = "#2026-8-1 12:33:45"

assert worksheet["A1"].value == date(2026, 8, 1)
assert worksheet["A2"].value == datetime(2026, 8, 1, 12, 33)
```

日期分隔符必须一致。无效日期或时间抛出 `ValueError`，不会覆盖原值。

## Cell 类型转换

写回型转换会修改工作簿内存，并影响之后保存的文件：

```python
number = worksheet["A1"].set_value("123").as_int()
assert number == 123
assert worksheet["A1"].value == 123
```

只读型转换通过 `read()` 获取快照，不修改单元格：

```python
worksheet["A1"].value = "123"
number = worksheet["A1"].read().as_int()
assert number == 123
assert worksheet["A1"].value == "123"
```

<p style="color:#C00000"><strong>🔴 重要区别：cell.as_*() 会把转换结果写回单元格；cell.read().as_*() 只转换读取快照，绝不会影响以后保存的 Excel 文件。</strong></p>

六种方法共用一份严格转换规则，不保留 `as_str()` 等同义别名。`as_int()` 不会
截断带小数的浮点数；`as_float()` 拒绝 NaN 和无穷大；`as_bool()` 仅接受
`true/false`、`yes/no`、`是/否`、`1/0` 及等价布尔或数值。

## 单元格样式

```python
from excelkit.style import Alignment, Border, BorderSide, Fill, Font, Style

worksheet["A1"].style = Style(
    font=Font(name="微软雅黑", size=12, bold=True, color="FFFFFF"),
    fill=Fill("4472C4"),
    border=Border(bottom=BorderSide(Border.THIN, "000000")),
    alignment=Alignment(
        horizontal=Alignment.HORIZONTAL_CENTER,
        vertical=Alignment.VERTICAL_CENTER,
        wrap_text=True,
    ),
    number_format="0.00",
)
```

样式对象不可变，可安全复用于多个单元格。XLSX 支持样式完整往返；XLS 会映射基础
样式，但受旧格式调色板和对齐表示能力限制。

只复制样式时以目标单元格调用：

```python
worksheet["B1"].copy_style(worksheet["A1"])
```

## 公式计算与缓存结果

```python
worksheet["A1"] = 10
worksheet["A2"] = 20
worksheet["A3"].formula = "=SUM(A1:A2)"

workbook.calculate()
assert worksheet["A3"].value == 30
assert worksheet["A3"].formula_status == "calculated"
assert worksheet["A3"].read().as_int() == 30
```

当前计算器支持 `+ - * / // % ^`、比较、括号、单格与区域引用、跨表引用，以及
`SUM`、`AVERAGE`、`MIN`、`MAX`、`COUNT`、`COUNTA`、`IF`、`AND`、`OR`、
`NOT`、`ABS`、`INT`、`ROUND`、`CONCAT`、`LEN`、`LEFT`、`RIGHT`、`MID`。
`strict=False` 会把逐格错误记录到 `calculation_error` 并继续；`strict=True` 在
首个错误处抛出 `FormulaCalculationError`。这不是完整 Excel 公式引擎，未支持的
函数仍应交给 Excel/WPS 计算。

## 命名区域和数据表

```python
amounts = workbook.add_named_range(
    "SalesAmount", worksheet.range("C2:C100")
)
assert workbook.named_range("salesamount") is amounts

table = worksheet.add_table(
    "A1:C100",
    name="SalesTable",
    style="TableStyleMedium9",
    has_header=True,
    show_row_stripes=True,
    show_column_stripes=False,
)
```

命名区域属于 Workbook，名称大小写不敏感；工作表重命名后引用自动跟随，删除工作表
时相关命名区域也会删除。Table 属于 Worksheet，但名称在整个工作簿内唯一，且同一
工作表中的 Table 区域不能重叠。两者都支持 XLSX 保存和读取；旧版 `.xls` 只保留
区域内的单元格数据和可表达的基础样式，不保留命名区域或 Table 定义。

## 读取文件

```python
from excelkit import Workbook

workbook = Workbook.load("input.xlsx")
legacy = Workbook.load("input.xls")
csv_book = Workbook.load("input.csv")
```

XLS 读取由 `xlrd` 完成、写出由 `xlwt` 完成；XLSX 使用 ExcelKit 的 Open XML
实现。XLS 公式读取只能取得文件中已有的缓存计算结果，不能恢复公式表达式。

## Excel 模板渲染

先在 Excel 模板中写入标量标签：

```text
客户：{customer.name}
日期：{report_date}
```

需要重复的行块使用独立的开始行和结束行：

```text
{loop items}
{items.@index + 1}    {items.name}    {items.quantity}
{items.price * items.quantity | format:",.2f"}
{/loop}
```

然后加载、渲染并保存：

```python
from excelkit import Workbook

Workbook.load("report_template.xlsx").render({
    "customer": {"name": "示例公司"},
    "report_date": "#2026-8-1 12:33",
    "items": [
        {"name": "产品 A", "quantity": 2, "price": 19.5},
        {"name": "产品 B", "quantity": 3, "price": 8},
    ],
}).save("report.xlsx")
```

多张工作表可以分别接收独立根数据，公共字段会合并到每张目标表，同名字段由工作表
独立数据覆盖：

```python
workbook = Workbook.load("multi_sheet_template.xlsx")
workbook.render(
    {"company": "示例公司", "created_at": "#2026-8-1 12:33"},
    sheet_data={
        "封面": {"title": "销售报表"},
        "销售明细": {"items": [...]},
        2: {"total": 1000},  # 0-based工作表索引。
    },
).save("multi_sheet_result.xlsx")
```

使用 `sheet_data` 时只渲染其中列出的工作表，其他表保持不变；所有目标表原子提交，
任意一张失败都不会留下部分渲染结果。

- 循环体必须用集合名称作为前缀，如 `{items.name}`；`{name}` 始终指根数据。
- `{items.@index}` 返回当前元素的 0-based 索引；`{items.@index + 1}` 从 1 显示。
- 数值表达式支持 `+ - * / // %` 和括号，不允许函数调用或任意 Python 代码。
- `| format:"格式"` 使用 Python 格式规则生成显示字符串，例如 `",.2f"`。
- 整个单元格只有一个标签时保留原始类型，混合文字时转换为字符串。
- 循环复制会保留样式，并调整复制公式及移动公式的相对行引用。
- `strict=False` 是默认值：整格缺失标签清空，混合文本删除标签，缺失循环集合按
  空数组处理，公式缺失模板数据时删除整条公式。
- `strict=True` 会在任意字段或循环集合缺失时抛出 `TemplateError`；循环结构、
  非法表达式、除零及无效格式字符串在两种模式下都会报错。

## 工作表布局与打印

```python
from excelkit.page_setup import HeaderFooter, PageMargins

worksheet.range("A1:F1").merge()
worksheet.row(0).height = 28
worksheet.column(1).width = 24
worksheet.freeze_panes = "A3"
worksheet.auto_filter.range = "A2:F100"
worksheet.show_gridlines = False

page = worksheet.page
page.orientation = "landscape"
page.paper_size = "A4"
page.fit(width=1)            # 一页宽，高度不限。
page.print_area = "A1:F100"
page.repeat_rows = (0, 1)   # 0-based，包含结束行。
page.margins = PageMargins(left=1.5, right=1.5, top=2, bottom=2)
page.header = HeaderFooter(center="销售报表", right="&D")
page.footer = HeaderFooter(center="第 &P 页，共 &N 页", right="&F")
```

设置 `page.fit(...)` 会关闭百分比缩放；重新设置 `page.scale = 90` 会清除适应
页数。边距统一使用厘米。`show_gridlines` 控制屏幕，`page.print_gridlines` 控制打印。

XLSX 能完整往返上述布局与打印设置。XLS 能往返合并、行列尺寸和冻结窗格，并能
写出 `xlwt` 支持的常用页面设置；由于 `xlrd` 不暴露打印设置，XLS 重新读取时不能
完整恢复打印属性。XLS 也不支持本版本的自动筛选、打印区域和重复打印标题写出。

## 支持的普通值

- `None`：清除普通值。
- `str`：XLSX 内联字符串。
- `bool`：Excel 布尔值。
- `int`、有限 `float`：Excel 数值。
- `date`、`datetime`：Open XML 原生 ISO 日期；XLS 中为日期序列及数字格式。
- 其他对象：通过 `str(value)` 写成内联字符串。

## 高级写出入口

业务代码推荐使用 `workbook.save()`。底层或二次开发代码可以使用唯一的写出器：

```python
from excelkit.writer.xlsx import XlsxWriter

XlsxWriter(workbook).write("demo.xlsx")
```

## 可运行示例

`examples/` 目录提供覆盖核心对象、公式、样式、模板、读写、打印和数据整理的独立示例。
其中 [16_search_replace_and_export.py](examples/16_search_replace_and_export.py) 演示查找、
替换和 CSV 导出；[17_visual_sort_filter.py](examples/17_visual_sort_filter.py) 演示图表、
图片、批注、排序、筛选和可见性；[18_business_report.py](examples/18_business_report.py)
演示字典记录、业务 Table、批量公式、自动填充、条件格式、汇总和打印分页。

## 0.8.2 业务报表快捷接口

```python
from excelkit import Workbook
from excelkit.autofill import AutoFillMode
from excelkit.style import NumberFormat, ReportStyle
from excelkit.table import TotalFunction

wb = Workbook()
ws = wb.active
orders = [{"订单号": "SO-001", "数量": 2, "单价": 99.5, "金额": None}]
table = ws.write_table(0, 0, orders, name="Orders", freeze_header=True, auto_fit=True)
ws["E2"].formula = "=C2*D2"
ws.range("E2:E2").auto_fill("E2:E100")
ws.range("E2:E100").format.number = NumberFormat.CURRENCY
ws.range("A1:E1").apply_style(ReportStyle.HEADER)
ws.range("G1:G2").auto_fill("G1:G100", mode=AutoFillMode.SERIES)
table.set_total("金额", TotalFunction.SUM)

# 0-based、包含首尾索引；collapsed=True 时打开文件即为折叠状态。
ws.group_rows(1, 9, collapsed=True)
ws.group_columns(1, 3)
```

`write_records()` 用于把字典列表写入 Excel；区域和 Table 统一通过 `to_records()`
转回字典列表：

```python
records = ws.range("A1:D2").to_records()
records = ws.range("A5:D20").to_records(header_row=1)  # 字段位于工作表第2行。
records = ws.to_records(header_row=1)                  # 整张表只有一个主数据区。
records = table.to_records()
```

`Range.to_records()` 还可传入显式字段名序列；`headers=False` 时自动生成
`Column1`、`Column2` 等字段名。所有 `header_row` 都是工作表绝对 0-based 行索引。

批量模板与分表任务直接使用普通 Python 循环组合 `copy_sheet()`、`render()`、
`add_sheet()` 和 `write_table()`，不再增加只包装循环的专用 API。完整参数和示例见中文
API 手册。

## 0.8.2 能力边界

本版本包含工作表生命周期管理、合并单元格、行列尺寸、冻结窗格、自动筛选、页面
打印设置、普通值、类型转换、日期时间、公式保存与常用公式计算、公式缓存、区域
批量操作、基础样式、命名区域、基础 Table、模板安全数值表达式、XLS/XLSX 读写及
CSV/TSV 读写、超链接、传统批注、数据验证、条件格式、筛选条件、保护、文档属性、
基础图表及 PNG/JPEG 图片写出。Python 公式计算器不是 Excel 全函数兼容引擎；本版本
不包含从外部 XLSX 恢复图表或图片、结构化 Table 公式或流式读写。

## 开发验证

```bash
python -m unittest discover -s tests -v
python -m build --wheel --outdir dist
```
