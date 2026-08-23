# ExcelKit 0.2.1

ExcelKit 是一个使用清晰对象模型读写 XLSX 与 XLS 文件的轻量级库。

逐项参数、返回值、异常及示例请参阅
[《ExcelKit 0.2.1 完整中文使用与 API 手册》](docs/API完整使用手册.md)。

## 安装

```bash
pip install excelkit-0.2.1-py3-none-any.whl
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
```

A1 字符串仍遵循 Excel 原生表示，所以第一格写作 `A1`。`MAX_ROW` 和
`MAX_COLUMN` 表示可用数量，分别为 1048576 和 16384，并非最大索引。空工作表的
`max_row`、`max_column` 都是 `-1`。

## 核心 API

### Workbook

- `Workbook()`：创建空工作簿。
- `add_sheet(name)`：创建并返回工作表。
- `sheet(name_or_index)`：按名称或 0-based 索引返回工作表；负索引遵循 Python 规则。
- `remove_sheet(name_or_index)`：删除工作表并返回当前工作簿。
- `move_sheet(name_or_index, index)`：移动到指定 0-based 最终位置。
- `copy_sheet(name_or_index, new_name)`：完整复制内容、布局和打印设置。
- `sheets`：按创建顺序返回工作表 tuple。
- `active`：返回第一张工作表；空工作簿会创建 `Sheet1`。
- `Workbook.load(filename)`：读取 XLS、XLSX、XLSM、XLTX、CSV 或 TSV。
- `render(data=None, *, by_sheet=None, strict=False)`：使用公共或分工作表数据
  替换模板标签并展开循环行块。
- `save(filename)`：按扩展名保存 XLSX 或 XLS，并返回当前工作簿。

### Worksheet

- `worksheet["A1"]`：日常 A1 单元格访问。
- `worksheet.cell("A1")`：显式 A1 单元格访问。
- `worksheet.cell(row, column)`：0-based 动态行列访问，先行后列。
- `worksheet.range("A1:C10")`：创建连续矩形区域。
- `append(values)`：追加一行；空表从索引 0、即 A1 开始。
- `append_rows(rows)`：连续追加二维数据。
- `values`：返回从 A1 到已触及边界的全部普通值二维列表。
- `max_row`、`max_column`：已经触及的最大 0-based 索引；空表为 `-1`。
- `label`：读取或设置工作表标签名称；设置时同步 Workbook 名称索引。
- `label_color`：读取、设置或清除工作表标签颜色。
- `row(index)`、`column(index)`：设置 0-based 行高、列宽和隐藏状态。
- `merged_ranges`：返回全部合并区域的只读 tuple。
- `freeze`：设置冻结后的第一个可滚动 A1 单元格，`None` 清除。
- `filter_range`：设置或清除连续自动筛选区域。
- `show_gridlines`：控制屏幕网格线。
- `page`：页面布局与打印设置唯一入口。

`cell_at` 已彻底删除。数字坐标和 A1 地址统一由 `cell()` 处理。

工作表名称和标签颜色使用普通属性设置：

```python
worksheet.label = "销售明细"
worksheet.label_color = "4472C4"

assert workbook.sheet("销售明细") is worksheet
assert worksheet.label_color == "FF4472C4"
```

标签颜色可在 XLSX 中保存和读取；旧版 XLS 写出后端不支持标签颜色。

### Cell

- `row`、`column`：0-based 行列索引。
- `address`：对应的规范化 A1 地址。
- `value`：普通值；写入普通值会清除同位置的公式。
- `formula`：公式；可包含或省略 `=`，读取时始终带 `=`；写入公式会清除普通值。
- `style`：完整不可变样式，支持字体、填充、边框、对齐和数字格式。
- `set_value(value)`：写入值并返回当前 Cell，用于链式类型转换。
- `as_string()`、`as_int()`、`as_float()`、`as_bool()`、`as_date()`、
  `as_datetime()`：转换、写回并直接返回目标类型。
- `read()`：取得只读值快照，随后使用同一组 `as_*()` 但不写回。

ExcelKit 只保存公式表达式，不在 Python 中计算公式。

### Range

- `min_row`、`min_column`、`max_row`、`max_column`：0-based 边界索引。
- `values`：以二维 list 读取普通值；公式单元格显示为 `None`。
- `set_values(values)`：写入等形状二维数据并返回当前区域；普通值会覆盖原公式。
- `address`：规范化 A1 区域地址。
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
    parse_range,
    cell_address,
)

assert column_to_index("A") == 0
assert column_to_index("AA") == 26
assert index_to_column(26) == "AA"
assert cell_index("C8") == (7, 2)
assert parse_range("B3:D8") == (2, 1, 7, 3)
assert cell_address(7, 2) == "C8"
```

## 异常

异常只从 `excelkit.errors` 导入：

```python
from excelkit.errors import (
    ExcelKitError,
    InvalidAddressError,
    InvalidWorksheetNameError,
    InvalidFileError,
)
```

- `ExcelKitError`：ExcelKit 自定义异常的共同基类。
- `InvalidAddressError`：A1 地址或 0-based 行列索引无效。
- `InvalidWorksheetNameError`：工作表名称无效。
- `InvalidFileError`：读取的文件格式不支持或结构损坏。

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
from excelkit import Alignment, Border, Fill, Font, Side, Style

worksheet["A1"].style = Style(
    font=Font(name="微软雅黑", size=12, bold=True, color="FFFFFF"),
    fill=Fill("4472C4"),
    border=Border(bottom=Side("thin", "000000")),
    alignment=Alignment(horizontal="center", vertical="center", wrap_text=True),
    number_format="0.00",
)
```

样式对象不可变，可安全复用于多个单元格。XLSX 支持样式完整往返；XLS 会映射基础
样式，但受旧格式调色板和对齐表示能力限制。

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
    by_sheet={
        "封面": {"title": "销售报表"},
        "销售明细": {"items": [...]},
        2: {"total": 1000},  # 0-based工作表索引。
    },
).save("multi_sheet_result.xlsx")
```

使用 `by_sheet` 时只渲染其中列出的工作表，其他表保持不变；所有目标表原子提交，
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
from excelkit import HeaderFooter, PageMargins

worksheet.range("A1:F1").merge()
worksheet.row(0).height = 28
worksheet.column(1).width = 24
worksheet.freeze = "A3"
worksheet.filter_range = "A2:F100"
worksheet.show_gridlines = False

page = worksheet.page
page.orientation = "landscape"
page.paper_size = "A4"
page.fit(width=1)            # 一页宽，高度不限。
page.area = "A1:F100"
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

## 0.2.1 能力边界

本版本包含工作表生命周期管理、合并单元格、行列尺寸、冻结窗格、自动筛选、页面
打印设置、普通值、类型转换、日期时间、公式保存、区域批量写入、基础样式、模板
安全数值表达式、XLS/XLSX 读写及 CSV/TSV 读取。不包含公式计算、条件格式、数据
验证、超链接对象、批注、图表、图片或流式读写。

## 开发验证

```bash
python -m unittest discover -s tests -v
python -m build --wheel --outdir dist
```
