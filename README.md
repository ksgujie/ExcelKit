# ExcelKit 0.1.2

ExcelKit 是一个使用清晰对象模型读写 XLSX 与 XLS 文件的轻量级库。

逐项参数、返回值、异常及示例请参阅
[《ExcelKit 0.1.2 完整中文使用与 API 手册》](docs/API完整使用手册.md)。

## 安装

```bash
pip install excelkit-0.1.2-py3-none-any.whl
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
- `sheets`：按创建顺序返回工作表 tuple。
- `active`：返回第一张工作表；空工作簿会创建 `Sheet1`。
- `Workbook.load(filename)`：读取 XLS、XLSX、XLSM、XLTX、CSV 或 TSV。
- `render(data, strict=True)`：替换模板标签并展开循环行块。
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

ExcelKit 只保存公式表达式，不在 Python 中计算公式。

### Range

- `min_row`、`min_column`、`max_row`、`max_column`：0-based 边界索引。
- `values`：以二维 list 读取普通值；公式单元格显示为 `None`。
- `set_values(values)`：写入等形状二维数据并返回当前区域；普通值会覆盖原公式。

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
{items.@index}    {items.name}    {items.quantity}    {items.price}
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

- 循环体必须用集合名称作为前缀，如 `{items.name}`；`{name}` 始终指根数据。
- `{items.@index}` 返回当前元素的 0-based 索引。
- 整个单元格只有一个标签时保留原始类型，混合文字时转换为字符串。
- 循环复制会保留样式，并调整复制公式及移动公式的相对行引用。
- `strict=False` 会保留缺失的普通标签；循环集合缺失或结构错误始终抛出
  `TemplateError`。

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

## 0.1.2 能力边界

本版本包含工作簿和工作表管理、普通值、日期时间、公式保存、区域读取与批量写入、
行追加、地址转换、基础样式、模板渲染、XLS/XLSX 读写及 CSV/TSV 读取。不包含公式计算、
图表、图片、合并单元格、数据验证或流式读写。

## 开发验证

```bash
python -m unittest discover -s tests -v
python -m build --wheel --outdir dist
```
