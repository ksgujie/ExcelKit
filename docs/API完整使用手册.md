# ExcelKit 0.1.2 完整中文使用与 API 手册

版本：0.1.2  
适用对象：ExcelKit 使用者、二次开发者和维护者

## 1. 安装与导入

安装 wheel：

```bash
pip install excelkit-0.1.2-py3-none-any.whl
```

稳定核心对象从顶层导入：

```python
from excelkit import (
    Workbook, Worksheet, Cell, Range,
    Style, Font, Fill, Side, Border, Alignment,
    __version__,
)
```

地址和异常分别从唯一模块导入：

```python
from excelkit.address import cell_index, cell_address
from excelkit.errors import InvalidAddressError
```

## 2. 必须先了解的索引规则

ExcelKit 的所有数字索引统一从 0 开始，所有行列参数统一为先行后列：

```python
worksheet.cell(0, 0)  # A1
worksheet.cell(0, 1)  # B1
worksheet.cell(1, 0)  # A2
worksheet.cell(7, 2)  # C8
```

A1 字符串是 Excel 文件格式的原生表示，仍从 `A1` 开始。转换关系如下：

| A1 地址 | 行索引 | 列索引 |
|---|---:|---:|
| A1 | 0 | 0 |
| B1 | 0 | 1 |
| A2 | 1 | 0 |
| C8 | 7 | 2 |

`MAX_ROW = 1048576` 和 `MAX_COLUMN = 16384` 表示可用数量，不是最大索引。
合法最大索引分别为 `1048575` 和 `16383`。

## 3. 顶层版本 API

### `excelkit.__version__`

功能：返回当前安装版本字符串。

类型：`str`。

示例：

```python
import excelkit

assert excelkit.__version__ == "0.1.2"
```

## 4. Workbook 工作簿

### `Workbook()`

功能：创建空工作簿。构造时不自动创建工作表，第一次访问 `active` 或保存空工作簿
时才创建 `Sheet1`。

参数：无。

返回：`Workbook` 实例。

示例：

```python
from excelkit import Workbook

workbook = Workbook()
assert workbook.sheets == ()
```

### `Workbook.add_sheet(name)`

功能：在工作簿末尾创建工作表。

参数：

- `name: str`：工作表名称，长度为 1～31；不能包含 `:`、`\\`、`/`、`?`、
  `*`、`[`、`]` 或 XML 非法字符；名称按不区分大小写的方式去重。

返回：新创建的 `Worksheet`。

异常：名称无效时抛出 `InvalidWorksheetNameError`；名称重复时抛出 `ValueError`。

示例：

```python
scores = workbook.add_sheet("成绩")
summary = workbook.add_sheet("统计")
```

### `Workbook.sheet(name_or_index)`

功能：按名称或创建顺序索引取得工作表。两种查询由同一方法完成。

参数：

- `name_or_index: str | int`：字符串按名称查询；整数按 0-based 索引查询。
  整数支持 Python 负索引，例如 `-1` 表示最后一张。布尔值不作为整数索引。

返回：匹配的 `Worksheet`。

异常：名称不存在时抛出 `KeyError`；索引越界时抛出 `IndexError`；类型错误时抛出
`TypeError`。

示例：

```python
assert workbook.sheet("成绩") is scores
assert workbook.sheet(0) is scores
assert workbook.sheet(-1) is summary
```

### `Workbook.sheets`

功能：取得全部工作表的只读顺序快照。

参数：无，这是只读属性。

返回：`tuple[Worksheet, ...]`。

示例：

```python
for index, worksheet in enumerate(workbook.sheets):
    print(index, worksheet.label)
```

### `Workbook.active`

功能：返回第一张工作表。空工作簿会创建并返回 `Sheet1`。

参数：无，这是只读属性。

返回：`Worksheet`。

示例：

```python
empty_workbook = Workbook()
worksheet = empty_workbook.active
assert worksheet.label == "Sheet1"
assert empty_workbook.active is worksheet
```

### `Workbook.load(filename)`

功能：类方法；从已有表格文件创建新的工作簿。这是唯一公开读取入口，必须通过类
调用，不需要先构造空工作簿。

参数：

- `filename: str | os.PathLike`：源文件路径。支持 `.xls`、`.xlsx`、`.xlsm`、
  `.xltx`、`.csv` 和 `.tsv`，扩展名不区分大小写。

返回：新的 `Workbook`。对子类调用时返回该子类实例。

异常：文件不存在时抛出 `FileNotFoundError`；文件损坏、加密、结构无效或格式不受
支持时抛出 `InvalidFileError`。

示例：

```python
workbook = Workbook.load("input.xlsx")
legacy_workbook = Workbook.load("history.xls")
worksheet = workbook.active
print(worksheet.values)
```

读取能力：

| 格式 | 值 | 日期 | 公式表达式 | 基础样式 |
|---|---|---|---|---|
| XLSX / XLSM / XLTX | 是 | 是 | 是 | 是 |
| XLS | 是 | 是 | 否，仅能取得文件内缓存结果 | 是，受旧格式限制 |
| CSV / TSV | 是 | `#...` 字面量会转换 | 不适用 | 不适用 |

XLSM 中的宏不会执行；0.1.2 也不提供宏对象模型。

### `Workbook.render(data, *, strict=True)`

功能：把当前工作簿作为 Excel 模板，替换普通标签并展开一个或多个循环行块。渲染
直接作用于当前工作簿，成功后可以链式保存；任意工作表失败时整本工作簿保持原状。

参数：

- `data: Mapping[str, Any]`：根模板数据，通常是字典；支持点分路径读取嵌套映射、
  列表的 0-based 数字下标及对象公开属性。
- `strict: bool = True`：严格模式缺少普通标签时抛出 `TemplateError`；设为
  `False` 时原样保留缺失的普通标签。循环集合缺失始终报错。

返回：当前 `Workbook`，支持 `load().render().save()` 链式调用。

异常：`data` 不是映射或 `strict` 不是布尔值时抛出 `TypeError`；标签缺失、循环
结构错误、循环数据不是非映射可迭代对象、公式渲染结果无效或展开超过行数上限时
抛出 `TemplateError`。

完整调用示例：

```python
from excelkit import Workbook

data = {
    "title": "销售明细",
    "customer": {"name": "示例公司"},
    "created_at": "#2026-8-1 12:33",
    "items": [
        {"name": "产品 A", "quantity": 2, "price": 19.5},
        {"name": "产品 B", "quantity": 3, "price": 8},
    ],
}

result = (
    Workbook.load("销售模板.xlsx")
    .render(data)
    .save("销售结果.xlsx")
)
```

#### 普通标签

模板单元格可以只包含标签，也可以把标签嵌入文本：

```text
{title}
客户名称：{customer.name}
生成时间：{created_at}
{customers.0.name}
```

- 整格为 `{quantity}` 时，整数、浮点数、布尔值、日期和 `None` 等类型原样写入。
- `数量：{quantity}` 这种混合内容会生成字符串。
- 标签支持中文、英文、数字、下划线和连字符组成的路径段。
- 渲染结果是 `#2026-8-1` 或带时间形式时，会继续经过统一日期转换。

#### 循环行块

循环开始标签和结束标签必须分别独占一整行，中间可以包含一行或多行模板内容：

```text
{loop items}
{items.@index}    {items.name}    {items.quantity}    {items.price}
                  备注：{items.note}
{/loop}
```

规则：

- 开始标签固定为 `{loop 集合路径}`，结束标签固定为 `{/loop}`；不提供其他别名。
- 当前元素必须使用循环集合最后一个路径段作为前缀。`{loop items}` 对应
  `{items.name}`；`{loop order.items}` 仍对应 `{items.name}`。
- `{items.@index}` 是当前元素的 0-based 索引。
- `{name}` 在循环体内仍表示根数据字段，不会隐式切换为当前元素。
- 空集合会删除开始行、模板内容和结束行，不留下空白模板行。
- 当前版本不允许循环嵌套；同一工作表可以按顺序放置多个互不重叠的循环块。
- 循环标记行可带样式，但不能包含其他普通值或公式。

#### 样式和公式

循环展开会复制模板行中的 `Style`。公式中也可使用数据标签，例如：

```text
=C3*{items.price}
```

复制到其他行时，相对行引用会随目标行调整，`$C$3` 形式的绝对行引用保持不变；
双引号字符串中的 `"C3"` 和 `LOG10(...)` 等函数名不会被当成单元格地址。

模板公式调整是基础 A1 行引用转换，不是完整 Excel 公式解析器。复杂外部引用、结构化
Table 引用、动态数组或需要自动扩张合计区域的场景，应在最终模板中使用绝对引用、
预留范围或由 Excel 打开后重新计算。

### `Workbook.save(filename)`

功能：按扩展名把工作簿原子写出为 XLSX 或 XLS 文件。序列化失败时不会损坏已有
目标文件。

参数：

- `filename: str | os.PathLike`：目标文件路径；`.xls` 选择 Excel 97–2003
  二进制格式，其他扩展名按 XLSX 写出；父目录必须存在。

返回：当前 `Workbook`，可以链式调用。

异常：路径类型错误时抛出 `TypeError`；父目录不存在、无权限或文件系统失败时透传
对应异常。

示例：

```python
from pathlib import Path

result = workbook.save(Path("成绩.xlsx"))
assert result is workbook

workbook.save("兼容旧版.xls")
```

XLS 限制：最多 65536 行、256 列和 56 种自定义调色板颜色；超过限制时抛出
`ValueError`。XLSX 使用 1048576 行、16384 列上限。

### `len(workbook)`

功能：返回当前工作表数量，不触发 `Sheet1` 延迟创建。

参数：无。

返回：非负 `int`。

```python
workbook = Workbook()
assert len(workbook) == 0
workbook.active
assert len(workbook) == 1
```

## 5. Worksheet 工作表

工作表通常由 `Workbook.add_sheet()` 或 `Workbook.active` 获得，不直接构造。

### `Worksheet.label`

功能：读取或修改工作表名称。修改时会校验名称，并同步更新所属 Workbook 的名称
查询索引；工作表对象、顺序、值、公式和样式不会改变。

参数：读取时无参数；设置值必须是长度 1～31 的字符串，不能包含 `:`、`\`、`/`、
`?`、`*`、`[`、`]` 或 XML 非法字符。

返回：读取时返回 `str`；设置时返回 `None`。

异常：名称无效时抛出 `InvalidWorksheetNameError`；与同一工作簿其他工作表名称
大小写不敏感重复时抛出 `ValueError`。失败时原名称和查询索引保持不变。

```python
worksheet = workbook.add_sheet("原名称")
worksheet.label = "新名称"

assert worksheet.label == "新名称"
assert workbook.sheet("新名称") is worksheet
```

### `Worksheet.label_color`

功能：读取、设置或清除 Excel 工作表底部标签颜色。

参数：读取时无参数；设置值可以是 6 位 `RRGGBB`、8 位 `AARRGGBB` 字符串或
`None`。6 位 RGB 自动补 `FF` 不透明度，字母统一转换为大写；`None` 清除颜色。

返回：读取时返回规范化的 8 位 ARGB 字符串；没有颜色时返回 `None`。设置时返回
`None`。

异常：颜色类型、长度或十六进制字符无效时抛出 `ValueError`，原颜色保持不变。

```python
worksheet.label_color = "4472c4"
assert worksheet.label_color == "FF4472C4"

worksheet.label_color = "804472C4"
worksheet.label_color = None
assert worksheet.label_color is None
```

XLSX 保存和加载会保留标签颜色。Excel 97–2003 XLS 的当前 `xlwt` 写出后端没有
标签颜色接口，因此保存为 `.xls` 时忽略该属性；内存中的属性不会因此改变。

### `Worksheet.__getitem__(address)` / `worksheet["A1"]`

功能：按 A1 地址取得 `Cell`，是固定地址的推荐访问方式。

参数：

- `address: str`：合法 A1 单元格地址，例如 `A1`、`C8`。

返回：`Cell`。

异常：地址无效时抛出 `InvalidAddressError`。

```python
cell = worksheet["C8"]
assert cell.row == 7
assert cell.column == 2
```

### `Worksheet.__setitem__(address, value)` / `worksheet["A1"] = value`

功能：按 A1 地址写入普通值。写入会清除同一位置已有公式。

参数：

- `address: str`：合法 A1 地址。
- `value: Any`：普通 Python 值；`None` 删除普通值。

返回：无。

```python
worksheet["A1"] = "姓名"
worksheet["B1"] = 95
worksheet["C1"] = None
```

### `Worksheet.cell(address)`

功能：显式按 A1 地址取得 `Cell`。

参数：

- `address: str`：合法 A1 单元格地址。

返回：`Cell`。

```python
worksheet.cell("B2").value = 88
```

### `Worksheet.cell(row, column)`

功能：按动态数字索引取得 `Cell`。这是唯一的数字行列访问方法，项目不提供
`cell_at`。

参数：

- `row: int`：0-based 行索引，0～1048575。
- `column: int`：0-based 列索引，0～16383。

返回：`Cell`。

异常：参数组合错误时抛出 `TypeError`；索引无效时抛出 `InvalidAddressError`。

```python
for row in range(10):
    for column in range(5):
        worksheet.cell(row, column).value = row * column
```

### `Worksheet.range(address)`

功能：取得连续矩形 `Range`。

参数：

- `address: str`：包含冒号的标准 A1 区域，例如 `A1:C10`。

返回：`Range`。

异常：单个地址、越界地址或反向边界会抛出 `InvalidAddressError`。

```python
area = worksheet.range("A1:C10")
```

### `Worksheet.max_row`

功能：返回已经触及的最大 0-based 行索引。清除值不会缩小历史最大索引。

参数：无，只读属性。

返回：`int`；空表为 `-1`。

```python
worksheet = Workbook().active
assert worksheet.max_row == -1
worksheet["A10"] = "值"
assert worksheet.max_row == 9
```

### `Worksheet.max_column`

功能：返回已经触及的最大 0-based 列索引。

参数：无，只读属性。

返回：`int`；空表为 `-1`。

```python
worksheet["F1"] = "值"
assert worksheet.max_column == 5
```

### `Worksheet.append(values)`

功能：在 `max_row + 1` 处追加一行普通值。空表从行索引 0、即 Excel A1 行开始。

参数：

- `values: Iterable[Any]`：一维行数据；列索引从 0 依次递增。字符串和字节对象
  不能直接作为整行。空可迭代对象不推进最大索引。

返回：当前 `Worksheet`。

异常：参数不是适当可迭代对象时抛出 `TypeError`；超过 Excel 上限时抛出
`InvalidAddressError`。

```python
worksheet.append(["姓名", "成绩"])
worksheet.append(["张三", 95])
```

### `Worksheet.append_rows(rows)`

功能：连续追加二维普通值数据，内部逐行使用 `append()`。

参数：

- `rows: Iterable[Iterable[Any]]`：二维可迭代数据。

返回：当前 `Worksheet`。

异常：数据结构无效时抛出 `TypeError`；超出 Excel 上限时抛出
`InvalidAddressError`。前面已经成功写入的行不会回滚。

```python
worksheet.append_rows([
    ["张三", 95],
    ["李四", 92],
])
```

### `Worksheet.values`

功能：一次读取工作表从 A1 到 `max_row`、`max_column` 的全部普通值。

参数：无，只读属性；不提供整体赋值，批量写入使用 `Range.set_values()`、
`append()` 或 `append_rows()`。

返回：`list[list[Any]]`；空工作表返回 `[]`，空单元格和公式单元格返回 `None`。

示例：

```python
worksheet["B2"] = 10
assert worksheet.values == [
    [None, None],
    [None, 10],
]
```

## 6. Cell 单元格

### `Cell.row`

功能：返回单元格 0-based 行索引。

参数：无，只读属性。

返回：`int`。

```python
assert worksheet["C8"].row == 7
```

### `Cell.column`

功能：返回单元格 0-based 列索引。

参数：无，只读属性。

返回：`int`。

```python
assert worksheet["C8"].column == 2
```

### `Cell.address`

功能：返回规范化 A1 地址。

参数：无，只读属性。

返回：`str`。

```python
assert worksheet.cell(7, 2).address == "C8"
```

### `Cell.value`

功能：读取或设置普通值。公式单元格的普通值为 `None`；设置任何普通值（包括
`None`）都会清除同一位置的公式。

参数：属性写入值为任意 Python 对象。

返回：属性读取时返回普通值或 `None`。

```python
cell = worksheet["A1"]
cell.value = 100
assert cell.value == 100
cell.value = None
assert cell.value is None
```

日期和日期时间字面量会在所有普通值入口统一转换：

```python
from datetime import date, datetime

worksheet["A1"] = "#2026-8-1"
worksheet["A2"] = "#2026/8/1"
worksheet["A3"] = "#2026-8-1 12:33"
worksheet["A4"] = "#2026/8/1 12:33:45"

assert worksheet["A1"].value == date(2026, 8, 1)
assert worksheet["A3"].value == datetime(2026, 8, 1, 12, 33)
```

日期部分可用 `-` 或 `/`，同一个字面量中的分隔符必须一致；时间支持小时、分钟及
可选秒。非法公历日期或时间抛出 `ValueError`，并且不会清除原值或公式。

### `Cell.formula`

功能：读取或设置公式。写入公式会清除同一位置的普通值。ExcelKit 保存表达式，
不计算结果。

参数：写入值必须是包含表达式的非空 `str`；前导 `=` 可省略。

返回：读取时返回带前导 `=` 的标准化公式；无公式时返回 `None`。

异常：空字符串、只有 `=` 或非字符串会抛出 `TypeError`。

```python
worksheet["A1"] = 10
worksheet["A2"] = 20
worksheet["A3"].formula = "SUM(A1:A2)"
assert worksheet["A3"].formula == "=SUM(A1:A2)"
assert worksheet["A3"].value is None
```

### `Cell.style`

功能：读取或设置单元格完整样式。样式包括字体、填充、四边边框、对齐和数字格式。

参数：写入值必须是 `Style`，不接受含义不明确的字典。将 `Style()` 赋给单元格可
恢复默认样式。

返回：读取时返回不可变 `Style`；未设置时返回默认 `Style()`。

异常：赋值不是 `Style` 时抛出 `TypeError`；颜色、字号、边框线型、对齐方式或
数字格式无效时，在构造对应样式对象时抛出 `TypeError` 或 `ValueError`。

```python
from excelkit import Alignment, Border, Fill, Font, Side, Style

title_style = Style(
    font=Font(name="微软雅黑", size=12, bold=True, color="FFFFFF"),
    fill=Fill(color="4472C4"),
    border=Border(bottom=Side(style="thin", color="000000")),
    alignment=Alignment(horizontal="center", vertical="center", wrap_text=True),
    number_format="0.00",
)
worksheet["A1"].style = title_style
assert worksheet["A1"].style is title_style
```

设置纯样式单元格也会更新 `max_row` 和 `max_column`。样式对象不可变且可哈希，推荐
创建一次后复用于多个单元格。

## 7. 样式类型

### `Font(name="Calibri", size=11, bold=False, italic=False, underline=False, color=None)`

功能：定义字体名称、字号、粗体、斜体、下划线和颜色。

参数：`name` 为非空字符串；`size` 为正数；三个字形开关为布尔值；`color` 为
`RRGGBB`、`AARRGGBB` 或 `None`。6 位颜色自动补为不透明 ARGB。

返回：不可变 `Font`。

### `Fill(color=None)`

功能：定义单色实心填充。`None` 表示无填充。

参数：`color` 与 `Font.color` 使用相同颜色规则。

返回：不可变 `Fill`。

### `Side(style=None, color=None)`

功能：定义一条边框边。

参数：`style` 支持 `thin`、`medium`、`thick`、`dashed`、`dotted`、`double`、
`hair`、`dashDot`、`dashDotDot`、`mediumDashed`、`mediumDashDot`、
`mediumDashDotDot`、`slantDashDot` 或 `None`；`color` 为颜色。

返回：不可变 `Side`。

### `Border(left=Side(), right=Side(), top=Side(), bottom=Side())`

功能：组合单元格左、右、上、下四条边。

参数：四个参数都必须是 `Side`。

返回：不可变 `Border`。

### `Alignment(horizontal=None, vertical=None, wrap_text=False)`

功能：定义水平、垂直对齐和自动换行。

参数：水平支持 `general`、`left`、`center`、`right`、`fill`、`justify`、
`centerContinuous`、`distributed`；垂直支持 `top`、`center`、`bottom`、
`justify`、`distributed`；也都可为 `None`。`wrap_text` 为布尔值。

返回：不可变 `Alignment`。

### `Style(font=Font(), fill=Fill(), border=Border(), alignment=Alignment(), number_format="General")`

功能：组合完整单元格样式。

参数：四个组件必须是相应类型；`number_format` 为非空 Excel 数字格式字符串。

返回：不可变 `Style`。

> XLSX 能完整往返上述字段。XLS 只能使用 56 色调色板，ARGB 透明度会被忽略；
> 某些“未设置”对齐值读取后可能成为等价的显式默认值。

## 8. Range 区域

### `Range.min_row` / `min_column` / `max_row` / `max_column`

功能：返回区域四个 0-based 边界索引，属性顺序和所有元组均遵循先行后列。

参数：无，均为只读属性。

返回：`int`。

```python
area = worksheet.range("B3:D8")
assert (area.min_row, area.min_column) == (2, 1)
assert (area.max_row, area.max_column) == (7, 3)
```

### `Range.values`

功能：以二维 list 读取矩形区域的普通值。

参数：无，只读属性；批量写入统一使用 `set_values()`。

返回：`list[list[Any]]`。空单元格和公式单元格均为 `None`。

```python
values = worksheet.range("A1:C2").values
```

### `Range.set_values(values)`

功能：一次性写入与区域形状完全一致的二维普通值。方法先验证完整形状，再开始
写入，所以尺寸错误不会造成部分覆盖。普通值会清除同位置公式。

参数：

- `values: Iterable[Iterable[Any]]`：二维数据，行数和每行列数必须与区域一致。

返回：当前 `Range`。

异常：不可迭代、不是二维结构或形状不一致时抛出 `ValueError`。

```python
area = worksheet.range("A1:C2")
result = area.set_values([
    [1, 2, 3],
    [4, 5, 6],
])
assert result is area
```

## 9. address 地址工具

### `MAX_ROW` / `MAX_COLUMN`

功能：Excel 可用行数和列数常量。

值：`MAX_ROW == 1048576`，`MAX_COLUMN == 16384`。

### `column_to_index(column)`

功能：把 Excel 列字母转换为 0-based 列索引。

参数：`column: str`，仅含 ASCII 字母，不区分大小写，最大为 `XFD`。

返回：0～16383 的 `int`。

异常：无效时抛出 `InvalidAddressError`。

```python
assert column_to_index("A") == 0
assert column_to_index("AA") == 26
```

### `index_to_column(index)`

功能：把 0-based 列索引转换为 Excel 列字母。

参数：`index: int`，范围 0～16383；布尔值被拒绝。

返回：大写列字母 `str`。

```python
assert index_to_column(0) == "A"
assert index_to_column(26) == "AA"
```

### `cell_index(address)`

功能：解析 A1 地址。

参数：`address: str`。

返回：`(row, column)` 0-based 整数元组，顺序为先行后列。

异常：无效时抛出 `InvalidAddressError`。

```python
assert cell_index("C8") == (7, 2)
```

### `parse_range(address)`

功能：解析连续矩形 A1 区域。

参数：`address: str`，必须包含起点和终点。

返回：`(min_row, min_column, max_row, max_column)` 0-based 整数元组。

异常：无效或边界反向时抛出 `InvalidAddressError`。

```python
assert parse_range("B3:D8") == (2, 1, 7, 3)
```

### `cell_address(row, column)`

功能：把 0-based 行列索引转换为 A1 地址。

参数：`row: int`、`column: int`，顺序固定为先行后列。

返回：规范化 A1 地址 `str`。

异常：无效时抛出 `InvalidAddressError`。

```python
assert cell_address(7, 2) == "C8"
```

## 10. errors 异常

### `ExcelKitError`

功能：ExcelKit 自定义异常的共同基类，适合统一捕获库错误。

### `InvalidAddressError`

功能：表示 A1 地址、区域地址或 0-based 行列索引无效。同时继承
`ExcelKitError` 和 `ValueError`。

### `InvalidWorksheetNameError`

功能：表示工作表名称不符合当前 Excel 和 XML 规则。同时继承
`ExcelKitError` 和 `ValueError`。

### `InvalidFileError`

功能：表示读取的文件格式不支持、文件损坏、加密或缺少必要结构。同时继承
`ExcelKitError` 和 `ValueError`。文件路径不存在仍使用标准 `FileNotFoundError`。

### `TemplateError`

功能：表示模板标签缺失、循环标记结构错误、循环数据类型错误、公式渲染无效或循环
展开越界。同时继承 `ExcelKitError` 和 `ValueError`。

```python
from excelkit.errors import TemplateError

try:
    workbook.render({"items": []})
except TemplateError as error:
    print("模板无法渲染：", error)
```

```python
from excelkit.errors import ExcelKitError, InvalidAddressError

try:
    worksheet.cell(-1, 0)
except InvalidAddressError as error:
    print(error)
except ExcelKitError as error:
    print("其他 ExcelKit 错误", error)
```

## 11. XlsxWriter 高级写出 API

模块：`excelkit.writer.xlsx`。业务代码优先使用 `Workbook.save()`。

### `XlsxWriter(workbook)`

功能：创建绑定到指定工作簿的唯一 XLSX 写出器。

参数：`workbook: Workbook`。

返回：`XlsxWriter`。

### `XlsxWriter.write(filename)`

功能：原子写出 XLSX ZIP 包。

参数：`filename: str | os.PathLike`，父目录必须存在。

返回：`None`。

```python
from excelkit.writer.xlsx import XlsxWriter

XlsxWriter(workbook).write("高级写出.xlsx")
```

### `content_types(sheet_count)`

功能：生成 XLSX 根部件 `[Content_Types].xml`。

参数：`sheet_count: int`，大于等于 1 的工作表数量；这是计数而非索引。

返回：UTF-8 XML `bytes`。

异常：不是正整数或传入布尔值时抛出 `ValueError`。

```python
from excelkit.writer.xlsx import content_types

xml_data = content_types(2)
```

### `workbook_xml(sheets)`

功能：生成 `xl/workbook.xml`，包含工作表名称、顺序、sheetId 和关系编号。

参数：`sheets: Sequence[Worksheet]`，按创建顺序排列。

返回：UTF-8 XML `bytes`。

```python
from excelkit.writer.xlsx import workbook_xml

xml_data = workbook_xml(workbook.sheets)
```

### `workbook_rels(sheets)`

功能：生成 `xl/_rels/workbook.xml.rels` 工作表关系清单。

参数：`sheets: Sequence[Worksheet]`。

返回：UTF-8 XML `bytes`。

```python
from excelkit.writer.xlsx import workbook_rels

xml_data = workbook_rels(workbook.sheets)
```

### `cell_xml(address, value)`

功能：把普通 Python 值转换为 SpreadsheetML `c` 元素。

参数：`address: str` 为 A1 地址；`value: Any` 为普通值。

返回：`xml.etree.ElementTree.Element`；`value is None` 时返回 `None`。

```python
from excelkit.writer.xlsx import cell_xml

element = cell_xml("A1", "标题")
```

### `formula_xml(address, formula)`

功能：生成公式单元格 `c` 元素，内部 `f` 文本自动去除前导 `=`。

参数：`address: str` 为 A1 地址；`formula: str` 为非空公式表达式。

返回：`xml.etree.ElementTree.Element`。

异常：公式类型错误或没有表达式时抛出 `TypeError`。

```python
from excelkit.writer.xlsx import formula_xml

element = formula_xml("C2", "=SUM(A2:B2)")
```

### `sheet_xml(sheet)`

功能：把普通值和公式合并排序，生成单张工作表 XML。

参数：`sheet: Worksheet`。

返回：UTF-8 XML `bytes`。

```python
from excelkit.writer.xlsx import sheet_xml

xml_data = sheet_xml(workbook.active)
```

这些底层函数供二次开发和内部测试使用，不建议业务代码依赖。A1 参数遵循 Excel
表示法，工作表内部数字位置仍统一使用 0-based 索引。

### `XlsWriter(workbook)` / `XlsWriter.write(filename)`

模块：`excelkit.writer.xls`。这是旧版二进制格式的高级入口，业务代码仍应优先
使用 `workbook.save("output.xls")`。

功能：`XlsWriter(workbook)` 绑定工作簿；`write(filename)` 原子写出 Excel
97–2003 XLS 文件。

参数：构造参数 `workbook` 为 `Workbook`；`filename` 为字符串或 `os.PathLike`。

返回：构造函数返回 `XlsWriter`，`write()` 返回 `None`。

异常：超过 65536 行、256 列或 56 种自定义颜色时抛出 `ValueError`；路径和 xlwt
序列化错误按原异常透传。

```python
from excelkit.writer.xls import XlsWriter

XlsWriter(workbook).write("兼容旧版.xls")
```

## 12. ValueStore 内部 API

模块：`excelkit.storage`。它不是稳定业务 API。

### `ValueStore()`

功能：创建 0-based 稀疏普通值存储。

参数：无。

返回：`ValueStore` 实例。

### `ValueStore.get(row, column)`

功能：读取指定普通值；不存在时返回 `None`。

参数：0-based `row: int`、`column: int`，顺序为先行后列。

返回：保存的任意 Python 值或 `None`。

### `ValueStore.set(row, column, value)`

功能：设置普通值；`value is None` 时删除该位置。

参数：0-based `row`、`column` 和任意 `value`。

返回：`None`。

### `ValueStore.items()`

功能：按行优先顺序迭代 `((row, column), value)`。

参数：无。

返回：项目迭代器。

```python
from excelkit.storage import ValueStore

store = ValueStore()
store.set(0, 0, "A1")
assert store.get(0, 0) == "A1"
assert list(store.items()) == [((0, 0), "A1")]
```

## 13. 普通值写出规则

| Python 值 | XLSX 表示 |
|---|---|
| `None` | 不生成普通值单元格 |
| `str` | inline string |
| `bool` | Excel Boolean |
| `int` | 数值 |
| 有限 `float` | 数值 |
| `date` / `datetime` | Open XML ISO 日期类型 |
| 其他对象 | `str(value)` 后写成 inline string |

公式单元格写入 `<f>`，不包含 Python 端计算结果。打开文件后由 Excel、WPS 或其他
兼容软件计算公式。

## 14. 可运行示例文件

项目 `examples/` 目录提供：

- `01_basic_workbook.py`：工作簿、工作表和普通值。
- `02_cell_formula.py`：0-based 单元格索引和公式覆盖规则。
- `03_range.py`：区域边界、读取和二维批量写入。
- `04_append.py`：使用 `append()` 和 `append_rows()` 追加单行及多行。
- `05_multi_sheet.py`：名称与 0-based 工作表索引。
- `06_address.py`：全部稳定地址工具。
- `07_xlsx_writer.py`：高级 `XlsxWriter`。
- `08_date_and_values.py`：日期、日期时间字面量和 `Worksheet.values`。
- `09_style.py`：字体、填充、边框、对齐和数字格式。
- `10_load_and_xls.py`：`Workbook.load()` 及 XLS/XLSX 读写。
- `11_template.py`：创建模板并演示标量标签与循环行块渲染。
- `12_worksheet_properties.py`：工作表重命名和标签颜色。
- `create_excel.py`：组合示例。

在项目根目录执行，例如：

```bash
python -m examples.02_cell_formula
```

## 15. 0.1.2 能力边界

0.1.2 不提供模板循环嵌套、完整公式语法重写、公式计算、XLS 公式表达式恢复、
行高、列宽、合并单元格、冻结窗格、
图片、图表、条件格式、数据验证、Table、筛选、打印设置、宏对象模型或流式大文件
处理。XLSM 中的宏只会被忽略，不会执行；保存为其他文件时不会保留宏。
