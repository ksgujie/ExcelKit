# ExcelKit 0.2.2 完整中文使用与 API 手册

版本：0.2.2
适用对象：ExcelKit 使用者、二次开发者和维护者

## 1. 安装与导入

安装 wheel：

```bash
pip install excelkit-0.2.2-py3-none-any.whl
```

稳定核心对象从顶层导入：

```python
from excelkit import (
    Workbook, Worksheet, Cell, CellValue, Range,
    RowDimension, ColumnDimension,
    PageSettings, PageMargins, HeaderFooter,
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

assert excelkit.__version__ == "0.2.2"
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

### `Workbook.remove_sheet(name_or_index)`

功能：按名称或 0-based 索引删除一张工作表。删除后其余工作表保持原有相对顺序；
如果删除的是最后一张工作表，工作簿会变为空，下一次访问 `active` 时再创建
`Sheet1`。

参数：

- `name_or_index: str | int`：字符串按标签名称查找；整数按当前顺序查找，允许
  Python 负索引。布尔值不作为整数。

返回：当前 `Workbook`，可继续链式调用。

异常：名称不存在时抛出 `KeyError`；索引越界时抛出 `IndexError`；类型不正确时
抛出 `TypeError`。查找失败不会修改工作簿。

```python
workbook.remove_sheet("统计")
workbook.remove_sheet(0).save("删除后.xlsx")
```

### `Workbook.move_sheet(name_or_index, index)`

功能：把已有工作表移动到指定的最终位置。索引表示移动完成后的 0-based 位置，
不采用插入前位置，因此调用者不需要自行修正向后移动时的偏移量。

参数：

- `name_or_index: str | int`：要移动的工作表名称或当前索引，查询规则与
  `sheet()` 相同。
- `index: int`：移动完成后的非负 0-based 索引，范围为 `0` 到
  `len(workbook) - 1`；这里不接受负索引和布尔值。

返回：当前 `Workbook`。

异常：目标位置无效时抛出 `IndexError` 或 `TypeError`；工作表查询异常与
`sheet()` 相同。失败时顺序保持不变。

```python
workbook = Workbook()
workbook.add_sheet("一")
workbook.add_sheet("二")
workbook.add_sheet("三")
workbook.move_sheet("三", 0)
assert tuple(sheet.label for sheet in workbook.sheets) == ("三", "一", "二")
```

### `Workbook.copy_sheet(name_or_index, new_name)`

功能：复制一张工作表，并将副本追加到工作簿末尾。普通值、公式、不可变单元格
样式、标签颜色、行列尺寸、合并区域、冻结窗格、筛选、网格线以及全部页面设置
都会复制；副本之后可以独立修改，不会反向影响源表。

参数：

- `name_or_index: str | int`：源工作表名称或索引。
- `new_name: str`：副本名称，使用与 `add_sheet()` 完全相同的校验和大小写不敏感
  去重规则。

返回：新创建的 `Worksheet`。

异常：源表查询异常与 `sheet()` 相同；新名称无效时抛出
`InvalidWorksheetNameError`，重复时抛出 `ValueError`。失败时不会留下半成品副本。

```python
source = workbook.sheet("月报")
copy = workbook.copy_sheet(source.label, "月报副本")
copy["A1"] = "仅修改副本"
assert source["A1"].value != copy["A1"].value
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

XLSM 中的宏不会执行；0.2.2 也不提供宏对象模型。

### `Workbook.render(data=None, *, by_sheet=None, strict=False)`

功能：把当前工作簿作为 Excel 模板，使用一份公共数据渲染全部工作表，或者给多张
指定工作表分别提供独立根数据。渲染直接作用于当前工作簿，成功后可以链式保存；
任意目标工作表失败时，全部目标工作表保持渲染前状态。

参数：

- `data: Mapping[str, Any] | None = None`：所有目标工作表共享的根数据；通常是
  字典，支持点分路径读取嵌套映射、列表的 0-based 数字下标及对象公开属性。
  `None` 等价于空字典。
- `by_sheet: Mapping[str | int, Mapping[str, Any]] | None = None`：可选的分工作表
  数据。键为工作表名称或当前顺序的 0-based 索引，值为该表自己的根字典。同名
  字段以工作表数据为准。指定该参数时只渲染列出的工作表，其他工作表完全不变。
- `strict: bool = False`：默认非严格模式把缺失标签当成空值；设为 `True` 后，
  任意缺失普通字段或循环集合都会抛出 `TemplateError`。

签名中的 `*` 不是参数，它表示 `by_sheet` 和 `strict` 必须按参数名称传递：

```python
workbook.render(data, by_sheet=sheet_data, strict=True)  # 正确
# workbook.render(data, sheet_data, True)                # TypeError
```

返回：当前 `Workbook`，支持 `load().render().save()` 链式调用。

异常：`data`、`by_sheet`、独立根数据或 `strict` 类型错误时抛出 `TypeError`；
工作表名称不存在抛出 `KeyError`，索引越界抛出 `IndexError`，同一张表同时被名称
和索引重复指定时抛出 `ValueError`；严格模式字段缺失、循环结构错误、循环数据不是
非映射可迭代对象、表达式无效或展开超过行数上限时抛出 `TemplateError`。

全部工作表共用一份根数据：

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

每张工作表使用独立数据：

```python
workbook = Workbook.load("多表模板.xlsx")
workbook.render(
    by_sheet={
        "封面": {
            "title": "2026 年销售报表",
            "customer": {"name": "示例公司"},
        },
        "销售明细": {
            "items": [
                {"name": "产品 A", "quantity": 2, "price": 19.5},
                {"name": "产品 B", "quantity": 3, "price": 8},
            ],
        },
        2: {"total": 63},  # 当前第3张工作表，使用0-based索引。
    },
).save("多表结果.xlsx")
```

公共数据与独立数据组合：

```python
workbook.render(
    {
        "company": "示例公司",
        "created_at": "#2026-8-1 12:33",
        "title": "公共标题",
    },
    by_sheet={
        "封面": {"title": "封面标题"},
        "销售明细": {"items": [...]},
    },
)
```

“封面”可直接使用 `{company}`、`{created_at}` 和 `{title}`，其中 `{title}` 的
结果为工作表独立值“封面标题”。“销售明细”也可使用公共字段，并额外使用自己的
`{items...}`。这里是浅层合并，不会复制 `items` 等大型对象。

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

#### 缺失数据和严格模式

默认 `strict=False`，生成结果中不会残留无法识别的 `{标签}`：

| 模板内容 | 数据缺失时的默认结果 |
|---|---|
| `{name}` | 清空整个单元格，普通值为 `None` |
| `姓名：{name}` | 保留固定文字，结果为 `姓名：` |
| `{price * quantity}` | 清空整个单元格 |
| `金额：{price * quantity}` | 结果为 `金额：` |
| 缺少 `{loop items}` 的 `items` | 当作空数组，删除整个循环块 |
| 公式中任一模板字段缺失或为 `None` | 删除整条公式，避免产生无效 Excel 公式 |

显式使用 `strict=True` 时，上述字段缺失全部抛出 `TemplateError`，且所有目标
工作表原子回滚。循环标记未配对、循环嵌套、非数组循环数据、非法表达式、除数为零、
无效格式字符串等真正的模板错误不受 `strict` 影响，两种模式下一律报错。

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

#### 模板数值表达式

标签中可以直接进行受限数值计算。表达式支持整数、有限浮点数、数据路径、圆括号、
一元正负号和 `+`、`-`、`*`、`/`、`//`、`%`；不执行 Python 函数、属性方法、
下标表达式、比较、逻辑运算或任意代码。

```text
{items.@index + 1}
{items.quantity * items.price}
{(items.price - items.discount) * items.quantity}
{items.total / 100}
```

把第一条数据的显示序号从 0 改为 1，直接使用：

```text
{items.@index + 1}
```

表达式中的所有运算数都必须是数值，布尔值不作为整数参与计算。除数为零、非数值
字段、过长或过于复杂的表达式都会抛出 `TemplateError`。默认模式下缺少字段的
表达式按空值处理；`strict=True` 时缺少字段报错。已经找到字段但计算本身无效时，
两种模式都会报错。

#### 模板显示格式

表达式末尾可追加唯一的 `format` 过滤器，格式字符串遵循 Python 内置
`format(value, spec)` 的格式规范：

```text
{items.price | format:",.2f"}
{items.quantity * items.price | format:",.2f"}
{items.ratio | format:".1%"}
{items.code | format:"04d"}
```

参数说明：`format:` 后面的内容必须使用成对单引号或双引号包围；格式无效时抛出
`TemplateError`。使用过滤器后结果一定是显示字符串。如果需要真正的 Excel 数值
以便继续参与公式计算，应让整格只写数值表达式，并通过 `Cell.style.number_format`
控制 Excel 显示格式。

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

### `Worksheet.row(index)`

功能：取得指定行的持久化尺寸对象，用于设置行高和隐藏状态。重复传入同一索引会
返回同一个 `RowDimension` 对象。

参数：`index: int`，0-based 行索引，范围为 0～1048575；布尔值无效。

返回：`RowDimension`。仅取得对象不会扩大 `max_row`，实际设置行尺寸也不创建
普通值单元格。

异常：索引类型或范围无效时抛出 `InvalidAddressError`。

```python
row = worksheet.row(0)
row.height = 28       # 磅
row.hidden = False
```

### `Worksheet.column(index)`

功能：取得指定列的持久化尺寸对象，用于设置 Excel 列宽和隐藏状态。

参数：`index: int`，0-based 列索引，范围为 0～16383；布尔值无效。

返回：`ColumnDimension`。重复访问同一索引返回同一个对象。

异常：索引类型或范围无效时抛出 `InvalidAddressError`。

```python
column = worksheet.column(1)
column.width = 24
column.hidden = False
```

### `Worksheet.merged_ranges`

功能：取得当前工作表全部合并区域的只读快照，顺序按左上角位置排列。

参数：无，只读属性。

返回：`tuple[Range, ...]`；没有合并时返回空 tuple。返回的 `Range` 可读取
`address`，也可调用 `unmerge()`。

```python
worksheet.range("A1:F1").merge()
assert tuple(area.address for area in worksheet.merged_ranges) == ("A1:F1",)
```

### `Worksheet.freeze`

功能：读取、设置或清除冻结窗格。地址表示冻结后左上角第一个仍可滚动的单元格；
因此 `"A3"` 冻结前两行，`"C1"` 冻结前两列，`"C3"` 同时冻结前两行两列。

参数：设置值为单个 A1 地址或 `None`。`"A1"` 等价于 `None`，因为其上方和
左侧都没有可冻结内容。

返回：读取时返回规范化 A1 地址或 `None`；设置时返回 `None`。

异常：地址无效时抛出 `InvalidAddressError`；非字符串且非 `None` 时抛出
`TypeError`。

```python
worksheet.freeze = "A3"
assert worksheet.freeze == "A3"
worksheet.freeze = None
```

### `Worksheet.filter_range`

功能：设置连续矩形区域的自动筛选按钮，或读取、清除现有筛选区域。它只定义
筛选范围，不在 Python 内存中隐藏不符合条件的数据行。

参数：设置值为 A1 区域字符串（例如 `"A1:F100"`）或 `None`。

返回：读取时返回规范化区域地址或 `None`；设置时返回 `None`。

异常：单格地址、反向区域或越界地址抛出 `InvalidAddressError`。

```python
worksheet.filter_range = "A2:F100"
worksheet.filter_range = None
```

### `Worksheet.show_gridlines`

功能：控制工作表在 Excel/WPS 窗口中的屏幕网格线。它与打印网格线
`worksheet.page.print_gridlines` 是两个互不影响的设置。

参数：设置值必须是 `bool`。

返回：读取时返回 `bool`，默认 `True`；设置时返回 `None`。

异常：非布尔值抛出 `TypeError`。

```python
worksheet.show_gridlines = False
worksheet.page.print_gridlines = True
```

### `Worksheet.page`

功能：返回该工作表唯一的页面布局与打印设置对象。属性本身只读；其内部字段可
修改，详见“页面布局与打印 API”。

参数：无，只读属性。

返回：`PageSettings`；每次访问同一工作表都返回同一个对象。

```python
worksheet.page.orientation = "landscape"
worksheet.page.fit(width=1)
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

### `RowDimension.index` / `height` / `hidden`

功能：`index` 返回所属行的只读 0-based 索引；`height` 读取或设置行高（单位为
磅），`None` 恢复应用程序默认行高；`hidden` 控制是否隐藏整行。

参数：`height` 接受大于 0 且不超过 409 的有限 `int | float` 或 `None`；
`hidden` 必须是 `bool`。

返回：`index` 为 `int`，`height` 为 `float | None`，`hidden` 为 `bool`；属性设置
返回 `None`。

异常：尺寸不是有效有限数或超出范围时抛出 `ValueError`；类型错误抛出
`TypeError`。

```python
row = worksheet.row(2)
assert row.index == 2
row.height = 30
row.hidden = True
row.height = None
```

### `ColumnDimension.index` / `width` / `hidden`

功能：`index` 返回只读 0-based 列索引；`width` 使用 Excel 字符宽度单位设置列宽，
`None` 恢复默认列宽；`hidden` 控制是否隐藏整列。

参数：`width` 接受大于 0 且不超过 255 的有限 `int | float` 或 `None`；
`hidden` 必须是 `bool`。

返回和异常：与 `RowDimension` 对应属性相同。

```python
column = worksheet.column(1)
assert column.index == 1
column.width = 24
column.hidden = True
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

### `Cell.index`

功能：一次返回单元格的 0-based 行列组合索引。该属性由单元格现有行列位置动态
组成，不创建或维护第二套坐标数据。

参数：无，只读属性；结果元素顺序固定为先行、后列。

返回：`tuple[int, int]`，内容为 `(row, column)`。

异常：属性没有设置器，尝试对 `cell.index` 赋值会抛出 `AttributeError`。需要访问
其他坐标时，应通过 `worksheet.cell(row, column)` 或 A1 地址取得另一个单元格对象。

```python
cell = worksheet["D3"]

assert cell.index == (2, 3)
assert cell.row == 2
assert cell.column == 3
```

`index` 适合传递、解包或比较完整坐标；只需要一个维度时，直接使用 `row` 或
`column` 更清楚：

```python
row, column = worksheet["D3"].index
assert (row, column) == (2, 3)
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

### `Cell.set_value(value)`

功能：写入普通值并返回当前 `Cell`，是需要立即调用写回型 `as_*()` 时的链式入口。
它与 `cell.value = value` 使用同一套写入和日期字面量识别规则；写入会清除公式。

参数：`value: Any`，任意普通 Python 值。

返回：当前 `Cell`。

异常：日期字面量非法时抛出 `ValueError`；合并区域非左上角单元格禁止写入时抛出
`ValueError`。写入失败不会改变原值。

```python
result = worksheet["A1"].set_value("2026-8-1").as_date()
assert result == date(2026, 8, 1)
assert worksheet["A1"].value == result
```

### `Cell.as_string()` / `as_int()` / `as_float()` / `as_bool()` / `as_date()` / `as_datetime()`

功能：把当前普通值转换成指定类型，**把转换结果写回同一单元格**，然后直接返回
目标类型的结果。写回意味着结果会出现在之后保存的 XLS/XLSX 文件中。

参数：六个方法均无参数。

返回：依次为 `str`、`int`、`float`、`bool`、`datetime.date` 或
`datetime.datetime`。

转换规则：

| 方法 | 接受的常用输入 | 关键限制 |
|---|---|---|
| `as_string()` | 任意值 | `None` 明确转换为空字符串 |
| `as_int()` | 整数、整数值浮点数、整数字符串 | 不接受布尔值；不截断 `1.2`，拒绝 NaN/无穷大 |
| `as_float()` | 有限整数、浮点数或数字字符串 | 不接受布尔值；结果必须有限 |
| `as_bool()` | 布尔值、0/1、`true/false`、`yes/no`、`是/否` | 其他数字和文本无效，忽略文本首尾空格和大小写 |
| `as_date()` | `date`、`datetime`、`YYYY-M-D` 或 `YYYY/M/D` 文本 | `datetime` 只取日期部分 |
| `as_datetime()` | `datetime`、`date`、日期或日期时间文本 | `date` 转换为当天 00:00:00 |

日期转换文本既可以带 `#`，也可以不带；日期时间支持 `HH:MM` 和 `HH:MM:SS`。
转换失败抛出 `TypeError` 或 `ValueError`，原单元格保持不变。

```python
worksheet["A1"].set_value("123").as_int()
worksheet["A2"].set_value("是").as_bool()
worksheet["A3"].set_value("2026/8/1 12:33").as_datetime()
```

### `Cell.read()`

功能：读取当前普通值并创建一个独立的 `CellValue` 快照。快照随后可使用同名
`as_*()` 转换，但不会写回工作表。

参数：无。

返回：`CellValue`。创建后即使原单元格改变，快照仍保留读取时的值。

```python
worksheet["A1"].value = "123"
snapshot = worksheet["A1"].read()
worksheet["A1"].value = "456"
assert snapshot.as_int() == 123
```

### `CellValue.as_string()` / `as_int()` / `as_float()` / `as_bool()` / `as_date()` / `as_datetime()`

功能：按照与 `Cell.as_*()` 完全相同的严格规则转换只读快照，只返回临时结果，
**绝不修改工作簿**。

参数：六个方法均无参数。

返回和异常：与对应 `Cell.as_*()` 完全相同。

```python
worksheet["A1"].value = "123"
temporary = worksheet["A1"].read().as_int()
assert temporary == 123
assert worksheet["A1"].value == "123"
```

<p style="color:#C00000"><strong>🔴 重要区别：cell.as_*() 会把转换结果写回单元格并影响保存文件；cell.read().as_*() 只转换读取快照，绝不会改变 Excel 文件。需要写回时使用 set_value(...).as_*()，只为当前 Python 代码临时取值时使用 read().as_*()。</strong></p>

API 只保留语义明确的 `as_string()`，不提供 `as_str()`；也不提供与 `value` 属性
冲突的 `value()` 方法。

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

### `Range.address`

功能：返回由四个边界组成的规范化大写 A1 区域地址。

参数：无，只读属性。

返回：`str`，例如 `"B3:D8"`。

```python
assert worksheet.range("b3:d8").address == "B3:D8"
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

### `Range.merge()`

功能：合并当前矩形区域。合并后只允许左上角锚点保存普通值或公式；为防止数据
丢失，合并前区域内除左上角外必须没有普通值或公式。各位置样式保持不变。对同一
地址重复调用是幂等的。

参数：无。

返回：当前 `Range`，支持链式调用。

异常：与已有合并区域重叠但地址不完全相同时抛出 `ValueError`，原结构保持不变。

```python
title = worksheet.range("A1:F1").merge()
worksheet["A1"] = "销售报表"
assert title.address == "A1:F1"
```

### `Range.unmerge()`

功能：取消与当前地址完全相同的合并区域。取消后原锚点值仍在左上角，其余格保持
空白。

参数：无。

返回：当前 `Range`。当前地址不是一个完整合并区域时抛出 `ValueError`。

```python
worksheet.range("A1:F1").unmerge()
```

## 9. 页面布局与打印 API

页面设置从 `worksheet.page` 进入，不另外创建或替换 `PageSettings`。默认值为 A4、
纵向、100% 缩放、常用厘米边距，不打印网格线和标题。

### `PageSettings.orientation`

功能：读取或设置打印方向。

参数：设置值为 `"portrait"`（纵向）或 `"landscape"`（横向）。

返回：读取时为 `str`；设置时为 `None`。无效枚举值抛出 `ValueError`。

```python
worksheet.page.orientation = "landscape"
```

### `PageSettings.paper_size`

功能：读取或设置常用纸张规格。

参数：`"A3"`、`"A4"`、`"A5"`、`"Letter"` 或 `"Legal"`，输入不区分
大小写，读取时返回表中规范写法。

返回：读取时为 `str`；设置时为 `None`。不支持的规格抛出 `ValueError`。

```python
worksheet.page.paper_size = "A4"
```

### `PageSettings.scale`

功能：读取或设置打印缩放百分比。设置具体百分比会自动清除先前的适应页数设置，
避免两套互斥配置同时生效。

参数：10～400 的整数或 `None`；布尔值无效。

返回：读取时为 `int | None`；设置时为 `None`。范围无效抛出 `ValueError`。

```python
worksheet.page.scale = 90
```

### `PageSettings.fit(width=1, height=None)`

功能：一次设置“将内容缩放到几页宽、几页高”。它代替两个容易漏配的
`fit_width`、`fit_height` 属性；调用后自动把 `scale` 设为 `None`。

参数：

- `width: int | None = 1`：横向页数，默认 1；`None` 表示宽度不限。
- `height: int | None = None`：纵向页数；`None` 表示高度不限。
- 两者至少一项非 `None`，具体值必须是正整数，布尔值无效。

返回：当前 `PageSettings`，支持链式使用。参数无效抛出 `ValueError`。

```python
page = worksheet.page
page.fit()                    # 等价于 fit(width=1, height=None)：一页宽，高度不限
page.fit(width=1, height=1)  # 整张打印区域适应一页
page.fit(width=None, height=2)  # 高度两页，宽度不限
```

`fit()` 最常用的“一页宽”场景只需一次调用，不需要同时写两个属性。若随后执行
`page.scale = 90`，适应页数模式会被清除。

### `PageSettings.first_page_number`

功能：设置打印时显示的起始页码；`None` 让 Excel 自动从 1 开始。

参数：正整数或 `None`。返回：`int | None`。非正整数抛出 `ValueError`。

```python
worksheet.page.first_page_number = 5
```

### `PageSettings.black_and_white` / `draft`

功能：分别控制黑白打印和草稿质量打印。

参数：两个属性都只接受 `bool`。返回：读取时为 `bool`，默认 `False`。

```python
worksheet.page.black_and_white = True
worksheet.page.draft = False
```

### `PageSettings.order`

功能：当打印区域横向和纵向都跨页时，指定页面编号和打印顺序。

参数：`"down_then_over"` 表示先向下再向右；`"over_then_down"` 表示先向右
再向下。返回：规范字符串。无效值抛出 `ValueError`。

```python
worksheet.page.order = "down_then_over"
```

### `PageMargins(left, right, top, bottom, header, footer)`

功能：创建不可变打印边距值对象。ExcelKit 的公开单位统一为厘米，写出 XLSX 时
自动换算为文件格式要求的英寸。

参数：六个字段均为有限的非负 `int | float`，默认依次为 1.78、1.78、1.91、
1.91、0.76、0.76 厘米。布尔值、负数、NaN 和无穷大无效。

返回：不可变 `PageMargins`。错误类型或范围抛出 `TypeError` 或 `ValueError`。

```python
from excelkit import PageMargins

worksheet.page.margins = PageMargins(
    left=1.5, right=1.5, top=2.0, bottom=2.0,
    header=0.8, footer=0.8,
)
```

`PageSettings.margins` 只接受完整的 `PageMargins`，这样六项设置作为一个不可变值
整体替换，不会出现部分更新失败。

### `PageSettings.area`

功能：读取、设置或清除打印区域。

参数：标准 A1 矩形区域字符串或 `None`。返回：规范化大写地址或 `None`。
无效区域抛出 `InvalidAddressError`。

```python
worksheet.page.area = "A1:F100"
worksheet.page.area = None
```

### `PageSettings.repeat_rows` / `repeat_columns`

功能：设置每一打印页顶部重复的标题行，或每页左侧重复的标题列。

参数：包含式起止索引二元组 `(start, end)` 或 `None`，所有索引均为 0-based；
`repeat_rows` 校验行范围，`repeat_columns` 校验列范围，且起点不得大于终点。

返回：对应 tuple 或 `None`。结构、类型、范围或顺序无效时抛出 `TypeError`、
`InvalidAddressError` 或 `ValueError`。

```python
worksheet.page.repeat_rows = (0, 1)     # 每页重复第 1～2 行
worksheet.page.repeat_columns = (0, 0)  # 每页重复 A 列
```

### `PageSettings.center_horizontal` / `center_vertical`

功能：控制打印内容是否在纸张的水平或垂直方向居中。

参数：只接受 `bool`。返回：读取时为 `bool`，默认 `False`。

### `PageSettings.print_gridlines` / `print_headings`

功能：分别控制是否打印单元格网格线，以及是否打印 A/B/C 列标和 1/2/3 行号。
这两个设置不改变屏幕显示。

参数：只接受 `bool`。返回：读取时为 `bool`，默认 `False`。

```python
worksheet.page.print_gridlines = True
worksheet.page.print_headings = True
```

### `HeaderFooter(left="", center="", right="")`

功能：创建不可变的页眉或页脚三区域值对象。ExcelKit 在序列化时自动加入
`&L`、`&C`、`&R` 区域标记，调用者只在 `left`、`center`、`right` 中填写要显示
的文本和动态控制符。

参数：三个字段都必须是 `str`，默认空字符串。返回：不可变 `HeaderFooter`；
类型错误抛出 `TypeError`。

```python
from excelkit import HeaderFooter

page = worksheet.page
page.header = HeaderFooter(
    left="ExcelKit",
    center="销售报表",
    right="&D &T",
)
page.footer = HeaderFooter(
    left="&F",
    center="第 &P 页，共 &N 页",
    right="&A",
)
```

`PageSettings.header` 和 `footer` 只接受 `HeaderFooter`；传入空对象
`HeaderFooter()` 可清空对应内容。

#### 页眉/页脚全部特殊控制符

下表按 Microsoft Excel 公布的页眉/页脚代码完整列出。ExcelKit 会把这些文本代码
写入 XLSX/XLS，由打开文件的 Excel、WPS 等应用在打印或预览时解释；显示细节可能
随应用而异。

| 控制符 | 功能 | 示例 | ExcelKit 0.2.2 |
|---|---|---|---|
| `&L` | 后续内容进入左侧区域 | `&L公司` | 自动生成；通常不要手写 |
| `&C` | 后续内容进入中间区域 | `&C月报` | 自动生成；通常不要手写 |
| `&R` | 后续内容进入右侧区域 | `&R&D` | 自动生成；通常不要手写 |
| `&P` | 当前页码 | `第 &P 页` | 支持 |
| `&P+数字` | 当前页码加指定数 | `&P+1` | 支持，由表格应用解释 |
| `&P-数字` | 当前页码减指定数 | `&P-1` | 支持，由表格应用解释 |
| `&N` | 当前文档总页数 | `共 &N 页` | 支持 |
| `&D` | 打印时的当前日期 | `打印日期：&D` | 支持 |
| `&T` | 打印时的当前时间 | `打印时间：&T` | 支持 |
| `&F` | 工作簿文件名 | `文件：&F` | 支持 |
| `&A` | 当前工作表标签名称 | `工作表：&A` | 支持 |
| `&Z` | 工作簿文件路径 | `路径：&Z` | 支持 |
| `&&` | 显示一个普通 `&` 字符 | `研发 && 销售` | 支持 |
| `&B` | 开启/关闭粗体；再次出现即关闭 | `&B重要&B` | 支持 |
| `&I` | 开启/关闭斜体 | `&I斜体&I` | 支持 |
| `&U` | 开启/关闭单下划线 | `&U下划线&U` | 支持 |
| `&E` | 开启/关闭双下划线 | `&E双下划线&E` | 支持 |
| `&S` | 开启/关闭删除线 | `&S作废&S` | 支持 |
| `&X` | 开启/关闭上标 | `m&X2&X` | 支持 |
| `&Y` | 开启/关闭下标 | `H&Y2&YO` | 支持 |
| `&"字体名,字形"` | 设置后续字体及字形 | `&"微软雅黑,Bold"标题` | 支持，具体字体需系统存在 |
| `&nn` | 设置后续字号，常用两位磅值 | `&14标题` | 支持 |
| `&Krrggbb` / `&color` | 使用六位十六进制文字颜色 | `&KFF0000红色` | XLSX 支持；旧版 XLS/应用兼容性有限 |
| `&"+"` | 使用当前主题的标题字体 | `&"+"标题` | 支持，由表格应用解释 |
| `&"-"` | 使用当前主题的正文字体 | `&"-"正文` | 支持，由表格应用解释 |
| `&Kxx.Snnn` | 使用主题颜色；`xx` 为 01～12，`S` 为 `+`/`-`，`nnn` 为 000～100 的明暗百分比 | `&K04.+050文字` | XLSX 支持；由表格应用解释 |
| `&G` | 插入页眉/页脚图片 | `&G` | **暂不支持**；0.2.2 不创建图片关系和媒体文件 |

格式开关是切换式的。例如 `&B重要&B普通` 只让“重要”变粗。要显示普通 `&`，必须
写成 `&&`。`HeaderFooter` 的 `left`、`center`、`right` 已经代表三个区域，所以
不要在字段内容中再嵌套 `&L`、`&C` 或 `&R`。

控制符含义依据 Microsoft 官方
[Excel 页眉页脚格式与 VBA 代码](https://learn.microsoft.com/en-us/office/vba/excel/concepts/workbooks-and-worksheets/formatting-and-vba-codes-for-headers-and-footers)
及 [MS-XLS Header 记录规范](https://learn.microsoft.com/en-us/openspecs/office_file_formats/ms-xls/b64cf6b8-9472-4f97-9a69-d839f0fa1089)。

最常见的完整页脚：

```python
page.footer = HeaderFooter(
    left="ExcelKit",
    center="第 &P 页，共 &N 页",
    right="&D",
)
```

### 页面设置格式兼容性

XLSX 可保存并重新读取本节全部非图片设置。XLS 写出支持方向、纸张、缩放/适应
页数、页序、黑白/草稿、居中、网格线/标题、边距、页眉页脚、行列尺寸、合并和冻结；
受 `xlwt` 接口限制，XLS 暂不写出打印区域、重复标题和自定义起始页码。`xlrd`
不会公开 XLS 页面设置，因此从 XLS 加载时这些打印属性不能恢复。`&G` 图片在两种
格式中均暂不支持。

## 10. address 地址工具

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

## 11. errors 异常

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

## 12. XlsxWriter 高级写出 API

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

## 13. ValueStore 内部 API

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

## 14. 普通值写出规则

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

## 15. 可运行示例文件

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
- `13_layout_and_print.py`：工作表管理、合并、尺寸、视图与完整打印设置，同时生成
  XLSX 和 XLS 示例文件。
- `create_excel.py`：组合示例。

在项目根目录执行，例如：

```bash
python -m examples.02_cell_formula
```

## 16. 0.2.2 能力边界

0.2.2 不提供模板循环嵌套、完整公式语法重写、Python 端公式计算、XLS 公式表达式
恢复、页眉页脚图片、普通图片、图表、条件格式、数据验证、Table、筛选条件执行、
宏对象模型或流式大文件处理。

模板循环展开会复制单元格值、公式和样式，但不会自动移动或扩张模板中已有的合并
区域、冻结位置、筛选范围和打印区域。需要动态结构时，应在渲染后通过对应 0.2.2
API 显式设置。

XLSM 中的宏只会被忽略，不会执行；保存为其他文件时不会保留宏。旧版 XLS 受
65536 行、256 列、56 色调色板和第三方后端能力限制；页面设置兼容性详见本手册
“页面设置格式兼容性”。

## 17. API 选择指南

| 场景 | 推荐 API | 原因 |
|---|---|---|
| 固定 A1 单元格 | `worksheet["A1"]` | 最短、最直观 |
| 动态行列坐标 | `worksheet.cell(row, column)` | 统一 0-based、先行后列 |
| 连续矩形数据 | `worksheet.range("A1:C10")` | 集中读取、写入、合并 |
| 连续追加二维数据 | `worksheet.append_rows(rows)` | 自动从下一空行开始 |
| 按名称或顺序取表 | `workbook.sheet(name_or_index)` | 一个方法覆盖两种清晰参数类型 |
| 永久改变单元格类型 | `cell.set_value(value).as_int()` 等 | 转换结果写回并保存 |
| 只在 Python 中临时转换 | `cell.read().as_int()` 等 | 不改变工作簿 |
| 打印适应一页宽 | `worksheet.page.fit()` | 一次调用，不必维护两个属性 |
| Excel 模板批量生成 | `Workbook.load(...).render(...).save(...)` | 保留模板内容和样式 |

不提供 `Workbook.create()`、`Worksheet.cell_at()`、`as_str()`、`append_many()`、
`fit_width` 或 `fit_height` 等重复入口。相同能力只保留一处明确实现。

## 18. 0.2.2 API 速查表

| 对象/模块 | 稳定公开 API |
|---|---|
| `Workbook` | `add_sheet`、`sheet`、`remove_sheet`、`move_sheet`、`copy_sheet`、`sheets`、`active`、`load`、`render`、`save`、`len()` |
| `Worksheet` | `label`、`label_color`、`cell`、`range`、`row`、`column`、`merged_ranges`、`freeze`、`filter_range`、`show_gridlines`、`page`、`max_row`、`max_column`、`values`、`append`、`append_rows`、`[]` |
| `Cell` | `row`、`column`、`index`、`address`、`value`、`formula`、`style`、`set_value`、`read`、六种 `as_*` |
| `CellValue` | `as_string`、`as_int`、`as_float`、`as_bool`、`as_date`、`as_datetime` |
| `Range` | 四个 0-based 边界、`address`、`values`、`set_values`、`merge`、`unmerge` |
| 行列尺寸 | `RowDimension.index/height/hidden`、`ColumnDimension.index/width/hidden` |
| 页面 | `PageSettings`、`PageMargins`、`HeaderFooter` 及本手册第 9 节全部属性 |
| `excelkit.address` | `MAX_ROW`、`MAX_COLUMN`、`column_to_index`、`index_to_column`、`cell_index`、`parse_range`、`cell_address` |
| `excelkit.errors` | `ExcelKitError`、`InvalidAddressError`、`InvalidWorksheetNameError`、`InvalidFileError`、`TemplateError` |

0.2.x 内保持上述公开名称和参数语义兼容；以下划线开头的属性、方法和模块属于内部
实现，不纳入稳定性承诺。新增向后兼容能力使用补丁或次版本号；破坏性公开 API
变更只在新的主版本中进行并写入更新日志。
