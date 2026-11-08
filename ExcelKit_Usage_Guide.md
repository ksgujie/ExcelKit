# ExcelKit精简版 使用文档

> 版本：v0.1.0  
> 项目根目录：`W:\ExcelKit`  
> 适用环境：Python 3.8+

---

## 1. 项目简介

ExcelKit 是一个以实用为主的轻量级 Excel 读写库，基于原 GitHub 项目 **ksgujie/ExcelKit** 精简而来。

只保留日常办公中高频使用的核心功能，去掉图表、图片、条件格式、数据验证、工作簿/工作表保护等装饰性功能，代码精简、易维护、无第三方强制依赖（XLSX 读写纯标准库实现）。

### 核心功能

| 功能类别 | 说明 |
|---------|------|
| 工作簿管理 | 创建、添加/删除/移动/复制工作表、按名称或索引查询 |
| 单元格读写 | `ws["A1"] = 值`、公式读写、0-based 行列访问 |
| 批量写入 | `append()` 追加一行、`append_rows()` 追加多行 |
| 区域操作 | `ws.range("A1:C10")`、值读写、清除、复制、合并、转记录 |
| 单元格样式 | Font、Fill、Border、Alignment、Style（不可变，可复用） |
| 公式计算 | `wb.calculate()` 在 Python 端计算常用公式并缓存结果 |
| 文件读写 | XLSX（纯标准库）、XLS（需 xlrd/xlwt）、CSV/TSV（标准库） |
| 模板渲染 | 标量标签 `{title}`、`{loop items}` 循环行块、数值表达式、格式过滤 |
| 辅助功能 | 合并单元格、冻结窗格、自动筛选、页面打印设置、命名区域、基础数据表 |

### 明确不做（已精简掉）

- 图表、图片、单元格批注、超链接
- 条件格式、数据验证
- 工作簿 / 工作表保护
- 行列插入/删除、文档属性
- `auto_fill` 序列填充、`find/replace`
- `render_many` / `export_pages` 等包装循环的 API

---

## 2. 安装

### 方式一：从源码安装（推荐）

```bash
# 在 W:\ExcelKit 目录下
pip install .
```

### 方式二：XLS 格式支持（可选）

XLSX/CSV/TSV 使用 Python 标准库，无需额外安装。如需读写 `.xls` 格式，需安装 `xlrd`、`xlwt`：

```bash
pip install xlrd xlwt
# 或
pip install .[xls]
```

---

## 3. 快速入门

```python
from excelkit import Workbook

# 1. 创建工作簿，获取活动工作表
wb = Workbook()
ws = wb.active          # 默认名为 "Sheet1"
ws.name = "成绩表"

# 2. 写入单元格（A1 地址）
ws["A1"] = "姓名"
ws["B1"] = "成绩"
ws["A2"] = "张三"
ws["B2"] = 95
ws["A3"] = "李四"
ws["B3"] = 88

# 3. 写入公式并计算
ws["C1"] = "总分"
ws["C2"].formula = "=B2*2"
ws["C3"].formula = "=B3*2"
wb.calculate()          # 在 Python 端计算公式
print(ws["C2"].value)   # 190.0

# 4. 保存
wb.save("成绩.xlsx")

# 5. 重新加载
wb2 = Workbook.load("成绩.xlsx")
ws2 = wb2.active
print(ws2["A2"].value)  # 张三
print(ws2["C2"].value)  # 190.0（公式计算结果）
```

---

## 4. 核心概念

### 4.1 索引规则

- **所有数字索引统一为 0-based**
- **行列参数统一为先行、后列**：`cell(0, 0)` 对应 A1，`cell(7, 2)` 对应 C8
- **A1 地址字符串**遵循 Excel 原生表示（1-based 列字母 + 1-based 行号），如 `"A1"`、`"C8"`、`"B3:D8"`

| 表达式 | 含义 |
|--------|------|
| `ws["A1"]` | 第一行第一列（A1） |
| `ws.cell(0, 0)` | 0-based (0, 0)，即 A1 |
| `ws.cell(7, 2)` | 0-based (7, 2)，即 C8 |
| `ws.range("A1:C10")` | A1 到 C10 的矩形区域 |

### 4.2 顶层入口

```python
from excelkit import Workbook, Worksheet, Cell, CellValue, Range
```

日常使用只需导入 `Workbook`，其余对象通过工作簿/工作表自动获取。

---

## 5. Workbook（工作簿）用法

### 5.1 创建与工作表管理

```python
wb = Workbook()                     # 自动创建一个名为 "Sheet1" 的工作表
ws = wb.active                      # 获取当前活动工作表

ws2 = wb.add_sheet("销售明细")      # 新建工作表并返回
ws2 = wb.sheet("销售明细")          # 按名称查询（不区分大小写）
ws2 = wb.sheet(0)                   # 按 0-based 索引查询

wb.remove_sheet("Sheet1")           # 删除工作表
wb.move_sheet("销售明细", 0)        # 移动到第 0 位
wb.copy_sheet("销售明细", "副本")   # 完整复制工作表（含内容、布局）

print(wb.sheets)                    # 所有工作表的 tuple
print(wb.active.name)               # 当前活动工作表名称
```

### 5.2 工作表名称与颜色

```python
ws.name = "月报"                    # 重命名，自动同步工作簿索引
ws.color = "4472C4"                 # 设置标签颜色（6 位 RGB 或 8 位 ARGB）
print(ws.color)                     # 读取颜色
ws.color = None                     # 清除颜色
```

### 5.3 保存与加载

```python
# 保存 XLSX
wb.save("report.xlsx")

# 保存 CSV/TSV（仅导出当前活动工作表）
wb.save("report.csv")
wb.save("report.tsv")

# 保存前结构校验（发现合并区域重叠、数据表区域重叠等问题）
wb.save("report.xlsx", validate=True)

# 加载文件
wb2 = Workbook.load("input.xlsx")   # XLSX / XLSM / XLTX
wb3 = Workbook.load("input.xls")    # XLS（需 xlrd）
wb4 = Workbook.load("input.csv")    # CSV
wb5 = Workbook.load("input.tsv")    # TSV
```

### 5.4 公式计算

```python
ws["A1"] = 10
ws["A2"] = 20
ws["A3"].formula = "=SUM(A1:A2)"
ws["A4"].formula = "=IF(A1>A2,\"大\",\"小\")"

wb.calculate()                      # 计算全部公式，结果缓存到单元格
wb.calculate(strict=True)           # 严格模式：任一公式出错即抛出异常

print(ws["A3"].value)               # 30.0
print(ws["A4"].value)               # 小
```

支持的公式函数见第 10 节。

### 5.5 结构校验

```python
problems = wb.validate()            # 返回中文问题列表，空表示无问题
if problems:
    for p in problems:
        print(p)
```

校验内容包括：合并区域重叠、数据表区域重叠、重复的工作表名称、重复的数据表名称。

### 5.6 命名区域（Workbook 级）

```python
rng = ws.range("C2:C100")
nr = wb.add_named_range("SalesAmount", rng)
print(wb.named_range("salesamount").ref)   # 不区分大小写
print([n.name for n in wb.named_ranges])   # 枚举全部
wb.remove_named_range("SalesAmount")
```

---

## 6. Worksheet（工作表）用法

### 6.1 单元格读写

```python
# A1 地址访问
ws["A1"] = "姓名"
print(ws["A1"].value)               # 姓名

# 0-based 行列访问（先行后列）
ws.cell(0, 0).value = "标题"        # (0, 0) = A1
ws.cell(1, 0).value = "张三"        # (1, 0) = A2

# 显式 Cell 对象
cell = ws.cell("A1")
cell.value = "新值"
```

### 6.2 值与公式

```python
ws["B1"] = 95                       # 写入普通值
ws["B2"].formula = "=SUM(B1:B1)"    # 写入公式（可带或不带 '='）
print(ws["B2"].formula)             # =SUM(B1:B1)
print(ws["B2"].value)               # 计算前为 None（无缓存结果）

wb.calculate()
print(ws["B2"].value)               # 计算后的结果

ws["B2"].formula = None             # 清除公式
ws["B2"] = 100                      # 写入值会自动清除同位置公式
```

### 6.3 批量写入

```python
# 追加一行（空表从 A1 开始，否则在最后一行之后）
ws.append(["姓名", "成绩"])
ws.append(["张三", 95])

# 原子校验并追加多行（列数不一致会抛异常，失败不留下部分数据）
ws.append_rows([
    ["李四", 88],
    ["王五", 92],
])

print(ws.values)                    # [[...], [...], ...] 从 A1 到已触及边界
print(ws.max_row)                   # 已触及的最大行索引（0-based）
print(ws.max_column)                # 已触及的最大列索引（0-based）
print(ws.used_range.address)        # 已使用区域地址，空表为 None
```

### 6.4 区域操作（Range）

```python
rng = ws.range("A1:C3")

# 读取区域值（二维列表）
vals = rng.values                   # [[...], [...], ...]

# 写入等形状二维数据
rng.set_values([
    [1, 2, 3],
    [4, 5, 6],
])

# 清除区域（可选是否清除值/样式）
rng.clear(values=True, styles=True)

# 复制到等尺寸目标区域
rng.copy_to(ws.range("E1:G3"), values=True, formulas=True, styles=True)

# 合并 / 取消合并
rng.merge()
rng.unmerge()
print(ws.merged_ranges)             # 只读 tuple

# 区域转字典记录
records = ws.range("A1:D4").to_records()              # 首行作为字段名
records = ws.range("A5:D20").to_records(header_row=1) # 指定绝对行作为字段行
records = ws.range("A1:D4").to_records(headers=False) # 自动生成 Column1, Column2...
```

### 6.5 合并单元格、冻结窗格、自动筛选

```python
ws.range("A1:D1").merge()           # 合并标题行

ws.freeze_panes = "A3"              # 冻结前两行，A3 为首个可滚动单元格
ws.freeze_panes = None              # 清除冻结

ws.auto_filter.range = "A1:D100"    # 设置筛选区域
ws.auto_filter.clear()              # 清除筛选
```

### 6.6 行高、列宽、隐藏

```python
ws.row(0).height = 28               # 第 1 行行高
ws.column(0).width = 24             # 第 1 列列宽
ws.row(1).hidden = True             # 隐藏第 2 行
ws.column(2).hidden = True          # 隐藏第 3 列
```

### 6.7 页面打印设置

```python
page = ws.page
page.orientation = "landscape"      # 横向
page.paper_size = "A4"
page.fit(width=1)                   # 适应一页宽
page.print_area = "A1:F100"         # 打印区域
page.repeat_rows = (0, 1)           # 重复打印标题行（0-based，包含结束行）
page.margins.left = 1.5             # 边距（厘米）
page.show_gridlines = False         # 屏幕不显示网格线
page.print_gridlines = True         # 打印时显示网格线
```

### 6.8 数据表（基础 Table）

```python
table = ws.add_table(
    "A1:D100",
    name="Orders",
    style="TableStyleMedium9",
    has_header=True,
    show_row_stripes=True,
)
table.set_total("金额", "sum")        # 设置汇总行函数

print(ws.tables)                    # 枚举本表全部表
t = ws.table("Orders")              # 按名称查询
ws.remove_table("Orders")           # 删除定义（不删除单元格数据）
```

### 6.9 查找与替换

```python
# 查找（返回 Cell 列表）
for cell in ws.find("待审核", whole=True):
    print(cell.address, cell.value)

# 替换（返回替换数量）
count = ws.replace("待审核", "已审核", whole=True)
print(f"替换了 {count} 个单元格")
```

参数：`whole` 是否完全匹配，`case_sensitive` 是否区分大小写，`limit` 替换数量上限。

### 6.10 导出单表为 CSV/TSV

```python
ws.export("sheet.csv")              # 默认逗号分隔
ws.export("sheet.tsv", delimiter="\t")
```

---

## 7. Cell（单元格）用法

### 7.1 获取 Cell

```python
cell = ws["A1"]                     # A1 地址
cell = ws.cell("A1")                # 显式 A1
cell = ws.cell(0, 0)                # 0-based 行列
```

### 7.2 值与公式属性

| 属性 | 说明 |
|------|------|
| `cell.value` | 读写普通值；公式单元格返回当前缓存结果（无则为 None） |
| `cell.formula` | 读写公式，读取时始终带 `=`；赋值 `None` 清除公式 |
| `cell.address` | 规范化 A1 地址，如 `"A1"` |
| `cell.index` | 0-based `(row, column)` 元组，先行后列 |
| `cell.row` / `cell.column` | 0-based 行/列索引 |

### 7.3 样式

```python
from excelkit.style import Font, Fill, Border, BorderSide, Alignment, Style

ws["A1"].style = Style(
    font=Font(name="微软雅黑", size=12, bold=True, color="FFFFFF"),
    fill=Fill("4472C4"),
    border=Border(bottom=BorderSide(Border.THIN, "000000")),
    alignment=Alignment(horizontal="center", vertical="center"),
    number_format="0.00",
)

# 只复制样式，不复制值和公式
ws["B1"].copy_style(ws["A1"])
```

### 7.4 类型转换

**写回型**（`as_*` 会修改单元格值）：
```python
ws["A1"] = "123"
n = ws["A1"].set_value("123").as_int()   # 写回并返回 123
assert ws["A1"].value == 123
```

**只读型**（`read().as_*` 不修改原单元格）：
```python
ws["A1"] = "123"
n = ws["A1"].read().as_int()             # 只转换快照，原值保持 "123"
assert ws["A1"].value == "123"
```

| 方法 | 说明 |
|------|------|
| `as_string()` | 转字符串 |
| `as_int()` | 转整数（不截断小数，NaN/Inf 报错） |
| `as_float()` | 转浮点（拒绝 NaN/Inf） |
| `as_bool()` | 接受 true/false/yes/no/是/否/1/0 |
| `as_date()` | 支持 `#2026-8-1`、`2026-8-1`、`2026/8/1` |
| `as_datetime()` | 支持 `#2026/8/1 12:33` 等 |

> ⚠️ 重要区别：`cell.as_*()` 会把转换结果**写回**单元格；`cell.read().as_*()` 只转换**读取快照**，绝不会影响后续保存的文件。

### 7.5 日期字面量

```python
ws["A1"] = "#2026-8-1"              # date
ws["A2"] = "#2026/8/1 12:33"        # datetime
ws["A3"] = "#2026-8-1 12:33:45"     # datetime
```

---

## 8. Range（区域）用法详解

```python
rng = ws.range("A1:C10")

# 属性
rng.address        # "A1:C10"
rng.rows           # 10
rng.columns        # 3

# 值
rng.values         # 二维列表
rng.set_values([[1,2],[3,4]])   # 写入等形状数据

# 清除
rng.clear(values=True, styles=True)

# 复制
rng.copy_to(target_range, values=True, formulas=True, styles=True)

# 合并
rng.merge() / rng.unmerge()

# 转记录
records = rng.to_records()                          # 首行做字段名
records = rng.to_records(headers=False)             # 自动 Column1...
records = rng.to_records(header_row=5)              # 指定绝对行做字段名
```

---

## 9. 样式系统

所有样式对象为 **不可变**（frozen），可安全复用于多个单元格。

```python
from excelkit.style import Font, Fill, BorderSide, Border, Alignment, Style, NumberFormat, ReportStyle

font = Font(name="Calibri", size=11, bold=True, color="FF000000")
fill = Fill(fg_color="4472C4", pattern_type="solid")
border = Border(bottom=BorderSide(Border.THIN, "000000"))
alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

style = Style(font=font, fill=fill, border=border, alignment=alignment, number_format="0.00")

ws["A1"].style = style

# 内置预设
NumberFormat.CURRENCY      # '"$"#,##0.00'
NumberFormat.PERCENT_2DEC  # '0.00%'
ReportStyle.HEADER         # 白字蓝底居中标题样式
```

---

## 10. 公式计算器

`wb.calculate()` 在 Python 端计算当前支持的公式并缓存结果。这不是完整 Excel 引擎，复杂公式仍应交给 Excel/WPS 计算。

### 支持的运算符

`+ - * / // % ^`、比较 `= <> < <= > >=`、括号、单引号字符串、`NOT`。

### 支持的函数

| 函数 | 说明 |
|------|------|
| `SUM` | 求和（支持区域引用） |
| `AVERAGE` | 平均值 |
| `MIN` / `MAX` | 最小/最大 |
| `COUNT` | 数字计数 |
| `COUNTA` | 非空计数 |
| `IF` | 条件判断 |
| `AND` / `OR` / `NOT` | 逻辑 |
| `ABS` | 绝对值 |
| `INT` | 取整（向下取整） |
| `ROUND` | 四舍五入 |
| `CONCAT` | 字符串拼接 |
| `LEN` | 字符串长度 |
| `LEFT` / `RIGHT` / `MID` | 字符串截取 |

### 引用方式

- 单格：`A1`、`Sheet1!A1`
- 区域：`A1:B2`、`Sheet1!A1:B2`
- 跨表：`=Sheet1!A1 + Sheet2!B1`

```python
wb.calculate(strict=False)   # 非严格：出错单元格设为 None，继续
wb.calculate(strict=True)    # 严格：首个错误抛出 FormulaCalculationError
```

---

## 11. 模板渲染

先在 Excel 中准备模板，写入标签，再用代码渲染并保存。

### 11.1 标量标签

```python
ws["A1"] = "{title}"
ws["A2"] = "客户：{customer.name}"
ws["A3"] = "日期：{report_date}"

wb.render({"title": "月报", "customer": {"name": "示例公司"}, "report_date": "#2026-8-1"})
```

### 11.2 循环行块（{loop} / {/loop}）

```python
ws["A4"] = "{loop items}"
ws["A5"] = "{items.@index + 1}"   # 序号（从 1 开始）
ws["B5"] = "{items.name}"
ws["C5"] = "{items.quantity}"
ws["D5"] = "{items.quantity * items.price}"      # 数值表达式
ws["E5"] = '￥{items.quantity * items.price | format:",.2f"}'  # 格式化
ws["A6"] = "{/loop}"

wb.render({
    "title": "销售明细",
    "items": [
        {"name": "产品 A", "quantity": 2, "price": 19.5},
        {"name": "产品 B", "quantity": 3, "price": 8},
    ],
})
```

### 11.3 多工作表数据

```python
wb.render(
    {"company": "示例公司"},                    # 公共数据
    sheet_data={
        "封面": {"title": "月报"},
        "明细": {"items": [...]},
        1: {"total": 1000},                      # 也可用 0-based 索引
    },
)
```

### 11.4 strict 模式

```python
wb.render(data, strict=False)   # 默认：缺失标签清空，缺失循环视为空数组
wb.render(data, strict=True)    # 缺失字段抛出 TemplateError
```

规则：
- 整个单元格只有一个标签时保留原始类型（数字/日期等）；混合文字时转为字符串
- `| format:",.2f"` 使用 Python 格式规则
- `{items.@index}` 返回 0-based 索引

---

## 12. 文件格式支持

| 格式 | 读取 | 写出 | 依赖 |
|------|------|------|------|
| `.xlsx` | ✅ | ✅ | 无（标准库） |
| `.xls` | ✅ | ✅ | `xlrd` / `xlwt` |
| `.csv` | ✅ | ✅ | 无（标准库） |
| `.tsv` | ✅ | ✅ | 无（标准库） |

XLS 公式只能取得文件中已有的缓存计算结果，无法恢复公式表达式。XLS 标签颜色因格式限制不写出。

---

## 13. 常见问题与注意事项

1. **索引从 0 开始**：`ws.cell(0, 0)` 是 A1，与 Excel 界面显示的行列号不同。
2. **公式不自动重算**：读取公式单元格时返回缓存结果；调用 `wb.calculate()` 后才会在 Python 端重新计算。
3. **样式不可变**：`Style`/`Font` 等为 frozendataclass，可安全复用；修改需新建对象。
4. **XLS 依赖**：读写 `.xls` 需安装 `xlrd`、`xlwt`，否则只影响 XLS 功能，XLSX/CSV/TSV 不受影响。
5. **保存前校验**：重要报表建议 `wb.save(..., validate=True)`，可提前发现合并区域/数据表区域重叠等问题。
6. **路径建议**：使用绝对路径或 `A_ScriptDir` 风格的相对路径，避免工作目录变化导致文件找不到。

---

## 14. API 快速查询

| 对象 | 常用方法/属性 |
|------|--------------|
| `Workbook()` | `add_sheet()`, `sheet()`, `remove_sheet()`, `move_sheet()`, `copy_sheet()`, `active`, `sheets`, `save()`, `load()`, `calculate()`, `validate()`, `render()`, `add_named_range()` |
| `Worksheet` | `["A1"]`, `cell()`, `range()`, `append()`, `append_rows()`, `values`, `used_range`, `merge()`, `freeze_panes`, `auto_filter`, `page`, `add_table()`, `find()`, `replace()`, `export()` |
| `Cell` | `value`, `formula`, `style`, `read()`, `as_int()`, `as_float()`, `as_bool()`, `as_date()`, `as_datetime()`, `set_value()`, `copy_style()` |
| `Range` | `values`, `set_values()`, `clear()`, `copy_to()`, `merge()`, `unmerge()`, `to_records()`, `address` |

---

## 15. 运行测试

```bash
# 在 W:\ExcelKit 目录下
python -m unittest tests.test_address tests.test_workbook tests.test_cell tests.test_append tests.test_range tests.test_formula tests.test_file tests.test_style tests.test_template -v
```

预期：`Ran 42 tests, OK`，退出码为 0。

---

## 16. 版本隔离说明

本项目遵循版本隔离保存规范：每次修改入口或核心库时，递增版本号另存为新文件，旧版本保留。当前版本为 `v0.1.0`，详见 `W:\ExcelKit\ExcelKit项目整体情况概览.md`。
