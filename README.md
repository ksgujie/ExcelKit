# ExcelKit精简版

ExcelKit 是一个以实用为主的轻量级 Excel 读写库，基于原 GitHub 项目 ksgujie/ExcelKit 精简而来。

只保留日常办公高频功能，去掉图表/条件格式/保护等复杂功能；支持在 XLSX 中插入图片（浮动/嵌入单元格）。

## 核心功能

- Workbook / Worksheet / Cell / Range 对象模型
- A1 地址与 0-based 索引（先行、后列）
- XLSX / XLS / CSV / TSV 读写（XLSX/CSV/TSV 纯标准库；XLS 使用 xlrd/xlwt，默认安装）
- 图片插入（XLSX）：PNG/JPEG，`mode="float"` 浮动 / `mode="embed"` 随单元格长宽缩放，支持读写往返
- 地址与索引互转：`cell_to_index`/`index_to_cell`/`range_to_index`/`index_to_range`/`column_to_index`/`index_to_column`（支持 `$A$1`）
- 单元格值、公式、基础样式（Font/Fill/Border/Alignment/Style）
- 区域批量操作（range / values / set_values / clear / copy_to / merge / to_records）
- 模板渲染（标量标签 + {loop items} / {/loop} 循环行块）
- 基础公式计算（SUM/AVERAGE/MIN/MAX/COUNT/COUNTA/IF/AND/OR/NOT/ABS/INT/ROUND/CONCAT/LEN/LEFT/RIGHT/MID 等）
- 合并单元格、冻结窗格、自动筛选（仅启用区域，无 apply() 别名；XLSX/XLS 持久化）、页面打印设置

- XLSX 字符串使用标准 inline string（`<is>` 元素）写出，兼容 Excel/WPS；读取端兼容旧版 `<inlineStr>` 文件
- XLS 写出/读取支持自动筛选区域（Excel/WPS 中显示筛选下拉）；XLS 仍不写图片、合并单元格、冻结窗格、列宽

## 快速开始

```python
from excelkit import Workbook

wb = Workbook()
ws = wb.active
ws["A1"] = "姓名"
ws["B1"] = "成绩"
ws.append(["张三", 95])
ws.append(["李四", 88])
ws["C1"] = "总分"
ws["C2"].formula = "=SUM(B2:B2)"
wb.calculate()
wb.save("成绩.xlsx")

# 重新读取
wb2 = Workbook.load("成绩.xlsx")
print(wb2.active["C2"].value)   # 183.0
```

## 安装

### 从 Wheel 安装（推荐）

```bash
pip install --force-reinstall dist\excelkit-0.1.2-py3-none-any.whl
# 需要 XLS 支持时额外安装：
pip install xlrd xlwt
```

### 从源码安装

```bash
pip install .
# 需要 XLS 支持时：
pip install .[xls]
# 或者
pip install xlrd xlwt
```

安装后确认实际导入的是哪个版本/路径：

```bash
python -c "import excelkit; print(excelkit.__version__, excelkit.__file__)"
```

## 运行测试

```bash
python -m unittest tests.test_address tests.test_workbook tests.test_cell tests.test_append tests.test_range tests.test_formula tests.test_file tests.test_style tests.test_template tests.test_image -v
```
