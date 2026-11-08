# ExcelKit精简版

ExcelKit 是一个以实用为主的轻量级 Excel 读写库，基于原 GitHub 项目 ksgujie/ExcelKit 精简而来。

只保留日常办公高频功能，去掉图表/图片/条件格式/保护等装饰性功能。

## 核心功能

- Workbook / Worksheet / Cell / Range 对象模型
- A1 地址与 0-based 索引（先行、后列）
- XLSX / XLS / CSV / TSV 读写（XLSX 纯标准库，无外部依赖；XLS 需 xlrd/xlwt）
- 单元格值、公式、基础样式（Font/Fill/Border/Alignment/Style）
- 区域批量操作（range / values / set_values / clear / copy_to / merge / to_records）
- 模板渲染（标量标签 + {loop items} / {/loop} 循环行块）
- 基础公式计算（SUM/AVERAGE/MIN/MAX/COUNT/COUNTA/IF/AND/OR/NOT/ABS/INT/ROUND/CONCAT/LEN/LEFT/RIGHT/MID 等）
- 合并单元格、冻结窗格、自动筛选、页面打印设置、基础数据表

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

```bash
pip install .
# 需要 XLS 支持时：
pip install .[xls]
# 或者
pip install xlrd xlwt
```

## 运行测试

```bash
python -m unittest discover -s tests -v
```
