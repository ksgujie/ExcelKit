"""生成包含数据表、公式、条件格式和打印分页的业务报表示例。"""

from excelkit import Workbook
from excelkit.autofill import AutoFillMode
from excelkit.conditional import IconSet
from excelkit.style import NumberFormat, ReportStyle
from excelkit.table import TotalFunction


orders = [
    {"订单号": "SO-001", "客户": "张三", "数量": 2, "单价": 99.5},
    {"订单号": "SO-002", "客户": "李四", "数量": 5, "单价": 48.0},
    {"订单号": "SO-003", "客户": "王五", "数量": 1, "单价": 320.0},
]

workbook = Workbook()
worksheet = workbook.active
worksheet.name = "订单明细"

# 从第 0 行、第 0 列写入记录，并自动创建带筛选按钮的 Excel 数据表。
table = worksheet.write_table(
    0, 0, orders, name="Orders", freeze_header=True, auto_fit=True
)

# 使用相对引用公式批量计算金额；区域格式和条件格式均作用于业务数据行。
worksheet.fill_formula("E2:E4", "=C2*D2")
worksheet.range("E2:E4").format.number = NumberFormat.CURRENCY
worksheet.range("A1:E1").apply_style(ReportStyle.HEADER)
worksheet.add_data_bar("E2:E4", color="5B9BD5")
worksheet.add_icon_set("C2:C4", style=IconSet.THREE_TRAFFIC_LIGHTS)
table.set_total("数量", TotalFunction.SUM)

# 自动填充也可用于编号、日期、公式和样式的向下扩展。
worksheet.range("G1:G2").set_values([[1], [2]]).auto_fill(
    "G1:G10", mode=AutoFillMode.SERIES
)

# 宽表打印时使用横向、适应一页宽，并在第 50 行前新起一页。
worksheet.page.orientation = worksheet.page.LANDSCAPE
worksheet.page.fit(width=1)
worksheet.add_horizontal_page_break(49)

workbook.save("18_business_report.xlsx")
