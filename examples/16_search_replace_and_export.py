"""演示 0.7.0 的查找、替换、清除附属信息与 CSV 导出。"""

from excelkit import Workbook
from excelkit.note import Note


def main() -> None:
    """功能：创建示例表，执行数据整理后导出 CSV 文件。

    使用方法：在项目根目录运行 ``python examples/16_search_replace_and_export.py``。
    参数：无。
    返回：``None``；生成 ``search_replace.xlsx`` 与 ``search_replace.csv``。
    """
    workbook = Workbook()
    worksheet = workbook.active
    worksheet.append(["姓名", "状态", "成绩"])
    worksheet.append(["张三", "待审核", 95])
    worksheet.append(["李四", "待审核", 88])
    worksheet["A2"].note = Note("已完成核对", author="教务处")

    for cell in worksheet.find("待审核", whole=True):
        print(cell.address, cell.value)
    worksheet.replace("待审核", "已审核", whole=True)
    worksheet.range("A2:A2").clear(values=False, styles=False, notes=True)

    worksheet.export("search_replace.csv")
    workbook.save("search_replace.xlsx")


if __name__ == "__main__":
    main()
