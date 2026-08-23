"""创建一个小型多工作表 XLSX 文件的示例。"""

from pathlib import Path

from excelkit import Workbook


def main() -> None:
    """功能：创建成绩和统计两张工作表并保存示例文件。

    使用方法：在项目根目录执行 ``python -m examples.create_excel``。
    参数：无。
    返回：``None``；成功后在当前目录生成 ``example.xlsx``。
    """
    workbook = Workbook()
    scores = workbook.add_sheet("成绩")
    scores.range("A1:B3").set_values([
        ["姓名", "成绩"],
        ["张三", 95],
        ["李四", 88],
    ])

    summary = workbook.add_sheet("统计")
    summary.range("A1:B2").set_values([["人数", 2], ["平均分", 91.5]])

    output = Path("example.xlsx")
    workbook.save(output)
    print(f"已生成：{output.resolve()}")


if __name__ == "__main__":
    main()
