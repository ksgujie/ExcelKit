"""多工作表创建、标签查询和非负 0-based 索引查询示例。"""

from pathlib import Path

from excelkit import Workbook


def main() -> None:
    """功能：创建多张工作表并演示标签和非负 0-based 索引查询。

    使用方法：在项目根目录执行 ``python -m examples.05_multi_sheet``。
    参数：无。
    返回：``None``；在当前目录生成 ``05_multi_sheet.xlsx``。
    """
    workbook = Workbook()
    students = workbook.add_sheet("学生")
    scores = workbook.add_sheet("成绩")
    summary = workbook.add_sheet("统计")

    students.append(["姓名", "班级"])
    scores.append(["姓名", "数学"])
    summary.append(["项目", "值"])

    assert workbook.sheet(0) is students
    assert workbook.sheet(1) is scores
    assert workbook.sheet("成绩") is scores
    assert workbook.sheet("统计") is summary

    output = Path("05_multi_sheet.xlsx")
    workbook.save(output)
    print(f"已生成：{output.resolve()}")


if __name__ == "__main__":
    main()
