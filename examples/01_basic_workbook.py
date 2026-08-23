"""工作簿、工作表和普通值基础示例。"""

from pathlib import Path

from excelkit import Workbook


def main() -> None:
    """功能：创建基础成绩表并演示工作表名称与 0-based 索引查询。

    使用方法：在项目根目录执行 ``python -m examples.01_basic_workbook``。
    参数：无。
    返回：``None``；在当前目录生成 ``01_basic_workbook.xlsx``。
    """
    workbook = Workbook()
    scores = workbook.add_sheet("成绩")
    scores["A1"] = "姓名"
    scores["B1"] = "成绩"
    scores["A2"] = "张三"
    scores["B2"] = 95

    assert workbook.sheet("成绩") is scores
    assert workbook.sheet(0) is scores

    output = Path("01_basic_workbook.xlsx")
    workbook.save(output)
    print(f"已生成：{output.resolve()}")


if __name__ == "__main__":
    main()
