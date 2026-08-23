"""ExcelKit 对外公开的异常类型。"""


class ExcelKitError(Exception):
    """ExcelKit 自定义异常的共同基类。"""


class InvalidAddressError(ExcelKitError, ValueError):
    """A1 地址或 0-based 行列索引无效时抛出。"""


class InvalidWorksheetNameError(ExcelKitError, ValueError):
    """工作表名称不符合 Excel 或 XML 规则时抛出。"""


class InvalidFileError(ExcelKitError, ValueError):
    """输入文件格式不受支持或内部结构损坏时抛出。"""


class TemplateError(ExcelKitError, ValueError):
    """模板标签、循环结构或渲染数据不符合要求时抛出。"""


__all__ = [
    "ExcelKitError",
    "InvalidAddressError",
    "InvalidWorksheetNameError",
    "InvalidFileError",
    "TemplateError",
]
