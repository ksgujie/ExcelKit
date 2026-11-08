"""ExcelKit 自定义异常，所有异常均继承自 ExcelKitError。"""


class ExcelKitError(Exception):
    """ExcelKit 自定义异常的共同基类。"""


class InvalidAddressError(ExcelKitError):
    """A1 地址或 0-based 行列索引无效。"""


class InvalidWorksheetNameError(ExcelKitError):
    """工作表名称无效。"""


class InvalidFileError(ExcelKitError):
    """读取的文件格式不支持或结构损坏。"""


class FormulaCalculationError(ExcelKitError):
    """Python 端公式解析、依赖或计算失败。"""


class TemplateError(ExcelKitError):
    """Excel 模板渲染失败。"""


class ValidationError(ExcelKitError):
    """工作簿或工作表结构校验失败。"""


__all__ = [
    "ExcelKitError",
    "InvalidAddressError",
    "InvalidWorksheetNameError",
    "InvalidFileError",
    "FormulaCalculationError",
    "TemplateError",
    "ValidationError",
]
