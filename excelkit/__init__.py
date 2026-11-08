"""ExcelKit 核心公共 API 的唯一顶层入口。"""

from excelkit.address import (
    XLS_MAX_COLUMNS,
    XLS_MAX_ROWS,
    XLSX_MAX_COLUMNS,
    XLSX_MAX_ROWS,
    cell_address,
    cell_index,
    column_to_index,
    index_to_column,
    range_address,
    range_index,
)
from excelkit.core import (
    Cell,
    CellValue,
    Image,
    Range,
    Workbook,
    Worksheet,
)
from excelkit.errors import UnsupportedFeatureError

__version__ = "0.1.0"

__all__ = [
    "__version__",
    "Workbook",
    "Worksheet",
    "Cell",
    "CellValue",
    "Image",
    "Range",
    "XLS_MAX_ROWS",
    "XLS_MAX_COLUMNS",
    "XLSX_MAX_ROWS",
    "XLSX_MAX_COLUMNS",
    "column_to_index",
    "index_to_column",
    "cell_index",
    "cell_address",
    "range_index",
    "range_address",
    "UnsupportedFeatureError",
]
