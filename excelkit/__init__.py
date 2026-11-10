"""ExcelKit 核心公共 API 的唯一顶层入口。"""

from excelkit.address import (
    XLS_MAX_COLUMNS,
    XLS_MAX_ROWS,
    XLSX_MAX_COLUMNS,
    XLSX_MAX_ROWS,
    index_to_cell,
    cell_to_index,
    column_to_index,
    index_to_column,
    index_to_range,
    range_to_index,
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

__version__ = "0.1.2"

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
    "cell_to_index",
    "index_to_cell",
    "range_to_index",
    "index_to_range",
    "UnsupportedFeatureError",
]
