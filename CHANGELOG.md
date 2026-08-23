# 更新日志

## 0.1.2 - 2026-08-23

- 将工作表标签属性统一命名为 `Worksheet.label` 和
  `Worksheet.label_color`。
- 彻底删除 `Worksheet.name`、`Worksheet.tab_name` 和
  `Worksheet.tab_color`，避免重复或不一致的公开 API。
- 重命名继续原子同步 `Workbook.sheet()` 名称索引，标签颜色继续支持 XLSX
  保存与读取。
- 同步更新完整 API 手册、README、示例和回归测试。

## 0.1.1 - 2026-08-23

- 增加 `Worksheet.name` 设置器，可通过 `worksheet.name = "新名称"` 重命名，
  并原子同步 `Workbook.sheet()` 的名称索引。
- 增加 `Worksheet.tab_color`，支持 6 位 RGB、8 位 ARGB 和 `None` 清除颜色。
- XLSX 写出和读取支持工作表标签颜色；旧版 XLS 因格式后端限制不写出标签颜色。
- 增加工作簿模板渲染：`Workbook.render()`、标量标签、循环行块、样式复制及
  基础公式行引用调整。
- 同步更新完整 API 手册、README、示例与自动化测试。

## 0.1.0 - 2026-08-22

- 提供首版 `Workbook`、`Worksheet`、`Cell` 和 `Range` 对象模型。
- 提供 A1 地址解析与校验。
- 使用稀疏结构存储单元格值。
- 提供 XLSX Open XML 写出器，以及由 `xlrd`、`xlwt` 支持的 XLS 读写。
- `Workbook.sheet()` 同时支持名称查询和 Python 风格整数索引。
- `Worksheet.cell()` 同时支持 A1 地址和数字行列坐标。
- 源码注释和函数文档统一为中文。
- 所有数字索引统一为 0-based，行列参数统一使用先行后列顺序。
- 彻底删除重复的 `Worksheet.cell_at()`。
- 增加公式保存、`append()`、`append_rows()`、统一地址 API 和 `XlsxWriter`。
- 单元格地址转换最终统一命名为 `cell_index()` 和 `cell_address()`。
- `Worksheet.values` 返回从 A1 到已触及边界的全部普通值。
- `Workbook.load()` 支持 XLS、XLSX、XLSM、XLTX、CSV 和 TSV。
- `cell.value` 自动将 `#YYYY-M-D`、`#YYYY/M/D` 及带时分秒形式转换为
  `date` 或 `datetime`。
- 增加不可变 `Style`、`Font`、`Fill`、`Side`、`Border` 和 `Alignment`，
  支持 XLSX 样式往返及 XLS 基础样式映射。
- 增加 `Workbook.render(data, strict=True)` Excel 模板渲染；支持标量标签、
  点分路径、`{loop items}` / `{/loop}` 循环行块、显式 `{items.name}` 作用域、
  0-based `{items.@index}`、样式复制和基础公式行引用调整。
