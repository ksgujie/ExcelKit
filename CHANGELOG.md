# 更新日志

## 0.2.3 - 2026-08-23

- 工作表标签查询改为大小写不敏感，并让所有公开工作表整数索引统一拒绝负数。
- `Workbook.save()` 只接受 `.xlsx` 和 `.xls`，错误扩展名抛出
  `InvalidFileError`，不会创建伪装格式文件或改变空工作簿。
- `append()`、`append_rows()` 改为写入前完整验证，失败不再留下部分数据。
- `copy_sheet()` 深复制嵌套可变普通值，并保持页面适应页数模式可独立复制。
- `PageSettings.scale` 禁止直接设置 `None`；适应页数统一通过 `fit()` 切换。
- `CellValue` 增加只读 `value`；`Cell.formula = None` 可直接清除公式。
- 更新完整中文 API 手册、README、示例和自动化回归测试。

## 0.2.2 - 2026-08-23

- `Cell` 增加只读 `index` 属性，以先行后列的 `(row, column)` 元组返回 0-based
  组合索引；现有 `row`、`column` 和 `address` 功能保持不变。
- 更新完整中文 API 手册、README、单元格示例和自动化回归测试。

## 0.2.1 - 2026-08-23

- 将 `Workbook.render()` 默认参数改为 `strict=False`；缺失的整格标签清空，混合
  文本删除缺失标签，结果文件不再残留未解析的 `{标签}`。
- 非严格模式下缺失循环集合按空数组处理并删除循环块；公式缺少模板数据时删除
  整条公式，避免生成无效公式。
- `Workbook.render()` 增加 `by_sheet`：可按工作表名称或0-based索引为多张表
  分别提供独立根数据，并与公共 `data` 浅层合并，独立字段覆盖同名公共字段。
- 分工作表渲染只修改列出的目标表，并保持跨目标工作表的原子提交语义。
- 更新完整中文 API 手册、README、多工作表示例和自动化回归测试。

## 0.2.0 - 2026-08-23

- 增加 `Workbook.remove_sheet()`、`move_sheet()` 和 `copy_sheet()`，工作表复制包含
  内容、样式、布局与打印设置。
- 增加 `Range.merge()`、`unmerge()`、`Worksheet.merged_ranges` 及合并区域安全
  写入规则。
- 增加 0-based `Worksheet.row()`、`column()`，支持行高、列宽和隐藏状态。
- 增加冻结窗格、屏幕网格线和 XLSX 自动筛选区域。
- 增加统一 `Worksheet.page` 打印 API：方向、纸张、百分比缩放、单次 `fit()`、
  打印区域、重复标题、边距、居中、页序、黑白、草稿、网格线、行列标题、页眉页脚。
- XLSX 支持新布局和打印属性保存及往返读取；XLS 写出支持后端可表达的常用子集，
  并支持合并、尺寸和冻结窗格往返。
- 新增布局打印完整示例、格式兼容测试和页眉页脚全部控制符参考表。

## 0.1.3 - 2026-08-23

- 模板标签增加安全数值表达式，支持 `+`、`-`、`*`、`/`、`//`、`%`、括号及
  `{items.@index + 1}` 等用法，不使用 `eval()`。
- 模板增加 `| format:"..."` 显示格式过滤器。
- 增加 `Cell.set_value()` 与写回型 `Cell.as_string()`、`as_int()`、
  `as_float()`、`as_bool()`、`as_date()`、`as_datetime()`。
- 增加 `Cell.read()` 和只读快照 `CellValue`；同名转换只返回临时值，不修改文件。
- 日期字面量与显式类型转换共用唯一转换实现，并补充完整中文手册和示例。

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
