# 更新日志

## 0.7.0 - 2026-08-29

- 新增工作表数据查找与替换：`Worksheet.find()`、`Worksheet.replace()` 支持字符串、
  数值、完整匹配、大小写匹配、公式文本和替换数量上限。
- 新增 `Worksheet.export()`，可将单张工作表导出为 CSV 或 TSV；`Workbook.save()`
  同步支持单工作表 CSV/TSV 输出，并拒绝静默丢弃多工作表数据。
- 增强 `Range.clear()`：保留原有清除值和样式的默认行为，按需清除超链接与批注。

## 0.6.0 - 2026-08-29

- 新增基础 XLSX 图表：`Worksheet.add_chart()` 支持柱状、条形、折线与饼图，以及
  系列、标题、图例和尺寸设置。
- 新增 PNG/JPEG 图片插入：`Worksheet.add_image()`，支持锚点、尺寸、偏移和替代文本。
- 新增传统批注：`Cell.note` 和 `excelkit.note.Note` 支持 XLSX 读写。
- 新增区域排序 `Worksheet.sort()` / `SortKey`、实际筛选 `AutoFilter.apply()`，以及
  `Worksheet.visibility` 的普通隐藏和非常隐藏状态。
- 公式计算器新增命名区域引用、`SUMIF`、`COUNTIF`、`AVERAGEIF`、`IFERROR`、
  `ROUNDUP`、`ROUNDDOWN`、日期函数、`VLOOKUP`、`HLOOKUP` 和 `XLOOKUP`。
- 同步更新中文 API 手册、示例说明与回归测试。

## 0.5.0 - 2026-08-29

- 增加行列插入/删除、公式引用同步、依赖关系查询（`Cell.dependencies`、`dependents`）。
- 增加超链接、工作簿文档属性、CSV/TSV 编码/分隔符/表头参数及 `worksheet.headers`。
- 增强数据表列名、调整区域、追加行、清空数据和汇总行设置。
- 增加数据有效性、条件格式、自动筛选条件、工作簿和工作表保护的 XLSX 往返。
- XLSX 核心属性写入 `docProps/core.xml`；继续使用标准库实现 XLSX 读写。
- 更新中文 API 手册、示例和自动化测试。

## 0.4.0 - 2026-08-23

- 增加 `Workbook.calculate()` 受控公式计算器，支持基础算术、比较、跨工作表
  A1 引用、递归依赖和常用数学、文本、逻辑函数；严格模式统一抛出
  `FormulaCalculationError`。
- 公式结果与普通值彻底分离：增加 `Cell.cached_value`、`formula_status` 和
  `calculation_error`；XLSX 可读取和写出 `<v>` 缓存，并声明由 Excel/WPS 自动重算。
- 增加 `Cell.copy_style(source)`，只复制完整不可变样式，不复制值、公式或缓存。
- 增加 `Range.clear()`、`clear_values()`、`clear_styles()` 和 `copy_to()`；区域复制
  支持独立控制值、公式、样式，并按行列偏移转换相对 A1 引用。
- 增加工作簿级命名区域：`add_named_range()`、`named_range()`、`named_ranges`、
  `remove_named_range()`，支持 XLSX 保存和读取。
- 增加基础 Excel 数据表：`Worksheet.add_table()`、`table()`、`tables`、
  `remove_table()`，支持名称、区域、表头、样式和行列条纹的 XLSX 往返。
- 同步更新完整中文 API 手册、README、可运行示例和自动化回归测试。

## 0.3.0 - 2026-08-23

- 收敛顶层导出：`excelkit` 仅保留 `Workbook`、`Worksheet`、`Cell`、`CellValue`、
  `Range` 和版本号；样式、页面、地址、异常分别从 `excelkit.style`、
  `excelkit.page_setup`、`excelkit.address`、`excelkit.errors` 导入。
- 统一公开命名：工作表使用 `name`、`color`，冻结和筛选使用
  `freeze_panes`、`auto_filter_range`，页面使用 `print_order`、`print_area`，
  模板分表数据参数使用 `sheet_data`。
- 删除 0.3.0 中的旧公开别名，避免同一功能多处实现；`Side` 改为
  `BorderSide`，`parse_range` 改为 `range_index`。
- 为 `Border`、`Alignment`、`PageSettings` 增加可由 IDE 自动补全的常量，减少
  手写枚举字符串错误。
- 新增 `range_address()`，与 `range_index()` 成对提供区域地址解析和生成。
- 同步更新中文 API 手册、README、示例、迁移说明和回归测试。

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
