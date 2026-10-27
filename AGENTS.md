# AGENTS.md

ExcelKit：纯 Python 的 Excel 读写库（XLSX/XLS/CSV/TSV），对外只有 `excelkit` 一个包。
本文件只记录"看代码不容易猜到、猜错会出问题"的事实。

## 环境

- 仓库根 `W:\ExcelKit`；`.git` 存在，但 `python`、`py`、`git` 均不在 PATH，运行命令前先确认解释器位置，不要假设可用；git 安装在 `W:\git`。
- 要求 Python >= 3.9（代码大量使用 `X | Y` 注解，靠 `from __future__ import annotations` 兜底，新文件必须带上）。
- 运行期依赖只有 `xlrd`/`xlwt`，**且只服务 `.xls`**。XLSX 的读写是标准库 `zipfile` + `ElementTree` 手写实现的，**不要引入** openpyxl / xlsxwriter / pandas。

## 命令（均在仓库根执行）

- 全量测试：`python -m unittest discover -s tests -v`
- 单个测试文件：`python -m unittest discover -s tests -p "test_release_100.py" -v`
- 不要用 `python tests/test_core.py`：那样 `sys.path[0]` 是 `tests/`，`import excelkit` 会失败（除非先 `pip install -e .`）。
- `pytest` 也能跑（`pyproject.toml` 有 `[tool.pytest.ini_options]`，`addopts = "-q"`），但没额外收益。
- 打包：`python -m build --wheel --outdir dist`
- 没有 lint / format / typecheck / CI 配置，别假装有，也别新增工具链；以测试通过为准。

## 版本与文档同步（用户硬性要求，最容易漏）

改动程序行为后必须：

1. 升版本号：新增功能升 minor（0.10.0 → 0.11.0），纯修复升 patch。
2. 同步 **5 处**版本号：`pyproject.toml`、`excelkit/__init__.py` 的 `__version__`、`README.md` 标题、`docs/API完整使用手册.md` 标题、`CHANGELOG.md` 顶部新条目（格式 `## x.y.z - YYYY-MM-DD`）。
3. 中文写更新日志；README 与手册中对应章节一起改，新 API 还要补手册第 18 节速查表。
4. 更新本文件末尾的「当前状态」章节（版本、新增能力、新发现的坑），不另建总览文件。

## 数据模型（理解错就会写错整个功能）

- `Worksheet` 用稀疏字典按 `(row, column)` 存 7 类单元格状态：`_values`（ValueStore）、`_formulas`、`_formula_values`、`_formula_errors`、`_hyperlinks`、`_notes`、`_styles`。没有二维数组。
- 写值只能走 `Worksheet._set_value()` 或 `Cell.value`，写公式走 `_set_formula()`：它们负责日期字面量规范化、清掉同格公式/缓存、更新 `_max_row`/`_max_column`。绕过它们直接改字典会让样式、缓存、使用范围全部失配。
- 值或公式变化后调用 `workbook._invalidate_formula_caches()`；公式结果只由 `calculate()` 或读取 XLSX 时填充，`Cell.value` 读公式格**只返回缓存，不会自动计算**。
- `_max_row`/`_max_column` 是"触及过"的最大索引：写值只增不减，清空单元格也不回退（只有 `insert_rows`/`delete_rows` 这类整体编辑才会调整）。真正的内容范围请看 `used_range`，不要拿 `max_row` 当"最后一行有数据"。
- 所有索引 0-based、先行后列；A1 字符串只出现在 API 边界，转换集中在 `address.py`。
- 索引语法只有 `ws["A1"]` 单格；区域一律 `ws.range("A1:B2")`，没有切片语法。

## 改结构时必须同步的位置（漏了会静默丢数据）

新增任何"按坐标存储的状态"，都要在这些地方成对处理：

- `Worksheet._edit_axis()`：插入/删除行列的坐标重映射，同时还要搬 merged ranges、tables、named ranges、validations、conditionals、images、freeze、filter、print area。
- `Range._rewrite_rows()`（去重 / 删空行）
- `Worksheet.sort()`
- `Workbook.copy_sheet()`
- `Worksheet.used_range`

已知缺口：图表 anchor 不在 `_edit_axis` 中重映射，动这块时顺手补齐。

## 代码风格

- 注释与 docstring 全中文，固定五段式：`功能：`、`使用方法：`、`参数：`、`返回：`、`异常：`（无异常可省）。新方法照抄周围格式，不要写英文 docstring。
- 内部辅助方法用 `_` 前缀，并在 docstring 注明"由 XX 内部调用"；只有公开 API 才写进 README/手册。
- 异常类型沿用现有惯例：类型不对 `TypeError`，取值不合法 `ValueError`，地址问题 `InvalidAddressError`，文件格式 `InvalidFileError`。
- 魔法字符串走常量类：`Border.THIN`、`NumberFormat.CURRENCY`、`ImagePlacement.CELL`、`ChartType.COLUMN` 等。
- 核心对象都定义了 `__slots__`，加字段必须同步改 `__slots__`。

## 测试

- 每个版本的新功能回归放 `tests/test_release_<去掉点的版本号>.py`（0.10.0 → `test_release_100.py`）；基础行为在 `test_core`、`test_address`、`test_xlsx_writer`、`test_layout` 等。
- 用标准库 `unittest`（`TestCase`），不用 pytest fixture；文件末尾保留 `if __name__ == "__main__": unittest.main()`。
- 典型写法是临时目录写出 xlsx 再读回校验往返一致性；XLS 用例依赖 `xlwt`/`xlrd`。

## AI 协作约定（用户明确要求）

- 全程使用中文交流。
- 先做规划，只有用户明确说"可以写代码"后才开始写代码；非代码改动的对话不输出"已完成"。
- 修改程序行为后，必须同步修改版本号（`pyproject.toml`、`excelkit/__init__.py` 的 `__version__`、`README.md` 标题、`docs/API完整使用手册.md` 标题、`CHANGELOG.md` 顶部），并用中文写更新日志。
- 代码中需要有详细的中文注释；新方法遵循五段式 docstring（`功能：`、`使用方法：`、`参数：`、`返回：`、`异常：`）。
- 修改代码完成后，回复最后一行输出"已完成"。非代码改动不输出"已完成"。
- git 安装路径为 `W:\git`。
- 修改代码后，在当前目录生成一个文件，内容是该程序的整体情况，方便下次换另一个 AI 模型修改代码时参考。

## 示例

`examples/` 是自包含脚本，会把 `.xlsx` 写到当前工作目录（仓库里已有几个历史产物被跟踪）。跑完示例产生的文件不要顺手提交。

## 当前状态（每次改完代码更新这里）

**版本**：0.11.0（2026-10-27）｜约 10500 行 Python｜`excelkit/` 下 core/reader/writer 三层加顶层横切模块

**分层**

- `excelkit/__init__.py`：只导出 `Workbook`、`Worksheet`、`Cell`、`CellValue`、`Range`、`Image`、`ImageFit`、`ImagePlacement`、`__version__`。
- `excelkit/core/`：对象模型与算法——`workbook`、`worksheet`、`range`、`cell`、`table`、`named_range`、`dimension`（行列尺寸/分组）、`page`（打印设置）、`calculation`（公式求值器）、`conversion`（类型转换）。
- `excelkit/reader/`：按扩展名分派——`delimited`（CSV/TSV）、`xls`（xlrd）、`xlsx`（zipfile + ElementTree）、`styles`。
- `excelkit/writer/`：`xlsx`（手写 OOXML，约 1400 行，改动最集中的地方）、`xls`（xlwt）、`delimited`、`styles`。
- 横切模块：`style`、`validation`、`conditional`、`chart`、`image`、`hyperlink`、`note`、`filter`、`sort`、`protection`、`properties`、`template`、`address`、`storage`、`errors`。
- 兼容转发模块：`excelkit/table.py`、`page_setup.py`、`autofill.py` 只有 re-export（3～9 行），是为旧导入路径保留的，不要往里加逻辑。

**能力范围**：XLSX/XLSM/XLTX/XLS 与 CSV/TSV 读写、样式、合并、行列尺寸与分组、冻结窗格、自动筛选、打印设置、公式保存与常用函数计算、区域批量操作、命名区域、数据表（Table）、模板渲染、超链接、批注、数据验证、条件格式、图表、PNG/JPEG 图片、文档属性、保护。

**明确不做**（README「能力边界」）：Excel 全函数兼容的计算引擎；从外部 XLSX 恢复图表；结构化 Table 公式；流式读写；XLS 的自动筛选/打印区域/重复标题写出、CELL 锚定图片。

**已知待改进**（审查所得，未修改，动手前先与用户确认优先级）

1. 顶层导出面太窄：`Style`、`Validation`、`Table`、`Chart`、`SortKey`、`NumberFormat` 等都要从子模块导入，与 `from excelkit import Workbook` 的便利性不匹配。
2. `Chart`/`Image` 的公开属性（`width`、`height`、`title`、`offset_x/y`、`alt_text`、`filename`）无类型/范围校验，而 `Style`、`Validation` 是严格校验的，风格不一致，错误会推迟到写出时才爆。
3. `remove_*` 语义不统一：`remove_validation`/`remove_conditional_format` 收对象，`remove_table`/`remove_named_range`/`remove_sheet` 收名称。
4. `Worksheet.range()` 只接受 A1 字符串，而 `cell()` 同时支持 `"B3"` 与 `(2, 1)`，两者不对称。
5. `find()`、`auto_fit_columns()`、`auto_fit_rows()` 按 `max_row × max_column` 全表扫描，稀疏大表上代价高。
6. `Validation.kind` 与 `ConditionalFormat.rule` 用 Excel 风格裸字符串（`"textLength"`、`"cellIs"`），没有枚举；`operator` 甚至不校验取值。
7. 图表 `anchor` 不随插入/删除行列重映射（见上文「已知缺口」）。
8. 没有上下文管理器（`with Workbook.load(...)`）、没有整表清空方法（目前只能 `Range.clear()`）。

**AI 模型切换参考**：修改代码后，会在当前目录生成一份程序整体情况文件（含版本、分层、能力范围、已知坑），供下次换另一个 AI 模型修改代码时参考；请勿随意删除。
