"""ExcelKit 发布资料一致性测试。"""

from __future__ import annotations

import ast
import io
import re
import tempfile
import tokenize
import unittest
from pathlib import Path

import excelkit
from excelkit import Workbook


_ROOT = Path(__file__).resolve().parents[1]


class ReleaseDocumentationTests(unittest.TestCase):
    """验证版本元数据、文档和示例清单不会在发布时失步。"""

    def test_release_version_is_consistent_across_public_files(self) -> None:
        """功能：验证包版本、构建配置、README 和完整手册中的版本一致。"""
        version = excelkit.__version__
        pyproject = (_ROOT / "pyproject.toml").read_text(encoding="utf-8")
        readme = (_ROOT / "README.md").read_text(encoding="utf-8")
        manual = (_ROOT / "docs" / "API完整使用手册.md").read_text(encoding="utf-8")
        self.assertIn(f'version = "{version}"', pyproject)
        self.assertIn(f"# ExcelKit {version}", readme)
        self.assertIn(f"版本：{version}", manual)

    def test_manual_lists_every_runnable_example(self) -> None:
        """功能：验证完整手册列出 examples 目录中的每个可运行 Python 示例。"""
        manual = (_ROOT / "docs" / "API完整使用手册.md").read_text(encoding="utf-8")
        names = sorted(
            path.name for path in (_ROOT / "examples").glob("*.py")
            if path.name != "__init__.py"
        )
        for name in names:
            self.assertRegex(manual, re.escape(f"`{name}`"))

    def test_manual_has_navigable_contents(self) -> None:
        """功能：验证完整手册开头提供覆盖主要章节的可点击目录。"""
        manual = (_ROOT / "docs" / "API完整使用手册.md").read_text(encoding="utf-8")
        self.assertIn("## 目录", manual)
        for entry in (
            "[4. Workbook 工作簿](#4-workbook-工作簿)",
            "[5. Worksheet 工作表](#5-worksheet-工作表)",
            "[6. Cell 单元格](#6-cell-单元格)",
            "[8. Range 区域](#8-range-区域)",
            "[15. 可运行示例文件](#15-可运行示例文件)",
            "[17. API 选择指南](#17-api-选择指南)",
            "[18. 0.8.2 API 速查表](#18-082-api-速查表)",
        ):
            self.assertIn(entry, manual)

    def test_source_functions_have_complete_chinese_documentation(self) -> None:
        """功能：验证每个函数都有中文功能、用法、参数和返回说明。"""
        required = ("功能", "使用方法", "参数", "返回")
        missing: list[str] = []
        for path in sorted((_ROOT / "excelkit").rglob("*.py")):
            tree = ast.parse(path.read_text(encoding="utf-8"))
            for node in ast.walk(tree):
                if not isinstance(node, (ast.FunctionDef, ast.AsyncFunctionDef)):
                    continue
                if node.name.startswith("__"):
                    continue
                document = ast.get_docstring(node) or ""
                absent = [field for field in required if field not in document]
                if absent:
                    missing.append(
                        f"{path.relative_to(_ROOT)}:{node.lineno} {node.name} 缺少 {absent}"
                    )
        self.assertEqual(missing, [])

    def test_source_line_comments_are_chinese(self) -> None:
        """功能：验证代码行注释含中文说明，避免重新引入纯英文注释。"""
        invalid: list[str] = []
        for path in sorted((_ROOT / "excelkit").rglob("*.py")):
            source = path.read_text(encoding="utf-8")
            tokens = tokenize.generate_tokens(io.StringIO(source).readline)
            for token in tokens:
                if token.type != tokenize.COMMENT:
                    continue
                has_english = re.search(r"[A-Za-z]", token.string) is not None
                has_chinese = re.search(r"[\u4e00-\u9fff]", token.string) is not None
                if has_english and not has_chinese:
                    invalid.append(
                        f"{path.relative_to(_ROOT)}:{token.start[0]} {token.string}"
                    )
        self.assertEqual(invalid, [])

    def test_add_image_accepts_pathlike_filename(self) -> None:
        """功能：验证图片 API 与其他文件 API 一致地接受 pathlib.Path。"""
        png = (
            "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVQIHWP4z8DwHwAFgAI/"
            "9kL7UAAAAABJRU5ErkJggg=="
        )
        import base64
        with tempfile.TemporaryDirectory() as directory:
            filename = Path(directory) / "pixel.png"
            filename.write_bytes(base64.b64decode(png))
            image = Workbook().active.add_image(filename, anchor="A1")
        self.assertEqual(image.format, "png")


if __name__ == "__main__":
    unittest.main()
