"""ExcelKit 发布资料一致性测试。"""

from __future__ import annotations

import re
import tempfile
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
            "[18. 0.8.0 API 速查表](#18-080-api-速查表)",
        ):
            self.assertIn(entry, manual)

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
