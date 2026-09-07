"""0.10.0 图片单元格锚定与读回回归测试。"""

from __future__ import annotations

import base64
import pathlib
import tempfile
import unittest
import zipfile
import xml.etree.ElementTree as ET

from excelkit import ImageFit, ImagePlacement, Workbook


_PNG = base64.b64decode(
    "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNk+A8AAQUBAScY42YAAAAASUVORK5CYII="
)
_NS = {"xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"}


class ImageReleaseTests(unittest.TestCase):
    """验证 0.10.0 图片 API、DrawingML 写出和读回。"""

    def test_cell_anchor_modes_and_roundtrip(self) -> None:
        """功能：验证区域锚定、三种填充策略和 XLSX 读回保持模型属性。"""
        for fit in (ImageFit.STRETCH, ImageFit.CONTAIN, ImageFit.COVER):
            with self.subTest(fit=fit), tempfile.TemporaryDirectory() as directory:
                path = pathlib.Path(directory) / "image.xlsx"
                workbook = Workbook()
                worksheet = workbook.active
                image = worksheet.add_image(
                    _PNG,
                    anchor="b2:f8",
                    name="pixel.png",
                    placement=ImagePlacement.CELL,
                    fit=fit,
                )
                self.assertIs(worksheet.image("B2"), image)
                self.assertIs(worksheet.image(0), image)
                workbook.save(path)
                with zipfile.ZipFile(path) as package:
                    root = ET.fromstring(package.read("xl/drawings/drawing1.xml"))
                    anchor = root.find("xdr:twoCellAnchor", _NS)
                    self.assertIsNotNone(anchor)
                    self.assertEqual(anchor.findtext("xdr:from/xdr:row", namespaces=_NS), "1")
                    self.assertEqual(anchor.findtext("xdr:from/xdr:col", namespaces=_NS), "1")
                    self.assertEqual(anchor.findtext("xdr:to/xdr:row", namespaces=_NS), "8")
                    self.assertEqual(anchor.findtext("xdr:to/xdr:col", namespaces=_NS), "6")
                loaded = Workbook.load(path)
                loaded_image = loaded.active.images[0]
                self.assertEqual(loaded_image.anchor, "B2:F8")
                self.assertEqual(loaded_image.placement, ImagePlacement.CELL)
                self.assertEqual(loaded_image.fit, fit)

    def test_floating_compatibility_and_axis_edit(self) -> None:
        """功能：验证旧版浮动图片和行列编辑后的锚点同步。"""
        workbook = Workbook()
        worksheet = workbook.active
        image = worksheet.add_image(_PNG, anchor="B2", name="pixel.png")
        worksheet.insert_rows(0)
        worksheet.insert_columns(0)
        self.assertEqual(image.anchor, "C3")
        self.assertEqual(worksheet.image("C3"), image)

    def test_invalid_combinations(self) -> None:
        """功能：验证浮动图片不能使用多单元格区域。"""
        workbook = Workbook()
        with self.assertRaises(ValueError):
            workbook.active.add_image(
                _PNG,
                anchor="A1:B2",
                name="pixel.png",
                placement=ImagePlacement.FLOATING,
            )


if __name__ == "__main__":
    unittest.main()
