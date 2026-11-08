"""图片插入测试（PNG/JPEG，浮动/嵌入）。"""

import base64
import os
import sys
import tempfile
import unittest
import zipfile
import zlib

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from excelkit import UnsupportedFeatureError, Workbook
from excelkit.errors import InvalidFileError


def make_png(w=2, h=3):
    sig = b"\x89PNG\r\n\x1a\n"
    def chunk(ctype, data):
        c = ctype + data
        return len(data).to_bytes(4, "big") + c + (zlib.crc32(c) & 0xFFFFFFFF).to_bytes(4, "big")
    ihdr = chunk(b"IHDR", w.to_bytes(4, "big") + h.to_bytes(4, "big") + b"\x08\x02\x00\x00\x00")
    raw = b"".join(b"\x00" + b"\xff\x00\x00" * w for _ in range(h))
    idat = chunk(b"IDAT", zlib.compress(raw))
    iend = chunk(b"IEND", b"")
    return sig + ihdr + idat + iend


# 最小可解析 JPEG 结构（含 SOF0），用于尺寸检测
JPEG_4x5 = bytes.fromhex(
    "ffd8ffc00011080005000403010122000211030311ffd9"
)


class TestImage(unittest.TestCase):
    def _save_load(self, wb):
        fd, path = tempfile.mkstemp(suffix=".xlsx")
        os.close(fd)
        try:
            wb.save(path)
            return path, Workbook.load(path)
        finally:
            try:
                pass
            except Exception:
                pass

    def test_add_float_image_and_roundtrip(self):
        wb = Workbook()
        ws = wb.active
        ws["A1"] = "x"
        png = make_png(10, 20)
        img = ws.add_image("B2", png, mode="float", alt_text="logo")
        self.assertEqual(len(ws.images), 1)
        self.assertEqual(img.mode, "float")
        self.assertEqual(img.natural_width, 10)
        self.assertEqual(img.natural_height, 20)
        path, wb2 = self._save_load(wb)
        try:
            ws2 = wb2.active
            self.assertEqual(len(ws2.images), 1)
            im2 = ws2.images[0]
            self.assertEqual(im2.anchor, "B2")
            self.assertEqual(im2.mode, "float")
            self.assertEqual(im2.natural_width, 10)
            self.assertEqual(im2.natural_height, 20)
            with zipfile.ZipFile(path) as z:
                names = z.namelist()
                self.assertTrue(any("xl/media/" in n for n in names))
                self.assertTrue(any("xl/drawings/drawing" in n for n in names))
                sheet = z.read("xl/worksheets/sheet1.xml").decode("utf-8", "ignore")
                self.assertIn("drawing", sheet)
        finally:
            os.unlink(path)

    def test_add_embed_image(self):
        wb = Workbook()
        ws = wb.active
        ws.add_image("B2:D4", make_png(3, 3), mode="embed")
        self.assertEqual(ws.images[0].mode, "embed")
        path, wb2 = self._save_load(wb)
        try:
            im2 = wb2.active.images[0]
            self.assertEqual(im2.mode, "embed")
            self.assertIn("B2", im2.anchor)
            self.assertIn("D4", im2.anchor)
        finally:
            os.unlink(path)

    def test_embed_expands_merged_anchor(self):
        wb = Workbook()
        ws = wb.active
        ws.range("B2:C3").merge()
        ws.add_image("B2", make_png(4, 4), mode="embed")
        self.assertEqual(ws.images[0].anchor, "B2:C3")

    def test_jpeg_size_and_roundtrip(self):
        wb = Workbook()
        ws = wb.active
        ws.add_image("A1", JPEG_4x5, mode="float")
        self.assertEqual(ws.images[0].natural_width, 4)
        self.assertEqual(ws.images[0].natural_height, 5)
        path, wb2 = self._save_load(wb)
        try:
            self.assertEqual(wb2.active.images[0].natural_width, 4)
        finally:
            os.unlink(path)

    def test_xls_rejects_images(self):
        wb = Workbook()
        wb.active.add_image("A1", make_png(1, 1))
        fd, path = tempfile.mkstemp(suffix=".xls")
        os.close(fd)
        try:
            with self.assertRaises(UnsupportedFeatureError):
                wb.save(path)
        finally:
            if os.path.exists(path):
                os.unlink(path)

    def test_embed_rejects_size(self):
        wb = Workbook()
        with self.assertRaises(ValueError):
            wb.active.add_image("A1", make_png(1, 1), mode="embed", width=10)

    def test_remove_and_clear(self):
        wb = Workbook()
        ws = wb.active
        im = ws.add_image("A1", make_png(1, 1))
        ws.remove_image(im)
        self.assertEqual(len(ws.images), 0)
        ws.add_image("A1", make_png(1, 1))
        ws.add_image("B1", make_png(1, 1))
        ws.clear_images()
        self.assertEqual(len(ws.images), 0)


if __name__ == "__main__":
    unittest.main()
