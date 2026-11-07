import tempfile
import unittest
from pathlib import Path

import fitz
from PIL import Image

import core_pdf2img
import core_split


def write_pdf(path, page_count):
    with fitz.open() as doc:
        for index in range(page_count):
            page = doc.new_page()
            page.insert_text((30, 30), f"Page {index + 1}")
        doc.save(path)


class PDFExportTests(unittest.TestCase):
    def test_split_keeps_existing_export(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "source.pdf"
            write_pdf(source, 3)
            old_dir = root / "source_split_fixed"
            old_dir.mkdir()
            old_result = old_dir / "source_Part1_P1-2.pdf"
            write_pdf(old_result, 1)
            original_bytes = old_result.read_bytes()

            result = core_split.run_split(source, "split_fixed", 2, 1, 1, 1)

            self.assertEqual(result["status"], "success")
            self.assertEqual(old_result.read_bytes(), original_bytes)
            new_dir = root / "source_split_fixed (1)"
            with fitz.open(new_dir / old_result.name) as doc:
                self.assertEqual(len(doc), 2)
            with fitz.open(new_dir / "source_Part2_P3-3.pdf") as doc:
                self.assertIn("Page 3", doc[0].get_text())

    def test_image_export_uses_separate_directory_without_stale_pages(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "source.pdf"
            write_pdf(source, 2)
            first = core_pdf2img.run_pdf2img(source, 200, 95)
            self.assertEqual(first["status"], "success")
            first_dir = root / "source_images"
            old_page = first_dir / "0002.jpg"
            original_bytes = old_page.read_bytes()

            write_pdf(source, 1)
            second = core_pdf2img.run_pdf2img(source, 200, 95)

            self.assertEqual(second["status"], "success")
            self.assertEqual(old_page.read_bytes(), original_bytes)
            self.assertEqual({p.name for p in (root / "source_images (1)").iterdir()}, {"0001.jpg"})

    def test_image_export_honors_offset_crop_and_rotation(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            source = root / "source.pdf"
            with fitz.open() as doc:
                for rotation in (0, 90):
                    page = doc.new_page(width=600, height=800)
                    page.draw_rect(fitz.Rect(465, 600, 495, 640), color=None, fill=(1, 0, 0))
                    page.set_cropbox(fitz.Rect(100, 200, 500, 650))
                    page.set_rotation(rotation)
                doc.save(source)

            result = core_pdf2img.run_pdf2img(source, 900, 95)

            self.assertEqual(result["status"], "success")
            out_dir = root / "source_images"
            with Image.open(out_dir / "0001.jpg") as image:
                self.assertEqual(image.size, (800, 900))
                self.assertEqual(image.info["dpi"], (144, 144))
                red, green, blue = image.getpixel((760, 840))
                self.assertGreater(red, 200)
                self.assertLess(max(green, blue), 50)
            with Image.open(out_dir / "0002.jpg") as image:
                self.assertEqual(image.size, (900, 800))
                self.assertEqual(image.info["dpi"], (144, 144))

    def test_blank_path_is_rejected(self):
        self.assertEqual(core_split.run_split("  ", "split_fixed", 1, 1, 1, 1)["status"], "error")
        self.assertEqual(core_pdf2img.run_pdf2img("  ", 100, 95)["status"], "error")


if __name__ == "__main__":
    unittest.main()
