import errno
import tempfile
import unittest
from pathlib import Path
from unittest import mock

import fitz
import core_img2pdf
import core_ocr
import core_pdf2word
import core_word2pdf
import safe_output
from PIL import Image


class OutputSafetyTests(unittest.TestCase):
    def test_existing_result_is_not_overwritten(self):
        with tempfile.TemporaryDirectory() as directory:
            folder = Path(directory)
            old = folder / "out.pdf"
            old.write_bytes(b"old")
            tmp = folder / "tmp.pdf"
            tmp.write_bytes(b"new")
            actual = safe_output.publish_output(tmp, old)
            self.assertEqual(old.read_bytes(), b"old")
            self.assertEqual(actual.name, "out (1).pdf")
            self.assertEqual(actual.read_bytes(), b"new")

    def test_fallback_copy_failure_removes_incomplete_output(self):
        with tempfile.TemporaryDirectory() as directory:
            folder = Path(directory)
            tmp = folder / "tmp.pdf"
            tmp.write_bytes(b"new")
            desired = folder / "out.pdf"
            desired.write_bytes(b"old")
            def fail_copy(source, output):
                output.write(b"partial")
                raise OSError("disk full")
            with mock.patch.object(safe_output.os, "link", side_effect=OSError(errno.EXDEV, "different volume")), \
                 mock.patch.object(safe_output.shutil, "copyfileobj", side_effect=fail_copy):
                with self.assertRaises(OSError):
                    safe_output.publish_output(tmp, desired)
            self.assertEqual(desired.read_bytes(), b"old")
            self.assertEqual(tmp.read_bytes(), b"new")
            self.assertFalse((folder / "out (1).pdf").exists())

    def test_image_conversion_failure_preserves_old_output(self):
        with tempfile.TemporaryDirectory() as directory:
            folder = Path(directory)
            Image.new("RGB", (4, 4), "white").save(folder / "a.jpg")
            old = folder / (folder.name + ".pdf")
            old.write_bytes(b"old")
            with mock.patch.object(core_img2pdf.img2pdf, "convert", side_effect=ValueError("conversion failed")):
                result = core_img2pdf.run_img2pdf(str(folder), False, False)
            self.assertEqual(result["status"], "error")
            self.assertEqual(old.read_bytes(), b"old")
            self.assertEqual(list(folder.glob(".img2pdf_*.pdf")), [])

    def test_image_conversion_uses_new_name(self):
        with tempfile.TemporaryDirectory() as directory:
            folder = Path(directory)
            Image.new("RGB", (4, 4), "white").save(folder / "a.jpg")
            old = folder / (folder.name + ".pdf")
            old.write_bytes(b"old")
            result = core_img2pdf.run_img2pdf(str(folder), False, False)
            self.assertEqual(result["status"], "success")
            self.assertEqual(old.read_bytes(), b"old")
            self.assertTrue((folder / (folder.name + " (1).pdf")).read_bytes().startswith(b"%PDF"))

    def test_pdf2word_preserves_existing_docx_and_cleans_temp_on_failure(self):
        with tempfile.TemporaryDirectory() as directory:
            folder = Path(directory)
            source = folder / "sample.pdf"
            with fitz.open() as doc:
                doc.new_page().insert_text((36, 36), "Hello")
                doc.save(source)
            existing = folder / "sample.docx"
            existing.write_bytes(b"existing docx")

            res = core_pdf2word.run_pdf2word(str(source), "image", 72)
            self.assertEqual(res["status"], "success", res)
            self.assertEqual(existing.read_bytes(), b"existing docx")
            self.assertTrue((folder / "sample (1).docx").exists())

    def test_ocr_preserves_existing_output(self):
        with tempfile.TemporaryDirectory() as directory:
            folder = Path(directory)
            for name in ("chi_sim.traineddata", "eng.traineddata"):
                (folder / name).touch()
            source = folder / "source.pdf"
            source.write_bytes(b"source")
            old = folder / "source_可搜索.pdf"
            old.write_bytes(b"old")
            def fake_ocr(_source, output, **kwargs):
                Path(output).write_bytes(b"new")
            with mock.patch.object(core_ocr, "TESSDATA_DIR", folder), \
                 mock.patch.object(core_ocr, "_heartbeat_worker"), \
                 mock.patch.object(core_ocr.ocrmypdf, "ocr", side_effect=fake_ocr):
                result = core_ocr._run_single_ocr(source, 1)
            self.assertEqual(result["status"], "success")
            self.assertEqual(old.read_bytes(), b"old")
            self.assertEqual(Path(result["data"]).read_bytes(), b"new")
            self.assertNotEqual(Path(result["data"]), old)

    def test_word_and_excel_same_stem_and_existing_pdf(self):
        with tempfile.TemporaryDirectory() as directory:
            folder = Path(directory)
            (folder / "same.docx").touch()
            (folder / "same.xlsx").touch()
            old = folder / "same.pdf"
            old.write_bytes(b"old")

            class Document:
                def ExportAsFixedFormat(self, path, *args, **kwargs):
                    Path(path).write_bytes(b"word")
                def Repaginate(self): pass
                def Close(self, *args): pass
                TablesOfContents = type("ToCs", (), {"Count": 0})()
            class Workbook:
                def ExportAsFixedFormat(self, fmt, path, **kwargs):
                    Path(path).write_bytes(b"excel")
                def Close(self, *args): pass
            class App:
                def __init__(self, kind):
                    self.kind = kind
                    self.AutomationSecurity = None
                    self.Options = type("Options", (), {})()
                    self.Documents = self
                    self.Workbooks = self
                def Open(self, *args, **kwargs):
                    assert self.AutomationSecurity == 3
                    return Document() if self.kind == "Word.Application" else Workbook()
                def Quit(self, *args): pass
            with mock.patch.object(core_word2pdf.ComManager, "init"), \
                 mock.patch.object(core_word2pdf.ComManager, "uninit"), \
                 mock.patch.object(core_word2pdf.win32com.client, "DispatchEx", side_effect=App):
                result = core_word2pdf.run_word2pdf(str(folder), False, False, True, "", False)
            self.assertEqual(result["status"], "success")
            self.assertEqual(old.read_bytes(), b"old")
            self.assertEqual({(folder / "same (1).pdf").read_bytes(), (folder / "same (2).pdf").read_bytes()}, {b"word", b"excel"})


if __name__ == "__main__":
    unittest.main()
