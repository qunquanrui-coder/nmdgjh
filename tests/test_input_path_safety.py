import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import fitz

import core_blank_page
import core_compress
import core_diff
import core_img2pdf
import core_invoice
import core_ocr
import core_pdf2word
import core_pdf_cleaner
import core_word2pdf
import core_word_split


class InputPathSafetyTests(unittest.TestCase):
    def test_blank_cleanup_does_not_touch_files_in_working_directory(self):
        with tempfile.TemporaryDirectory() as directory:
            source = Path(directory) / "unselected.pdf"
            with fitz.open() as doc:
                doc.new_page().insert_text((72, 72), "Unselected document")
                doc.new_page()
                doc.save(source)
            original = source.read_bytes()
            previous_directory = Path.cwd()
            os.chdir(directory)
            try:
                for path in ("", " \t "):
                    with self.subTest(path=path):
                        result = core_blank_page.run_rm_blank(path)
                        self.assertEqual(result["status"], "error")
                        self.assertEqual(source.read_bytes(), original)
                        self.assertEqual(list(Path(directory).iterdir()), [source])
            finally:
                os.chdir(previous_directory)

    def test_batch_tools_do_not_resolve_empty_paths(self):
        tasks = (
            (core_img2pdf, core_img2pdf.run_img2pdf, (True, True)),
            (core_invoice, core_invoice.run_invoice, (True,)),
            (core_ocr, core_ocr.run_ocr, ()),
            (core_pdf2word, core_pdf2word.run_pdf2word, ("image", 150)),
            (core_pdf_cleaner, core_pdf_cleaner.run_pdf_cleaner, ()),
            (core_word2pdf, core_word2pdf.run_word2pdf, (True, True, True, "标题", False)),
            (core_compress, core_compress.run_compress, (100, "MB")),
        )
        for module, task, args in tasks:
            for path in ("", " \t "):
                with self.subTest(task=task.__name__, path=path), patch.object(module, "Path") as resolve_path:
                    result = task(path, *args)
                    self.assertEqual(result["status"], "error")
                    resolve_path.assert_not_called()

    def test_diff_rejects_empty_paths_and_legacy_xls(self):
        self.assertEqual(core_diff.run_diff("  ", "b.docx", False)["status"], "error")
        with tempfile.TemporaryDirectory() as tmp:
            a = Path(tmp) / "a.xls"
            b = Path(tmp) / "b.xls"
            a.write_bytes(b"xls1")
            b.write_bytes(b"xls2")
            res = core_diff.run_diff(str(a), str(b), False)
            self.assertEqual(res["status"], "error")
            self.assertIn(".xlsx", res["msg"])

    def test_compress_rejects_non_positive_target_size(self):
        with tempfile.TemporaryDirectory() as tmp:
            pdf = Path(tmp) / "sample.pdf"
            with fitz.open() as doc:
                doc.new_page()
                doc.save(pdf)
            for bad_size in (0, -5, "abc"):
                res = core_compress.run_compress(str(pdf), bad_size, "MB")
                self.assertEqual(res["status"], "error")

    def test_word_split_requires_output_location(self):
        with patch.object(core_word_split._engine, "split_document") as split_document:
            result = core_word_split.run_word_split("selected.docx", " \t ", "1", 1)
            self.assertEqual(result["status"], "error")
            split_document.assert_not_called()
