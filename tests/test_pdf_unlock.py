import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import fitz

import core_unlock


def make_pdf(path, password):
    with fitz.open() as doc:
        doc.new_page().insert_text((50, 50), "PDF fixture")
        doc.save(path, encryption=fitz.PDF_ENCRYPT_AES_256, user_pw=password, owner_pw="owner")


class PDFUnlockTests(unittest.TestCase):
    def test_wrong_password_returns_error_without_result(self):
        with tempfile.TemporaryDirectory() as temp:
            source = Path(temp) / "protected.pdf"
            make_pdf(source, "correct")
            result = core_unlock.run_unlock(str(source), "wrong", False, True, 1, False)
            self.assertEqual(result["status"], "error", result)
            self.assertEqual(result["data"], {"succeeded": 0, "failed": 1})
            self.assertIn("密码错误", result["msg"])
            self.assertEqual(list(source.parent.glob("*_unlocked*.pdf")), [])

    def test_mixed_batch_reports_partial_and_preserves_existing_output(self):
        with tempfile.TemporaryDirectory() as temp:
            folder = Path(temp)
            make_pdf(folder / "good.pdf", "correct")
            make_pdf(folder / "bad.pdf", "different")
            old = folder / "good_unlocked.pdf"
            old.write_bytes(b"previous result")

            result = core_unlock.run_unlock(str(folder), "correct", False, True, 1, False)

            self.assertEqual(result["status"], "partial", result)
            self.assertEqual(result["data"], {"succeeded": 1, "failed": 1})
            self.assertEqual(old.read_bytes(), b"previous result")
            with fitz.open(folder / "good_unlocked (1).pdf") as output:
                self.assertFalse(output.needs_pass)
                self.assertIn("PDF fixture", output[0].get_text())

    def test_empty_path_does_not_start_worker(self):
        with patch.object(core_unlock.mp, "get_context") as get_context:
            result = core_unlock.run_unlock("  ", "", False, False, 1, False)
        self.assertEqual(result["status"], "error")
        get_context.assert_not_called()

    def test_failed_in_place_replace_preserves_original(self):
        with tempfile.TemporaryDirectory() as temp:
            source = Path(temp) / "protected.pdf"
            make_pdf(source, "correct")
            original = source.read_bytes()
            with patch.object(core_unlock.os, "replace", side_effect=PermissionError("locked")) as replace:
                result = core_unlock._process_one_pdf(str(source), "correct", False, False, 1)
            self.assertFalse(result["ok"])
            self.assertEqual(source.read_bytes(), original)
            self.assertEqual(replace.call_args.args[0].parent, source.parent)
            self.assertEqual(list(source.parent.glob(".unlock-*.pdf")), [])


if __name__ == "__main__":
    unittest.main()
