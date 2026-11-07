import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import core_word_merge


class WordMergeSafetyTests(unittest.TestCase):
    def test_missing_folder_and_unsafe_names_do_not_launch_word(self):
        with tempfile.TemporaryDirectory() as tmp, patch.object(core_word_merge.win32, "DispatchEx") as dispatch:
            for folder, name in (("   ", "merged.docx"), (tmp, " "), (tmp, "../outside.docx"),
                                 (tmp, r"C:\outside.docx"), (tmp, "NUL.docx")):
                with self.subTest(folder=folder, name=name):
                    result = core_word_merge.run_word_merge(folder, name)
                    self.assertEqual(result["status"], "error", result)
            dispatch.assert_not_called()

    @patch.object(core_word_merge.ctypes, "windll")
    @patch.object(core_word_merge, "pythoncom")
    @patch.object(core_word_merge.win32, "DispatchEx")
    def test_rerun_preserves_prior_outputs_and_merges_only_sources(self, dispatch, _com, _clipboard):
        with tempfile.TemporaryDirectory(dir=Path.cwd()) as tmp:
            root = Path(tmp)
            (root / "section2.docx").write_bytes(b"section 2")
            (root / "section10.docx").write_bytes(b"section 10")
            (root / "merged.docx").write_bytes(b"prior merge")
            (root / "merged (1).docx").write_bytes(b"second prior merge")
            word = dispatch.return_value
            merged = word.Documents.Add.return_value

            def save(path, file_format):
                self.assertEqual(file_format, 16)
                self.assertNotEqual(Path(path).parent, root)
                Path(path).write_bytes(b"complete merged result")

            merged.SaveAs2.side_effect = save
            result = core_word_merge.run_word_merge(root, "merged")

            self.assertEqual(result["status"], "success", result)
            self.assertIn("merged (2).docx", result["msg"])
            opened_names = [Path(call.args[0]).name for call in word.Documents.Open.call_args_list]
            self.assertEqual(opened_names, ["section2.docx", "section10.docx"])
            self.assertEqual((root / "merged.docx").read_bytes(), b"prior merge")
            self.assertEqual((root / "merged (1).docx").read_bytes(), b"second prior merge")
            self.assertEqual((root / "merged (2).docx").read_bytes(), b"complete merged result")

    @patch.object(core_word_merge.ctypes, "windll")
    @patch.object(core_word_merge, "pythoncom")
    @patch.object(core_word_merge.win32, "DispatchEx")
    def test_failed_save_does_not_publish_partial_output(self, dispatch, _com, _clipboard):
        with tempfile.TemporaryDirectory(dir=Path.cwd()) as tmp:
            root = Path(tmp)
            (root / "source.docx").write_bytes(b"source")
            (root / "merged.docx").write_bytes(b"prior merge")
            merged = dispatch.return_value.Documents.Add.return_value

            def save(path, _file_format):
                Path(path).write_bytes(b"partial output")
                raise OSError("write failed")

            merged.SaveAs2.side_effect = save
            result = core_word_merge.run_word_merge(root, "merged.docx")

            self.assertEqual(result["status"], "error", result)
            self.assertEqual((root / "merged.docx").read_bytes(), b"prior merge")
            self.assertEqual({p.name for p in root.iterdir()}, {"source.docx", "merged.docx"})
            merged.Close.assert_called_once_with(False)


if __name__ == "__main__":
    unittest.main()
