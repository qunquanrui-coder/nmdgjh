import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

import build_modern


class BuildSafetyTests(unittest.TestCase):
    def test_previous_build_and_other_dist_files_are_preserved(self):
        with tempfile.TemporaryDirectory() as tmp:
            dist = Path(tmp)
            old = dist / build_modern.APP_NAME
            old.mkdir()
            (old / "old.txt").write_text("old")
            (dist / "old-package.zip").write_text("keep")
            stage = dist / "stage"
            stage.mkdir()
            (stage / "new.txt").write_text("new")
            with patch.object(build_modern, "DIST_DIR", dist):
                build_modern.install_build(stage)
            self.assertEqual((dist / f"{build_modern.APP_NAME}_previous" / "old.txt").read_text(), "old")
            self.assertEqual((old / "new.txt").read_text(), "new")
            self.assertEqual((dist / "old-package.zip").read_text(), "keep")

    def test_failure_before_install_leaves_old_package_in_place(self):
        with tempfile.TemporaryDirectory() as tmp:
            root = Path(tmp)
            dist = root / "dist"
            old = dist / build_modern.APP_NAME
            old.mkdir(parents=True)
            (old / "old.txt").write_text("old")
            with patch.multiple(build_modern, DIST_DIR=dist, BUILD_DIR=root / "build"), \
                 patch.object(build_modern, "ensure_exists"), \
                 patch.object(build_modern, "create_version_file"), \
                 patch.object(build_modern.subprocess, "run", side_effect=OSError("failed")):
                with self.assertRaises(OSError):
                    build_modern.build()
            self.assertEqual((old / "old.txt").read_text(), "old")


if __name__ == "__main__":
    unittest.main()
