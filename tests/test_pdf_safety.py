import tempfile
import unittest
from pathlib import Path

import fitz

import core_blank_page
import core_compress
import core_pdf_cleaner
import core_pdf_replace


class PDFSafetyTests(unittest.TestCase):
    def test_cleaner_does_not_whiten_page_with_enclosing_border(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / 'source.pdf'
            doc = fitz.open()
            page = doc.new_page(width=300, height=300)
            page.draw_rect(fitz.Rect(0, 0, 300, 300), color=(0, 0, 0), width=8)
            page.insert_text((80, 150), 'IMPORTANT DATA', fontsize=14)
            doc.save(source)
            doc.close()
            result = core_pdf_cleaner.run_pdf_cleaner(str(source))
            self.assertEqual(result['status'], 'success')
            with fitz.open(Path(tmp) / 'source_去边.pdf') as out:
                pix = out[0].get_pixmap()
                self.assertLess(min(pix.samples), 128)
                self.assertIn('IMPORTANT DATA', out[0].get_text())

    def test_page_replacement_restores_links_to_replaced_pages(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp = Path(tmp)
            doc = fitz.open()
            for _ in range(4):
                doc.new_page()
            doc[0].insert_link({'kind': fitz.LINK_GOTO, 'from': fitz.Rect(20, 20, 50, 50), 'page': 1, 'to': fitz.Point(0, 0)})
            doc[3].insert_link({'kind': fitz.LINK_GOTO, 'from': fitz.Rect(20, 20, 50, 50), 'page': 2, 'to': fitz.Point(0, 0)})
            doc.save(tmp / 'source.pdf')
            doc.close()
            replacement = fitz.open()
            replacement.new_page()
            replacement.save(tmp / 'replacement.pdf')
            replacement.close()
            core_pdf_replace._replace_pages_in_one_pdf(tmp / 'source.pdf', tmp / 'out.pdf', '2-3', tmp / 'replacement.pdf', '1')
            with fitz.open(tmp / 'out.pdf') as out:
                self.assertEqual(len(out), 3)
                self.assertEqual(out[0].get_links()[0]['page'], 1)
                self.assertEqual(out[2].get_links()[0]['page'], 1)

    def test_backups_are_excluded_from_batch_cleanup(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp = Path(tmp)
            for name in ['document.pdf', 'document.blankbak.pdf', 'document.blankbak_123.pdf', 'document.blankbak_123 (1).pdf', 'another.docx']:
                (tmp / name).touch()
            self.assertEqual({p.name for p in core_blank_page.collect_target_files(tmp)}, {'document.pdf', 'another.docx'})

    def test_backup_is_independent_from_original(self):
        with tempfile.TemporaryDirectory() as tmp:
            source = Path(tmp) / 'test.docx'
            source.write_bytes(b'original')
            ok, backup, error = core_blank_page.make_backup_before_overwrite(source)
            self.assertTrue(ok, error)
            source.write_bytes(b'changed')
            self.assertEqual(backup.read_bytes(), b'original')

    def test_cleanup_uses_same_volume_and_saves_backup(self):
        with tempfile.TemporaryDirectory(dir=Path.cwd()) as tmp:
            tmp = Path(tmp)
            doc = fitz.open()
            page = doc.new_page()
            page.insert_text((30, 30), 'keep')
            doc.new_page()
            source = tmp / 'source.pdf'
            doc.save(source)
            doc.close()
            original = source.read_bytes()
            status, removed = core_blank_page.process_pdf_core(source)
            self.assertEqual(removed, [2], status)
            self.assertEqual((tmp / 'source.blankbak.pdf').read_bytes(), original)
            with fitz.open(source) as result:
                self.assertEqual(len(result), 1)
            self.assertEqual(core_blank_page.collect_target_files(tmp), [source])

    def test_filename_mapping_restores_links_to_replaced_pages(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp = Path(tmp)
            doc = fitz.open()
            for _ in range(3):
                doc.new_page()
            doc[0].insert_link({'kind': fitz.LINK_GOTO, 'from': fitz.Rect(20, 20, 50, 50), 'page': 1, 'to': fitz.Point(0, 0)})
            doc.save(tmp / 'source.pdf')
            doc.close()
            replacement = fitz.open()
            replacement.new_page()
            replacement.save(tmp / 'replacement.pdf')
            replacement.close()
            core_pdf_replace._replace_pages_by_filename_pdf_mapping(tmp / 'source.pdf', tmp / 'out.pdf', tmp / 'replacement.pdf', [(1, 0)])
            with fitz.open(tmp / 'out.pdf') as out:
                self.assertEqual(out[0].get_links()[0]['page'], 1)

    def test_compress_output_preserves_existing_file(self):
        with tempfile.TemporaryDirectory() as tmp:
            tmp = Path(tmp)
            (tmp / 'output.pdf').write_bytes(b'old')
            (tmp / 'candidate.pdf').write_bytes(b'new')
            result = core_compress._copy_to_unique_output(tmp / 'candidate.pdf', tmp / 'output.pdf')
            self.assertEqual((tmp / 'output.pdf').read_bytes(), b'old')
            self.assertEqual(result.read_bytes(), b'new')
            self.assertNotEqual(result, tmp / 'output.pdf')


if __name__ == '__main__':
    unittest.main()
