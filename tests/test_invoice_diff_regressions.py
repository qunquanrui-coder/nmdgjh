import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch

from docx import Document
from openpyxl import Workbook, load_workbook

import core_diff
import core_invoice


class InvoiceRegressionTests(unittest.TestCase):
    def test_negative_invoice_amounts_keep_their_sign(self):
        for minus in ("-", "−", "－"):
            with self.subTest(minus=minus):
                fields = core_invoice.get_info_strict_code(
                    f"价税合计{minus}113.00 不含税金额{minus}100.00 税额{minus}13.00 税率13%"
                )
                self.assertEqual(fields[3:], ("13%", "-100.00", "-13.00", "-113.00"))

    def test_missing_tax_rate_is_not_fabricated(self):
        fields = core_invoice.get_info_strict_code("价税合计101.00")
        self.assertEqual(fields[3:], ("", "", "", "101.00"))

    def test_existing_invoice_output_is_preserved(self):
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            (folder / "input.pdf").write_bytes(b"not needed")
            existing = folder / "发票汇总表.xlsx"
            existing.write_bytes(b"original")
            with patch.object(core_invoice, "extract_text_from_pdf", return_value="价税合计113.00不含税金额100.00税额13.00税率13%"):
                result = core_invoice.run_invoice(str(folder), False)
            self.assertEqual(result["status"], "success", result)
            self.assertEqual(existing.read_bytes(), b"original")
            self.assertTrue((folder / "发票汇总表 (1).xlsx").exists())


class DiffRegressionTests(unittest.TestCase):
    def test_word_table_text_is_compared(self):
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            for name, text in (("a.docx", "old"), ("b.docx", "new")):
                doc = Document()
                doc.add_paragraph("same")
                doc.add_table(rows=1, cols=1).cell(0, 0).text = text
                doc.save(folder / name)
            result = core_diff.run_diff(str(folder / "a.docx"), str(folder / "b.docx"), False)
            self.assertEqual(result["status"], "success", result)
            workbook = load_workbook(folder / "b_Word对比报告.xlsx")
            rows = list(workbook.active.values)
            self.assertTrue(any("old" in str(row) and "new" in str(row) for row in rows), rows)

    def test_excel_formula_changes_are_literal_and_existing_report_survives(self):
        with tempfile.TemporaryDirectory() as tmp:
            folder = Path(tmp)
            for name, formula in (("a.xlsx", "=1+1"), ("b.xlsx", "=2+2")):
                workbook = Workbook()
                workbook.active["A1"] = formula
                workbook.save(folder / name)
            existing = folder / "b_Excel对比报告.xlsx"
            existing.write_bytes(b"original")
            result = core_diff.run_diff(str(folder / "a.xlsx"), str(folder / "b.xlsx"), False)
            self.assertEqual(result["status"], "success", result)
            self.assertEqual(existing.read_bytes(), b"original")
            report = load_workbook(folder / "b_Excel对比报告 (1).xlsx", data_only=False)
            row = list(report.active.iter_rows())[1]
            self.assertEqual([cell.value for cell in row], ["Sheet", "A1 (公式)", "=1+1", "=2+2"])
            self.assertEqual([row[2].data_type, row[3].data_type], ["s", "s"])


if __name__ == "__main__":
    unittest.main()
