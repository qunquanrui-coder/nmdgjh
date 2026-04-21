# -*- coding: utf-8 -*-
import eel
import difflib
import pandas as pd
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

try:
    import docx
except Exception:
    docx = None

def _read_docx_lines(path: Path):
    d = docx.Document(str(path))
    lines = []
    for i, p in enumerate(d.paragraphs, 1):
        t = (p.text or "").strip()
        if t: lines.append((i, t))
    return lines

def _diff_text(a_lines, b_lines):
    a = [x[1] for x in a_lines]
    b = [x[1] for x in b_lines]
    sm = difflib.SequenceMatcher(a=a, b=b)
    changes = []
    for tag, i1, i2, j1, j2 in sm.get_opcodes():
        if tag == "equal": continue
        changes.append({
            "类型": tag,
            "原段落区间": f"{i1+1}-{i2}",
            "新段落区间": f"{j1+1}-{j2}",
            "原内容": "\n".join(a[i1:i2]),
            "修改后": "\n".join(b[j1:j2]),
        })
    return changes

def _write_styled_report(out: Path, df: pd.DataFrame, is_word: bool):
    if df.empty:
        if is_word: df = pd.DataFrame([{"类型":"无差异","原段落区间":"","新段落区间":"","原内容":"","修改后":""}])
        else: df = pd.DataFrame([{"Sheet":"","Cell":"","Old":"","New":"无差异"}])
        
    with pd.ExcelWriter(out, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="差异报告")
        
    wb = load_workbook(out)
    ws = wb["差异报告"]
    ws.freeze_panes = "A2"
    
    header_fill = PatternFill("solid", fgColor="1F4E79")
    header_font = Font(color="FFFFFF", bold=True)
    
    for c in range(1, ws.max_column + 1):
        cell = ws.cell(1, c)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        if is_word:
            ws.column_dimensions[get_column_letter(c)].width = 24 if c < 4 else 60
        else:
            ws.column_dimensions[get_column_letter(c)].width = 18 if c <= 2 else 60
    wb.save(out)

@eel.expose
def run_diff(a_path, b_path, strict):
    try:
        a, b = Path(a_path), Path(b_path)
        if a.suffix.lower() != b.suffix.lower():
            return {"status": "error", "msg": "两个文件类型必须一致（同为 docx 或 xlsx）"}
            
        ext = a.suffix.lower()
        if ext == ".docx":
            if docx is None: return {"status": "error", "msg": "未安装 python-docx 依赖"}
            eel.update_terminal("[*] 正在进行 Word 段落深度比对...")
            changes = _diff_text(_read_docx_lines(a), _read_docx_lines(b))
            out = b.parent / f"{b.stem}_Word对比报告.xlsx"
            _write_styled_report(out, pd.DataFrame(changes), is_word=True)
            return {"status": "success", "msg": f"报告已生成: {out.name}"}
            
        elif ext in (".xlsx", ".xls"):
            eel.update_terminal("[*] 正在进行 Excel 逐单元格比对...")
            wa, wb = load_workbook(a, data_only=True), load_workbook(b, data_only=True)
            changes = []
            sheets = set(wa.sheetnames) | set(wb.sheetnames)
            for s in sorted(sheets):
                sa = wa[s] if s in wa.sheetnames else None
                sb = wb[s] if s in wb.sheetnames else None
                max_r = max(sa.max_row if sa else 1, sb.max_row if sb else 1)
                max_c = max(sa.max_column if sa else 1, sb.max_column if sb else 1)
                for r in range(1, max_r + 1):
                    for c in range(1, max_c + 1):
                        va = sa.cell(r, c).value if sa else None
                        vb = sb.cell(r, c).value if sb else None
                        if va != vb:
                            changes.append({"Sheet": s, "Cell": f"{get_column_letter(c)}{r}", "Old": str(va) if va is not None else "", "New": str(vb) if vb is not None else ""})
                            
            out = b.parent / f"{b.stem}_Excel对比报告.xlsx"
            _write_styled_report(out, pd.DataFrame(changes), is_word=False)
            return {"status": "success", "msg": f"报告已生成: {out.name}"}
        else:
            return {"status": "error", "msg": "仅支持 docx 或 xlsx 比对"}
            
    except Exception as e: 
        return {"status": "error", "msg": str(e)}
