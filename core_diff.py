# -*- coding: utf-8 -*-
import bridge
import difflib
import threading
import time
import tempfile
import pandas as pd
from pathlib import Path
from openpyxl import load_workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter
from safe_output import publish_output

try:
    import docx
except Exception:
    docx = None


def _start_heartbeat(label, state, interval=5):
    stop_event = threading.Event()
    start_time = time.time()

    def worker():
        while not stop_event.wait(interval):
            stage = state.get("stage", "processing")
            current = state.get("current")
            total = state.get("total")
            elapsed = int(time.time() - start_time)
            if current is not None and total:
                bridge.update_terminal(f"HB [{label}] {stage}: {current}/{total}, elapsed {elapsed}s")
            else:
                bridge.update_terminal(f"HB [{label}] {stage}, elapsed {elapsed}s")

    thread = threading.Thread(target=worker, daemon=True)
    thread.start()
    return stop_event, thread


def _stop_heartbeat(stop_event, thread):
    stop_event.set()
    try:
        thread.join(timeout=1)
    except Exception:
        pass


def _read_docx_lines(path: Path, progress=None):
    d = docx.Document(str(path))
    lines = []
    from docx.oxml.table import CT_Tbl
    from docx.oxml.text.paragraph import CT_P
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    def read_blocks(container, element, location=""):
        table_no = 0
        for child in element.iterchildren():
            if isinstance(child, CT_P):
                value = (Paragraph(child, container).text or "").strip()
                if value:
                    lines.append((len(lines) + 1, f"{location}{value}"))
            elif isinstance(child, CT_Tbl):
                table_no += 1
                table = Table(child, container)
                for row_no, row in enumerate(table.rows, 1):
                    for col_no, cell in enumerate(row.cells, 1):
                        label = f"{location}表格{table_no}第{row_no}行第{col_no}列："
                        read_blocks(cell, cell._tc, label)
            if progress and len(lines) and len(lines) % 200 == 0:
                progress("read Word paragraphs and tables", len(lines), None)

    read_blocks(d, d.element.body)
    return lines


def _diff_text(a_lines, b_lines, progress=None):
    a = [x[1] for x in a_lines]
    b = [x[1] for x in b_lines]
    if progress:
        progress("compare paragraphs", None, None)
    sm = difflib.SequenceMatcher(a=a, b=b)
    changes = []
    opcodes = sm.get_opcodes()
    total = len(opcodes)
    for idx, (tag, i1, i2, j1, j2) in enumerate(opcodes, 1):
        if progress and (idx % 100 == 0 or idx == total):
            progress("scan paragraph changes", idx, total)
        if tag == "equal":
            continue
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
        if is_word:
            df = pd.DataFrame([{"类型":"无差异","原段落区间":"","新段落区间":"","原内容":"","修改后":""}])
        else:
            df = pd.DataFrame([{"Sheet":"","Cell":"","Old":"","New":"无差异"}])

    with pd.ExcelWriter(out, engine="openpyxl") as w:
        df.to_excel(w, index=False, sheet_name="差异报告")

    wb = load_workbook(out)
    ws = wb["差异报告"]
    # Source document text is data, even when it starts with '='.
    for row in ws.iter_rows():
        for cell in row:
            if cell.data_type == "f":
                cell.data_type = "s"
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


@bridge.expose
def run_diff(a_path, b_path, strict):
    state = {"stage": "prepare", "current": None, "total": None}
    stop_event, heartbeat_thread = _start_heartbeat("diff", state)
    try:
        def progress(stage, current=None, total=None):
            state["stage"] = stage
            state["current"] = current
            state["total"] = total

        a, b = Path(a_path), Path(b_path)
        if a.suffix.lower() != b.suffix.lower():
            return {"status": "error", "msg": "两个文件类型必须一致（同为 docx 或 xlsx）"}

        ext = a.suffix.lower()
        if ext == ".docx":
            if docx is None:
                return {"status": "error", "msg": "未安装 python-docx 依赖"}
            bridge.update_terminal("[*] 正在进行 Word 段落深度比对...")
            progress(f"read {a.name}")
            a_lines = _read_docx_lines(a, progress)
            progress(f"read {b.name}")
            b_lines = _read_docx_lines(b, progress)
            changes = _diff_text(a_lines, b_lines, progress)
            out = b.parent / f"{b.stem}_Word对比报告.xlsx"
            progress("write report")
            with tempfile.TemporaryDirectory(prefix="diff_") as tmp_dir:
                tmp_output = Path(tmp_dir) / out.name
                _write_styled_report(tmp_output, pd.DataFrame(changes), is_word=True)
                out = publish_output(tmp_output, out)
            return {"status": "success", "msg": f"报告已生成: {out.name}"}

        elif ext in (".xlsx", ".xls"):
            bridge.update_terminal("[*] 正在进行 Excel 逐单元格比对...")
            progress(f"open {a.name}")
            wa = load_workbook(a, data_only=True)
            progress(f"open {b.name}")
            wb = load_workbook(b, data_only=True)
            fa = load_workbook(a, data_only=False)
            fb = load_workbook(b, data_only=False)
            changes = []
            sheets = sorted(set(wa.sheetnames) | set(wb.sheetnames))
            for sheet_idx, s in enumerate(sheets, 1):
                sa = wa[s] if s in wa.sheetnames else None
                sb = wb[s] if s in wb.sheetnames else None
                fsa = fa[s] if s in fa.sheetnames else None
                fsb = fb[s] if s in fb.sheetnames else None
                max_r = max(sa.max_row if sa else 1, sb.max_row if sb else 1)
                max_c = max(sa.max_column if sa else 1, sb.max_column if sb else 1)
                progress(f"scan sheet {s}", sheet_idx, len(sheets))
                for r in range(1, max_r + 1):
                    if r % 500 == 0 or r == max_r:
                        bridge.update_terminal(f"  └─ Sheet {s}: row {r}/{max_r}")
                    for c in range(1, max_c + 1):
                        va = sa.cell(r, c).value if sa else None
                        vb = sb.cell(r, c).value if sb else None
                        if va != vb:
                            changes.append({"Sheet": s, "Cell": f"{get_column_letter(c)}{r}", "Old": str(va) if va is not None else "", "New": str(vb) if vb is not None else ""})
                        ca = fsa.cell(r, c) if fsa else None
                        cb = fsb.cell(r, c) if fsb else None
                        old_formula = ca.value if ca and ca.data_type == "f" else None
                        new_formula = cb.value if cb and cb.data_type == "f" else None
                        if old_formula != new_formula:
                            changes.append({"Sheet": s, "Cell": f"{get_column_letter(c)}{r} (公式)", "Old": old_formula or "", "New": new_formula or ""})

            out = b.parent / f"{b.stem}_Excel对比报告.xlsx"
            progress("write report")
            with tempfile.TemporaryDirectory(prefix="diff_") as tmp_dir:
                tmp_output = Path(tmp_dir) / out.name
                _write_styled_report(tmp_output, pd.DataFrame(changes), is_word=False)
                out = publish_output(tmp_output, out)
            return {"status": "success", "msg": f"报告已生成: {out.name}"}
        else:
            return {"status": "error", "msg": "仅支持 docx 或 xlsx 比对"}

    except Exception as e:
        return {"status": "error", "msg": str(e)}
    finally:
        _stop_heartbeat(stop_event, heartbeat_thread)
