#!/usr/bin/env python3
"""Gộp engine + writer + glue thành web/engine_bundle.py cho Pyodide."""
from pathlib import Path

root = Path(__file__).parent
engine = (root / "core" / "engine.py").read_text(encoding="utf-8")
writer = (root / "core" / "writer.py").read_text(encoding="utf-8")

# bỏ import tương đối trong writer (đã có sẵn trong cùng namespace)
writer = writer.replace(
    "from .engine import SyncResult, COL_PRICE, repair_shopee_xlsx", ""
)
writer = writer.replace("from __future__ import annotations", "")
writer = writer.replace("import shutil\n", "")
# bỏ import shutil trùng, giữ openpyxl imports
GLUE = '''

# ============================================================================
# Glue cho Pyodide — process_all() được gọi từ JavaScript
# ============================================================================
import base64 as _b64
import tempfile as _tf
from pathlib import Path as _P


def process_all():
    """Đọc XLSX_BYTES / CSV_BYTES (globals từ JS), trả dict kết quả + file b64."""
    xb = bytes(XLSX_BYTES)
    cb = bytes(CSV_BYTES)
    up = float(P_UP); down = float(P_DOWN); ratio = float(P_RATIO)
    rstep = int(P_ROUND); clear = bool(P_CLEAR); unlock = bool(P_UNLOCK)

    global ROUND_STEP
    ROUND_STEP = rstep

    d = _P(_tf.mkdtemp())
    xin = d / "in.xlsx"
    xin.write_bytes(xb)

    pm = load_price_map_from_csv_bytes(cb)
    result, _ = run_sync(
        str(xin), pm,
        max_up_pct=up, max_down_pct=down, ratio_limit=ratio,
        clear_unmatched=clear,
    )

    out = d / "import_shopee.xlsx"
    write_output(str(xin), result, str(out), unlock_protection=unlock)
    rep = d / "report.xlsx"
    write_report(result, str(rep))

    payload = result.to_dict()
    payload["import_b64"] = _b64.b64encode(out.read_bytes()).decode()
    payload["report_b64"] = _b64.b64encode(rep.read_bytes()).decode()
    # rows quá lớn, không cần gửi sang JS
    payload.pop("rows", None)
    payload.pop("groups", None)
    return payload
'''

bundle = engine + "\n\n# ====== WRITER ======\n" + writer + GLUE
outp = root / "web" / "engine_bundle.py"
outp.write_text(bundle, encoding="utf-8")
print("wrote", outp, len(bundle), "bytes")
