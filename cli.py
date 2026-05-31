#!/usr/bin/env python3
"""
Shopee Price Sync — CLI
=======================

Ví dụ:
    python cli.py --xlsx SHOPEE.xlsx --csv BANG_GIA.csv --out import_shopee.xlsx

Tuỳ chọn:
    --up 0.5        Tăng tối đa (mặc định 0.50 = 50%)
    --down 0.02     Giảm tối đa (mặc định 0.02 = 2%)
    --ratio 5       Giới hạn max/min (mặc định 5)
    --no-clear      KHÔNG xoá giá sản phẩm ngoài batch (chế độ full-list)
    --report r.xlsx Xuất báo cáo điều chỉnh
"""
import argparse
import json
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent))

from core import (
    run_sync, load_price_map_from_path, write_output, write_report,
    DEFAULT_MAX_UP_PCT, DEFAULT_MAX_DOWN_PCT, RATIO_LIMIT,
)


def main():
    ap = argparse.ArgumentParser(description="Đồng bộ giá Shopee")
    ap.add_argument("--xlsx", required=True, help="File Excel Shopee gốc")
    ap.add_argument("--csv", required=True, help="Bảng giá mới (CSV)")
    ap.add_argument("--out", default="import_shopee.xlsx", help="File Excel xuất ra")
    ap.add_argument("--report", default=None, help="File báo cáo (xlsx)")
    ap.add_argument("--up", type=float, default=DEFAULT_MAX_UP_PCT)
    ap.add_argument("--down", type=float, default=DEFAULT_MAX_DOWN_PCT)
    ap.add_argument("--ratio", type=float, default=RATIO_LIMIT)
    ap.add_argument("--no-clear", action="store_true",
                    help="Không xoá giá sản phẩm ngoài batch")
    ap.add_argument("--keep-lock", action="store_true",
                    help="Giữ nguyên sheet protection")
    args = ap.parse_args()

    pm = load_price_map_from_path(args.csv)
    print(f"[i] Đọc {len(pm)} dòng giá từ CSV")

    result, _ = run_sync(
        args.xlsx, pm,
        max_up_pct=args.up, max_down_pct=args.down, ratio_limit=args.ratio,
        clear_unmatched=not args.no_clear,
    )

    write_output(args.xlsx, result, args.out,
                 unlock_protection=not args.keep_lock)
    print(f"[✓] Đã ghi file import: {args.out}")

    if args.report:
        write_report(result, args.report)
        print(f"[✓] Đã ghi báo cáo: {args.report}")

    s = result.stats
    print(json.dumps(s, ensure_ascii=False, indent=2))
    if result.unresolved:
        print(f"[!] {len(result.unresolved)} nhóm KHÔNG xử lý được — xem báo cáo.")


if __name__ == "__main__":
    main()
