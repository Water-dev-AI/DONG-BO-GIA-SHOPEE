"""Test bộ giải ràng buộc và pipeline."""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))

from core.engine import (
    solve_group, round_thousand, parse_price, is_circle,
    run_sync, load_price_map_from_path,
)

EX = Path(__file__).parent.parent / "examples"


def test_round_thousand():
    assert round_thousand(10000) == 10000
    assert round_thousand(10500) == 10000  # nearest
    assert round_thousand(10600) == 11000
    assert round_thousand(10100, "up") == 11000
    assert round_thousand(10900, "down") == 10000


def test_parse_price():
    assert parse_price("69000") == 69000
    assert parse_price("69,000") == 69000
    assert parse_price("69.000₫") == 69000
    assert parse_price("") is None
    assert parse_price(None) is None


def test_is_circle():
    assert is_circle("⭕")
    assert not is_circle("WM-e-SM-1GB")
    assert not is_circle(None)


def test_solve_no_violation():
    finals, ok, _ = solve_group([100000, 200000, 400000], 0.5, 0.02)
    assert ok
    assert finals == [100000, 200000, 400000]


def test_solve_lift_min():
    finals, ok, _ = solve_group([100000, 200000, 600000], 0.5, 0.02)
    assert ok
    assert max(finals) <= min(finals) * 5 + 1
    assert min(finals) > 100000  # min đã được nâng


def test_solve_unsolvable():
    finals, ok, note = solve_group([100000, 2000000], 0.5, 0.02)
    assert not ok
    assert "không thể" in note.lower()


def test_solve_rounds_to_thousand():
    finals, ok, _ = solve_group([100000, 200000, 560000], 0.5, 0.02)
    assert all(f % 1000 == 0 for f in finals)


def test_full_pipeline():
    pm = load_price_map_from_path(EX / "bang_gia_mau.csv")
    assert len(pm) > 100
    result, _ = run_sync(EX / "shopee_mau.xlsx", pm)
    s = result.stats
    assert s["total_rows"] > 0
    assert s["matched"] > 0
    # mọi giá cuối đều tròn nghìn
    for rec in result.rows:
        if rec.final_price is not None:
            assert rec.final_price % 1000 == 0
    # ⭕ = max nhóm
    from collections import defaultdict
    g = defaultdict(list)
    for rec in result.rows:
        g[rec.product].append(rec)
    for pid, recs in g.items():
        circ = [r for r in recs if r.is_circle and r.final_price]
        priced = [r.final_price for r in recs if not r.is_circle and r.final_price]
        if circ and priced:
            assert circ[0].final_price == max(priced)


def test_partial_clears_out_of_batch():
    pm = load_price_map_from_path(EX / "bang_gia_mau.csv")
    result, _ = run_sync(EX / "shopee_mau.xlsx", pm, clear_unmatched=True)
    cleared = [r for r in result.rows if r.cleared]
    assert len(cleared) > 0


if __name__ == "__main__":
    import subprocess
    sys.exit(subprocess.call(["python3", "-m", "pytest", __file__, "-v"]))
