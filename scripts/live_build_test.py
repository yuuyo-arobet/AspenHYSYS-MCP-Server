"""構築系 write ツールの実機検証 (空ケース推奨)。

Windows venv python で実行:
    ./venv/Scripts/python.exe scripts/live_build_test.py

各ステップは独立して try し、結果を出力。最後に作成物を必ず後始末する。
ドライラン (confirm=False) → 実行 (confirm=True) の順で各操作を確認する。
テスト用の名前は接頭辞 MCPT_ を付け、衝突を避ける。
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from hysys_mcp.hysys_client import HysysClient

S = "MCPT_S1"          # テストストリーム
OP = "MCPT_E100"       # テスト装置 (Cooler)


def step(label, fn):
    print(f"\n[{label}]")
    try:
        out = fn()
        print("    OK ->", out)
        return out
    except Exception as e:  # noqa: BLE001
        print(f"    ERROR {type(e).__name__}: {e}")
        return None


def main() -> None:
    c = HysysClient()
    c.connect()
    print("connected. active case =",
          [x for x in c.list_cases() if x.get("is_active")])

    # ── 1. create_stream: dry-run → confirm ──────────────
    step("1a create_stream DRY", lambda: c.create_stream(S, "material", confirm=False))
    step("1b create_stream CONFIRM", lambda: c.create_stream(S, "material", confirm=True))
    step("1c verify list_streams", lambda: c.list_streams())

    # ── 2. create_unit_op: Cooler ────────────────────────
    step("2a create_unit_op DRY", lambda: c.create_unit_op("coolerop", OP, confirm=False))
    step("2b create_unit_op CONFIRM", lambda: c.create_unit_op("coolerop", OP, confirm=True))
    step("2c verify list_unit_ops",
         lambda: [getattr(o, "name", o) for o in c.list_unit_ops()])

    # ── 3. list_ports (read) ─────────────────────────────
    step("3 list_ports", lambda: c.list_ports(OP))

    # ── 4. connect_stream: S を Cooler feed へ ───────────
    step("4a connect_stream DRY", lambda: c.connect_stream(OP, S, port="feed", confirm=False))
    step("4b connect_stream CONFIRM", lambda: c.connect_stream(OP, S, port="feed", confirm=True))
    step("4c re-list_ports", lambda: c.list_ports(OP))

    # ── 5. disconnect_stream ─────────────────────────────
    step("5 disconnect_stream CONFIRM",
         lambda: c.disconnect_stream(OP, S, port="feed", confirm=True))

    # ── 6. delete_object (後始末も兼ねる) ────────────────
    step("6a delete unit_op CONFIRM",
         lambda: c.delete_object(OP, kind="unit_op", confirm=True))
    step("6b delete stream CONFIRM",
         lambda: c.delete_object(S, kind="stream", confirm=True))

    # ── 最終状態 ─────────────────────────────────────────
    print("\n[final] streams =", c.list_streams())
    print("[final] unit_ops =",
          [getattr(o, "name", o) for o in c.list_unit_ops()])
    print("\nDONE. (保存はしていない。HYSYS で Save しなければモデルは残らない)")


if __name__ == "__main__":
    main()
