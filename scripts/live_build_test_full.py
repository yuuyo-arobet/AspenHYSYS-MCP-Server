"""構築系 write の網羅実機検証 (HYSYS V14・成分+FP 定義済みケース)。

Windows venv python で実行:
    ./venv/Scripts/python.exe scripts/live_build_test_full.py

基本 (live_build_test.py) でカバー済みの material+feed+cooler に加え、
未検証だった以下を網羅する:
  1. energy ストリームの作成
  2. 複数の装置型 (mixer / heater / separator / valve)
  3. product ポートへの接続
  4. energy ポートへの接続 (Cooler の Duty)
各ブロックは独立 try。最後に全テスト物を必ず削除して後始末する。
結果は UTF-8 で scripts/live_build_test_full_result.txt に書き出す
(Windows stdout は cp932 なので文字化け回避)。
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from hysys_mcp.hysys_client import HysysClient

LOG: list[str] = []
def log(m): LOG.append(m)


def step(label, fn):
    try:
        out = fn()
        log(f"[OK]  {label} -> {out}")
        return out
    except Exception as e:  # noqa: BLE001
        log(f"[ERR] {label} -> {type(e).__name__}: {e}")
        return None


# テスト対象名 (接頭辞 FT_ で衝突回避)
MAT = "FT_MAT"        # material stream
ENE = "FT_ENE"        # energy stream
PROD = "FT_PROD"      # product 用 material stream


def main():
    c = HysysClient()
    fs_names = []
    c.connect()
    log("=== 網羅実機検証 (HYSYS V14) ===")
    log(f"active case = {[x['name'] for x in c.list_cases() if x.get('is_active')]}")

    # 後始末を先に (前回残骸対策)
    def cleanup():
        for nm in ("FT_E_MIX", "FT_E_HEAT", "FT_E_SEP", "FT_E_VALVE", "FT_E_COOL"):
            try: c.delete_object(nm, kind="unit_op", confirm=True)
            except Exception: pass
        for nm in (MAT, ENE, PROD):
            try: c.delete_object(nm, kind="stream", confirm=True)
            except Exception: pass
    cleanup()

    # ── 1. energy ストリーム作成 ──
    log("\n-- 1. energy stream --")
    step("create_stream energy DRY", lambda: c.create_stream(ENE, "energy", confirm=False))
    step("create_stream energy CONFIRM", lambda: c.create_stream(ENE, "energy", confirm=True))
    step("verify list_streams", lambda: c.list_streams())

    # material も用意 (接続テスト用)
    step("create_stream material(MAT)", lambda: c.create_stream(MAT, "material", confirm=True))
    step("create_stream material(PROD)", lambda: c.create_stream(PROD, "material", confirm=True))

    # ── 2. 複数の装置型 ──
    log("\n-- 2. unit op types --")
    for typ, nm in [("mixerop", "FT_E_MIX"), ("heaterop", "FT_E_HEAT"),
                    ("separator", "FT_E_SEP"), ("valveop", "FT_E_VALVE"),
                    ("coolerop", "FT_E_COOL")]:
        step(f"create_unit_op {typ}", lambda typ=typ, nm=nm: c.create_unit_op(typ, nm, confirm=True))
    step("verify list_unit_ops", lambda: [getattr(o, "name", o) for o in c.list_unit_ops()])

    # ── 3. feed / product 接続 (Cooler) ──
    log("\n-- 3. connect feed/product (Cooler) --")
    step("list_ports Cooler", lambda: c.list_ports("FT_E_COOL"))
    step("connect MAT -> Cooler feed",
         lambda: c.connect_stream("FT_E_COOL", MAT, port="feed", confirm=True))
    step("connect PROD -> Cooler product",
         lambda: c.connect_stream("FT_E_COOL", PROD, port="product", confirm=True))

    # ── 4. energy ポート接続 (Cooler Duty) ──
    log("\n-- 4. connect energy port (Cooler) --")
    step("connect ENE -> Cooler energy",
         lambda: c.connect_stream("FT_E_COOL", ENE, port="energy", confirm=True))

    # ── 後始末 ──
    log("\n-- cleanup --")
    cleanup()
    log(f"final streams = {c.list_streams()}")
    log(f"final unit_ops = {[getattr(o,'name',o) for o in c.list_unit_ops()]}")
    log("DONE")

    (Path(__file__).parent / "live_build_test_full_result.txt").write_text(
        "\n".join(LOG), encoding="utf-8")


if __name__ == "__main__":
    main()
