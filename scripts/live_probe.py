"""実機接続プローブ (非破壊・読取のみ)。

Windows venv python で実行:
    ./venv/Scripts/python.exe scripts/live_probe.py

HYSYS が起動していなければ接続に失敗する。何も壊さない。
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent / "src"))

from hysys_mcp.hysys_client import HysysClient


def main() -> None:
    client = HysysClient()
    print("[1] connect() ...")
    client.connect()
    print("    app =", client._app is not None)

    print("[2] list_cases() ...")
    try:
        cases = client.list_cases()
        print("    cases =", cases)
    except Exception as e:  # noqa: BLE001
        print("    list_cases ERROR:", type(e).__name__, e)

    print("[3] get_solver_status() ...")
    try:
        st = client.get_solver_status()
        print("    status =", st)
    except Exception as e:  # noqa: BLE001
        print("    get_status ERROR:", type(e).__name__, e)

    print("[4] list_streams() (先頭10件) ...")
    try:
        streams = client.list_streams()
        names = [getattr(s, "name", s) for s in (streams or [])]
        print("    count =", len(names))
        print("    head  =", names[:10])
    except Exception as e:  # noqa: BLE001
        print("    list_streams ERROR:", type(e).__name__, e)

    print("[5] list_unit_ops() (先頭10件) ...")
    try:
        ops = client.list_unit_ops()
        names = [getattr(o, "name", o) for o in (ops or [])]
        print("    count =", len(names))
        print("    head  =", names[:10])
    except Exception as e:  # noqa: BLE001
        print("    list_unit_ops ERROR:", type(e).__name__, e)

    print("DONE (no writes performed)")


if __name__ == "__main__":
    main()
