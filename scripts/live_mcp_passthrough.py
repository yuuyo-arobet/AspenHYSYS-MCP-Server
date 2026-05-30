"""MCP 通し実機検証 (server.py → registry → handler → HYSYS)。

server.py はモードを **import 時に1度だけ** _MODE へ確定する (list_tools と
call_tool が必ず同じモードを見るための設計)。よって各モードは別プロセスで
検証する。実運用も .mcp.json の env でプロセス起動前にモードが決まるため、
これが実態に即した検証になる。

使い方:
    親 (検証ランナー):
        ./venv/Scripts/python.exe scripts/live_mcp_passthrough.py
      → 自分自身を HYSYS_MCP_MODE=default / =enhanced の2プロセスで呼び直し、
        結果を集約して scripts/live_mcp_passthrough_result.txt に書く。
    子 (1モード分・内部用):
        HYSYS_MCP_MODE=<mode> ... live_mcp_passthrough.py --child
"""

import asyncio
import json
import os
import subprocess
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent / "src"))


def _text(result):
    try:
        return "\n".join(getattr(x, "text", str(x)) for x in result)
    except Exception:
        return str(result)


async def _child() -> dict:
    """1モード分の検証。結果 dict を返す (import 時に _MODE が env から確定)。"""
    from hysys_mcp import server

    out: dict = {"mode": server._MODE}
    tools = await server.list_tools()
    names = {t.name for t in tools}
    out["tool_count"] = len(tools)
    out["create_stream_visible"] = "hysys_create_stream" in names

    # write 系を call (default では拒否, enhanced では実行)
    res = await server.call_tool("hysys_create_stream",
                                 {"name": "MCPP_PROBE", "confirm": False})
    out["create_stream_call_head"] = _text(res)[:80]

    if server._MODE == "enhanced":
        async def call(name, args):
            return _text(await server.call_tool(name, args))
        seq = {}
        seq["create_stream"] = (await call("hysys_create_stream",
            {"name": "MCPP_S", "stream_kind": "material", "confirm": True}))[:100]
        seq["create_unit_op"] = (await call("hysys_create_unit_op",
            {"type_name": "coolerop", "name": "MCPP_E", "confirm": True}))[:100]
        seq["connect"] = (await call("hysys_connect_stream",
            {"op_name": "MCPP_E", "stream_name": "MCPP_S", "port": "feed",
             "confirm": True}))[:130]
        seq["delete_op"] = (await call("hysys_delete_object",
            {"name": "MCPP_E", "kind": "unit_op", "confirm": True}))[:100]
        seq["delete_stream"] = (await call("hysys_delete_object",
            {"name": "MCPP_S", "kind": "stream", "confirm": True}))[:100]
        out["enhanced_full_path"] = seq
    return out


def _run_child_mode(mode: str) -> dict:
    env = dict(os.environ, HYSYS_MCP_MODE=mode)
    proc = subprocess.run(
        [sys.executable, str(Path(__file__)), "--child"],
        env=env, capture_output=True, text=True)
    for line in reversed(proc.stdout.strip().splitlines()):
        line = line.strip()
        if line.startswith("{"):
            try:
                return json.loads(line)
            except Exception:
                break
    return {"mode": mode, "error": "parse failed",
            "stdout_tail": proc.stdout[-400:], "stderr_tail": proc.stderr[-400:]}


def main():
    if "--child" in sys.argv:
        print(json.dumps(asyncio.run(_child()), ensure_ascii=False))
        return

    log = ["=== MCP 通し検証 (モード別プロセス) ==="]
    for mode in ("default", "enhanced"):
        r = _run_child_mode(mode)
        log.append(f"\n--- mode={mode} ---")
        log.append(json.dumps(r, ensure_ascii=False, indent=2))
    (Path(__file__).parent / "live_mcp_passthrough_result.txt").write_text(
        "\n".join(log), encoding="utf-8")


if __name__ == "__main__":
    main()
