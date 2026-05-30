"""HYSYS MCP Server (薄い adapter)。

役割は3つだけ:
    1. registry に全ツールを登録させる (hysys_mcp.tools.register_all_tools)
    2. registry.ToolSpec を mcp.types.Tool へ変換して list_tools で公開
       (現在のモードで許可された tag のものだけ)
    3. call_tool で 1 行ディスパッチ (モードゲート → client → handler → JSON 化)

ツールの定義・実装は hysys_mcp/tools/*.py と hysys_mcp/hysys_client.py にある。

起動方法:
    python -m hysys_mcp.server

モード制御 (環境変数 HYSYS_MCP_MODE):
    readonly / default(既定) / enhanced  ← registry.py 参照
"""

from __future__ import annotations

import asyncio
import sys
from typing import Any

# MCP SDK (Python 公式)
try:
    from mcp.server import Server
    from mcp.server.stdio import stdio_server
    from mcp.types import Tool, TextContent
except ImportError:
    print(
        "MCP Python SDK がインストールされていません。"
        "`pip install mcp` を実行してください。",
        file=sys.stderr,
    )
    sys.exit(1)

from hysys_mcp.hysys_client import HysysClient
from hysys_mcp import registry
from hysys_mcp.tools import register_all_tools

# 起動時に全ツールを登録
register_all_tools()

# モードは「サーバ起動時に1度だけ」確定する。
# list_tools と call_tool が必ず同じモードを見るようにし、プロセス稼働中の
# 環境変数書き換えで両者がズレる (list には出ないのに呼べてしまう等の) のを防ぐ。
# 実運用でもモードは .mcp.json の env でプロセス起動前に決まるため、これが正しい。
_MODE = registry.server_mode()


# グローバル client(MCP server のライフサイクル中、1つだけ存在)
_client: HysysClient | None = None


def get_client() -> HysysClient:
    global _client
    if _client is None:
        _client = HysysClient()
        _client.connect()
    return _client


# ─────────────────────────────────────────────────
# Server インスタンスと tool 登録
# ─────────────────────────────────────────────────
app = Server("hysys-mcp")


@app.list_tools()
async def list_tools() -> list[Tool]:
    """現在のモードで公開すべき MCP ツール一覧を返す。"""
    return [
        Tool(
            name=spec.name,
            description=spec.description,
            inputSchema=spec.input_schema,
        )
        for spec in registry.specs_for_mode(_MODE)
    ]


@app.call_tool()
async def call_tool(name: str, arguments: dict[str, Any]) -> list[TextContent]:
    """ツール呼び出しのディスパッチ。"""
    spec = registry.get_spec(name)
    if spec is None:
        return [TextContent(type="text", text=f"Unknown tool: {name}")]

    # モードゲートは get_client() より前 (basic で拒否すべき呼び出しで
    # HYSYS 接続/起動を発生させないため)
    if spec.tag not in registry.allowed_tags(_MODE):
        return [TextContent(
            type="text",
            text=(
                f"Tool '{name}' (tag={spec.tag}) is disabled in mode '{_MODE}'. "
                f"これは安全のための制限です。書込/ソルバ系を使うには "
                f"環境変数 HYSYS_MCP_MODE=enhanced を設定してサーバを再起動してください。"
            ),
        )]

    try:
        client = get_client()
        payload = spec.handler(client, arguments)
        return [TextContent(type="text", text=registry.render_text(payload))]
    except Exception as e:  # noqa: BLE001 - 失敗内容を呼び出し側に返す
        return [TextContent(
            type="text",
            text=f"Error in {name}: {type(e).__name__}: {e}",
        )]


# ─────────────────────────────────────────────────
# Entry point
# ─────────────────────────────────────────────────
async def main_async() -> None:
    async with stdio_server() as (read_stream, write_stream):
        await app.run(read_stream, write_stream, app.create_initialization_options())


def main() -> None:
    asyncio.run(main_async())


if __name__ == "__main__":
    main()
