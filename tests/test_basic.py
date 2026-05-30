"""基本的なインポートテスト。

HYSYS COM 接続のテストは別途、実際に HYSYS が起動している状態で手動実行する。
mcp / pywin32 が無い環境 (WSL 等) では該当テストを skip する。
"""

import sys
import pytest


def test_import():
    """パッケージが正常に import できるか。"""
    import hysys_mcp
    assert hysys_mcp.__version__ == "0.1.0"


def test_client_import():
    """HysysClient が import できるか。"""
    from hysys_mcp.hysys_client import HysysClient, StreamInfo, UnitOpInfo, SolverStatus
    assert HysysClient is not None


def test_server_import():
    """MCP server module が import できるか (mcp 未インストール環境では skip)。

    server.py は mcp が無いと sys.exit(1) する設計なので、
    mcp が import できない環境ではこのテストを skip する (環境制約)。
    """
    pytest.importorskip("mcp")
    from hysys_mcp import server
    assert server.app.name == "hysys-mcp"


@pytest.mark.skipif(sys.platform != "win32", reason="HYSYS COM は Windows 専用")
def test_connect_to_hysys():
    """実際に HYSYS に接続できるか(HYSYS 起動済みで実行)。"""
    from hysys_mcp.hysys_client import HysysClient

    client = HysysClient()
    client.connect()
    # ケースが開かれていない場合もあるので is_connected ではなく app だけ確認
    assert client._app is not None


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
