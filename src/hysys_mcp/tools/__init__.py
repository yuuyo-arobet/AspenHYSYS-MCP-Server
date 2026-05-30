"""ツールモジュール集約。

register_all_tools() を1回呼ぶと、全ドメインの register_all() が走り
registry.REGISTRY に45ツールが登録される。冪等 (重複登録は registry 側で防止)。
"""

from __future__ import annotations

from hysys_mcp.tools import (
    build,
    columns,
    connection,
    fluid,
    generic,
    logical,
    solver,
    streams,
    unit_ops,
)

# 登録順 = list_tools での表示順
_MODULES = (
    connection,
    streams,
    unit_ops,
    columns,
    solver,
    logical,
    fluid,
    build,
    generic,
)

_registered = False


def register_all_tools() -> None:
    global _registered
    if _registered:
        return
    for mod in _MODULES:
        mod.register_all()
    _registered = True
