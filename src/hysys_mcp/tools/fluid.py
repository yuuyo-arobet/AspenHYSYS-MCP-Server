"""Fluid Package / コンポーネント関連ツール。"""

from __future__ import annotations

from hysys_mcp.registry import READ, register


def _get_fluid_package(client, a):
    return client.get_fluid_package()


def _list_components(client, a):
    return client.list_components()


def _get_component(client, a):
    return client.get_component(name=a["name"])


def register_all():
    register(
        "hysys_get_fluid_package",
        "現ケースの Fluid Package 情報を返す (Property Package 名、components、Reaction Package の有無)。",
        {"type": "object", "properties": {}},
        _get_fluid_package,
        READ,
    )
    register(
        "hysys_list_components",
        "Fluid Package のコンポーネント名一覧を返す。",
        {"type": "object", "properties": {}},
        _list_components,
        READ,
    )
    register(
        "hysys_get_component",
        "コンポーネント単体の物性値 (MW, NBP, Tc, Pc, ω, CAS) を返す。",
        {
            "type": "object",
            "properties": {"name": {"type": "string", "description": "コンポーネント名 (例: 'H2O', 'Methane')"}},
            "required": ["name"],
        },
        _get_component,
        READ,
    )
