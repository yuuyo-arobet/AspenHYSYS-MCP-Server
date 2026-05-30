"""装置 (unit operation) 関連ツール。"""

from __future__ import annotations

from hysys_mcp.registry import READ, WRITE, register


def _list_unit_ops(client, a):
    return client.list_unit_ops(a.get("flowsheet_path", "Main"))


def _set_unit_op_param(client, a):
    return client.set_unit_op_param(
        op_name=a["op_name"],
        parameter=a["parameter"],
        value=float(a["value"]),
        unit=a.get("unit", ""),
        confirm=bool(a.get("confirm", False)),
        flowsheet_path=a.get("flowsheet_path", "Main"),
    )


def _get_heat_op(client, a):
    return client.get_heat_op(
        name=a["name"], flowsheet_path=a.get("flowsheet_path", "Main")
    )


def _find_ops(client, a):
    return client.find_ops(
        type_name=a.get("type_name"),
        name_pattern=a.get("name_pattern"),
        flowsheet_path=a.get("flowsheet_path", "Main"),
    )


def register_all():
    register(
        "hysys_list_unit_ops",
        "フローシート内の全装置(operation)を返す。"
        "装置名・種類(Heater/Mixer/Column等)・収束状態を含む。",
        {
            "type": "object",
            "properties": {
                "flowsheet_path": {
                    "type": "string",
                    "description": "フローシートパス",
                    "default": "Main",
                },
            },
        },
        _list_unit_ops,
        READ,
    )
    register(
        "hysys_set_unit_op_param",
        "装置(unit operation)のパラメータを設定。"
        "parameter は HYSYS COM プロパティ名 (例: 'Duty', 'DeltaP', 'PressureDrop')。"
        "デフォルトはドライラン。",
        {
            "type": "object",
            "properties": {
                "op_name": {"type": "string", "description": "装置名"},
                "parameter": {"type": "string", "description": "HYSYS COM プロパティ名"},
                "value": {"type": "number", "description": "設定値"},
                "unit": {
                    "type": "string",
                    "description": "RealVariable の場合の単位 (例: 'kJ/h', 'kPa')。プリミティブ属性なら空文字で OK。",
                    "default": "",
                },
                "confirm": {"type": "boolean", "default": False},
                "flowsheet_path": {"type": "string", "default": "Main"},
            },
            "required": ["op_name", "parameter", "value"],
        },
        _set_unit_op_param,
        WRITE,
    )
    register(
        "hysys_get_heat_op",
        "Cooler/Heater/HeatExchanger 共通の主要パラメータ (Duty, Feed/Product 流, UA, LMTD など) を一括取得。HEX 固有値は HEX のときのみ埋まる。",
        {
            "type": "object",
            "properties": {
                "name": {"type": "string"},
                "flowsheet_path": {"type": "string", "default": "Main"},
            },
            "required": ["name"],
        },
        _get_heat_op,
        READ,
    )
    register(
        "hysys_find_ops",
        "装置を type_name (例: 'coolerop') / name_pattern (regex) で検索。",
        {
            "type": "object",
            "properties": {
                "type_name": {"type": "string"},
                "name_pattern": {"type": "string"},
                "flowsheet_path": {"type": "string", "default": "Main"},
            },
        },
        _find_ops,
        READ,
    )
