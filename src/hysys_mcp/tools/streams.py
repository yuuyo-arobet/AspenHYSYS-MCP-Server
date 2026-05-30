"""ストリーム関連ツール。"""

from __future__ import annotations

from hysys_mcp.registry import READ, WRITE, register


def _list_streams(client, a):
    path = a.get("flowsheet_path", "Main")
    return {"flowsheet": path, "streams": client.list_streams(path)}


def _get_stream(client, a):
    return client.get_stream(a["name"], a.get("flowsheet_path", "Main"))


def _set_stream(client, a):
    return client.set_stream(
        name=a["name"],
        temperature_C=a.get("temperature_C"),
        pressure_kPa=a.get("pressure_kPa"),
        molar_flow_kgmole_h=a.get("molar_flow_kgmole_h"),
        composition_molar=a.get("composition_molar"),
        component_molar_flows=a.get("component_molar_flows"),
        confirm=bool(a.get("confirm", False)),
        flowsheet_path=a.get("flowsheet_path", "Main"),
    )


def _find_streams(client, a):
    return client.find_streams(
        contains_components=a.get("contains_components"),
        min_mol_fraction=a.get("min_mol_fraction"),
        vapor_fraction_min=a.get("vapor_fraction_min"),
        vapor_fraction_max=a.get("vapor_fraction_max"),
        molar_flow_min_kgmole_h=a.get("molar_flow_min_kgmole_h"),
        name_pattern=a.get("name_pattern"),
        flowsheet_path=a.get("flowsheet_path", "Main"),
    )


def _get_stream_phys(client, a):
    return client.get_stream_phys(
        name=a["name"], flowsheet_path=a.get("flowsheet_path", "Main")
    )


def _add_internal_stream(client, a):
    return client.add_internal_stream(
        column_name=a["column_name"],
        stage=int(a["stage"]),
        kind=a.get("kind", "Liquid"),
        stream_name=a.get("stream_name"),
        confirm=bool(a.get("confirm", False)),
        flowsheet_path=a.get("flowsheet_path", "Main"),
    )


def register_all():
    register(
        "hysys_list_streams",
        "フローシート内の全マテリアルストリーム名を返す。",
        {
            "type": "object",
            "properties": {
                "flowsheet_path": {
                    "type": "string",
                    "description": "フローシートパス(例: 'Main' or 'Main/CO2 Absorption Tower')。"
                                   "省略するとメインフロー。",
                    "default": "Main",
                },
            },
        },
        _list_streams,
        READ,
    )
    register(
        "hysys_get_stream",
        "1つのストリームの温度・圧力・流量・組成・相状態を返す。",
        {
            "type": "object",
            "properties": {
                "name": {
                    "type": "string",
                    "description": "ストリーム名(例: stream1, Feed)",
                },
                "flowsheet_path": {
                    "type": "string",
                    "description": "フローシートパス",
                    "default": "Main",
                },
            },
            "required": ["name"],
        },
        _get_stream,
        READ,
    )
    register(
        "hysys_set_stream",
        "ストリームの温度・圧力・モル流量・組成を設定。"
        "デフォルトはドライラン(before/after 表示のみ)で confirm=true 指定時のみ実書き込み。"
        "組成は部分更新可で合計1に正規化される。",
        {
            "type": "object",
            "properties": {
                "name": {"type": "string", "description": "ストリーム名"},
                "temperature_C": {"type": "number", "description": "温度 [℃]"},
                "pressure_kPa": {"type": "number", "description": "圧力 [kPa]"},
                "molar_flow_kgmole_h": {"type": "number", "description": "モル流量 [kgmole/h]"},
                "composition_molar": {
                    "type": "object",
                    "description": "成分名 → モル分率の部分マッピング。指定した成分のみ更新し残りは現行値を保持後に正規化。",
                    "additionalProperties": {"type": "number"},
                },
                "component_molar_flows": {
                    "type": "object",
                    "description": "成分名 → モル流量 [kgmole/h] の部分マッピング。指定した成分のみ更新、未指定の成分は現値維持。合計が新しい総流量になる。",
                    "additionalProperties": {"type": "number"},
                },
                "confirm": {
                    "type": "boolean",
                    "description": "true で実書き込み。省略時 false (ドライラン)。",
                    "default": False,
                },
                "flowsheet_path": {"type": "string", "default": "Main"},
            },
            "required": ["name"],
        },
        _set_stream,
        WRITE,
    )
    register(
        "hysys_find_streams",
        "条件に合うストリームを検索。"
        "contains_components + min_mol_fraction で組成フィルタ、"
        "vapor_fraction_min/max, molar_flow_min_kgmole_h, name_pattern (regex) も使える。",
        {
            "type": "object",
            "properties": {
                "contains_components": {
                    "type": "array", "items": {"type": "string"},
                    "description": "これらの成分を含むストリーム",
                },
                "min_mol_fraction": {"type": "number", "description": "上記成分の下限モル分率"},
                "vapor_fraction_min": {"type": "number"},
                "vapor_fraction_max": {"type": "number"},
                "molar_flow_min_kgmole_h": {"type": "number"},
                "name_pattern": {"type": "string", "description": "正規表現 (大文字小文字無視)"},
                "flowsheet_path": {"type": "string", "default": "Main"},
            },
        },
        _find_streams,
        READ,
    )
    register(
        "hysys_get_stream_phys",
        "ストリームの bulk 物性 (密度・粘度・Cp・熱伝導率・MW・エンタルピー・エントロピー等) を一括取得。HYSYS 内部単位ベース。",
        {
            "type": "object",
            "properties": {
                "name": {"type": "string"},
                "flowsheet_path": {"type": "string", "default": "Main"},
            },
            "required": ["name"],
        },
        _get_stream_phys,
        READ,
    )
    register(
        "hysys_add_internal_stream",
        "塔の Tray Section に内部 Liquid/Vapour Draw stream を追加。"
        "段別流量を MaterialStream として外側から取得するため。"
        "confirm=true 必須 (破壊的)。",
        {
            "type": "object",
            "properties": {
                "column_name": {"type": "string"},
                "stage": {"type": "integer"},
                "kind": {"type": "string", "enum": ["Liquid", "Vapor", "Vapour"], "default": "Liquid"},
                "stream_name": {"type": "string", "description": "省略時は自動命名 '<col>_S<stage>_<kind>'"},
                "confirm": {"type": "boolean", "default": False},
                "flowsheet_path": {"type": "string", "default": "Main"},
            },
            "required": ["column_name", "stage"],
        },
        _add_internal_stream,
        WRITE,
    )
