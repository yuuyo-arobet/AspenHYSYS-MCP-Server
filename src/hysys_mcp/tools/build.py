"""フローシート構築ツール (create / connect / disconnect / delete / ports)。

AspenPlus-MCP の enhanced(構築)モード相当の穴埋め。HYSYS COM の構築系は
バージョン差が大きいため、client 側で confirm ドライラン + 複数シグネチャ
防御 + Solver Hold を実装している。ここはその薄い登録層。

tag:
    list_ports = read (接続前のポート探索)
    create/connect/disconnect/delete = write (enhanced モードでのみ公開)
"""

from __future__ import annotations

from hysys_mcp.registry import READ, WRITE, register


def _create_stream(client, a):
    return client.create_stream(
        name=a["name"],
        stream_kind=a.get("stream_kind", "material"),
        confirm=bool(a.get("confirm", False)),
        flowsheet_path=a.get("flowsheet_path", "Main"),
    )


def _create_unit_op(client, a):
    return client.create_unit_op(
        type_name=a["type_name"],
        name=a.get("name"),
        confirm=bool(a.get("confirm", False)),
        flowsheet_path=a.get("flowsheet_path", "Main"),
    )


def _connect_stream(client, a):
    return client.connect_stream(
        op_name=a["op_name"],
        stream_name=a["stream_name"],
        port=a.get("port", "feed"),
        confirm=bool(a.get("confirm", False)),
        flowsheet_path=a.get("flowsheet_path", "Main"),
    )


def _disconnect_stream(client, a):
    return client.disconnect_stream(
        op_name=a["op_name"],
        stream_name=a["stream_name"],
        port=a.get("port", "feed"),
        confirm=bool(a.get("confirm", False)),
        flowsheet_path=a.get("flowsheet_path", "Main"),
    )


def _delete_object(client, a):
    return client.delete_object(
        name=a["name"],
        kind=a["kind"],
        confirm=bool(a.get("confirm", False)),
        flowsheet_path=a.get("flowsheet_path", "Main"),
    )


def _list_ports(client, a):
    return client.list_ports(
        op_name=a["op_name"], flowsheet_path=a.get("flowsheet_path", "Main")
    )


def register_all():
    register(
        "hysys_create_stream",
        "マテリアル / エネルギーストリームを新規作成。"
        "stream_kind='material'(既定) または 'energy'。"
        "デフォルトはドライラン (confirm=true で実作成)。"
        "作成は Solver Hold 中に行われる。",
        {
            "type": "object",
            "properties": {
                "name": {"type": "string", "description": "新規ストリーム名"},
                "stream_kind": {
                    "type": "string",
                    "enum": ["material", "energy"],
                    "default": "material",
                },
                "confirm": {"type": "boolean", "default": False},
                "flowsheet_path": {"type": "string", "default": "Main"},
            },
            "required": ["name"],
        },
        _create_stream,
        WRITE,
    )
    register(
        "hysys_create_unit_op",
        "装置 (unit operation) を新規作成。"
        "type_name は内部 TypeName ('coolerop' 等) か GUI 名 ('Cooler' 等) のどちらでも可 "
        "(alias で吸収、未知ならそのまま使用)。通らない場合は hysys_find_ops で既存装置の "
        "type_name を確認すること。引数順の環境差は client 側で吸収。"
        "デフォルトはドライラン。Solver Hold 中に作成。",
        {
            "type": "object",
            "properties": {
                "type_name": {
                    "type": "string",
                    "description": "内部 TypeName ('coolerop','heaterop','mixerop','teeop',"
                                   "'separatorop','pumpop','compressop','valveop','distillation' 等) "
                                   "または GUI 名",
                },
                "name": {"type": "string", "description": "装置名 (省略時は自動命名)"},
                "confirm": {"type": "boolean", "default": False},
                "flowsheet_path": {"type": "string", "default": "Main"},
            },
            "required": ["type_name"],
        },
        _create_unit_op,
        WRITE,
    )
    register(
        "hysys_connect_stream",
        "ストリームを装置の Feed / Product / Energy ポートに接続。"
        "port='feed'(既定) | 'product' | 'energy'。"
        "ストリームは material→energy の順で自動解決。"
        "デフォルトはドライラン。Solver Hold 中に接続 (op.<Port>Stream へ代入)。"
        "※当 COM ビルドでは feed/product/energy 以外の任意名ポートは未対応。",
        {
            "type": "object",
            "properties": {
                "op_name": {"type": "string", "description": "接続先の装置名"},
                "stream_name": {"type": "string"},
                "port": {
                    "type": "string",
                    "enum": ["feed", "product", "energy"],
                    "description": "feed | product | energy",
                    "default": "feed",
                },
                "confirm": {"type": "boolean", "default": False},
                "flowsheet_path": {"type": "string", "default": "Main"},
            },
            "required": ["op_name", "stream_name"],
        },
        _connect_stream,
        WRITE,
    )
    register(
        "hysys_disconnect_stream",
        "装置からストリームを切断。※当 HYSYS COM ビルドには接続点を空にする API が "
        "無いため、このツールは何も変更せず status:'unsupported' と代替手段を返す。"
        "繋ぎ替えは hysys_connect_stream、除去は hysys_delete_object (接続中でも可)、"
        "完全な切断は GUI を使うこと。",
        {
            "type": "object",
            "properties": {
                "op_name": {"type": "string"},
                "stream_name": {"type": "string"},
                "port": {
                    "type": "string",
                    "enum": ["feed", "product", "energy"],
                    "default": "feed",
                },
                "confirm": {"type": "boolean", "default": False},
                "flowsheet_path": {"type": "string", "default": "Main"},
            },
            "required": ["op_name", "stream_name"],
        },
        _disconnect_stream,
        WRITE,
    )
    register(
        "hysys_delete_object",
        "ストリーム / 装置を削除 (obj.Delete())。"
        "kind='stream'(material/energy 自動判定) | 'material_stream' | 'energy_stream' | 'unit_op'。"
        "接続中のオブジェクトは先に hysys_disconnect_stream で切断推奨。"
        "デフォルトはドライラン。Solver Hold 中に削除。",
        {
            "type": "object",
            "properties": {
                "name": {"type": "string", "description": "削除対象の名前"},
                "kind": {
                    "type": "string",
                    "enum": ["stream", "material_stream", "energy_stream", "unit_op"],
                },
                "confirm": {"type": "boolean", "default": False},
                "flowsheet_path": {"type": "string", "default": "Main"},
            },
            "required": ["name", "kind"],
        },
        _delete_object,
        WRITE,
    )
    register(
        "hysys_list_ports",
        "装置のポート (Ports / MaterialPorts / EnergyPorts / Feeds / Products) を列挙。"
        "hysys_connect_stream で名前付きポートに繋ぐ前に、正しいポート名を調べる読取専用ツール。",
        {
            "type": "object",
            "properties": {
                "op_name": {"type": "string"},
                "flowsheet_path": {"type": "string", "default": "Main"},
            },
            "required": ["op_name"],
        },
        _list_ports,
        READ,
    )
