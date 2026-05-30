"""Logical Operation (Adjust/Set/Recycle 等) 関連ツール。"""

from __future__ import annotations

from hysys_mcp.registry import READ, WRITE, register


def _list_logical_ops(client, a):
    return client.list_logical_ops(flowsheet_path=a.get("flowsheet_path", "Main"))


def _get_logical_op(client, a):
    return client.get_logical_op(
        name=a["name"], flowsheet_path=a.get("flowsheet_path", "Main")
    )


def _set_adjust_target(client, a):
    return client.set_adjust_target(
        name=a["name"],
        target_value=a.get("target_value"),
        tolerance=a.get("tolerance"),
        step_size=a.get("step_size"),
        max_value=a.get("max_value"),
        min_value=a.get("min_value"),
        confirm=bool(a.get("confirm", False)),
        flowsheet_path=a.get("flowsheet_path", "Main"),
    )


def _reset_recycle(client, a):
    return client.reset_recycle(
        name=a["name"],
        confirm=bool(a.get("confirm", False)),
        flowsheet_path=a.get("flowsheet_path", "Main"),
    )


def register_all():
    register(
        "hysys_list_logical_ops",
        "Adjust/Set/Recycle/Spreadsheet 等の Logical Operation を全列挙。各 op の type と type-specific 詳細 (Adjust の TargetValue, Recycle の IsConverged 等) を含む。",
        {
            "type": "object",
            "properties": {
                "flowsheet_path": {"type": "string", "default": "Main"},
            },
        },
        _list_logical_ops,
        READ,
    )
    register(
        "hysys_get_logical_op",
        "単一 Logical Op の詳細を返す。",
        {
            "type": "object",
            "properties": {
                "name": {"type": "string"},
                "flowsheet_path": {"type": "string", "default": "Main"},
            },
            "required": ["name"],
        },
        _get_logical_op,
        READ,
    )
    register(
        "hysys_set_adjust_target",
        "Adjust ブロックの TargetValue / Tolerance / StepSize / Min / Max を設定。"
        "プロパティ名は HYSYS V14 で固定的に確認されていないため、"
        "hasattr で動的判定する (一部 prop は applied=False で返る可能性あり)。"
        "confirm=true 必須。",
        {
            "type": "object",
            "properties": {
                "name": {"type": "string"},
                "target_value": {"type": "number"},
                "tolerance": {"type": "number"},
                "step_size": {"type": "number"},
                "max_value": {"type": "number"},
                "min_value": {"type": "number"},
                "confirm": {"type": "boolean", "default": False},
                "flowsheet_path": {"type": "string", "default": "Main"},
            },
            "required": ["name"],
        },
        _set_adjust_target,
        WRITE,
    )
    register(
        "hysys_reset_recycle",
        "Recycle ブロックの Reset() を呼ぶ。リサイクル収束が崩れている時の復帰用。confirm=true 必須。",
        {
            "type": "object",
            "properties": {
                "name": {"type": "string"},
                "confirm": {"type": "boolean", "default": False},
                "flowsheet_path": {"type": "string", "default": "Main"},
            },
            "required": ["name"],
        },
        _reset_recycle,
        WRITE,
    )
