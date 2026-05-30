"""ソルバー・収支・感度解析ツール。"""

from __future__ import annotations

from hysys_mcp.registry import READ, WRITE, register


def _get_status(client, a):
    return client.get_solver_status()


def _run(client, a):
    return client.run()


def _reset(client, a):
    return client.reset(confirm=bool(a.get("confirm", False)))


def _set_solver_can_solve(client, a):
    return client.set_solver_can_solve(
        value=bool(a["value"]), confirm=bool(a.get("confirm", False))
    )


def _balance_check(client, a):
    return client.balance_check(
        op_names=a.get("op_names"), flowsheet_path=a.get("flowsheet_path", "Main")
    )


def _case_study(client, a):
    return client.case_study(
        target_kind=a["target_kind"],
        target_name=a["target_name"],
        target_field=a["target_field"],
        sweep_values=[float(v) for v in a["sweep_values"]],
        observe=list(a["observe"]),
        target_unit=a.get("target_unit"),
        confirm=bool(a.get("confirm", False)),
        flowsheet_path=a.get("flowsheet_path", "Main"),
    )


def register_all():
    register(
        "hysys_get_status",
        "ソルバー全体の状態(計算可能か、進行中か、エラー有無)を返す。",
        {"type": "object", "properties": {}},
        _get_status,
        READ,
    )
    register(
        "hysys_run",
        "ソルバーを実行。CanSolve=False の場合エラー。実行後の SolverStatus を返す。",
        {"type": "object", "properties": {}},
        _run,
        WRITE,
    )
    register(
        "hysys_reset",
        "ソルバーをリセットして全変数を未計算状態に戻す。"
        "破壊的操作のため confirm=true 必須。",
        {
            "type": "object",
            "properties": {
                "confirm": {"type": "boolean", "default": False},
            },
        },
        _reset,
        WRITE,
    )
    register(
        "hysys_set_solver_can_solve",
        "フローシート全体ソルバーの CanSolve (Active/Hold) を切替。"
        "value=false は Stop Run 相当 (停滞中の塔ソルバー停止)。"
        "value=true で再開。confirm=true 必須。",
        {
            "type": "object",
            "properties": {
                "value": {"type": "boolean", "description": "true=Active, false=Hold"},
                "confirm": {"type": "boolean", "default": False},
            },
            "required": ["value"],
        },
        _set_solver_can_solve,
        WRITE,
    )
    register(
        "hysys_balance_check",
        "物質収支・熱収支を計算。op_names で範囲指定可 (省略時はフローシート全体)。"
        "境界 (集合内 op の Feed/Product から内部接続を除いたもの) の質量流入・流出と熱流入・流出を集計して closure を計算。",
        {
            "type": "object",
            "properties": {
                "op_names": {
                    "type": "array", "items": {"type": "string"},
                    "description": "対象 ops。省略時は全装置。",
                },
                "flowsheet_path": {"type": "string", "default": "Main"},
            },
        },
        _balance_check,
        READ,
    )
    register(
        "hysys_case_study",
        "感度解析: target を sweep_values で振り、各ケースで observe を観測する。"
        "confirm=false (デフォルト) はドライラン (sweep 値だけ返す)、"
        "confirm=true で実際に書き込み→ソルバー実行→観測→最後に元の値に復元。",
        {
            "type": "object",
            "properties": {
                "target_kind": {
                    "type": "string",
                    "enum": ["stream", "unit_op"],
                    "description": "対象が stream か unit_op か",
                },
                "target_name": {"type": "string"},
                "target_field": {
                    "type": "string",
                    "description": (
                        "stream の場合: 'temperature_C'|'pressure_kPa'|'molar_flow_kgmole_h'"
                        "|'mass_flow_kg_h'|'vapor_fraction'。"
                        "unit_op の場合: HYSYS COM プロパティ名 (例: 'Duty', 'PressureDrop')"
                    ),
                },
                "target_unit": {
                    "type": "string",
                    "description": "unit_op のとき必須 (例: 'kJ/h', 'kPa')。stream のときは無視。",
                },
                "sweep_values": {
                    "type": "array",
                    "items": {"type": "number"},
                    "description": "振る値のリスト",
                },
                "observe": {
                    "type": "array",
                    "description": (
                        "各観測点。[{kind,name,field,unit?}]。"
                        "kind: 'stream'|'unit_op'。"
                        "stream の場合 field は上記の同集合。"
                        "unit_op の場合 field=HYSYS COM プロパティ名, unit=単位文字列を指定。"
                    ),
                    "items": {
                        "type": "object",
                        "properties": {
                            "kind": {"type": "string", "enum": ["stream", "unit_op"]},
                            "name": {"type": "string"},
                            "field": {"type": "string"},
                            "unit": {"type": "string"},
                        },
                        "required": ["kind", "name", "field"],
                    },
                },
                "confirm": {"type": "boolean", "default": False},
                "flowsheet_path": {"type": "string", "default": "Main"},
            },
            "required": ["target_kind", "target_name", "target_field", "sweep_values", "observe"],
        },
        _case_study,
        WRITE,
    )
