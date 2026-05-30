"""塔 (蒸留塔/吸収塔) 関連ツール。"""

from __future__ import annotations

from hysys_mcp.registry import READ, WRITE, register


def _get_column_profile(client, a):
    return client.get_column_profile(
        column_name=a["column_name"], flowsheet_path=a.get("flowsheet_path", "Main")
    )


def _list_column_specs(client, a):
    return client.list_column_specs(
        column_name=a["column_name"], flowsheet_path=a.get("flowsheet_path", "Main")
    )


def _set_column_spec(client, a):
    return client.set_column_spec(
        column_name=a["column_name"],
        spec_name=a["spec_name"],
        goal=float(a["goal"]),
        unit=a.get("unit", ""),
        confirm=bool(a.get("confirm", False)),
        flowsheet_path=a.get("flowsheet_path", "Main"),
    )


def _set_column_spec_active(client, a):
    return client.set_column_spec_active(
        column_name=a["column_name"],
        spec_name=a["spec_name"],
        active=bool(a["active"]),
        confirm=bool(a.get("confirm", False)),
        flowsheet_path=a.get("flowsheet_path", "Main"),
    )


def _get_column_spec_detail(client, a):
    return client.get_column_spec_detail(
        column_name=a["column_name"],
        spec_name=a["spec_name"],
        flowsheet_path=a.get("flowsheet_path", "Main"),
    )


def _clear_column_estimates(client, a):
    return client.clear_column_estimates(
        column_name=a["column_name"],
        scope=a.get("scope", "all"),
        confirm=bool(a.get("confirm", False)),
        flowsheet_path=a.get("flowsheet_path", "Main"),
    )


def _set_column_solver_param(client, a):
    return client.set_column_solver_param(
        column_name=a["column_name"],
        parameter=a["parameter"],
        value=a["value"],
        confirm=bool(a.get("confirm", False)),
        flowsheet_path=a.get("flowsheet_path", "Main"),
    )


def _remove_column_spec(client, a):
    return client.remove_column_spec(
        column_name=a["column_name"],
        spec_name=a["spec_name"],
        confirm=bool(a.get("confirm", False)),
        flowsheet_path=a.get("flowsheet_path", "Main"),
    )


def _add_column_spec(client, a):
    return client.add_column_spec(
        column_name=a["column_name"],
        spec_type=a["spec_type"],
        spec_name=a["spec_name"],
        goal=a.get("goal"),
        stage=a.get("stage"),
        draw=a.get("draw"),
        phase=a.get("phase"),
        components=a.get("components"),
        is_active=bool(a.get("is_active", True)),
        flow_basis=a.get("flow_basis"),
        confirm=bool(a.get("confirm", False)),
        flowsheet_path=a.get("flowsheet_path", "Main"),
    )


def _column_reset(client, a):
    return client.column_reset(
        column_name=a["column_name"],
        confirm=bool(a.get("confirm", False)),
        flowsheet_path=a.get("flowsheet_path", "Main"),
    )


def _column_run(client, a):
    return client.column_run(
        column_name=a["column_name"],
        flowsheet_path=a.get("flowsheet_path", "Main"),
    )


def register_all():
    register(
        "hysys_get_column_profile",
        "蒸留塔・吸収塔の段別プロファイル(温度・圧力・液モル流量・気モル流量)を返す。"
        "column_name はメインフローシート上の塔名。"
        "注意: V14 標準 COM では段別 L/V 流量は露出していない (null になる)。"
        "必要なら HYSYS GUI で Internal Streams を作成して各段流量を MaterialStream 化してください。",
        {
            "type": "object",
            "properties": {
                "column_name": {"type": "string", "description": "塔の名前 (例: 'CO2 Absorption Tower')"},
                "flowsheet_path": {"type": "string", "default": "Main"},
            },
            "required": ["column_name"],
        },
        _get_column_profile,
        READ,
    )
    register(
        "hysys_list_column_specs",
        "塔の全 Specification (Reflux Ratio, Comp Fraction, Draw Rate 等) を返す。is_active と goal/current/error を含む。",
        {
            "type": "object",
            "properties": {
                "column_name": {"type": "string"},
                "flowsheet_path": {"type": "string", "default": "Main"},
            },
            "required": ["column_name"],
        },
        _list_column_specs,
        READ,
    )
    register(
        "hysys_set_column_spec",
        "塔の Spec の Goal を変更。"
        "unit を指定すると Goal.SetValue(value, unit) を使い、空文字なら GoalValue を直接代入。"
        "confirm=true 必須。",
        {
            "type": "object",
            "properties": {
                "column_name": {"type": "string"},
                "spec_name": {"type": "string", "description": "list_column_specs で得られる name (例: 'Reflux Ratio')"},
                "goal": {"type": "number"},
                "unit": {"type": "string", "default": ""},
                "confirm": {"type": "boolean", "default": False},
                "flowsheet_path": {"type": "string", "default": "Main"},
            },
            "required": ["column_name", "spec_name", "goal"],
        },
        _set_column_spec,
        WRITE,
    )
    register(
        "hysys_set_column_spec_active",
        "塔の Spec の IsActive 状態を切り替える。"
        "Active な Spec の数 = 塔の Degrees of Freedom と一致する必要があるため、"
        "新しい Spec を Active にしたら別の Spec を Inactive にする等の整合性確保が呼び出し側責任。"
        "confirm=true 必須。",
        {
            "type": "object",
            "properties": {
                "column_name": {"type": "string"},
                "spec_name": {"type": "string"},
                "active": {"type": "boolean"},
                "confirm": {"type": "boolean", "default": False},
                "flowsheet_path": {"type": "string", "default": "Main"},
            },
            "required": ["column_name", "spec_name", "active"],
        },
        _set_column_spec_active,
        WRITE,
    )
    register(
        "hysys_get_column_spec_detail",
        "塔の Spec の詳細を返す。"
        "TypeName / SpecifiedStage / SpecifiedDraw / Phase / FlowBasis / "
        "IncludedComponents / IsUsedAsEstimate 等の Spec 種別依存の属性を含む。"
        "list_column_specs では分からない『何のSpecか』を把握する用。",
        {
            "type": "object",
            "properties": {
                "column_name": {"type": "string"},
                "spec_name": {"type": "string"},
                "flowsheet_path": {"type": "string", "default": "Main"},
            },
            "required": ["column_name", "spec_name"],
        },
        _get_column_spec_detail,
        READ,
    )
    register(
        "hysys_clear_column_estimates",
        "塔の推定値をクリア。物性パッケージ切替後や前回解が新条件で使えない場合に呼ぶ。"
        "scope='all' (デフォルト) は ClearAllEstimates(), "
        "'composition' は ClearAllCompositionEstimates(), "
        "'tray_composition' は ClearTrayCompositionEstimates()。"
        "confirm=true 必須。",
        {
            "type": "object",
            "properties": {
                "column_name": {"type": "string"},
                "scope": {"type": "string", "enum": ["all", "composition", "tray_composition"], "default": "all"},
                "confirm": {"type": "boolean", "default": False},
                "flowsheet_path": {"type": "string", "default": "Main"},
            },
            "required": ["column_name"],
        },
        _clear_column_estimates,
        WRITE,
    )
    register(
        "hysys_set_column_solver_param",
        "塔ソルバーパラメータを設定。"
        "parameter: IsUsingSolutionForEstimates / MaximumIterations / "
        "IsAdaptiveDamping / DampingFactor / AdaptiveDampingPeriod / "
        "EquilibriumErrorTolerance / HeatSpecErrorTolerance / TraceLevel。"
        "前回解を初期推定に流用しない設定や、反復上限の引き上げ等に使う。"
        "confirm=true 必須。",
        {
            "type": "object",
            "properties": {
                "column_name": {"type": "string"},
                "parameter": {"type": "string"},
                "value": {"description": "数値またはブール (parameter に応じて型キャスト)"},
                "confirm": {"type": "boolean", "default": False},
                "flowsheet_path": {"type": "string", "default": "Main"},
            },
            "required": ["column_name", "parameter", "value"],
        },
        _set_column_solver_param,
        WRITE,
    )
    register(
        "hysys_remove_column_spec",
        "塔の Specification を削除。"
        "クローン塔から継承された不要 Spec を一掃する等。"
        "DOF が崩れる点に注意 (呼び出し側責任)。confirm=true 必須。",
        {
            "type": "object",
            "properties": {
                "column_name": {"type": "string"},
                "spec_name": {"type": "string"},
                "confirm": {"type": "boolean", "default": False},
                "flowsheet_path": {"type": "string", "default": "Main"},
            },
            "required": ["column_name", "spec_name"],
        },
        _remove_column_spec,
        WRITE,
    )
    register(
        "hysys_add_column_spec",
        "塔に新規 Spec を追加。"
        "spec_type: reflux_ratio / comp_recovery / comp_fraction / draw_rate / "
        "temperature / duty / vapour_fraction / pressure。"
        "stage='Condenser'/'Reboiler' 等、draw='R'/'C' 等、phase=0(Liq)/1(Vap)/2、"
        "components は Comp Recovery/Fraction 用の成分名リスト。"
        "confirm=true 必須。",
        {
            "type": "object",
            "properties": {
                "column_name": {"type": "string"},
                "spec_type": {
                    "type": "string",
                    "enum": [
                        "reflux_ratio", "comp_recovery", "comp_fraction",
                        "draw_rate", "temperature", "duty",
                        "vapour_fraction", "pressure",
                    ],
                },
                "spec_name": {"type": "string"},
                "goal": {"type": "number"},
                "stage": {"type": "string", "description": "段名 (Condenser/Reboiler/1_Main TS 等)"},
                "draw": {"type": "string", "description": "留出名 (R/C 等)"},
                "phase": {"type": "integer", "description": "0=Liquid 1=Vapour 2=Mixed"},
                "components": {"type": "array", "items": {"type": "string"}},
                "is_active": {"type": "boolean", "default": True},
                "flow_basis": {"type": "integer"},
                "confirm": {"type": "boolean", "default": False},
                "flowsheet_path": {"type": "string", "default": "Main"},
            },
            "required": ["column_name", "spec_type", "spec_name"],
        },
        _add_column_spec,
        WRITE,
    )
    register(
        "hysys_column_reset",
        "塔個別 Reset (ColumnFlowsheet.Reset)。"
        "フローシート全体 reset と違い、塔だけを未計算状態に戻す。"
        "FP切替後の停滞解クリアや、Solved with non-convergence 凍結解除に使う。"
        "confirm=true 必須。",
        {
            "type": "object",
            "properties": {
                "column_name": {"type": "string"},
                "flowsheet_path": {"type": "string", "default": "Main"},
                "confirm": {"type": "boolean", "default": False},
            },
            "required": ["column_name"],
        },
        _column_reset,
        WRITE,
    )
    register(
        "hysys_column_run",
        "塔個別 Run (ColumnFlowsheet.Run)。"
        "メインソルバーが Hold でも塔単独で動かせる。"
        "返値に CurrentIteration / CfsConverged / SolvingStatus / EquilibriumError を含む。",
        {
            "type": "object",
            "properties": {
                "column_name": {"type": "string"},
                "flowsheet_path": {"type": "string", "default": "Main"},
            },
            "required": ["column_name"],
        },
        _column_run,
        WRITE,
    )
