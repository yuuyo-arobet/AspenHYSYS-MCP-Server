"""汎用 COM アクセスツール (introspect / call_method / set_property)。

call_method / set_property は任意の COM メンバへ書き込み得るため WRITE。
introspect は読取のみなので READ。
"""

from __future__ import annotations

from hysys_mcp.registry import READ, WRITE, register


def _introspect(client, a):
    return client.introspect(
        path=a["path"],
        filter_keyword=a.get("filter_keyword"),
        max_members=int(a.get("max_members", 200)),
    )


def _call_method(client, a):
    return client.call_method(
        path=a["path"],
        method=a["method"],
        args=a.get("args"),
        confirm=bool(a.get("confirm", False)),
    )


def _set_property(client, a):
    return client.set_property(
        path=a["path"],
        property_name=a["property_name"],
        value=a["value"],
        unit=a.get("unit"),
        confirm=bool(a.get("confirm", False)),
    )


def register_all():
    register(
        "hysys_introspect",
        "任意の COM オブジェクトのメンバを列挙 (デバッグ・新 API 探索用)。"
        "path は 'case'/'flowsheet'/'fp'/'app' を起点としたドット式。"
        "例: 'flowsheet.Operations.Item(\"Cooler1\")', 'fp.Components.Item(0)'。"
        "filter_keyword で名前部分マッチで絞り込み可。",
        {
            "type": "object",
            "properties": {
                "path": {"type": "string"},
                "filter_keyword": {"type": "string"},
                "max_members": {"type": "integer", "default": 200},
            },
            "required": ["path"],
        },
        _introspect,
        READ,
    )
    register(
        "hysys_call_method",
        "任意 COM オブジェクトのメソッドを呼ぶ汎用ツール。"
        "introspect でメソッド存在を確認してから使うこと。confirm=true 必須。"
        "args は順序付きの [{type:'literal'|'path', value:...}] リスト。"
        "type='path' のとき value は introspect と同じ namespace で eval されて COM オブジェクトに解決。"
        "例: 反応セットを Fluid Package にアタッチ → "
        "path=\"case.BasisManager.ReactionPackageManager.ReactionSets.Item('Set-1')\", "
        "method=\"AssociateFluidPackage\", "
        "args=[{\"type\":\"path\",\"value\":\"case.BasisManager.FluidPackages.Item(1)\"}]",
        {
            "type": "object",
            "properties": {
                "path": {"type": "string"},
                "method": {"type": "string"},
                "args": {
                    "type": "array",
                    "items": {
                        "type": "object",
                        "properties": {
                            "type": {"type": "string", "enum": ["literal", "path"]},
                            "value": {},
                        },
                        "required": ["type"],
                    },
                    "default": [],
                },
                "confirm": {"type": "boolean", "default": False},
            },
            "required": ["path", "method"],
        },
        _call_method,
        WRITE,
    )
    register(
        "hysys_set_property",
        "任意 COM オブジェクトのプロパティを書き換える汎用ツール。confirm=true 必須。"
        "書込モードは自動分岐: "
        "(1) property が SetValue を持つ CDispatch → attr.SetValue(value [, unit]) "
        "  (HYSYS Variable: Temperature/Pressure 等。unit='C','kPa' 指定可)。"
        "(2) property が SetValues を持ち value が list → attr.SetValues(list) "
        "  (TemperatureEsts 等の配列 Variable)。"
        "(3) それ以外 → setattr(obj, property_name, value) "
        "  (PropPkgName, ColumnAlgorithm, MaximumIterations 等の素プロパティ)。"
        "confirm=false で dry-run (旧値と method 候補を返す)。"
        "path は call_method/introspect と同じく 'case'/'flowsheet'/'fp'/'app' 起点のドット式。"
        "例: 物性パッケージ切替 → "
        "path=\"case.BasisManager.FluidPackages.Item('Basis-1')\", "
        "property_name=\"PropPkgName\", value=\"Acid Gas - Chemical Solvents\"。"
        "塔アルゴリズム切替 → "
        "path=\"flowsheet.Operations.Item('tower').ColumnFlowsheet\", "
        "property_name=\"ColumnAlgorithm\", value=2。"
        "段別温度推定値 → "
        "path=\"flowsheet.Operations.Item('tower').ColumnFlowsheet.TemperatureEsts.Item(0)\", "
        "property_name=\"(SetValue対象として再取得)\", value=90.0, unit=\"C\"。",
        {
            "type": "object",
            "properties": {
                "path": {"type": "string"},
                "property_name": {"type": "string"},
                "value": {},
                "unit": {
                    "type": "string",
                    "description": "SetValue モード時のみ有効 (例: 'C', 'kPa')。省略可。",
                },
                "confirm": {"type": "boolean", "default": False},
            },
            "required": ["path", "property_name", "value"],
        },
        _set_property,
        WRITE,
    )
