"""接続・セッション・ケース管理ツール。

tag:
    open/close/reconnect/switch_instance/set_active_case/save = SESSION
    list_cases/list_instances = READ
"""

from __future__ import annotations

from hysys_mcp.registry import READ, SESSION, register


def _open(client, a):
    client.open_file(a["path"])
    return f"Opened: {a['path']}"


def _close(client, a):
    client.close_file()
    return "Closed."


def _reconnect(client, a):
    return client.reconnect()


def _list_cases(client, a):
    return client.list_cases()


def _list_instances(client, a):
    return client.list_instances()


def _switch_instance(client, a):
    return client.switch_instance(int(a["index"]))


def _set_active_case(client, a):
    return client.set_active_case(a["name_or_index"])


def _save(client, a):
    return client.save_as(path=a["path"], confirm=bool(a.get("confirm", False)))


def register_all():
    register(
        "hysys_open",
        "指定パスの .hsc / .tpl ファイルを開く。"
        "HYSYS がインストールされ起動済みである必要がある。"
        "開いたケースは自動的に ActiveDocument になる。",
        {
            "type": "object",
            "properties": {
                "path": {
                    "type": "string",
                    "description": "ファイルの絶対パス(例: C:\\\\path\\\\to\\\\file.hsc)",
                },
            },
            "required": ["path"],
        },
        _open,
        SESSION,
    )
    register(
        "hysys_close",
        "現在の ActiveDocument (アクティブケース) を閉じる。"
        "他のケースが残っていればそのうちの1つが ActiveDocument に昇格する。",
        {"type": "object", "properties": {}},
        _close,
        SESSION,
    )
    register(
        "hysys_reconnect",
        "キャッシュした HYSYS.Application 参照を破棄して再 Dispatch する。"
        "HYSYS が落ちて再起動された場合 (ライセンス切れ→再起動など) に "
        "RPC エラーから復旧するために使う。返値は再接続後のケース数と "
        "ActiveDocument 名。",
        {"type": "object", "properties": {}},
        _reconnect,
        SESSION,
    )
    register(
        "hysys_list_cases",
        "HYSYS で開いている全 SimulationCase を列挙。"
        "各ケースの index / name / full_name / is_active を返す。"
        "複数ファイルを並行操作したい場合、まずこのツールで一覧確認 → "
        "hysys_set_active_case で切替 → 通常ツールを呼ぶ流れになる。",
        {"type": "object", "properties": {}},
        _list_cases,
        READ,
    )
    register(
        "hysys_list_instances",
        "起動中の HYSYS.Application インスタンスを ROT (Running Object Table) "
        "から全列挙。Explorer ダブルクリック等で別プロセスとして起動された "
        "HYSYS も含めて発見する。各インスタンスの index / 開いているケース名 / "
        "現在 MCP が接続中か (is_current) を返す。"
        "is_current=false のインスタンスに切替えるには hysys_switch_instance を使う。",
        {"type": "object", "properties": {}},
        _list_instances,
        READ,
    )
    register(
        "hysys_switch_instance",
        "MCP の接続先 HYSYS インスタンスを切替える。"
        "index は hysys_list_instances で得られる値を指定。"
        "切替後、以降のすべての MCP ツール呼び出しはそのインスタンス上の "
        "ActiveDocument に対して行われる。",
        {
            "type": "object",
            "properties": {
                "index": {
                    "description": "list_instances で得たインスタンス index (0始まり)",
                    "type": "integer",
                },
            },
            "required": ["index"],
        },
        _switch_instance,
        SESSION,
    )
    register(
        "hysys_set_active_case",
        "HYSYS の ActiveDocument を切り替える。"
        "name_or_index は整数(0始まり) または name / FullName末尾 で指定可。"
        "切替後、以降のすべての MCP ツール呼び出しは新しいケースを操作する。"
        "HYSYS GUI 側でも対応するケースが前面に来る。",
        {
            "type": "object",
            "properties": {
                "name_or_index": {
                    "description": "ケース index (0始まり整数) または name / FullName末尾 (文字列)",
                },
            },
            "required": ["name_or_index"],
        },
        _set_active_case,
        SESSION,
    )
    register(
        "hysys_save",
        "ケースを別名保存(SaveAs)。"
        "上書きは禁止 (既存ファイルパスを指定するとエラー)。"
        "デフォルトはドライラン。",
        {
            "type": "object",
            "properties": {
                "path": {
                    "type": "string",
                    "description": "保存先 Windows 絶対パス (例: C:\\\\path\\\\to\\\\new.hsc)。既存ファイル不可。",
                },
                "confirm": {"type": "boolean", "default": False},
            },
            "required": ["path"],
        },
        _save,
        SESSION,
    )
