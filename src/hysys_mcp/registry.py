"""ツールレジストリ (MCP SDK 非依存)。

設計方針:
    - この module は ``mcp`` パッケージに一切依存しない。
      → HYSYS / mcp が無い WSL 環境でも import でき、単体テストが回る。
    - 各ツールは ToolSpec(name, description, input_schema, handler, tag) で表現。
    - server.py が REGISTRY を読み、mcp.types.Tool への変換と
      stdio ディスパッチを行う薄い adapter になる。

モード (セキュリティゲート):
    環境変数 ``HYSYS_MCP_MODE`` で「どの tag のツールを公開するか」を制御。
      readonly : read のみ                 (最も安全。閲覧専用)
      default  : read + session            (既定。読取 + 保存/接続管理。モデル値は不変)
      enhanced : read + session + write     (書込/ソルバ起動を解禁)
    既定は ``default``。これによりユーザの運用ルール
    「HYSYS は MCP 読取専用、書込は GUI」が設定なしで担保される。
"""

from __future__ import annotations

import json
import os
from dataclasses import dataclass, fields, is_dataclass
from typing import Any, Callable

# ─────────────────────────────────────────────────
# tag (ツールの副作用レベル)
# ─────────────────────────────────────────────────
READ = "read"        # モデルもディスクも変更しない純読取
SESSION = "session"  # 状態遷移 / 永続化はするがモデル値は変えない (open/close/save 等)
WRITE = "write"      # モデル変更 or ソルバ起動 (フリーズ要因)

_VALID_TAGS = (READ, SESSION, WRITE)

# モード → 公開する tag 集合
_MODE_ALLOW: dict[str, set[str]] = {
    "readonly": {READ},
    "default": {READ, SESSION},
    "enhanced": {READ, SESSION, WRITE},
}

DEFAULT_MODE = "default"


def server_mode() -> str:
    """環境変数から現在のサーバモードを解決。

    - 未設定: ``default`` (読取 + 保存。書込は非公開)。
    - 設定済みだが不正値 (タイポ等): 最も安全な ``readonly`` に丸める。
      ``default`` に丸めると "readonly のつもりが session 解禁" という
      fail-open になるため、設定ミス時は安全側へ倒す (Codex review)。
    """
    raw = os.environ.get("HYSYS_MCP_MODE")
    if raw is None:
        return DEFAULT_MODE
    m = raw.strip().lower()
    if m in _MODE_ALLOW:
        return m
    # 不正値はデバッグの手がかりとして警告してから安全側へ (Codex review nit)
    import sys
    print(
        f"[hysys-mcp] WARNING: 不明な HYSYS_MCP_MODE={raw!r}。"
        f"安全側の 'readonly' で起動します (有効値: {sorted(_MODE_ALLOW)})。",
        file=sys.stderr,
    )
    return "readonly"


def allowed_tags(mode: str | None = None) -> set[str]:
    """指定モード (省略時は環境変数) で公開される tag 集合を返す。"""
    return _MODE_ALLOW[mode or server_mode()]


# ─────────────────────────────────────────────────
# ToolSpec とレジストリ
# ─────────────────────────────────────────────────
HandlerFn = Callable[[Any, dict], Any]  # (client, arguments) -> payload


@dataclass(frozen=True)
class ToolSpec:
    name: str
    description: str
    input_schema: dict
    handler: HandlerFn
    tag: str


REGISTRY: dict[str, ToolSpec] = {}


def register(
    name: str,
    description: str,
    input_schema: dict,
    handler: HandlerFn,
    tag: str,
) -> None:
    """ツールを登録する。名前重複・不正 tag は即エラー。"""
    if tag not in _VALID_TAGS:
        raise ValueError(f"invalid tag '{tag}' for tool '{name}' (must be one of {_VALID_TAGS})")
    if name in REGISTRY:
        raise ValueError(f"duplicate tool registration: {name}")
    REGISTRY[name] = ToolSpec(name, description, input_schema, handler, tag)


def get_spec(name: str) -> ToolSpec | None:
    return REGISTRY.get(name)


def specs_for_mode(mode: str | None = None) -> list[ToolSpec]:
    """現在のモードで公開すべき ToolSpec を登録順で返す。"""
    tags = allowed_tags(mode)
    return [s for s in REGISTRY.values() if s.tag in tags]


def is_allowed(name: str, mode: str | None = None) -> bool:
    spec = REGISTRY.get(name)
    return spec is not None and spec.tag in allowed_tags(mode)


# ─────────────────────────────────────────────────
# payload → JSON テキスト 変換 (再帰 normalizer)
# ─────────────────────────────────────────────────
def to_jsonable(obj: Any) -> Any:
    """dataclass / list / dict / プリミティブを再帰的に JSON 化可能な構造へ。

    - asdict() は使わない (deepcopy が COM オブジェクトで失敗し得るため、
      fields() で手動再帰)。
    - 未知のオブジェクト (COM 参照等) は str() フォールバック。
    """
    if obj is None or isinstance(obj, (str, int, float, bool)):
        return obj
    if is_dataclass(obj) and not isinstance(obj, type):
        return {f.name: to_jsonable(getattr(obj, f.name)) for f in fields(obj)}
    if isinstance(obj, dict):
        return {str(k): to_jsonable(v) for k, v in obj.items()}
    if isinstance(obj, (list, tuple, set)):
        return [to_jsonable(v) for v in obj]
    return str(obj)


def render_text(payload: Any) -> str:
    """handler の戻り値を MCP TextContent 用の文字列へ。

    - str はそのまま (open/close の "Opened: ..." 等、現行挙動を維持)。
    - それ以外は JSON 整形 (ensure_ascii=False, indent=2)。
    """
    if isinstance(payload, str):
        return payload
    return json.dumps(to_jsonable(payload), ensure_ascii=False, indent=2)
