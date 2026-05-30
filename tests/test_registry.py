"""レジストリ層のオフライン検証 (HYSYS / mcp 不要)。

registry.py と tools/*.py は mcp / pywin32 に依存しないため、
WSL/Linux でもここだけは完全に回せる。
"""

from __future__ import annotations

import os
from dataclasses import dataclass

import pytest

from hysys_mcp import registry
from hysys_mcp.tools import register_all_tools

register_all_tools()


# 旧 server.py の if/elif が処理していた全45ツール (parity の基準)
EXPECTED_TOOLS = {
    # connection (session/read)
    "hysys_open", "hysys_close", "hysys_reconnect", "hysys_list_cases",
    "hysys_list_instances", "hysys_switch_instance", "hysys_set_active_case",
    "hysys_save",
    # streams
    "hysys_list_streams", "hysys_get_stream", "hysys_set_stream",
    "hysys_find_streams", "hysys_get_stream_phys", "hysys_add_internal_stream",
    # unit ops
    "hysys_list_unit_ops", "hysys_set_unit_op_param", "hysys_get_heat_op",
    "hysys_find_ops",
    # columns
    "hysys_get_column_profile", "hysys_list_column_specs", "hysys_set_column_spec",
    "hysys_set_column_spec_active", "hysys_get_column_spec_detail",
    "hysys_clear_column_estimates", "hysys_set_column_solver_param",
    "hysys_remove_column_spec", "hysys_add_column_spec", "hysys_column_reset",
    "hysys_column_run",
    # solver
    "hysys_get_status", "hysys_run", "hysys_reset", "hysys_set_solver_can_solve",
    "hysys_balance_check", "hysys_case_study",
    # logical
    "hysys_list_logical_ops", "hysys_get_logical_op", "hysys_set_adjust_target",
    "hysys_reset_recycle",
    # fluid
    "hysys_get_fluid_package", "hysys_list_components", "hysys_get_component",
    # generic
    "hysys_introspect", "hysys_call_method", "hysys_set_property",
    # build (フローシート構築)
    "hysys_create_stream", "hysys_create_unit_op", "hysys_connect_stream",
    "hysys_disconnect_stream", "hysys_delete_object", "hysys_list_ports",
}


def test_all_tools_registered():
    assert set(registry.REGISTRY.keys()) == EXPECTED_TOOLS
    assert len(registry.REGISTRY) == 51


def test_no_orphan_or_missing():
    """全ツールが registry に存在し、余計なものが無い。"""
    assert len(EXPECTED_TOOLS) == 51
    for name in EXPECTED_TOOLS:
        assert name in registry.REGISTRY, f"missing tool: {name}"


def test_tag_distribution():
    counts = {registry.READ: 0, registry.SESSION: 0, registry.WRITE: 0}
    for spec in registry.REGISTRY.values():
        counts[spec.tag] += 1
    # 元45 (read20/session6/write19) + build6 (read1: list_ports / write5)
    assert counts == {registry.READ: 21, registry.SESSION: 6, registry.WRITE: 24}


def test_every_schema_is_object():
    for spec in registry.REGISTRY.values():
        assert spec.input_schema.get("type") == "object", spec.name
        assert "properties" in spec.input_schema, spec.name


def test_mode_filtering_counts():
    assert len(registry.specs_for_mode("readonly")) == 21      # read のみ
    assert len(registry.specs_for_mode("default")) == 27       # read + session
    assert len(registry.specs_for_mode("enhanced")) == 51      # 全部


def test_build_tools_gated_to_write():
    """構築系は list_ports を除き write tag = 既定 default では非公開。"""
    default_visible = {s.name for s in registry.specs_for_mode("default")}
    for name in ("hysys_create_stream", "hysys_create_unit_op",
                 "hysys_connect_stream", "hysys_delete_object"):
        assert name not in default_visible
    assert "hysys_list_ports" in default_visible        # read なので見える
    assert "hysys_create_stream" in {s.name for s in registry.specs_for_mode("enhanced")}


def test_default_mode_hides_writes():
    """既定 (default) では書込/ソルバ系が公開されない = 読取専用ルール担保。"""
    visible = {s.name for s in registry.specs_for_mode("default")}
    assert "hysys_set_stream" not in visible
    assert "hysys_run" not in visible
    assert "hysys_call_method" not in visible
    # 読取と保存は見える
    assert "hysys_get_stream" in visible
    assert "hysys_save" in visible
    assert "hysys_list_streams" in visible


def test_readonly_mode_hides_session():
    visible = {s.name for s in registry.specs_for_mode("readonly")}
    assert "hysys_save" not in visible
    assert "hysys_open" not in visible
    assert "hysys_get_stream" in visible


def test_is_allowed_gate():
    assert registry.is_allowed("hysys_get_stream", "readonly")
    assert not registry.is_allowed("hysys_save", "readonly")
    assert registry.is_allowed("hysys_save", "default")
    assert not registry.is_allowed("hysys_run", "default")
    assert registry.is_allowed("hysys_run", "enhanced")
    assert not registry.is_allowed("nonexistent", "enhanced")


def test_server_mode_env_resolution(monkeypatch):
    monkeypatch.setenv("HYSYS_MCP_MODE", "enhanced")
    assert registry.server_mode() == "enhanced"
    monkeypatch.setenv("HYSYS_MCP_MODE", "ENHANCED")  # 大小無視
    assert registry.server_mode() == "enhanced"
    # 設定済みだが不正値 (タイポ等) は安全側の readonly に丸める (fail-safe)
    monkeypatch.setenv("HYSYS_MCP_MODE", "bogus")
    assert registry.server_mode() == "readonly"
    # 未設定は既定 default
    monkeypatch.delenv("HYSYS_MCP_MODE", raising=False)
    assert registry.server_mode() == "default"


# ─────────────────────────────────────────────────
# render_text / to_jsonable (再帰 normalizer)
# ─────────────────────────────────────────────────
@dataclass
class _Leaf:
    x: int
    label: str


@dataclass
class _Nest:
    name: str
    leaves: list


def test_render_text_str_passthrough():
    assert registry.render_text("Opened: C:/a.hsc") == "Opened: C:/a.hsc"
    assert registry.render_text("Closed.") == "Closed."


def test_render_text_dataclass_and_nested():
    obj = _Nest(name="col", leaves=[_Leaf(1, "a"), _Leaf(2, "b")])
    out = registry.to_jsonable(obj)
    assert out == {"name": "col", "leaves": [{"x": 1, "label": "a"}, {"x": 2, "label": "b"}]}


def test_to_jsonable_unknown_object_stringified():
    class Weird:
        def __repr__(self):
            return "<COM ref>"
    assert registry.to_jsonable(Weird()) == "<COM ref>"


def test_render_text_list_of_dataclasses():
    import json
    payload = [_Leaf(1, "a"), _Leaf(2, "b")]
    txt = registry.render_text(payload)
    assert json.loads(txt) == [{"x": 1, "label": "a"}, {"x": 2, "label": "b"}]


# ─────────────────────────────────────────────────
# 全 handler のディスパッチ parity (FakeClient)
# ─────────────────────────────────────────────────
class _FakeClient:
    """任意メソッド呼び出しを記録し dict を返す偽 HysysClient。"""

    def __init__(self):
        self.calls = []

    def __getattr__(self, attr):
        def _rec(*args, **kwargs):
            self.calls.append((attr, args, kwargs))
            return {"method": attr, "args": list(args), "kwargs": kwargs}
        return _rec


# 全ツール共通で渡す permissive な引数 (required を全カバー)
_SAMPLE_ARGS = {
    "path": "case.Foo",
    "index": 0,
    "name_or_index": 0,
    "name": "stream1",
    "op_name": "Cooler1",
    "parameter": "Duty",
    "value": 1.0,
    "column_name": "Tower",
    "spec_name": "Reflux Ratio",
    "goal": 1.5,
    "active": True,
    "stage": 2,
    "spec_type": "reflux_ratio",
    "target_kind": "stream",
    "target_name": "s1",
    "target_field": "temperature_C",
    "sweep_values": [1, 2, 3],
    "observe": [{"kind": "stream", "name": "s1", "field": "temperature_C"}],
    "method": "Reset",
    "property_name": "PropPkgName",
    # build 系 (op_name は上で定義済みのものを共用)
    "type_name": "coolerop",
    "stream_name": "S-1",
    "kind": "unit_op",
}


@pytest.mark.parametrize("name", sorted(EXPECTED_TOOLS))
def test_every_handler_dispatches(name):
    """全45 handler が FakeClient で例外なく実行され、render 可能な payload を返す。"""
    spec = registry.REGISTRY[name]
    client = _FakeClient()
    payload = spec.handler(client, dict(_SAMPLE_ARGS))
    text = registry.render_text(payload)
    assert isinstance(text, str) and len(text) > 0
    # open/close 以外は client メソッドを最低1回呼ぶ
    assert client.calls, f"{name} did not call any client method"


if __name__ == "__main__":
    pytest.main([__file__, "-v"])
