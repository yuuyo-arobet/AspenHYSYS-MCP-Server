"""HYSYS COM Automation wrapper.

Windows ネイティブ Python 専用(pywin32 が WSL では動作しない)。

使い方:
    client = HysysClient()
    client.connect()  # 既存の HYSYS インスタンスに接続
    streams = client.list_streams()
    info = client.get_stream("Feed")
"""

from __future__ import annotations

import sys
from contextlib import contextmanager
from dataclasses import dataclass, asdict
from typing import Any

if sys.platform != "win32":
    # WSL/Linux/macOS では import 時点では失敗させない(stub だけ提供)
    win32com = None
    pythoncom = None
else:
    import win32com.client  # type: ignore[import]
    import pythoncom  # type: ignore[import]


@dataclass
class StreamInfo:
    """マテリアルストリーム情報のスナップショット。"""

    name: str
    flowsheet: str  # 所属フローシート名(Main / サブフロー)
    temperature_C: float | None
    pressure_kPa: float | None
    molar_flow_kgmole_h: float | None
    mass_flow_kg_h: float | None
    vapor_fraction: float | None
    is_known: bool
    composition_molar: dict[str, float]  # 成分 → モル分率


@dataclass
class UnitOpInfo:
    """装置操作の情報スナップショット。"""

    name: str
    type_name: str  # "Heater", "Mixer", "DistillationColumn", etc.
    flowsheet: str
    solve_complete: bool


@dataclass
class SolverStatus:
    """ソルバーの全体状態。"""

    can_solve: bool
    is_solving: bool
    error_messages: list[str]


@dataclass
class SetResult:
    """書き込み操作の結果。dry-run / 実行どちらも同じ形で返す。"""

    target: str  # ストリーム名 or 装置名
    field: str  # "temperature" / "pressure" / "molar_flow" / "composition_molar" / 任意のparam名
    before: Any
    after: Any
    applied: bool  # True なら HYSYS に書き込み済み、False ならドライラン
    unit: str | None = None


@dataclass
class ColumnTrayInfo:
    """蒸留塔・吸収塔の1段分の状態。"""

    stage: int  # 0始まり(condenser=0 etc, 実装依存)
    temperature_C: float | None
    pressure_kPa: float | None
    liquid_molar_flow_kgmole_h: float | None
    vapor_molar_flow_kgmole_h: float | None


@dataclass
class ColumnSpecInfo:
    """塔のSpecification 1件分。"""

    name: str
    is_active: bool
    goal: float | None  # 規定値(デフォルト単位)
    current: float | None  # 現状値
    error: float | None  # current - goal
    abs_tolerance: float | None


@dataclass
class CaseStudyRow:
    """case_study の1ケース分の結果。"""

    target_value: float
    observed: dict  # {observation_key: value} のフラットな辞書
    converged: bool
    notes: str = ""


@dataclass
class CaseStudyResult:
    """case_study 全体の結果。"""

    target_kind: str  # "stream" / "unit_op"
    target_name: str
    target_field: str
    target_unit: str | None
    sweep_values: list[float]
    observe: list[dict]
    rows: list[CaseStudyRow]
    applied: bool  # True なら HYSYS で実走した、False ならドライラン
    reverted: bool  # True なら最後に元の値に戻した


@dataclass
class LogicalOpInfo:
    """Logical Op (Adjust/Set/Recycle/Spreadsheet) のスナップショット。"""

    name: str
    type_name: str
    is_active: bool
    details: dict  # type-specific properties


@dataclass
class FluidPackageInfo:
    """物性パッケージのスナップショット。"""

    name: str
    property_package_name: str
    component_count: int
    components: list[str]
    has_reaction_package: bool


@dataclass
class ComponentInfo:
    """コンポーネント個別の物性スナップショット。"""

    name: str
    cas_number: str | None
    molecular_weight: float | None
    normal_boiling_point_C: float | None
    critical_temperature_C: float | None
    critical_pressure_kPa: float | None
    acentricity: float | None


@dataclass
class HeatOpInfo:
    """Cooler / Heater / HeatExchanger 汎用情報。HEX 固有の UA/LMTD は HEX のみ埋まる。"""

    name: str
    type_name: str
    duty_default_units: float | None  # internal SI base
    duty_kW: float | None
    feed_stream: str | None
    product_stream: str | None
    energy_stream: str | None
    feed_T_C: float | None
    feed_P_kPa: float | None
    product_T_C: float | None
    product_P_kPa: float | None
    pressure_drop_kPa: float | None
    ua_kJ_C_h: float | None
    lmtd_C: float | None
    min_approach_C: float | None
    hot_inlet: str | None
    hot_outlet: str | None
    cold_inlet: str | None
    cold_outlet: str | None


@dataclass
class BalanceSummary:
    """物質収支・熱収支サマリ。"""

    scope: str
    inlet_streams: list[str]
    outlet_streams: list[str]
    total_in_kg_h: float
    total_out_kg_h: float
    mass_closure_kg_h: float
    mass_closure_pct: float
    inlet_energy_streams: list[str]
    outlet_energy_streams: list[str]
    enthalpy_in_kJ_h: float | None
    enthalpy_out_kJ_h: float | None
    energy_closure_kJ_h: float | None


@dataclass
class StreamPhysProps:
    """ストリームの bulk 物性スナップショット (デフォルト単位は HYSYS 内部単位)。"""

    name: str
    flowsheet: str
    vapor_fraction: float | None
    liquid_fraction: float | None
    molecular_weight: float | None
    mass_density_kg_m3: float | None
    molar_density_kgmole_m3: float | None
    viscosity_cP: float | None
    kinematic_viscosity_cSt: float | None
    thermal_conductivity_W_mK: float | None
    mass_heat_capacity_kJ_kgK: float | None
    molar_heat_capacity_kJ_kgmoleK: float | None
    cp_cv_ratio: float | None
    mass_enthalpy_kJ_kg: float | None
    molar_enthalpy_kJ_kgmole: float | None
    mass_entropy_kJ_kgK: float | None
    heat_flow_kJ_h: float | None
    std_liquid_density_kg_m3: float | None


def _summarize_value(v):
    """COM/Python 値を JSON シリアライズ可能な簡潔表現に変換する内部ヘルパー。

    HYSYS の Variable 系 CDispatch は GetValue('') で素値が取れる。
    name 属性を持つ CDispatch はその名前を返す。
    """
    if v is None:
        return None
    t = type(v).__name__
    if t in ("int", "float", "str", "bool"):
        return v
    if t == "CDispatch":
        try:
            if hasattr(v, "GetValue"):
                try:
                    return v.GetValue("")
                except Exception:
                    pass
            if hasattr(v, "name"):
                try:
                    return f"CDispatch(name={v.name!r})"
                except Exception:
                    return "CDispatch"
            return "CDispatch"
        except Exception:
            return "CDispatch"
    if t == "tuple":
        try:
            return list(v) if len(v) <= 64 else f"tuple(len={len(v)})"
        except Exception:
            return f"<{t}>"
    return f"<{t}>"


class HysysClient:
    """HYSYS COM オブジェクトへのラッパー。

    複数 SimulationCase が開いている場合、デフォルトでは HYSYS の
    ActiveDocument (= GUI で現在選択されているケース) を毎回参照する。
    `set_active_case()` で明示的に切替も可能。
    """

    def __init__(self) -> None:
        self._app: Any | None = None  # HYSYS.Application

    # ─────────────────────────────────────────────────
    # 接続管理
    # ─────────────────────────────────────────────────
    def connect(self) -> None:
        """既存の HYSYS インスタンスに接続。

        HYSYS が起動していない場合は自動起動を試みる。
        """
        if sys.platform != "win32":
            raise RuntimeError(
                "HYSYS COM Automation は Windows ネイティブ Python でのみ動作します。"
                "WSL からは使えません(claude_desktop_config.json で Windows パスの "
                "python.exe を指定してください)。"
            )

        try:
            self._app = win32com.client.Dispatch("HYSYS.Application")
            self._app.Visible = True
        except Exception as e:
            raise RuntimeError(
                f"HYSYS.Application への接続失敗: {e}. "
                "HYSYS がインストールされ起動可能か確認してください。"
            )

    def reconnect(self) -> dict:
        """キャッシュした HYSYS.Application 参照を破棄して再 Dispatch。

        HYSYS プロセスが落ちて再起動された場合 (ライセンス切れ→再起動など)、
        _app が古いプロセスを指したまま RPC エラーになるのを修復する。
        """
        self._app = None
        self.connect()
        info: dict = {"reconnected": True}
        try:
            n = int(self._app.SimulationCases.Count)
            info["open_case_count"] = n
            ad = self._app.ActiveDocument
            info["active_case"] = getattr(ad, "name", None) if ad is not None else None
        except Exception as e:
            info["warning"] = f"接続は成功したがケース情報取得失敗: {e}"
        return info

    def list_instances(self) -> list[dict]:
        """ROT (Running Object Table) から起動中の HYSYS.Application を全列挙。

        HYSYS が複数プロセス起動している場合 (Explorer ダブルクリック等で
        別プロセスが立ち上がった場合)、`connect()` は最初の1つしか掴まない。
        このメソッドは ROT を走査して全 HYSYS インスタンスを発見し、
        現在 `self._app` がどれを指しているか + 各インスタンスが開いている
        ケース名一覧を返す。

        switch_instance() で接続先を切り替えられる。
        """
        if sys.platform != "win32":
            raise RuntimeError("Windows ネイティブ Python でのみ動作します。")
        rot = pythoncom.GetRunningObjectTable()
        bind_ctx = pythoncom.CreateBindCtx(0)
        instances: list[dict] = []
        seen_ids: set = set()
        for moniker in rot:
            try:
                display_name = moniker.GetDisplayName(bind_ctx, None)
            except Exception:
                continue
            # HYSYS の ROT 表示名は ProgID 系 (e.g. "Hysys.Application") か
            # ファイルパス系 (.hsc/.tpl) で現れる。両方カバー。
            dn_lower = display_name.lower() if display_name else ""
            if not (
                "hysys" in dn_lower
                or dn_lower.endswith(".hsc")
                or dn_lower.endswith(".tpl")
            ):
                continue
            try:
                obj = rot.GetObject(moniker)
                disp = obj.QueryInterface(pythoncom.IID_IDispatch)
                wrapped = win32com.client.Dispatch(disp)
            except Exception:
                continue
            # SimulationCases コレクションを持つ = HYSYS.Application。
            # 持たないなら個別ケース or 別オブジェクト→スキップ。
            try:
                sim_cases = wrapped.SimulationCases
                count = int(sim_cases.Count)
            except Exception:
                continue
            # 同一 Application が複数 Moniker で登録されることがあるので
            # IUnknown ポインタ的なIDで重複除去 (id() を使う)
            inst_id = id(wrapped)
            if inst_id in seen_ids:
                continue
            seen_ids.add(inst_id)
            # 各ケースのファイル名を集める
            case_names: list[str] = []
            try:
                for i in range(count):
                    c = sim_cases.Item(i)
                    fn = getattr(c, "FullName", None) or getattr(c, "name", None) or ""
                    case_names.append(str(fn))
            except Exception:
                pass
            # 現在の self._app と一致するか判定
            is_current = False
            if self._app is not None:
                try:
                    is_current = id(self._app) == inst_id
                except Exception:
                    pass
            instances.append({
                "index": len(instances),
                "display_name": display_name,
                "case_count": count,
                "case_names": case_names,
                "is_current": is_current,
            })
        return instances

    def switch_instance(self, index: int) -> dict:
        """list_instances() の index で示される HYSYS インスタンスに接続を切替。

        切替後、self._app は新インスタンスを指し、以降の全 MCP 操作は
        そのインスタンス上のケースに対して行われる。
        """
        if sys.platform != "win32":
            raise RuntimeError("Windows ネイティブ Python でのみ動作します。")
        rot = pythoncom.GetRunningObjectTable()
        bind_ctx = pythoncom.CreateBindCtx(0)
        candidates: list = []
        seen_ids: set = set()
        for moniker in rot:
            try:
                display_name = moniker.GetDisplayName(bind_ctx, None)
            except Exception:
                continue
            dn_lower = display_name.lower() if display_name else ""
            if not (
                "hysys" in dn_lower
                or dn_lower.endswith(".hsc")
                or dn_lower.endswith(".tpl")
            ):
                continue
            try:
                obj = rot.GetObject(moniker)
                disp = obj.QueryInterface(pythoncom.IID_IDispatch)
                wrapped = win32com.client.Dispatch(disp)
                _ = wrapped.SimulationCases.Count  # HYSYS app 検証
            except Exception:
                continue
            inst_id = id(wrapped)
            if inst_id in seen_ids:
                continue
            seen_ids.add(inst_id)
            candidates.append(wrapped)
        if not candidates:
            raise RuntimeError("ROT で HYSYS インスタンスが見つかりません。")
        if index < 0 or index >= len(candidates):
            raise RuntimeError(
                f"index={index} は範囲外 (0..{len(candidates)-1})。"
            )
        self._app = candidates[index]
        info: dict = {"switched_to": index}
        try:
            info["case_count"] = int(self._app.SimulationCases.Count)
            ad = self._app.ActiveDocument
            info["active_case"] = (
                getattr(ad, "FullName", None) or getattr(ad, "name", None)
                if ad is not None else None
            )
        except Exception as e:
            info["warning"] = f"切替後ケース情報取得失敗: {e}"
        return info

    @property
    def _case(self) -> Any | None:
        """現在の SimulationCase を返す (HYSYS の ActiveDocument)。

        毎回 COM 経由で取得するため、HYSYS GUI でユーザがアクティブケースを
        切り替えると MCP も自動で追従する。複数ファイルが開いていても
        ActiveDocument が指す1つだけが対象になる。
        """
        if self._app is None:
            return None
        try:
            return self._app.ActiveDocument
        except Exception:
            return None

    @_case.setter
    def _case(self, value: Any) -> None:
        # 後方互換のため setter は維持するが、ActiveDocument は外部から
        # 書き換えられないので実質 no-op。set_active_case() を使うこと。
        pass

    def open_file(self, path: str) -> None:
        """指定パスの .hsc / .tpl を開く。

        SimulationCases.Open() が ActiveDocument を新規ファイルに自動切替
        するため、明示的なキャッシュは持たない。
        """
        if self._app is None:
            self.connect()
        self._app.SimulationCases.Open(path)

    def close_file(self) -> None:
        case = self._case
        if case is not None:
            case.Close()

    @property
    def is_connected(self) -> bool:
        return self._app is not None and self._case is not None

    # ─────────────────────────────────────────────────
    # 複数ケース対応
    # ─────────────────────────────────────────────────
    def list_cases(self) -> list[dict]:
        """HYSYS で開いている全 SimulationCase を返す。

        各ケースの index / name / FullName / is_active を含む。
        """
        if self._app is None:
            self.connect()
        result: list[dict] = []
        try:
            n = int(self._app.SimulationCases.Count)
        except Exception:
            n = 0
        active_name: str | None = None
        try:
            ad = self._app.ActiveDocument
            if ad is not None:
                active_name = getattr(ad, "name", None)
        except Exception:
            pass
        for i in range(n):
            try:
                c = self._app.SimulationCases.Item(i)
                nm = getattr(c, "name", None)
                result.append({
                    "index": i,
                    "name": nm,
                    "full_name": getattr(c, "FullName", None),
                    "is_active": (nm == active_name),
                })
            except Exception as e:
                result.append({"index": i, "error": f"{type(e).__name__}: {e}"})
        return result

    def set_active_case(self, name_or_index: str | int) -> dict:
        """ActiveDocument を切り替える。

        name_or_index は整数(0始まりindex) または name / FullName末尾。
        Activate() メソッドがあれば呼んで HYSYS GUI 側も切替える。
        """
        if self._app is None:
            self.connect()

        cases = self._app.SimulationCases
        try:
            n = int(cases.Count)
        except Exception:
            n = 0

        target = None

        # 整数として解釈可能か
        try:
            idx = int(name_or_index)
            if 0 <= idx < n:
                target = cases.Item(idx)
        except (ValueError, TypeError):
            pass

        # 名前で直接 Item アクセス
        if target is None:
            try:
                target = cases.Item(str(name_or_index))
            except Exception:
                pass

        # 全走査(FullName 末尾一致 or name 一致)
        if target is None:
            key = str(name_or_index)
            for i in range(n):
                try:
                    c = cases.Item(i)
                    nm = getattr(c, "name", "") or ""
                    full = getattr(c, "FullName", "") or ""
                    if nm == key or full.endswith(key):
                        target = c
                        break
                except Exception:
                    continue

        if target is None:
            raise ValueError(
                f"Case not found: {name_or_index!r}. "
                f"list_cases で開いているケースを確認してください。"
            )

        activated_via: str | None = None
        for method_name in ("Activate", "Show", "BringToFront"):
            if hasattr(target, method_name):
                try:
                    getattr(target, method_name)()
                    activated_via = method_name
                    break
                except Exception:
                    continue

        ad_name: str | None = None
        try:
            ad = self._app.ActiveDocument
            ad_name = getattr(ad, "name", None) if ad is not None else None
        except Exception:
            pass

        return {
            "active_case": getattr(target, "name", None),
            "full_name": getattr(target, "FullName", None),
            "activated_via": activated_via,
            "verified_active": ad_name == getattr(target, "name", None),
        }

    # ─────────────────────────────────────────────────
    # ストリーム
    # ─────────────────────────────────────────────────
    def list_streams(self, flowsheet_path: str | None = None) -> list[str]:
        """メインフロー内の全ストリーム名を返す。

        flowsheet_path: "Main" or "Main/CO2 Absorption Tower" など。
                        None ならメインフロー。
        """
        self._require_case()
        fs = self._navigate_flowsheet(flowsheet_path)
        return [s.name for s in fs.MaterialStreams]

    def get_stream(self, name: str, flowsheet_path: str | None = None) -> StreamInfo:
        self._require_case()
        fs = self._navigate_flowsheet(flowsheet_path)
        s = fs.MaterialStreams.Item(name)

        # 安全に値を取得(値が未計算の場合はNone)
        def safe_get(getter, unit: str) -> float | None:
            try:
                v = getter.GetValue(unit)
                return float(v) if v is not None else None
            except Exception:
                return None

        # 組成
        composition: dict[str, float] = {}
        try:
            comps = self._case.Flowsheet.FluidPackage.Components
            fractions = s.ComponentMolarFractionValue
            for i, comp in enumerate(comps):
                composition[comp.name] = float(fractions[i]) if fractions[i] is not None else 0.0
        except Exception:
            pass

        return StreamInfo(
            name=s.name,
            flowsheet=flowsheet_path or "Main",
            temperature_C=safe_get(s.Temperature, "C"),
            pressure_kPa=safe_get(s.Pressure, "kPa"),
            molar_flow_kgmole_h=safe_get(s.MolarFlow, "kgmole/h"),
            mass_flow_kg_h=safe_get(s.MassFlow, "kg/h"),
            vapor_fraction=safe_get(s.VapourFraction, ""),
            is_known=bool(s.IsKnown) if hasattr(s, "IsKnown") else False,
            composition_molar=composition,
        )

    def set_stream(
        self,
        name: str,
        temperature_C: float | None = None,
        pressure_kPa: float | None = None,
        molar_flow_kgmole_h: float | None = None,
        composition_molar: dict[str, float] | None = None,
        component_molar_flows: dict[str, float] | None = None,
        confirm: bool = False,
        flowsheet_path: str | None = None,
    ) -> list[SetResult]:
        """ストリームの T / P / 流量 / 組成 / 成分別流量を設定。

        confirm=False (デフォルト) はドライラン: before/after を返すだけで実書き込みなし。
        confirm=True で初めて HYSYS に反映する。

        composition_molar: 成分名→モル分率。部分更新可で、合計1に正規化される。
        component_molar_flows: 成分名→モル流量 (kgmole/h)。部分更新可、指定しない成分は現値維持。
                              total = sum(全成分流量) が新しい総モル流量になる。
        """
        self._require_case()
        fs = self._navigate_flowsheet(flowsheet_path)
        s = fs.MaterialStreams.Item(name)
        results: list[SetResult] = []

        def safe_get(getter, unit: str) -> float | None:
            try:
                v = getter.GetValue(unit)
                return float(v) if v is not None else None
            except Exception:
                return None

        if temperature_C is not None:
            before = safe_get(s.Temperature, "C")
            if confirm:
                s.Temperature.SetValue(temperature_C, "C")
                after = safe_get(s.Temperature, "C")
            else:
                after = temperature_C
            results.append(SetResult(name, "temperature", before, after, confirm, "C"))

        if pressure_kPa is not None:
            before = safe_get(s.Pressure, "kPa")
            if confirm:
                s.Pressure.SetValue(pressure_kPa, "kPa")
                after = safe_get(s.Pressure, "kPa")
            else:
                after = pressure_kPa
            results.append(SetResult(name, "pressure", before, after, confirm, "kPa"))

        if molar_flow_kgmole_h is not None:
            before = safe_get(s.MolarFlow, "kgmole/h")
            if confirm:
                s.MolarFlow.SetValue(molar_flow_kgmole_h, "kgmole/h")
                after = safe_get(s.MolarFlow, "kgmole/h")
            else:
                after = molar_flow_kgmole_h
            results.append(SetResult(name, "molar_flow", before, after, confirm, "kgmole/h"))

        if composition_molar is not None:
            comps = self._case.Flowsheet.FluidPackage.Components
            comp_names = [c.name for c in comps]
            try:
                current = list(s.ComponentMolarFractionValue)
            except Exception:
                current = [0.0] * len(comp_names)

            for cname in composition_molar:
                if cname not in comp_names:
                    raise ValueError(
                        f"Component not found in fluid package: {cname!r}. "
                        f"Available: {comp_names}"
                    )

            new_fracs = list(current)
            for cname, val in composition_molar.items():
                new_fracs[comp_names.index(cname)] = float(val)

            total = sum(new_fracs)
            if total <= 0:
                raise ValueError("Composition sums to <= 0. At least one fraction must be positive.")
            new_fracs = [f / total for f in new_fracs]

            before_summary = {c: round(v, 6) for c, v in zip(comp_names, current) if v > 0}
            after_summary = {c: round(v, 6) for c, v in zip(comp_names, new_fracs) if v > 0}

            if confirm:
                s.ComponentMolarFractionValue = tuple(new_fracs)

            results.append(SetResult(
                name, "composition_molar",
                before_summary, after_summary, confirm, "mol_frac (normalized)",
            ))

        if component_molar_flows is not None:
            comps = self._case.Flowsheet.FluidPackage.Components
            comp_names = [c.name for c in comps]
            try:
                current = list(s.ComponentMolarFlow.GetValues("kgmole/h"))
            except Exception:
                current = [0.0] * len(comp_names)

            for cname in component_molar_flows:
                if cname not in comp_names:
                    raise ValueError(
                        f"Component not found: {cname!r}. Available: {comp_names}"
                    )

            new_flows = list(current)
            while len(new_flows) < len(comp_names):
                new_flows.append(0.0)
            for cname, val in component_molar_flows.items():
                new_flows[comp_names.index(cname)] = float(val)

            before_summary = {c: round(v, 6) for c, v in zip(comp_names, current[:len(comp_names)]) if v > 0}
            after_summary = {c: round(v, 6) for c, v in zip(comp_names, new_flows[:len(comp_names)]) if v > 0}

            if confirm:
                s.ComponentMolarFlow.SetValues(tuple(new_flows[:len(comp_names)]), "kgmole/h")

            results.append(SetResult(
                name, "component_molar_flows",
                before_summary, after_summary, confirm, "kgmole/h",
            ))

        return results

    # ─────────────────────────────────────────────────
    # 装置
    # ─────────────────────────────────────────────────
    def list_unit_ops(self, flowsheet_path: str | None = None) -> list[UnitOpInfo]:
        self._require_case()
        fs = self._navigate_flowsheet(flowsheet_path)
        result: list[UnitOpInfo] = []
        for op in fs.Operations:
            result.append(
                UnitOpInfo(
                    name=op.name,
                    type_name=op.TypeName if hasattr(op, "TypeName") else "Unknown",
                    flowsheet=flowsheet_path or "Main",
                    solve_complete=bool(op.SolveComplete) if hasattr(op, "SolveComplete") else False,
                )
            )
        return result

    def set_unit_op_param(
        self,
        op_name: str,
        parameter: str,
        value: float,
        unit: str = "",
        confirm: bool = False,
        flowsheet_path: str | None = None,
    ) -> SetResult:
        """装置のパラメータを設定。

        parameter は HYSYS COM のプロパティ名(例: 'Duty', 'DeltaP', 'PressureDrop')。
        HYSYS の RealVariable (GetValue/SetValue 持ち) なら unit 指定で書き込み、
        プリミティブ属性なら setattr。confirm=False はドライラン。
        """
        self._require_case()
        fs = self._navigate_flowsheet(flowsheet_path)
        op = fs.Operations.Item(op_name)
        if not hasattr(op, parameter):
            raise AttributeError(
                f"Parameter {parameter!r} not found on {op_name} "
                f"(TypeName={getattr(op, 'TypeName', 'Unknown')}). "
                f"Use hysys_list_unit_ops で装置種別を確認してください。"
            )
        prop = getattr(op, parameter)

        is_real_var = hasattr(prop, "GetValue") and hasattr(prop, "SetValue")
        if is_real_var:
            try:
                before = float(prop.GetValue(unit))
            except Exception:
                before = None
            if confirm:
                prop.SetValue(value, unit)
                try:
                    after = float(prop.GetValue(unit))
                except Exception:
                    after = value
            else:
                after = value
        else:
            before = prop
            if confirm:
                setattr(op, parameter, value)
                after = getattr(op, parameter)
            else:
                after = value

        return SetResult(
            target=op_name,
            field=parameter,
            before=before,
            after=after,
            applied=confirm,
            unit=unit or None,
        )

    def get_column_profile(
        self,
        column_name: str,
        flowsheet_path: str | None = None,
    ) -> list[ColumnTrayInfo]:
        """蒸留塔・吸収塔の段別プロファイル(T/P/L/V)を返す。

        column_name は メイン側 Operations にある塔の名前
        (例: 'CO2 Absorption Tower', 'Distillation Tower')。
        """
        self._require_case()
        fs = self._navigate_flowsheet(flowsheet_path)
        col_op = fs.Operations.Item(column_name)
        if not hasattr(col_op, "ColumnFlowsheet"):
            raise ValueError(
                f"{column_name!r} は塔ではない (TypeName="
                f"{getattr(col_op, 'TypeName', 'Unknown')})"
            )
        cfs = col_op.ColumnFlowsheet

        tray_section = None
        try:
            for inner_op in cfs.Operations:
                t = getattr(inner_op, "TypeName", "")
                if "Tray" in t or "tray" in t or "TS" in inner_op.name:
                    tray_section = inner_op
                    break
        except Exception:
            pass
        if tray_section is None:
            try:
                tray_section = cfs.Operations.Item("Main TS")
            except Exception:
                inner_names = []
                try:
                    inner_names = [op.name for op in cfs.Operations]
                except Exception:
                    pass
                raise RuntimeError(
                    f"Tray section が見つからない。塔内部 Operations: {inner_names}"
                )

        def array_or_none(obj, *names) -> list | None:
            for n in names:
                if hasattr(obj, n):
                    try:
                        v = getattr(obj, n)
                        return list(v) if v is not None else None
                    except Exception:
                        continue
            return None

        temps = array_or_none(tray_section, "TemperatureValue", "TemperaturesValue", "TemperatureValues")
        press = array_or_none(tray_section, "PressureValue", "PressuresValue", "PressureValues")
        liq = array_or_none(tray_section, "NetLiquidMolarFlowValue", "LiquidFlowsValue", "LiquidMolarFlowsValue")
        vap = array_or_none(tray_section, "NetVapourMolarFlowValue", "VapourFlowsValue", "VapourMolarFlowsValue")

        n = 0
        for arr in (temps, press, liq, vap):
            if arr is not None:
                n = max(n, len(arr))
        if n == 0:
            n_trays = None
            for nm in ("NumberOfTrays", "NumberOfStages"):
                if hasattr(tray_section, nm):
                    try:
                        n_trays = int(getattr(tray_section, nm))
                        break
                    except Exception:
                        pass
            if n_trays:
                n = n_trays
            else:
                raise RuntimeError(
                    "段数・プロファイル配列が取得できない。HYSYS V14 の COM API を要調査。"
                )

        def at(arr, i):
            try:
                return float(arr[i]) if arr is not None and i < len(arr) else None
            except Exception:
                return None

        return [
            ColumnTrayInfo(
                stage=i,
                temperature_C=at(temps, i),
                pressure_kPa=at(press, i),
                liquid_molar_flow_kgmole_h=at(liq, i),
                vapor_molar_flow_kgmole_h=at(vap, i),
            )
            for i in range(n)
        ]

    # ─────────────────────────────────────────────────
    # ソルバー
    # ─────────────────────────────────────────────────
    def get_solver_status(self) -> SolverStatus:
        """ソルバー状態を返す。

        HYSYS V14 では Trace Pane の READ API が提供されていない (Trace は write-only)。
        代わりに「収束していないオブジェクト」を列挙してメッセージとして返す。
        """
        self._require_case()
        solver = self._case.Solver
        messages: list[str] = []

        try:
            for op in self._case.Flowsheet.Operations:
                try:
                    is_ignored = bool(getattr(op, "IsIgnored", False))
                    if is_ignored:
                        continue
                    is_valid = bool(getattr(op, "IsValid", True))
                    solve_complete = bool(getattr(op, "SolveComplete", True))
                    if not solve_complete or not is_valid:
                        reason = []
                        if not is_valid:
                            reason.append("IsValid=False")
                        if not solve_complete:
                            reason.append("SolveComplete=False")
                        type_name = getattr(op, "TypeName", "?")
                        messages.append(
                            f"[op] {op.name} ({type_name}): {', '.join(reason)}"
                        )
                except Exception as e:
                    messages.append(f"[op] inspection error for {getattr(op, 'name', '?')}: {e}")
        except Exception as e:
            messages.append(f"Operations walk error: {e}")

        try:
            for op in self._case.Flowsheet.Operations:
                if not hasattr(op, "ColumnFlowsheet"):
                    continue
                try:
                    cfs = op.ColumnFlowsheet
                    if hasattr(cfs, "CfsConverged") and not bool(cfs.CfsConverged):
                        cur = getattr(cfs, "CurrentIteration", "?")
                        mx = getattr(cfs, "MaximumIterations", "?")
                        messages.append(
                            f"[column] {op.name}: NOT converged (iter {cur}/{mx})"
                        )
                except Exception as e:
                    messages.append(f"[column] {op.name} inspection error: {e}")
        except Exception:
            pass

        try:
            for s in self._case.Flowsheet.MaterialStreams:
                try:
                    is_known = bool(getattr(s, "IsKnown", True))
                    if not is_known:
                        messages.append(f"[stream] {s.name}: not calculated (IsKnown=False)")
                except Exception:
                    pass
        except Exception:
            pass

        return SolverStatus(
            can_solve=bool(solver.CanSolve) if hasattr(solver, "CanSolve") else False,
            is_solving=bool(solver.IsSolving) if hasattr(solver, "IsSolving") else False,
            error_messages=messages,
        )

    def run(self) -> SolverStatus:
        """ソルバー実行。CanSolve=False の場合は失敗する。"""
        self._require_case()
        solver = self._case.Solver
        if hasattr(solver, "CanSolve") and not solver.CanSolve:
            raise RuntimeError(
                "Solver.CanSolve=False。停止中の操作があるか入力不足の可能性。"
            )
        if hasattr(solver, "Run"):
            solver.Run()
        elif hasattr(solver, "RunBoth"):
            solver.RunBoth()
        else:
            solver.CanSolve = True
        return self.get_solver_status()

    def reset(self, confirm: bool = False) -> dict:
        """ソルバーをリセット(全変数を未計算状態に戻す)。

        破壊的操作なので confirm=True 必須。
        """
        self._require_case()
        if not confirm:
            return {
                "applied": False,
                "message": "Set confirm=True to actually reset the solver.",
            }
        self._case.Solver.Reset()
        return {"applied": True, "message": "Solver reset."}

    def set_solver_can_solve(
        self, value: bool, confirm: bool = False
    ) -> dict:
        """フローシート全体ソルバーの CanSolve (Active/Hold) を切替。

        Hold (value=False) は塔ソルバー停滞時に Stop Run 相当として使う。
        confirm=True 必須。
        """
        self._require_case()
        solver = self._case.Solver
        before = bool(solver.CanSolve) if hasattr(solver, "CanSolve") else None
        if not confirm:
            return {
                "applied": False,
                "before": before,
                "would_set_to": bool(value),
                "message": "Set confirm=True to actually toggle Solver.CanSolve.",
            }
        solver.CanSolve = bool(value)
        after = bool(solver.CanSolve)
        return {
            "applied": True,
            "before": before,
            "after": after,
            "is_solving": bool(solver.IsSolving) if hasattr(solver, "IsSolving") else None,
        }

    def column_reset(
        self,
        column_name: str,
        confirm: bool = False,
        flowsheet_path: str | None = None,
    ) -> dict:
        """塔の Column Environment Reset 相当 (ColumnFlowsheet.Reset)。

        フローシート全体 reset と違い、塔だけを未計算状態に戻す。
        FP切替後の停滞解クリアや、「Solved with non-convergence」凍結解除に使う。
        confirm=True 必須。
        """
        cfs = self._get_column_flowsheet(column_name, flowsheet_path)
        before_iter = int(cfs.CurrentIteration) if hasattr(cfs, "CurrentIteration") else None
        before_converged = bool(cfs.CfsConverged) if hasattr(cfs, "CfsConverged") else None
        if not confirm:
            return {
                "applied": False,
                "target": column_name,
                "before_iteration": before_iter,
                "before_converged": before_converged,
                "message": "Set confirm=True to actually reset the column.",
            }
        if not hasattr(cfs, "Reset"):
            raise AttributeError("ColumnFlowsheet に Reset() が存在しない")
        cfs.Reset()
        return {
            "applied": True,
            "target": column_name,
            "before_iteration": before_iter,
            "before_converged": before_converged,
            "after_iteration": int(cfs.CurrentIteration) if hasattr(cfs, "CurrentIteration") else None,
            "after_converged": bool(cfs.CfsConverged) if hasattr(cfs, "CfsConverged") else None,
        }

    def column_run(
        self,
        column_name: str,
        flowsheet_path: str | None = None,
    ) -> dict:
        """塔個別 Run (ColumnFlowsheet.Run)。

        メインソルバー全体ではなく塔ソルバーだけを起動。
        メインソルバーが Hold (CanSolve=False) でも塔単独で動かしたいときに使う。
        """
        cfs = self._get_column_flowsheet(column_name, flowsheet_path)
        if not hasattr(cfs, "Run"):
            raise AttributeError("ColumnFlowsheet に Run() が存在しない")
        cfs.Run()
        return {
            "applied": True,
            "target": column_name,
            "after_iteration": int(cfs.CurrentIteration) if hasattr(cfs, "CurrentIteration") else None,
            "after_converged": bool(cfs.CfsConverged) if hasattr(cfs, "CfsConverged") else None,
            "solving_status": bool(cfs.SolvingStatus) if hasattr(cfs, "SolvingStatus") else None,
            "equilibrium_error": float(cfs.EquilibriumError) if hasattr(cfs, "EquilibriumError") else None,
        }

    # ─────────────────────────────────────────────────
    # 感度解析 (Case Study)
    # ─────────────────────────────────────────────────
    _STREAM_FIELD_MAP = {
        "temperature_C": ("Temperature", "C"),
        "pressure_kPa": ("Pressure", "kPa"),
        "molar_flow_kgmole_h": ("MolarFlow", "kgmole/h"),
        "mass_flow_kg_h": ("MassFlow", "kg/h"),
        "vapor_fraction": ("VapourFraction", ""),
    }

    def _resolve_stream_var(self, name: str, field: str, flowsheet_path: str | None) -> tuple[Any, str]:
        fs = self._navigate_flowsheet(flowsheet_path)
        s = fs.MaterialStreams.Item(name)
        if field not in self._STREAM_FIELD_MAP:
            raise ValueError(
                f"Stream field {field!r} not supported in case_study. "
                f"Supported: {list(self._STREAM_FIELD_MAP)}"
            )
        attr, unit = self._STREAM_FIELD_MAP[field]
        return getattr(s, attr), unit

    def _resolve_op_var(
        self, name: str, parameter: str, unit: str, flowsheet_path: str | None
    ) -> Any:
        fs = self._navigate_flowsheet(flowsheet_path)
        op = fs.Operations.Item(name)
        if not hasattr(op, parameter):
            raise AttributeError(
                f"Parameter {parameter!r} not found on {name} "
                f"(TypeName={getattr(op, 'TypeName', '?')})"
            )
        return getattr(op, parameter)

    def case_study(
        self,
        target_kind: str,
        target_name: str,
        target_field: str,
        sweep_values: list[float],
        observe: list[dict],
        target_unit: str | None = None,
        confirm: bool = False,
        flowsheet_path: str | None = None,
    ) -> CaseStudyResult:
        """変数を sweep_values で振って、各ケースで observe を観測する。

        target_kind: "stream" or "unit_op"
        target_field:
            stream の場合: 'temperature_C', 'pressure_kPa', 'molar_flow_kgmole_h',
                          'mass_flow_kg_h', 'vapor_fraction'
            unit_op の場合: HYSYS COM プロパティ名(例: 'Duty', 'PressureDrop')
        target_unit: unit_op のとき必須(例: 'kJ/h', 'kPa')
        observe: 各観測点。
            [{"kind":"stream","name":"Feed","field":"molar_flow_kgmole_h"},
             {"kind":"unit_op","name":"Cooler1","field":"Duty","unit":"kJ/h"}]
        confirm=False (デフォルト): ドライラン。実走せず変数だけ表示。
        confirm=True: 各 sweep 値で実際に書き込み→ソルバー実行→観測。
                     完了後、元の target 値に自動復元する (reverted=True)。
        """
        self._require_case()

        if target_kind == "stream":
            target_var, unit_resolved = self._resolve_stream_var(
                target_name, target_field, flowsheet_path
            )
            use_unit = unit_resolved
        elif target_kind == "unit_op":
            if target_unit is None:
                raise ValueError("target_unit is required when target_kind='unit_op'")
            target_var = self._resolve_op_var(
                target_name, target_field, target_unit, flowsheet_path
            )
            use_unit = target_unit
        else:
            raise ValueError(f"target_kind must be 'stream' or 'unit_op', got {target_kind!r}")

        original_value: float | None = None
        try:
            original_value = float(target_var.GetValue(use_unit))
        except Exception:
            original_value = None

        if not confirm:
            return CaseStudyResult(
                target_kind=target_kind,
                target_name=target_name,
                target_field=target_field,
                target_unit=use_unit,
                sweep_values=sweep_values,
                observe=observe,
                rows=[
                    CaseStudyRow(
                        target_value=v, observed={}, converged=False,
                        notes="dry-run (confirm=False)"
                    )
                    for v in sweep_values
                ],
                applied=False,
                reverted=False,
            )

        rows: list[CaseStudyRow] = []
        solver = self._case.Solver

        def observe_value(item: dict) -> tuple[str, Any]:
            kind = item.get("kind")
            nm = item["name"]
            fld = item["field"]
            unit_o = item.get("unit", "")
            key = f"{kind}:{nm}.{fld}"
            try:
                if kind == "stream":
                    var, u = self._resolve_stream_var(nm, fld, flowsheet_path)
                    return key, float(var.GetValue(u))
                elif kind == "unit_op":
                    var = self._resolve_op_var(nm, fld, unit_o, flowsheet_path)
                    return key, float(var.GetValue(unit_o))
                else:
                    return key, None
            except Exception as e:
                return key, f"ERR: {type(e).__name__}: {e}"

        for v in sweep_values:
            note_parts: list[str] = []
            converged = False
            observed_dict: dict = {}

            try:
                target_var.SetValue(float(v), use_unit)
            except Exception as e:
                note_parts.append(f"SetValue failed: {e}")
                rows.append(CaseStudyRow(
                    target_value=v, observed={}, converged=False,
                    notes="; ".join(note_parts),
                ))
                continue

            try:
                if hasattr(solver, "Run"):
                    solver.Run()
                elif hasattr(solver, "RunBoth"):
                    solver.RunBoth()
            except Exception as e:
                note_parts.append(f"Run error: {e}")

            try:
                status = self.get_solver_status()
                converged = not status.error_messages
                if status.error_messages:
                    note_parts.append(f"{len(status.error_messages)} unsolved")
            except Exception as e:
                note_parts.append(f"status error: {e}")

            for item in observe:
                k, val = observe_value(item)
                observed_dict[k] = val

            rows.append(CaseStudyRow(
                target_value=v,
                observed=observed_dict,
                converged=converged,
                notes="; ".join(note_parts),
            ))

        reverted = False
        if original_value is not None:
            try:
                target_var.SetValue(original_value, use_unit)
                try:
                    if hasattr(solver, "Run"):
                        solver.Run()
                    elif hasattr(solver, "RunBoth"):
                        solver.RunBoth()
                except Exception:
                    pass
                reverted = True
            except Exception:
                reverted = False

        return CaseStudyResult(
            target_kind=target_kind,
            target_name=target_name,
            target_field=target_field,
            target_unit=use_unit,
            sweep_values=sweep_values,
            observe=observe,
            rows=rows,
            applied=True,
            reverted=reverted,
        )

    # ─────────────────────────────────────────────────
    # Column Spec 操作
    # ─────────────────────────────────────────────────
    def _get_column_flowsheet(self, column_name: str, flowsheet_path: str | None):
        fs = self._navigate_flowsheet(flowsheet_path)
        col = fs.Operations.Item(column_name)
        if not hasattr(col, "ColumnFlowsheet"):
            raise ValueError(
                f"{column_name!r} は塔ではない (TypeName={getattr(col, 'TypeName', '?')})"
            )
        return col.ColumnFlowsheet

    def list_column_specs(
        self,
        column_name: str,
        flowsheet_path: str | None = None,
    ) -> list[ColumnSpecInfo]:
        """塔の全 Specifications を返す。"""
        cfs = self._get_column_flowsheet(column_name, flowsheet_path)
        result: list[ColumnSpecInfo] = []
        for sp in cfs.Specifications:
            try:
                goal = float(sp.GoalValue) if hasattr(sp, "GoalValue") else None
            except Exception:
                goal = None
            try:
                current = float(sp.CurrentValue) if hasattr(sp, "CurrentValue") else None
            except Exception:
                current = None
            try:
                err = float(sp.ErrorValue) if hasattr(sp, "ErrorValue") else None
            except Exception:
                err = None
            try:
                tol = float(sp.AbsoluteToleranceValue) if hasattr(sp, "AbsoluteToleranceValue") else None
            except Exception:
                tol = None
            result.append(ColumnSpecInfo(
                name=sp.name,
                is_active=bool(sp.IsActive) if hasattr(sp, "IsActive") else False,
                goal=goal,
                current=current,
                error=err,
                abs_tolerance=tol,
            ))
        return result

    def set_column_spec(
        self,
        column_name: str,
        spec_name: str,
        goal: float,
        unit: str = "",
        confirm: bool = False,
        flowsheet_path: str | None = None,
    ) -> SetResult:
        """Spec の Goal を変更。unit を指定すると Goal.SetValue(value, unit)、空文字なら GoalValue を直接代入。"""
        cfs = self._get_column_flowsheet(column_name, flowsheet_path)
        sp = cfs.Specifications.Item(spec_name)
        before: float | None
        try:
            before = float(sp.Goal.GetValue(unit)) if unit else float(sp.GoalValue)
        except Exception:
            before = None
        if confirm:
            if unit:
                sp.Goal.SetValue(goal, unit)
                after = float(sp.Goal.GetValue(unit))
            else:
                sp.GoalValue = goal
                after = float(sp.GoalValue)
        else:
            after = goal
        return SetResult(
            target=f"{column_name}/{spec_name}",
            field="Goal",
            before=before,
            after=after,
            applied=confirm,
            unit=unit or None,
        )

    def set_column_spec_active(
        self,
        column_name: str,
        spec_name: str,
        active: bool,
        confirm: bool = False,
        flowsheet_path: str | None = None,
    ) -> SetResult:
        """Spec の Active 状態を切り替える。塔の Degrees of Freedom が崩れる可能性に注意。"""
        cfs = self._get_column_flowsheet(column_name, flowsheet_path)
        sp = cfs.Specifications.Item(spec_name)
        before = bool(sp.IsActive) if hasattr(sp, "IsActive") else False
        if confirm:
            sp.IsActive = active
            after = bool(sp.IsActive)
        else:
            after = active
        return SetResult(
            target=f"{column_name}/{spec_name}",
            field="IsActive",
            before=before,
            after=after,
            applied=confirm,
            unit=None,
        )

    def get_column_spec_detail(
        self,
        column_name: str,
        spec_name: str,
        flowsheet_path: str | None = None,
    ) -> dict:
        """Spec の詳細 (TypeName, SpecifiedStage/Draw, Phase, FlowBasis, IncludedComponents 等) を返す。

        Spec種別に応じて利用可能な属性が変わる。存在しない属性は None で返す。
        """
        cfs = self._get_column_flowsheet(column_name, flowsheet_path)
        sp = cfs.Specifications.Item(spec_name)

        def _safe(getter):
            try:
                return getter()
            except Exception:
                return None

        type_name = _safe(lambda: str(sp.TypeName))
        stage = _safe(lambda: str(sp.SpecifiedStage.name))
        draw = _safe(lambda: str(sp.SpecifiedDraw.name))
        phase = _safe(lambda: int(sp.Phase))
        flow_basis = _safe(lambda: int(sp.FlowBasis))
        is_used_as_estimate = _safe(lambda: bool(sp.IsUsedAsEstimate))
        include_vapour = _safe(lambda: bool(sp.IncludeVapour))

        components: list[str] | None = None
        try:
            if hasattr(sp, "IncludedComponents"):
                inc = sp.IncludedComponents
                if hasattr(inc, "Count") and inc.Count > 0:
                    components = []
                    for i in range(inc.Count):
                        try:
                            c = inc.Item(i)
                            components.append(str(c.name))
                        except Exception:
                            components.append(f"<idx{i}>")
        except Exception:
            components = None

        return {
            "name": sp.name,
            "type_name": type_name,
            "is_active": bool(sp.IsActive) if hasattr(sp, "IsActive") else None,
            "goal": float(sp.GoalValue) if hasattr(sp, "GoalValue") else None,
            "current": float(sp.CurrentValue) if hasattr(sp, "CurrentValue") else None,
            "specified_stage": stage,
            "specified_draw": draw,
            "phase": phase,
            "flow_basis": flow_basis,
            "include_vapour": include_vapour,
            "is_used_as_estimate": is_used_as_estimate,
            "included_components": components,
        }

    def clear_column_estimates(
        self,
        column_name: str,
        scope: str = "all",
        confirm: bool = False,
        flowsheet_path: str | None = None,
    ) -> SetResult:
        """塔の推定値をクリア。

        scope: 'all' → ClearAllEstimates(), 'composition' → ClearAllCompositionEstimates(),
               'tray_composition' → ClearTrayCompositionEstimates()
        物性パッケージ切替後や前回解が新条件で使えない場合に呼ぶ。confirm=True 必須。
        """
        cfs = self._get_column_flowsheet(column_name, flowsheet_path)
        method_map = {
            "all": "ClearAllEstimates",
            "composition": "ClearAllCompositionEstimates",
            "tray_composition": "ClearTrayCompositionEstimates",
        }
        if scope not in method_map:
            raise ValueError(
                f"scope must be one of {list(method_map)}, got {scope!r}"
            )
        method_name = method_map[scope]
        if not hasattr(cfs, method_name):
            raise AttributeError(
                f"ColumnFlowsheet に {method_name}() が存在しない"
            )
        if confirm:
            getattr(cfs, method_name)()
        return SetResult(
            target=f"{column_name}/estimates",
            field=scope,
            before="(not cleared)",
            after="(cleared)" if confirm else "(would clear)",
            applied=confirm,
            unit=None,
        )

    _COLUMN_SOLVER_PARAMS = {
        "IsUsingSolutionForEstimates": bool,
        "MaximumIterations": int,
        "IsAdaptiveDamping": bool,
        "DampingFactor": float,
        "AdaptiveDampingPeriod": int,
        "EquilibriumErrorTolerance": float,
        "HeatSpecErrorTolerance": float,
        "TraceLevel": int,
    }

    def set_column_solver_param(
        self,
        column_name: str,
        parameter: str,
        value: float | int | bool,
        confirm: bool = False,
        flowsheet_path: str | None = None,
    ) -> SetResult:
        """塔ソルバーのパラメータを設定。

        parameter は ColumnFlowsheet の直接プロパティ名:
          - IsUsingSolutionForEstimates (bool): 前回解を初期推定に使うか
          - MaximumIterations (int): 最大反復回数
          - IsAdaptiveDamping (bool): 適応ダンピング
          - DampingFactor (float): ダンピング係数
          - AdaptiveDampingPeriod (int): 適応ダンピングの周期
          - EquilibriumErrorTolerance (float): 平衡誤差許容値
          - HeatSpecErrorTolerance (float): 熱バランス誤差許容値
          - TraceLevel (int): トレースレベル
        """
        cfs = self._get_column_flowsheet(column_name, flowsheet_path)
        if parameter not in self._COLUMN_SOLVER_PARAMS:
            raise ValueError(
                f"parameter must be one of {list(self._COLUMN_SOLVER_PARAMS)}, "
                f"got {parameter!r}"
            )
        expected_type = self._COLUMN_SOLVER_PARAMS[parameter]
        try:
            cast_value: Any = expected_type(value)
        except Exception as exc:
            raise ValueError(f"value cast to {expected_type.__name__} failed: {exc}")

        before = getattr(cfs, parameter, None)
        if confirm:
            setattr(cfs, parameter, cast_value)
            after = getattr(cfs, parameter)
        else:
            after = cast_value
        return SetResult(
            target=f"{column_name}/solver",
            field=parameter,
            before=before,
            after=after,
            applied=confirm,
            unit=None,
        )

    def remove_column_spec(
        self,
        column_name: str,
        spec_name: str,
        confirm: bool = False,
        flowsheet_path: str | None = None,
    ) -> SetResult:
        """塔の Specification を削除。

        Specifications.Remove(spec) を呼ぶ。confirm=True 必須。
        削除によって DOF が崩れる点に注意 (呼び出し側責任)。
        """
        cfs = self._get_column_flowsheet(column_name, flowsheet_path)
        sp = cfs.Specifications.Item(spec_name)
        # 壊れた (stage/draw/components 未確定の) spec は IsActive / GoalValue 取得時に
        # COM E_FAIL を返す。Remove 自体は通すため、スナップショットは best-effort。
        try:
            before_active = bool(sp.IsActive) if hasattr(sp, "IsActive") else None
        except Exception:
            before_active = None
        try:
            before_goal = float(sp.GoalValue) if hasattr(sp, "GoalValue") else None
        except Exception:
            before_goal = None
        if confirm:
            cfs.Specifications.Remove(sp)
            after = "(removed)"
        else:
            after = "(would remove)"
        return SetResult(
            target=f"{column_name}/{spec_name}",
            field="remove",
            before={"is_active": before_active, "goal": before_goal},
            after=after,
            applied=confirm,
            unit=None,
        )

    # HYSYS V14 の Specifications.Add(name, type_name) は内部 TypeName (lowercase, 短縮形) を要求。
    # フレンドリ名 "Column Reflux Ratio Spec" 等は通らない (COM E_FAIL)。
    # 既存 spec の TypeName 観察で確認済 (clmrefluxspec / clmcomprecoveryspec / clmcompfracspec / clmdrawspec)。
    # 他4種は推定 — 通らなければ既存 spec を introspect して実値で更新。
    _COLUMN_SPEC_TYPE_MAP = {
        "reflux_ratio": "clmrefluxspec",
        "comp_recovery": "clmcomprecoveryspec",
        "comp_fraction": "clmcompfracspec",
        "draw_rate": "clmdrawspec",
        "temperature": "clmtempspec",
        "duty": "clmdutyspec",
        "vapour_fraction": "clmvapfracspec",
        "pressure": "clmpressurespec",
    }

    def add_column_spec(
        self,
        column_name: str,
        spec_type: str,
        spec_name: str,
        goal: float | None = None,
        stage: str | None = None,
        draw: str | None = None,
        phase: int | None = None,
        components: list[str] | None = None,
        is_active: bool = True,
        flow_basis: int | None = None,
        confirm: bool = False,
        flowsheet_path: str | None = None,
    ) -> dict:
        """塔に新規 Spec を追加。

        spec_type 候補: reflux_ratio / comp_recovery / comp_fraction / draw_rate /
                       temperature / duty / vapour_fraction / pressure

        - stage: 'Condenser' / 'Reboiler' / '1_Main TS' などの段名 (SpecifiedStage)
        - draw: 'R' / 'C' / 留出名 (SpecifiedDraw)
        - phase: 0/1/2 (相 0=Liquid)
        - components: 含める成分名のリスト (Comp Recovery/Fraction系)
        - flow_basis: 流量基準コード (Draw Rate系: 0=Mass/h, 3=Mole 等、HYSYS V14依存)
        """
        cfs = self._get_column_flowsheet(column_name, flowsheet_path)
        if spec_type not in self._COLUMN_SPEC_TYPE_MAP:
            raise ValueError(
                f"spec_type must be one of {list(self._COLUMN_SPEC_TYPE_MAP)}, "
                f"got {spec_type!r}"
            )
        hysys_type_name = self._COLUMN_SPEC_TYPE_MAP[spec_type]

        if not confirm:
            return {
                "applied": False,
                "would_add": {
                    "column": column_name,
                    "spec_type": spec_type,
                    "hysys_type": hysys_type_name,
                    "name": spec_name,
                    "goal": goal,
                    "stage": stage,
                    "draw": draw,
                    "phase": phase,
                    "components": components,
                    "is_active": is_active,
                    "flow_basis": flow_basis,
                },
            }

        sp = cfs.Specifications.Add(spec_name, hysys_type_name)
        applied_fields: dict = {"name": spec_name, "hysys_type": hysys_type_name}

        # 順序が重要: stage/draw/phase/components が確定する前に GoalValue を触ると
        # Comp Fraction Spec 等で COM E_FAIL になる (値の意味が決まらないため)。
        # 全段 try/except で保護し、エラーは applied_fields に *_error として残す。
        if stage is not None and hasattr(sp, "SpecifiedStageVar"):
            try:
                stage_obj = cfs.ColumnStages.Item(stage)
                sp.SpecifiedStage = stage_obj
                applied_fields["stage"] = stage
            except Exception as exc:
                applied_fields["stage_error"] = str(exc)

        if draw is not None and hasattr(sp, "SpecifiedDrawVar"):
            try:
                for prod_coll_name in ("LiquidProducts", "VapourProducts"):
                    prods = getattr(cfs, prod_coll_name, None)
                    if prods is None:
                        continue
                    try:
                        draw_obj = prods.Item(draw)
                        sp.SpecifiedDraw = draw_obj
                        applied_fields["draw"] = draw
                        break
                    except Exception:
                        continue
            except Exception as exc:
                applied_fields["draw_error"] = str(exc)

        if phase is not None and hasattr(sp, "Phase"):
            try:
                sp.Phase = int(phase)
                applied_fields["phase"] = int(sp.Phase)
            except Exception as exc:
                applied_fields["phase_error"] = str(exc)

        if components is not None and hasattr(sp, "IncludedComponents"):
            try:
                inc = sp.IncludedComponents
                if hasattr(inc, "RemoveAll"):
                    inc.RemoveAll()
                fp = cfs.FluidPackage
                for comp_name in components:
                    comp = fp.Components.Item(comp_name)
                    inc.Add(comp)
                applied_fields["components"] = list(components)
            except Exception as exc:
                applied_fields["components_error"] = str(exc)

        if flow_basis is not None and hasattr(sp, "FlowBasis"):
            try:
                sp.FlowBasis = int(flow_basis)
                applied_fields["flow_basis"] = int(sp.FlowBasis)
            except Exception as exc:
                applied_fields["flow_basis_error"] = str(exc)

        if goal is not None and hasattr(sp, "GoalValue"):
            try:
                sp.GoalValue = float(goal)
                applied_fields["goal"] = float(sp.GoalValue)
            except Exception as exc:
                applied_fields["goal_error"] = str(exc)

        if hasattr(sp, "IsActive"):
            try:
                sp.IsActive = bool(is_active)
                applied_fields["is_active"] = bool(sp.IsActive)
            except Exception as exc:
                applied_fields["is_active_error"] = str(exc)

        return {"applied": True, "spec": applied_fields}

    def save_as(self, path: str, confirm: bool = False) -> dict:
        """別名保存(SaveAs)。既存ファイル上書きは禁止。

        confirm=False はドライラン(保存先パスと衝突確認のみ)。
        """
        import os
        self._require_case()
        if os.path.exists(path):
            raise FileExistsError(
                f"File already exists: {path}. 上書き禁止。別のパスを指定してください。"
            )
        if not confirm:
            return {
                "applied": False,
                "would_save_to": path,
                "message": "Set confirm=True to actually save.",
            }
        self._case.SaveAs(path)
        return {"applied": True, "saved_to": path}

    # ─────────────────────────────────────────────────
    # Logical Operations (Adjust / Set / Recycle / Spreadsheet)
    # ─────────────────────────────────────────────────
    _LOGICAL_TYPES = {
        "adjust", "adjustop", "set", "setop", "recycle", "recycleop",
        "spreadsheet", "spreadsheetop", "balance", "balanceop",
    }

    def _logical_op_details(self, op: Any) -> dict:
        tn = getattr(op, "TypeName", "").lower()
        details: dict = {}

        def safe(name: str) -> Any:
            if not hasattr(op, name):
                return None
            try:
                v = getattr(op, name)
                if hasattr(v, "name"):
                    return v.name
                if hasattr(v, "GetValue"):
                    try:
                        return float(v.GetValue(""))
                    except Exception:
                        return str(v)
                if isinstance(v, (int, float, str, bool)):
                    return v
                return str(v)
            except Exception:
                return None

        if "adjust" in tn:
            for prop in ("TargetValue", "Tolerance", "StepSize", "MaxValue", "MinValue",
                         "IsActive", "SourceObject", "TargetObject", "SolutionMethod",
                         "AdjustedVariable", "TargetVariable"):
                val = safe(prop)
                if val is not None:
                    details[prop] = val
        elif "recycle" in tn:
            for prop in ("IsConverged", "CurrentIteration", "IterationsCompleted",
                         "MaximumIterations", "AbsoluteTolerance", "RelativeTolerance"):
                val = safe(prop)
                if val is not None:
                    details[prop] = val
        elif tn.startswith("set"):
            for prop in ("SourceObject", "TargetObject", "Multiplier", "Offset",
                         "SourceVariable", "TargetVariable"):
                val = safe(prop)
                if val is not None:
                    details[prop] = val
        elif "spreadsheet" in tn:
            for prop in ("NumberOfRows", "NumberOfColumns"):
                val = safe(prop)
                if val is not None:
                    details[prop] = val
        return details

    def list_logical_ops(self, flowsheet_path: str | None = None) -> list[LogicalOpInfo]:
        """Adjust/Set/Recycle/Spreadsheet 系 Logical Operation 全列挙。"""
        fs = self._navigate_flowsheet(flowsheet_path)
        result: list[LogicalOpInfo] = []
        for op in fs.Operations:
            tn = getattr(op, "TypeName", "").lower()
            if any(t in tn for t in ("adjust", "recycle", "spreadsheet", "balance")) or tn.startswith("set"):
                result.append(LogicalOpInfo(
                    name=op.name,
                    type_name=tn,
                    is_active=not bool(getattr(op, "IsIgnored", False)),
                    details=self._logical_op_details(op),
                ))
        return result

    def get_logical_op(self, name: str, flowsheet_path: str | None = None) -> LogicalOpInfo:
        fs = self._navigate_flowsheet(flowsheet_path)
        op = fs.Operations.Item(name)
        return LogicalOpInfo(
            name=op.name,
            type_name=getattr(op, "TypeName", "").lower(),
            is_active=not bool(getattr(op, "IsIgnored", False)),
            details=self._logical_op_details(op),
        )

    def set_adjust_target(
        self,
        name: str,
        target_value: float | None = None,
        tolerance: float | None = None,
        step_size: float | None = None,
        max_value: float | None = None,
        min_value: float | None = None,
        confirm: bool = False,
        flowsheet_path: str | None = None,
    ) -> list[SetResult]:
        """Adjust ブロックの target_value / tolerance / step / min / max を変更。"""
        fs = self._navigate_flowsheet(flowsheet_path)
        op = fs.Operations.Item(name)
        tn = getattr(op, "TypeName", "").lower()
        if "adjust" not in tn:
            raise ValueError(f"{name!r} is not an Adjust op (TypeName={tn!r})")

        results: list[SetResult] = []
        for prop, val in [
            ("TargetValue", target_value),
            ("Tolerance", tolerance),
            ("StepSize", step_size),
            ("MaxValue", max_value),
            ("MinValue", min_value),
        ]:
            if val is None:
                continue
            if not hasattr(op, prop):
                results.append(SetResult(name, prop, None, val, False, None))
                continue
            cur = getattr(op, prop)
            before = None
            if hasattr(cur, "GetValue"):
                try:
                    before = float(cur.GetValue(""))
                except Exception:
                    pass
                if confirm:
                    cur.SetValue(float(val), "")
            else:
                before = cur
                if confirm:
                    setattr(op, prop, val)
            results.append(SetResult(name, prop, before, val, confirm, None))
        return results

    def reset_recycle(
        self,
        name: str,
        confirm: bool = False,
        flowsheet_path: str | None = None,
    ) -> dict:
        """Recycle ブロックを Reset する。confirm=True 必須。"""
        fs = self._navigate_flowsheet(flowsheet_path)
        op = fs.Operations.Item(name)
        tn = getattr(op, "TypeName", "").lower()
        if "recycle" not in tn:
            raise ValueError(f"{name!r} is not a Recycle op (TypeName={tn!r})")
        if not confirm:
            return {"applied": False, "message": "Set confirm=True to actually reset."}
        if hasattr(op, "Reset"):
            op.Reset()
            return {"applied": True, "message": f"Recycle {name} reset."}
        raise RuntimeError("Reset method not found on Recycle op")

    # ─────────────────────────────────────────────────
    # Fluid Package / Components
    # ─────────────────────────────────────────────────
    def get_fluid_package(self) -> FluidPackageInfo:
        self._require_case()
        fp = self._case.Flowsheet.FluidPackage
        comps = [c.name for c in fp.Components]
        pp_name = ""
        try:
            pp_name = str(fp.PropertyPackageName)
        except Exception:
            try:
                pp_name = str(fp.PropertyPackage.name)
            except Exception:
                pass
        has_rxn = False
        try:
            has_rxn = bool(fp.ReactionPackage is not None)
        except Exception:
            pass
        return FluidPackageInfo(
            name=getattr(fp, "name", "Basis"),
            property_package_name=pp_name,
            component_count=len(comps),
            components=comps,
            has_reaction_package=has_rxn,
        )

    def list_components(self) -> list[str]:
        self._require_case()
        return [c.name for c in self._case.Flowsheet.FluidPackage.Components]

    def get_component(self, name: str) -> ComponentInfo:
        """コンポーネント単体の主要物性を返す。HYSYS V14 では *Value (内部単位 float) が安定。"""
        self._require_case()
        c = self._case.Flowsheet.FluidPackage.Components.Item(name)

        def safe_float(attr: str) -> float | None:
            if not hasattr(c, attr):
                return None
            try:
                v = getattr(c, attr)
                return float(v)
            except Exception:
                return None

        def safe_var(attr: str, unit: str) -> float | None:
            if not hasattr(c, attr):
                return None
            try:
                v = getattr(c, attr)
                if hasattr(v, "GetValue"):
                    return float(v.GetValue(unit))
                return float(v)
            except Exception:
                return None

        cas = None
        try:
            cas_raw = getattr(c, "CAS_Number", None)
            if cas_raw is not None:
                cas = str(cas_raw)
        except Exception:
            cas = None

        mw = safe_float("MolecularWeightValue") or safe_var("MolecularWeight", "")
        nbp = safe_var("NormalBoilingPoint", "C")
        if nbp is None:
            nbp_k = safe_float("NormalBoilingPointValue")
            nbp = (nbp_k - 273.15) if nbp_k is not None else None
        tc = safe_var("CriticalTemperature", "C")
        if tc is None:
            tc_k = safe_float("CriticalTemperatureValue")
            tc = (tc_k - 273.15) if tc_k is not None else None
        pc = safe_var("CriticalPressure", "kPa")
        if pc is None:
            pc_pa = safe_float("CriticalPressureValue")
            pc = (pc_pa / 1000.0) if pc_pa is not None else None
        ac = safe_float("AcentricityValue") or safe_var("Acentricity", "")

        return ComponentInfo(
            name=c.name,
            cas_number=cas,
            molecular_weight=mw,
            normal_boiling_point_C=nbp,
            critical_temperature_C=tc,
            critical_pressure_kPa=pc,
            acentricity=ac,
        )

    # ─────────────────────────────────────────────────
    # Internal Stream 作成 (段別流量取得用)
    # ─────────────────────────────────────────────────
    def add_internal_stream(
        self,
        column_name: str,
        stage: int,
        kind: str = "Liquid",
        stream_name: str | None = None,
        confirm: bool = False,
        flowsheet_path: str | None = None,
    ) -> dict:
        """塔の Tray Section に内部 stream (Liquid/Vapor draw) を追加する。

        kind: 'Liquid' | 'Vapor' (大文字小文字どちらでも)。
        confirm=True 必須 (破壊的)。
        """
        cfs = self._get_column_flowsheet(column_name, flowsheet_path)
        ts = None
        for op in cfs.Operations:
            if "tray" in getattr(op, "TypeName", "").lower():
                ts = op
                break
        if ts is None:
            raise RuntimeError(
                f"Tray section が見つからない (column={column_name})"
            )
        kind_norm = kind.capitalize()
        if kind_norm not in ("Liquid", "Vapor", "Vapour"):
            raise ValueError(f"kind must be 'Liquid' or 'Vapor', got {kind!r}")
        if kind_norm == "Vapor":
            kind_norm = "Vapour"
        if stream_name is None:
            stream_name = f"{column_name}_S{stage}_{kind_norm}"

        if not confirm:
            return {
                "applied": False,
                "would_add": stream_name,
                "on_column": column_name,
                "tray_section": ts.name,
                "stage": stage,
                "kind": kind_norm,
                "message": "Set confirm=True to actually add the internal stream.",
            }

        errors: list[str] = []
        try:
            ts.AddDrawStream(int(stage), kind_norm, stream_name)
            return {"applied": True, "added_stream": stream_name,
                    "stage": stage, "kind": kind_norm}
        except Exception as e:
            errors.append(f"AddDrawStream(stage,kind,name): {e}")
        try:
            ts.AddDrawStream(stream_name, int(stage), kind_norm)
            return {"applied": True, "added_stream": stream_name,
                    "stage": stage, "kind": kind_norm}
        except Exception as e:
            errors.append(f"AddDrawStream(name,stage,kind): {e}")
        raise RuntimeError(
            f"AddDrawStream failed for all signatures. errors: {errors}"
        )

    # ─────────────────────────────────────────────────
    # Stream / Op 検索
    # ─────────────────────────────────────────────────
    def find_streams(
        self,
        contains_components: list[str] | None = None,
        min_mol_fraction: float | None = None,
        vapor_fraction_min: float | None = None,
        vapor_fraction_max: float | None = None,
        molar_flow_min_kgmole_h: float | None = None,
        name_pattern: str | None = None,
        flowsheet_path: str | None = None,
    ) -> list[dict]:
        """条件に合うストリームを検索。

        contains_components + min_mol_fraction: 指定成分のモル分率が下限以上のもの。
        name_pattern: 正規表現 (大文字小文字無視)。
        """
        import re
        fs = self._navigate_flowsheet(flowsheet_path)
        fp = self._case.Flowsheet.FluidPackage
        comp_names = [c.name for c in fp.Components]
        pattern = re.compile(name_pattern, re.IGNORECASE) if name_pattern else None
        result: list[dict] = []
        for s in fs.MaterialStreams:
            if pattern and not pattern.search(s.name):
                continue
            vf = None
            try:
                vf = float(s.VapourFractionValue)
            except Exception:
                pass
            if vapor_fraction_min is not None and (vf is None or vf < vapor_fraction_min):
                continue
            if vapor_fraction_max is not None and (vf is None or vf > vapor_fraction_max):
                continue

            mf = None
            try:
                mf = float(s.MolarFlow.GetValue("kgmole/h"))
            except Exception:
                pass
            if molar_flow_min_kgmole_h is not None and (mf is None or mf < molar_flow_min_kgmole_h):
                continue

            matched_components: dict[str, float] = {}
            if contains_components:
                try:
                    fracs = list(s.ComponentMolarFractionValue)
                except Exception:
                    continue
                ok = True
                for cname in contains_components:
                    if cname not in comp_names:
                        ok = False
                        break
                    idx = comp_names.index(cname)
                    if idx >= len(fracs):
                        ok = False
                        break
                    frac = float(fracs[idx])
                    if min_mol_fraction is not None and frac < min_mol_fraction:
                        ok = False
                        break
                    if min_mol_fraction is None and frac <= 0:
                        ok = False
                        break
                    matched_components[cname] = round(frac, 6)
                if not ok:
                    continue
            result.append({
                "name": s.name,
                "vapor_fraction": vf,
                "molar_flow_kgmole_h": mf,
                **({"matched_components": matched_components} if matched_components else {}),
            })
        return result

    def find_ops(
        self,
        type_name: str | None = None,
        name_pattern: str | None = None,
        flowsheet_path: str | None = None,
    ) -> list[UnitOpInfo]:
        """装置を type_name / name 正規表現で検索。"""
        import re
        fs = self._navigate_flowsheet(flowsheet_path)
        pattern = re.compile(name_pattern, re.IGNORECASE) if name_pattern else None
        result: list[UnitOpInfo] = []
        for op in fs.Operations:
            tn = getattr(op, "TypeName", "")
            if type_name and tn != type_name:
                continue
            if pattern and not pattern.search(op.name):
                continue
            result.append(UnitOpInfo(
                name=op.name,
                type_name=tn or "Unknown",
                flowsheet=flowsheet_path or "Main",
                solve_complete=bool(getattr(op, "SolveComplete", False)),
            ))
        return result

    # ─────────────────────────────────────────────────
    # Generic COM method invocation (Phase 8)
    # ─────────────────────────────────────────────────
    def call_method(
        self,
        path: str,
        method: str,
        args: list | None = None,
        confirm: bool = False,
    ) -> dict:
        """任意 COM オブジェクトのメソッドを呼ぶ汎用ツール (破壊的のため confirm=True 必須)。

        path は introspect と同じく 'case'/'flowsheet'/'fp'/'app' 起点のドット式。
        method はそのオブジェクト上のメソッド名 (大文字小文字区別)。
        args は 順序付き引数リスト。各要素は dict:
          - {"type": "literal", "value": <str|int|float|bool|null>}
          - {"type": "path", "value": "<path expression>"}  → 内部で eval されて COM オブジェクト解決

        例 (反応セットを Fluid Package にアタッチ):
            call_method(
                path="case.BasisManager.ReactionPackageManager.ReactionSets.Item('Set-1')",
                method="AssociateFluidPackage",
                args=[{"type": "path", "value": "case.BasisManager.FluidPackages.Item(1)"}],
                confirm=True,
            )
        """
        self._require_case()
        if any(bad in path for bad in ("__", "import ", "exec(", "eval(", ";", "open(")):
            raise ValueError(f"Forbidden tokens in path: {path!r}")
        if any(bad in method for bad in ("__", " ", "(", ")", ";")):
            raise ValueError(f"Invalid method name: {method!r}")

        namespace = {
            "app": self._app,
            "case": self._case,
            "flowsheet": self._case.Flowsheet,
            "fp": self._case.Flowsheet.FluidPackage,
        }

        try:
            obj = eval(path, {"__builtins__": {}}, namespace)
        except Exception as e:
            raise RuntimeError(f"Path eval failed for {path!r}: {e}")

        if not hasattr(obj, method):
            raise AttributeError(f"{path} has no method {method!r}")
        fn = getattr(obj, method)
        if not callable(fn):
            raise TypeError(f"{path}.{method} is not callable (type={type(fn).__name__})")

        resolved_args = []
        for i, arg in enumerate(args or []):
            if not isinstance(arg, dict) or "type" not in arg:
                raise ValueError(f"args[{i}] must be {{type, value}} dict")
            t = arg["type"]
            v = arg.get("value")
            if t == "literal":
                resolved_args.append(v)
            elif t == "path":
                if any(bad in str(v) for bad in ("__", "import ", "exec(", "eval(", ";", "open(")):
                    raise ValueError(f"Forbidden tokens in args[{i}].value: {v!r}")
                try:
                    resolved_args.append(eval(v, {"__builtins__": {}}, namespace))
                except Exception as e:
                    raise RuntimeError(f"args[{i}] path eval failed for {v!r}: {e}")
            else:
                raise ValueError(f"args[{i}].type must be 'literal' or 'path', got {t!r}")

        if not confirm:
            return {
                "applied": False,
                "would_call": f"{path}.{method}",
                "arg_count": len(resolved_args),
                "message": "Set confirm=True to actually invoke this COM method.",
            }

        try:
            result = fn(*resolved_args)
        except Exception as e:
            return {
                "applied": False,
                "called": f"{path}.{method}",
                "arg_count": len(resolved_args),
                "error": f"{type(e).__name__}: {e}",
            }

        # Try to summarize the return value.
        rt = type(result).__name__
        preview: str
        if result is None:
            preview = "None"
        elif rt in ("int", "float", "str", "bool"):
            preview = repr(result)
        elif rt == "CDispatch":
            try:
                preview = f"CDispatch(name={getattr(result, 'name', '?')!r})"
            except Exception:
                preview = "CDispatch"
        else:
            preview = f"<{rt}>"

        return {
            "applied": True,
            "called": f"{path}.{method}",
            "arg_count": len(resolved_args),
            "result_type": rt,
            "result_preview": preview,
        }

    # ─────────────────────────────────────────────────
    # Generic COM property writer (Phase 9)
    # ─────────────────────────────────────────────────
    def set_property(
        self,
        path: str,
        property_name: str,
        value,
        unit: str | None = None,
        confirm: bool = False,
    ) -> dict:
        """任意 COM オブジェクトのプロパティを書き換える汎用ツール。

        以下3通りの書込モードを自動分岐する:
        1. 取得した property が CDispatch かつ ``SetValue`` を持つ場合 → ``attr.SetValue(value [, unit])``
           (HYSYS の Temperature/Pressure 等の Variable オブジェクト)
        2. 取得した property が CDispatch かつ ``SetValues`` を持ち value が list/tuple の場合 → ``attr.SetValues(value)``
           (TemperatureEsts 等の配列 Variable)
        3. それ以外 → ``setattr(obj, property_name, value)``
           (PropPkgName, ColumnAlgorithm, MaximumIterations 等の素プロパティ)

        confirm=False のときは dry-run で旧値と method 候補を返す。
        confirm=True 必須 (破壊的)。

        path は call_method/introspect と同じく 'case'/'flowsheet'/'fp'/'app' 起点のドット式。

        例:
            # 物性パッケージ切替
            set_property("case.BasisManager.FluidPackages.Item('Basis-1')",
                         "PropPkgName", "Acid Gas - Chemical Solvents", confirm=True)
            # 塔のアルゴリズム切替
            set_property("flowsheet.Operations.Item('tower').ColumnFlowsheet",
                         "ColumnAlgorithm", 2, confirm=True)
            # 段別温度推定値の投入 (単一段)
            set_property("flowsheet.Operations.Item('tower').ColumnFlowsheet.TemperatureEsts.Item(0)",
                         "(SetValue)", 90.0, unit="C", confirm=True)
            # 段別温度推定値の一括投入
            set_property("flowsheet.Operations.Item('tower').ColumnFlowsheet",
                         "TemperatureEsts", [96.0, 99.0, ...], confirm=True)
        """
        self._require_case()
        if any(bad in path for bad in ("__", "import ", "exec(", "eval(", ";", "open(")):
            raise ValueError(f"Forbidden tokens in path: {path!r}")
        if any(bad in property_name for bad in ("__", " ", "(", ")", ";")):
            raise ValueError(f"Invalid property name: {property_name!r}")

        namespace = {
            "app": self._app,
            "case": self._case,
            "flowsheet": self._case.Flowsheet,
            "fp": self._case.Flowsheet.FluidPackage,
        }

        try:
            obj = eval(path, {"__builtins__": {}}, namespace)
        except Exception as e:
            raise RuntimeError(f"Path eval failed for {path!r}: {e}")

        if not hasattr(obj, property_name):
            raise AttributeError(f"{path} has no property {property_name!r}")

        old_raw = getattr(obj, property_name)
        old_preview = _summarize_value(old_raw)

        # 書込モードを決定
        mode: str
        if hasattr(old_raw, "SetValue") and callable(getattr(old_raw, "SetValue", None)):
            mode = "SetValue"
        elif (
            isinstance(value, (list, tuple))
            and hasattr(old_raw, "SetValues")
            and callable(getattr(old_raw, "SetValues", None))
        ):
            mode = "SetValues"
        else:
            mode = "setattr"

        if not confirm:
            return {
                "applied": False,
                "target": f"{path}.{property_name}",
                "old_value": old_preview,
                "proposed_value": value,
                "proposed_method": mode,
                "proposed_unit": unit,
                "message": "Set confirm=True to actually write this property.",
            }

        try:
            if mode == "SetValue":
                if unit is not None:
                    old_raw.SetValue(value, unit)
                else:
                    old_raw.SetValue(value)
            elif mode == "SetValues":
                old_raw.SetValues(list(value))
            else:  # setattr
                setattr(obj, property_name, value)
        except Exception as e:
            return {
                "applied": False,
                "target": f"{path}.{property_name}",
                "method_attempted": mode,
                "unit": unit,
                "old_value": old_preview,
                "attempted_value": value,
                "error": f"{type(e).__name__}: {e}",
            }

        # 新値を再読
        try:
            new_raw = getattr(obj, property_name)
            new_preview = _summarize_value(new_raw)
        except Exception as e:
            new_preview = f"(re-read failed: {e})"

        return {
            "applied": True,
            "target": f"{path}.{property_name}",
            "method_used": mode,
            "unit": unit,
            "old_value": old_preview,
            "new_value": new_preview,
        }

    # ─────────────────────────────────────────────────
    # Introspect (任意 COM オブジェクトの dir 取得)
    # ─────────────────────────────────────────────────
    def introspect(
        self,
        path: str,
        filter_keyword: str | None = None,
        max_members: int = 200,
    ) -> dict:
        """安全な範囲で COM オブジェクトのメンバを列挙。

        path は 'case', 'flowsheet', 'fp', 'app' を起点としたドット式。
        例: 'flowsheet.Operations.Item(\"Cooler1\")',
            'flowsheet.MaterialStreams.Item(\"stream7\")',
            'fp.Components.Item(0)'
        """
        self._require_case()
        if any(bad in path for bad in ("__", "import ", "exec(", "eval(", ";", "open(")):
            raise ValueError(f"Forbidden tokens in path: {path!r}")

        namespace = {
            "app": self._app,
            "case": self._case,
            "flowsheet": self._case.Flowsheet,
            "fp": self._case.Flowsheet.FluidPackage,
        }
        try:
            obj = eval(path, {"__builtins__": {}}, namespace)
        except Exception as e:
            raise RuntimeError(f"Path eval failed for {path!r}: {e}")

        BLACKLIST = {"AddRef", "Application", "GetIDsOfNames", "GetTypeInfo",
                     "GetTypeInfoCount", "Invoke", "QueryInterface", "Release",
                     "Moniker", "Parent"}
        members: list[dict] = []
        for a in sorted(dir(obj)):
            if a.startswith("_") or a in BLACKLIST:
                continue
            if filter_keyword and filter_keyword.lower() not in a.lower():
                continue
            entry = {"name": a}
            try:
                v = getattr(obj, a)
                t = type(v).__name__
                entry["type"] = t
                if t == "float":
                    entry["preview"] = f"= {v}"
                elif t in ("int", "str", "bool"):
                    entry["preview"] = f"= {v!r}"
                elif t == "tuple":
                    entry["preview"] = f"(tuple len={len(v)})"
                elif t == "CDispatch":
                    try:
                        if hasattr(v, "name"):
                            entry["preview"] = f"name={v.name!r}"
                        elif hasattr(v, "GetValue"):
                            try:
                                entry["preview"] = f"GetValue('')={v.GetValue('')}"
                            except Exception:
                                entry["preview"] = "(CDispatch)"
                        else:
                            entry["preview"] = "(CDispatch)"
                    except Exception:
                        entry["preview"] = "(CDispatch, no name)"
                elif callable(v):
                    entry["preview"] = "(method)"
            except Exception as e:
                entry["type"] = "?"
                entry["preview"] = f"(access error: {e})"
            members.append(entry)
            if len(members) >= max_members:
                members.append({"name": "...", "type": "...", "preview": f"truncated at {max_members}"})
                break

        return {
            "path": path,
            "object_type": type(obj).__name__,
            "member_count": len(members),
            "members": members,
        }

    # ─────────────────────────────────────────────────
    # Heat Operation (Cooler/Heater/HeatExchanger 汎用)
    # ─────────────────────────────────────────────────
    def get_heat_op(self, name: str, flowsheet_path: str | None = None) -> HeatOpInfo:
        """Cooler/Heater/HeatExchanger の主要パラメータをまとめて返す。"""
        fs = self._navigate_flowsheet(flowsheet_path)
        op = fs.Operations.Item(name)
        tn = getattr(op, "TypeName", "").lower()

        def stream_name(attr: str) -> str | None:
            try:
                v = getattr(op, attr, None)
                if v is None:
                    return None
                if hasattr(v, "name"):
                    return v.name
            except Exception:
                pass
            return None

        def rv(attr: str, unit: str) -> float | None:
            if not hasattr(op, attr):
                return None
            try:
                v = getattr(op, attr)
                if hasattr(v, "GetValue"):
                    return float(v.GetValue(unit))
            except Exception:
                return None
            return None

        info = HeatOpInfo(
            name=op.name,
            type_name=tn,
            duty_default_units=rv("Duty", "") if hasattr(op, "Duty") else rv("Energy", ""),
            duty_kW=rv("Duty", "kW") if hasattr(op, "Duty") else rv("Energy", "kW"),
            feed_stream=stream_name("FeedStream"),
            product_stream=stream_name("ProductStream"),
            energy_stream=stream_name("EnergyStream"),
            feed_T_C=rv("FeedTemperature", "C"),
            feed_P_kPa=rv("FeedPressure", "kPa"),
            product_T_C=rv("ProductTemperature", "C"),
            product_P_kPa=rv("ProductPressure", "kPa"),
            pressure_drop_kPa=rv("PressureDrop", "kPa"),
            ua_kJ_C_h=rv("UA", "kJ/C-h"),
            lmtd_C=rv("LMTD", "C"),
            min_approach_C=rv("MinApproach", "C"),
            hot_inlet=None, hot_outlet=None, cold_inlet=None, cold_outlet=None,
        )
        if "exchanger" in tn or "hx" in tn:
            mapping = [
                ("TubeSideFeedStream", "cold_inlet"),
                ("TubeSideProductStream", "cold_outlet"),
                ("ShellSideFeedStream", "hot_inlet"),
                ("ShellSideProductStream", "hot_outlet"),
                ("HotInlet", "hot_inlet"),
                ("HotOutlet", "hot_outlet"),
                ("ColdInlet", "cold_inlet"),
                ("ColdOutlet", "cold_outlet"),
            ]
            for attr, slot in mapping:
                val = stream_name(attr)
                if val and getattr(info, slot) is None:
                    setattr(info, slot, val)
        return info

    # ─────────────────────────────────────────────────
    # Material / Energy Balance
    # ─────────────────────────────────────────────────
    def balance_check(
        self,
        op_names: list[str] | None = None,
        flowsheet_path: str | None = None,
    ) -> BalanceSummary:
        """指定 ops (None なら全 ops) の境界ストリームで物質収支・熱収支を計算。

        境界 = 集合内の op の Feed/Product から、集合内の op 同士で互いに繋がるストリーム
        (= 内部ストリーム) を除いたもの。
        """
        fs = self._navigate_flowsheet(flowsheet_path)
        all_ops = list(fs.Operations)
        if op_names is None:
            target_ops = all_ops
            scope = f"flowsheet (all {len(all_ops)} ops)"
        else:
            target_ops = [op for op in all_ops if op.name in set(op_names)]
            scope = f"ops:{','.join(op_names[:5])}{'…' if len(op_names) > 5 else ''}"

        all_feeds: dict[str, set[str]] = {}
        all_prods: dict[str, set[str]] = {}
        energy_feeds: dict[str, set[str]] = {}
        energy_prods: dict[str, set[str]] = {}

        def is_energy(stream_obj) -> bool:
            t = getattr(stream_obj, "TypeName", "").lower()
            if "energy" in t:
                return True
            if not hasattr(stream_obj, "MassFlow"):
                return True
            return False

        for op in target_ops:
            feeds_set = set()
            prods_set = set()
            efeeds = set()
            eprods = set()
            try:
                for f in op.AttachedFeeds:
                    try:
                        nm = f.name
                        if is_energy(f):
                            efeeds.add(nm)
                        else:
                            feeds_set.add(nm)
                    except Exception:
                        continue
            except Exception:
                pass
            try:
                for p in op.AttachedProducts:
                    try:
                        nm = p.name
                        if is_energy(p):
                            eprods.add(nm)
                        else:
                            prods_set.add(nm)
                    except Exception:
                        continue
            except Exception:
                pass
            all_feeds[op.name] = feeds_set
            all_prods[op.name] = prods_set
            energy_feeds[op.name] = efeeds
            energy_prods[op.name] = eprods

        feeds_total = set().union(*all_feeds.values()) if all_feeds else set()
        prods_total = set().union(*all_prods.values()) if all_prods else set()
        internal = feeds_total & prods_total
        boundary_in = feeds_total - internal
        boundary_out = prods_total - internal

        efeeds_total = set().union(*energy_feeds.values()) if energy_feeds else set()
        eprods_total = set().union(*energy_prods.values()) if energy_prods else set()
        einternal = efeeds_total & eprods_total
        eboundary_in = efeeds_total - einternal
        eboundary_out = eprods_total - einternal

        def mass_kg_h(stream_name_: str) -> float:
            try:
                s = fs.MaterialStreams.Item(stream_name_)
                return float(s.MassFlow.GetValue("kg/h"))
            except Exception:
                return 0.0

        def stream_enthalpy_kJ_h(stream_name_: str) -> float | None:
            try:
                s = fs.MaterialStreams.Item(stream_name_)
                return float(s.HeatFlow.GetValue("kJ/h"))
            except Exception:
                return None

        def energy_stream_duty_kJ_h(stream_name_: str) -> float | None:
            for op in target_ops:
                try:
                    es = getattr(op, "EnergyStream", None)
                    if es is not None and getattr(es, "name", None) == stream_name_:
                        if hasattr(op, "Duty"):
                            return float(op.Duty.GetValue("kJ/h"))
                        if hasattr(op, "Energy"):
                            return float(op.Energy.GetValue("kJ/h"))
                except Exception:
                    continue
            return None

        total_in = sum(mass_kg_h(s) for s in boundary_in)
        total_out = sum(mass_kg_h(s) for s in boundary_out)
        mass_closure = total_in - total_out
        mass_closure_pct = 100.0 * mass_closure / total_in if total_in > 0 else 0.0

        h_in = 0.0
        h_in_valid = True
        for s in boundary_in:
            h = stream_enthalpy_kJ_h(s)
            if h is None:
                h_in_valid = False
                break
            h_in += h
        for s in eboundary_in:
            d = energy_stream_duty_kJ_h(s)
            if d is None:
                h_in_valid = False
                break
            h_in += d
        h_out = 0.0
        h_out_valid = True
        for s in boundary_out:
            h = stream_enthalpy_kJ_h(s)
            if h is None:
                h_out_valid = False
                break
            h_out += h
        for s in eboundary_out:
            d = energy_stream_duty_kJ_h(s)
            if d is None:
                h_out_valid = False
                break
            h_out += d

        enthalpy_in = h_in if h_in_valid else None
        enthalpy_out = h_out if h_out_valid else None
        energy_closure = (
            enthalpy_in - enthalpy_out
            if enthalpy_in is not None and enthalpy_out is not None
            else None
        )

        return BalanceSummary(
            scope=scope,
            inlet_streams=sorted(boundary_in),
            outlet_streams=sorted(boundary_out),
            total_in_kg_h=total_in,
            total_out_kg_h=total_out,
            mass_closure_kg_h=mass_closure,
            mass_closure_pct=mass_closure_pct,
            inlet_energy_streams=sorted(eboundary_in),
            outlet_energy_streams=sorted(eboundary_out),
            enthalpy_in_kJ_h=enthalpy_in,
            enthalpy_out_kJ_h=enthalpy_out,
            energy_closure_kJ_h=energy_closure,
        )

    # ─────────────────────────────────────────────────
    # Stream 物性 (bulk)
    # ─────────────────────────────────────────────────
    def get_stream_phys(
        self,
        name: str,
        flowsheet_path: str | None = None,
    ) -> StreamPhysProps:
        """ストリームの bulk 物性 (密度・粘度・Cp・熱伝導率など) を返す。"""
        fs = self._navigate_flowsheet(flowsheet_path)
        s = fs.MaterialStreams.Item(name)

        def gv(attr: str) -> float | None:
            if not hasattr(s, attr):
                return None
            try:
                return float(getattr(s, attr))
            except Exception:
                return None

        return StreamPhysProps(
            name=s.name,
            flowsheet=flowsheet_path or "Main",
            vapor_fraction=gv("VapourFractionValue"),
            liquid_fraction=gv("LiquidFractionValue"),
            molecular_weight=gv("MolecularWeightValue"),
            mass_density_kg_m3=gv("MassDensityValue"),
            molar_density_kgmole_m3=gv("MolarDensityValue"),
            viscosity_cP=gv("ViscosityValue"),
            kinematic_viscosity_cSt=gv("KineticViscosityValue"),
            thermal_conductivity_W_mK=gv("ThermalConductivityValue"),
            mass_heat_capacity_kJ_kgK=gv("MassHeatCapacityValue"),
            molar_heat_capacity_kJ_kgmoleK=gv("MolarHeatCapacityValue"),
            cp_cv_ratio=gv("CpCvValue"),
            mass_enthalpy_kJ_kg=gv("MassEnthalpyValue"),
            molar_enthalpy_kJ_kgmole=gv("MolarEnthalpyValue"),
            mass_entropy_kJ_kgK=gv("MassEntropyValue"),
            heat_flow_kJ_h=gv("HeatFlowValue"),
            std_liquid_density_kg_m3=gv("StdLiqMassDensityValue"),
        )

    # ─────────────────────────────────────────────────
    # 内部ヘルパー
    # ─────────────────────────────────────────────────
    # ─────────────────────────────────────────────────
    # フローシート構築 (create / connect / disconnect / delete / ports)
    #
    # 注意: HYSYS COM の構築系シグネチャはバージョン差が大きいため、
    #   - confirm=False は必ずドライラン (副作用なし)
    #   - confirm=True は複数の COM シグネチャを順に試し、全滅なら error 集約
    #   - 構築中は Solver.CanSolve=False (Hold) にし、終了後に元の状態へ復元
    # の防御的方針で書く (WSL ではオフライン検証不可。実機確認前提)。
    # この codebase の Specifications.Add は (name, type) 順なので Operations.Add も
    # name 先を優先して試す。名前プロパティは既存コードに倣い小文字 .name 優先。
    # ─────────────────────────────────────────────────
    def _solver_can_solve(self) -> bool | None:
        try:
            return bool(self._case.Solver.CanSolve)
        except Exception:
            return None

    def _set_can_solve(self, value: bool) -> bool:
        try:
            self._case.Solver.CanSolve = bool(value)
            return True
        except Exception:
            return False

    @contextmanager
    def _solver_hold(self):
        """書込み中だけ Solver を Hold するコンテキストマネージャ。

        - Hold (CanSolve=False) に失敗したら例外で中断する。中途半端に
          モデルを変更してソルバが暴走するのを防ぐ (Codex review major)。
        - 元値を読めた場合のみ終了時に**確実に**復元する (例外発生時も finally)。
          元値が読めなかった場合は安全側に倒して Active(True) へ戻す。
        """
        prev = self._solver_can_solve()
        if not self._set_can_solve(False):
            raise RuntimeError(
                "Solver Hold (CanSolve=False) に失敗しました。"
                "モデルを変更せず中断します。HYSYS の状態を確認してください。"
            )
        try:
            yield
        finally:
            self._set_can_solve(prev if prev is not None else True)

    def _set_object_name(self, obj: Any, name: str) -> None:
        """生成直後のオブジェクト名を best-effort で設定 (.name → .Name)。"""
        if not name:
            return
        for attr in ("name", "Name"):
            try:
                if getattr(obj, attr, None) != name:
                    setattr(obj, attr, name)
                return
            except Exception:
                continue

    def _resolve_stream_obj(self, fs: Any, stream_name: str) -> tuple[Any, str]:
        """マテリアル → エネルギーの順で stream を解決。

        両コレクションで Item が失敗した場合、最後の COM エラーをメッセージに
        含めて投げる (RPC 切断等の本当の障害が「見つからない」に化けるのを防ぐ)。
        """
        last_err: Exception | None = None
        try:
            return fs.MaterialStreams.Item(stream_name), "material"
        except Exception as e:  # noqa: BLE001
            last_err = e
        try:
            return fs.EnergyStreams.Item(stream_name), "energy"
        except Exception as e:  # noqa: BLE001
            last_err = e
        raise RuntimeError(f"stream が見つからない: {stream_name!r} (last error: {last_err})")

    def _port_property(self, pl: str, p: str | None = None) -> str | None:
        """ポート種別 → 装置の接続点プロパティ名を返す (未対応は None)。

        HYSYS V14 実機確認: 接続は op.FeedStream / op.ProductStream /
        op.EnergyStream にストリームオブジェクトを代入して行う。これ以外の
        任意名ポートは当 COM ビルドで未対応のため None を返し、呼び出し側で
        明示的に ValueError にする (Codex review major)。
        """
        return {
            "feed": "FeedStream", "feeds": "FeedStream",
            "inlet": "FeedStream", "in": "FeedStream",
            "product": "ProductStream", "products": "ProductStream",
            "outlet": "ProductStream", "out": "ProductStream",
            "energy": "EnergyStream", "duty": "EnergyStream", "q": "EnergyStream",
        }.get(pl)

    # GUI 名 → 内部 TypeName。alias に無ければ渡された文字列をそのまま使う。
    # [実機確認済] は実プロセスモデル (HYSYS V14) の既存装置の TypeName を読んで
    # 突合済み。それ以外は未実機確認。
    _UNIT_OP_TYPE_ALIASES = {
        # ── 実機確認済 ──
        "cooler": "coolerop",            # [実機確認済]
        "heater": "heaterop",            # [実機確認済]
        "mixer": "mixerop",              # [実機確認済]
        "tee": "teeop",                  # [実機確認済]
        "pump": "pumpop",                # [実機確認済]
        "valve": "valveop",              # [実機確認済]
        "compressor": "compressor",      # [実機確認済] V14 は 'compressor' (compressop は不在)
        "distillation column": "distillation",  # [実機確認済]
        "column": "distillation",        # [実機確認済]
        "absorber": "absorber",          # [実機確認済] 吸収塔
        "shortcut column": "fractop",    # [実機確認済] ショートカット精製塔
        "spreadsheet": "spreadsheetop",  # [実機確認済] スプレッドシート
        # separator は V14 実機で TypeName='flashtank' を確認 (separatorop 等は不在)
        "separator": "flashtank",        # [実機確認済]
        "2 phase separator": "flashtank",
        "flash": "flashtank",
        "flash tank": "flashtank",
        "vessel": "flashtank",
        # ── 未実機確認 (検証に使ったモデルに該当装置が無く突合できていない) ──
        "3 phase separator": "separator3op",
        "expander": "expandop",
        "heat exchanger": "heatexop",
        "tank": "tankop",
        "conversion reactor": "conversionrxnop",
        "equilibrium reactor": "equilibriumrxnop",
        "cstr": "cstrrxnop",
        "pfr": "pfrrxnop",
    }

    def create_stream(
        self,
        name: str,
        stream_kind: str = "material",
        confirm: bool = False,
        flowsheet_path: str | None = None,
    ) -> dict:
        """マテリアル / エネルギーストリームを新規作成。confirm=True 必須。"""
        self._require_case()
        fs = self._navigate_flowsheet(flowsheet_path)
        kind = (stream_kind or "material").strip().lower()
        if kind in ("material", "mat", "m"):
            coll = fs.MaterialStreams
            kind_label = "material"
        elif kind in ("energy", "e", "q"):
            coll = fs.EnergyStreams
            kind_label = "energy"
        else:
            raise ValueError(
                f"stream_kind must be 'material' or 'energy', got {stream_kind!r}"
            )

        existing = []
        try:
            existing = [s.name for s in coll]
        except Exception:
            pass
        if name in existing:
            raise RuntimeError(f"stream '{name}' は既に存在します ({kind_label})")

        if not confirm:
            return {
                "applied": False,
                "would_create": name,
                "stream_kind": kind_label,
                "message": "Set confirm=True to actually create the stream.",
            }

        errors: list[str] = []
        with self._solver_hold():
            obj = None
            for desc, fn in (
                ("Add(name)", lambda: coll.Add(name)),
                ("Add()", lambda: coll.Add()),
            ):
                try:
                    obj = fn()
                    break
                except Exception as e:
                    errors.append(f"{desc}: {e}")
            if obj is None:
                raise RuntimeError(f"stream 作成失敗。errors: {errors}")
            self._set_object_name(obj, name)
            return {"applied": True, "created": name, "stream_kind": kind_label}

    def create_unit_op(
        self,
        type_name: str,
        name: str | None = None,
        confirm: bool = False,
        flowsheet_path: str | None = None,
    ) -> dict:
        """装置 (unit operation) を新規作成。confirm=True 必須。

        type_name は内部 TypeName (例 'coolerop') か GUI 名 (例 'Cooler') の
        どちらでも可 (alias で吸収。無ければそのまま使う)。引数順は環境差が
        あるため (name,type)/(type,name) の両方を試す。
        """
        self._require_case()
        fs = self._navigate_flowsheet(flowsheet_path)
        raw = (type_name or "").strip()
        if not raw:
            raise ValueError("type_name は必須です")
        internal = self._UNIT_OP_TYPE_ALIASES.get(raw.lower(), raw)
        candidates = [internal]
        if raw not in candidates:
            candidates.append(raw)

        if not confirm:
            return {
                "applied": False,
                "would_create": name,
                "type_name_input": raw,
                "type_name_resolved": internal,
                "message": "Set confirm=True to actually create the unit op. "
                           "型名が通らない場合は hysys_find_ops で既存装置の type_name を確認。",
            }

        errors: list[str] = []
        with self._solver_hold():
            obj = None
            for tn in candidates:
                attempts = []
                if name:
                    attempts.append((f"Add({name!r},{tn!r})", lambda tn=tn: fs.Operations.Add(name, tn)))
                    attempts.append((f"Add({tn!r},{name!r})", lambda tn=tn: fs.Operations.Add(tn, name)))
                attempts.append((f"Add({tn!r})", lambda tn=tn: fs.Operations.Add(tn)))
                for desc, fn in attempts:
                    try:
                        obj = fn()
                        break
                    except Exception as e:
                        errors.append(f"{desc}: {e}")
                if obj is not None:
                    break
            if obj is None:
                raise RuntimeError(f"装置作成失敗。errors: {errors}")
            if name:
                self._set_object_name(obj, name)
            created = None
            for attr in ("name", "Name"):
                try:
                    created = getattr(obj, attr)
                    break
                except Exception:
                    continue
            return {
                "applied": True,
                "created": created or name,
                "type_name_used": internal,
            }

    def connect_stream(
        self,
        op_name: str,
        stream_name: str,
        port: str = "feed",
        confirm: bool = False,
        flowsheet_path: str | None = None,
    ) -> dict:
        """ストリームを装置の Feed / Product / Energy ポートに接続。confirm=True 必須。

        port: 'feed' | 'product' | 'energy' のみ対応 (当 HYSYS COM ビルドでは
        op.FeedStream/ProductStream/EnergyStream への代入で接続する)。
        それ以外のポート名は ValueError。
        """
        self._require_case()
        fs = self._navigate_flowsheet(flowsheet_path)
        op = fs.Operations.Item(op_name)
        stream, skind = self._resolve_stream_obj(fs, stream_name)
        p = (port or "feed").strip()
        pl = p.lower()

        if not confirm:
            return {
                "applied": False,
                "would_connect": stream_name,
                "stream_kind": skind,
                "to_op": op_name,
                "port": p,
                "message": "Set confirm=True to actually connect. "
                           "ポート名が不明なら hysys_list_ports で列挙可。",
            }

        # HYSYS V14 実機確認: 接続点プロパティに「ストリームオブジェクト」を
        # 代入して接続する (例 op.FeedStream = stream)。文字列/None は
        # DISP_E_TYPEMISMATCH で弾かれるためオブジェクトのみ。Feeds.Add /
        # AttachFeed / Ports.ConnectTo は当環境に存在しない。
        # 対応ポートは feed / product / energy のみ (Codex review major #2:
        # 任意名ポートは実機未確認なので公開しない。schema も同様に絞る)。
        prop = self._port_property(pl, p)
        if prop is None:
            raise ValueError(
                f"未対応のポート '{p}'。対応は feed / product / energy のみ "
                "(任意名ポート接続は当 COM ビルドで未対応)。"
            )
        with self._solver_hold():
            try:
                setattr(op, prop, stream)
            except Exception as e:
                raise RuntimeError(
                    f"接続失敗 (op.{prop}=stream): {e}"
                ) from e
            return {
                "applied": True,
                "connected": stream_name,
                "to_op": op_name,
                "port": p,
                "via": f"op.{prop}=stream",
            }

    def disconnect_stream(
        self,
        op_name: str,
        stream_name: str,
        port: str = "feed",
        confirm: bool = False,
        flowsheet_path: str | None = None,
    ) -> dict:
        """装置からストリームを切断。confirm=True 必須。"""
        self._require_case()
        fs = self._navigate_flowsheet(flowsheet_path)
        op = fs.Operations.Item(op_name)
        stream, skind = self._resolve_stream_obj(fs, stream_name)
        p = (port or "feed").strip()
        pl = p.lower()

        if not confirm:
            return {
                "applied": False,
                "would_disconnect": stream_name,
                "from_op": op_name,
                "port": p,
                "message": "Set confirm=True to actually disconnect.",
            }

        # HYSYS V14 実機確認: このCOMビルドには接続点を空にする API が無い
        # (FeedStream への ''/None/Empty 代入は DISP_E_TYPEMISMATCH、
        #  AttachedFeeds/DownstreamOpers は読取専用で Remove 不可)。
        # 実際には何も変更しないので Solver には一切触れない (Hold すると計算中の
        # ソルバに割り込む副作用がある — Codex review major #2)。接続中でも
        # delete_object は可能なので、純粋な切断は対応不可として明示通知する。
        prop = self._port_property(pl, p)
        current = []
        try:
            coll = op.AttachedFeeds if prop == "FeedStream" else op.AttachedProducts
            current = list(coll.Names)
        except Exception:
            pass
        return {
            "applied": False,
            "status": "unsupported",
            "supported": False,
            "from_op": op_name,
            "stream": stream_name,
            "port": p,
            "currently_attached": current,
            "message": (
                "このHYSYSビルドのCOMには接続点を空にするAPIがありません。"
                "別ストリームへ繋ぎ替えるなら hysys_connect_stream を、"
                "オブジェクトごと消すなら hysys_delete_object を使ってください "
                "(削除は接続中でも可能)。完全な切断のみが目的なら GUI で行ってください。"
            ),
        }

    def delete_object(
        self,
        name: str,
        kind: str,
        confirm: bool = False,
        flowsheet_path: str | None = None,
    ) -> dict:
        """ストリーム / 装置を削除。confirm=True 必須。

        kind: 'stream' (material/energy 自動判定) | 'material_stream' |
              'energy_stream' | 'unit_op'。

        HYSYS V14 実機確認: obj.Delete() は存在しない。所属コレクションの
        Remove(name) で削除する。接続中でも削除可 (切断不要)。
        """
        self._require_case()
        fs = self._navigate_flowsheet(flowsheet_path)
        k = (kind or "").strip().lower()

        if k in ("unit_op", "op", "operation", "unitop"):
            colls = [("Operations", fs.Operations)]
            kind_label = "unit_op"
        elif k in ("material_stream", "material"):
            colls = [("MaterialStreams", fs.MaterialStreams)]
            kind_label = "material_stream"
        elif k in ("energy_stream", "energy"):
            colls = [("EnergyStreams", fs.EnergyStreams)]
            kind_label = "energy_stream"
        elif k in ("stream", "auto", ""):
            colls = [("MaterialStreams", fs.MaterialStreams),
                     ("EnergyStreams", fs.EnergyStreams)]
            kind_label = "stream"
        else:
            raise ValueError(f"kind must be stream/material_stream/energy_stream/unit_op, got {kind!r}")

        if not confirm:
            return {
                "applied": False,
                "would_delete": name,
                "kind": kind_label,
                "message": "Set confirm=True to actually delete.",
            }

        # コレクション名 → 実際に削除した種別 (Codex review minor #3:
        # kind='stream' の自動判定で material/energy のどちらを消したか返す)
        _resolved = {
            "Operations": "unit_op",
            "MaterialStreams": "material_stream",
            "EnergyStreams": "energy_stream",
        }
        errors: list[str] = []
        with self._solver_hold():
            for coll_name, coll in colls:
                try:
                    coll.Item(name)  # 存在確認
                except Exception:
                    errors.append(f"{coll_name}: '{name}' なし")
                    continue
                coll.Remove(name)
                return {
                    "applied": True,
                    "deleted": name,
                    "kind": kind_label,
                    "resolved_kind": _resolved.get(coll_name, kind_label),
                }
            raise RuntimeError(f"対象が見つからない: {name} (errors: {errors})")

    def list_ports(self, op_name: str, flowsheet_path: str | None = None) -> dict:
        """装置のポート / Feeds / Products を列挙 (接続前の探索用、読取専用)。"""
        self._require_case()
        fs = self._navigate_flowsheet(flowsheet_path)
        op = fs.Operations.Item(op_name)
        result: dict[str, Any] = {"op": op_name}
        try:
            result["type_name"] = getattr(op, "TypeName", None)
        except Exception:
            result["type_name"] = None

        def _enum(coll):
            names = []
            try:
                n = coll.Count
            except Exception:
                try:
                    return [getattr(it, "name", str(it)) for it in coll]
                except Exception:
                    return None
            for i in range(int(n)):
                try:
                    it = coll.Item(i)
                except Exception:
                    try:
                        it = coll.Item(i + 1)
                    except Exception:
                        continue
                nm = None
                for attr in ("name", "Name"):
                    try:
                        nm = getattr(it, attr)
                        break
                    except Exception:
                        continue
                names.append(nm if nm is not None else f"#{i}")
            return names

        for coll_name in ("Ports", "MaterialPorts", "EnergyPorts", "Feeds", "Products"):
            col = getattr(op, coll_name, None)
            if col is None:
                continue
            enumerated = _enum(col)
            if enumerated is not None:
                result[coll_name] = enumerated
        return result

    def _require_case(self) -> None:
        if self._case is None:
            raise RuntimeError(
                "HYSYS にケースが開かれていません。"
                "hysys_open でファイルを開くか、HYSYS GUI でファイルを開いてから接続してください。"
            )

    def _navigate_flowsheet(self, path: str | None) -> Any:
        """'Main' or 'Main/CO2 Absorption Tower' のようなパスから Flowsheet オブジェクトを取得。"""
        if path is None or path == "Main" or path == "":
            return self._case.Flowsheet
        # サブフローシート対応(TODO: 階層が深い場合の再帰)
        parts = path.split("/")
        fs = self._case.Flowsheet
        for p in parts[1:]:  # 先頭の Main をスキップ
            fs = fs.Flowsheets.Item(p)
        return fs


# 簡易テスト(直接実行時)
if __name__ == "__main__":
    client = HysysClient()
    try:
        client.connect()
        print(f"接続成功。ケース: {client._case.Title if client._case else '未オープン'}")
        if client.is_connected:
            streams = client.list_streams()
            print(f"ストリーム数: {len(streams)}")
            print(f"最初の5本: {streams[:5]}")
    except Exception as e:
        print(f"エラー: {e}", file=sys.stderr)
        sys.exit(1)
