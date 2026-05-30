# TODO

## Phase 0: 環境準備 (完了 2026-05-14)

- [x] Windows ネイティブ Python 3.12.4 (D:\anaconda3) で venv 作成、`pip install -e .` 完了
- [x] HYSYS V14 起動・ケース open 確認
- [x] PowerShell / WSL から COM 接続テスト成功 (`HYSYS.Application` Dispatch OK)

## Phase 1: HYSYS COM API 探索 (完了 2026-05-15)

- [x] `Hysys.Application` / `ActiveDocument` / `Flowsheet` のプロパティ列挙
- [x] MaterialStream の T/P/F/MolarFlow/Mass/ComponentMolarFlow/Fraction 確認
- [x] Operations 種別と TypeName 確認 (`distillation`, `traysection`, `coolerop`, `flashtank` 等)
- [x] サブフローシート (ColumnFlowsheet) API パス確認 (`col.ColumnFlowsheet.Operations.Item('Main Tower')`)
- [x] V14 固有プロパティ確認: `TemperatureValue`/`PressureValue` (単数+Value), Cooler は `PressureDrop` (DeltaP は無い)
- [x] Specifications API (`cfs.Specifications.Item(name).IsActive / Goal / GoalValue`) 確認
- [x] Trace pane が write-only であることを確認 (read API 公開なし)

## Phase 2: client実装 (完了 2026-05-15)

- [x] `connect()` / `open_file()` / `close_file()`
- [x] `list_streams()` / `get_stream()` (T/P/F/組成/相)
- [x] `list_unit_ops()` / `set_unit_op_param()`
- [x] `get_solver_status()` (拡張: 未収束 ops/cols/streams を error_messages に列挙)
- [x] `set_stream()` (T/P/F/組成/component_molar_flows, dry-run ガード)
- [x] `run()` / `reset()` (confirm=True 必須) / `save_as()` (上書き禁止)
- [x] `get_column_profile()` (T/P。L/V は V14 制約で null)
- [x] `case_study()` (sweep → run → observe → revert)
- [x] `list_column_specs()` / `set_column_spec()` / `set_column_spec_active()`

## Phase 3: MCP server 実装 (完了 2026-05-15)

- [x] tools/list 応答 (16 ツール登録)
- [x] tools/call 応答 (全ツール dispatcher)
- [x] stdio トランスポート (smoke test 通過)

## Phase 4: Claude Code 連携 (完了 2026-05-14)

- [x] Vault ルート `.mcp.json` にプロジェクトスコープで hysys 登録
- [x] WSL Claude Code から Windows python.exe 起動経由で MCP 認識

## Phase 5: 書き込み系拡張(4日目以降)

- [x] hysys_set_stream (2026-05-15, confirm=True 必須のドライランガード付き)
- [x] hysys_set_unit_op_param (V14 確認済みパラメータ名: `PressureDrop`, `Duty`, `DeltaT` etc)
- [x] hysys_run
- [x] hysys_reset (confirm=True 必須)
- [x] hysys_save (別名保存のみ、上書き禁止ガード)

## Phase 6: 高度なツール

- [x] hysys_get_column_profile (T/P 段別取得可、L/V 流量は V14 標準 COM 未露出のため null)
- [x] hysys_case_study (2026-05-15 感度解析: stream/op を sweep_values で振り observe 観測、confirm=True で実走後に自動復元)
- [x] hysys_list_column_specs (Spec 列挙: name/IsActive/Goal/Current/Error/Tolerance)
- [x] hysys_set_column_spec (Goal 値の変更、Goal.SetValue(v,unit) or GoalValue= に dispatch)
- [x] hysys_set_column_spec_active (IsActive 切替。Degrees of Freedom 整合は呼び出し側責任)
- [x] hysys_get_column_spec_detail (2026-05-15: TypeName/SpecifiedStage/SpecifiedDraw/Phase/FlowBasis/IncludedComponents 等の詳細取得)
- [x] hysys_clear_column_estimates (2026-05-15: ClearAllEstimates/ClearAllCompositionEstimates/ClearTrayCompositionEstimates、FP切替後の推定値リセット用)
- [x] hysys_set_column_solver_param (2026-05-15: IsUsingSolutionForEstimates/MaximumIterations/IsAdaptiveDamping/DampingFactor等のソルバー設定)
- [x] hysys_remove_column_spec (2026-05-15: クローン塔継承の不要Spec削除用、Specifications.Remove)
  - 2026-05-15 修正: 壊れた spec (stage/draw 未確定) では `IsActive`/`GoalValue` アクセスで COM E_FAIL になり Remove 自体が失敗 → `before_active`/`before_goal` 取得を try/except で保護
- [x] hysys_add_column_spec (2026-05-15: 新規Spec追加、reflux_ratio/comp_recovery/comp_fraction/draw_rate/temperature/duty/vapour_fraction/pressure)
  - 2026-05-15 修正: 属性設定順序を stage→draw→phase→components→goal→is_active に変更。Comp Fraction Spec 等で stage/draw/phase/components 未確定状態で `GoalValue` を触ると COM E_FAIL になるため。全段 try/except 化し、エラーは `*_error` キーで返却
- [x] hysys_set_stream に component_molar_flows 追加 (成分別 kgmole/h で部分更新)
- [x] hysys_get_status を拡張 (未収束 ops/columns/streams を error_messages として列挙)
- [ ] hysys_place_unit / hysys_connect_stream (新規装置配置・配線、未着手 — Tier 3)

## Phase 7: Tier 1+2 拡張 (完了 2026-05-15)

合計 14 ツール追加。MCP サーバー総数 **30 ツール**。

### Tier 1: 設計ループの土台

- [x] hysys_list_logical_ops / get_logical_op / set_adjust_target / reset_recycle
      (検証に使ったモデルに Logical Op が無く API 探索のみ完了、`AdjustOp`/`SpreadsheetOp` の TypeName は確認済み。プロパティ名は hasattr 判定で動的対応)
- [x] hysys_get_fluid_package / list_components / get_component
      (H2O: MW=18.01, NBP=99.998℃, Tc=374.15℃, Pc=22120 kPa, ω=0.344 をテキスト値一致で取得確認)
- [x] hysys_add_internal_stream (段別流量取得用、AddDrawStream 経由)
- [x] hysys_find_streams (組成・流量・vap_frac・name パターンで検索)
- [x] hysys_find_ops (type/name パターン)
- [x] hysys_introspect (任意 COM オブジェクトの dir() を安全に列挙)

### Tier 2: 設計判断情報

- [x] hysys_get_heat_op (Cooler/Heater/HeatExchanger 汎用、Duty/Feed/Product/UA/LMTD)
- [x] hysys_balance_check (op 範囲指定 or 全体、内部 stream 除外で境界収支計算)
- [x] hysys_get_stream_phys (bulk 物性、密度・粘度・Cp・熱伝導率・MW・エンタルピー等 17 項目)

### Tier 3 残作業 (次セッション以降)

- [ ] hysys_place_unit / hysys_connect_stream (新規装置追加、Operations.Add 系)
- [ ] Component 追加・削除
- [ ] Property Package 切替

### Tier 4 残作業 (後日)

- [ ] Workbook / Datatable Excel エクスポート
- [ ] PFD 画像エクスポート
- [ ] Optimizer (HYOPT) 接続
- [ ] Aspen Process Economic Analyzer 連携

### Phase 7 e2e 検証シナリオ (実プロセスモデルでの例)

1. Fluid Package 確認 → "UNIQUAC-理想", 8 成分 + Reaction Package あり
2. H2O > 80% ストリーム検索 → 26 本ヒット
3. 全 Cooler の Duty 集計 → 7 台合計 53.8 MW 冷却負荷
4. 全フローシート物質収支 → 936,513 kg/h、closure 0.00013% (essentially perfect)
5. introspect で stream7 の enthalpy 系列挙 → 8 メンバ取得

### Phase 6 残作業メモ

#### 段別 L/V 流量 (今回未解決)

HYSYS V14 の `traysection` オブジェクトに段別流量配列プロパティは無い。
公開資料 (Aspen Customization Guide, GitHub CAChemE/stochastic-optimization, Cheresources)
にも記述なしを Web 調査で確認 (2026-05-15)。

**回避策 (ユーザ操作)**: HYSYS GUI で各段に Internal Stream を作成し、
通常の MaterialStream として段流量を露出させる。
作成後は `hysys_get_stream("Stage_5_Net_Liquid")` で取得できる。

#### エラーメッセージ取得 (代替案で実装済み)

`Application.Trace` / `Case.Trace` は write-only で読み取り API 公開なし。
代わりに「未収束オブジェクト walk」で実装:
- Operations の IsValid=False or SolveComplete=False を列挙
- Column の CfsConverged=False を列挙
- Stream の IsKnown=False を列挙

これで「どこが赤いか」は分かるが、HYSYS GUI に出る詳細エラーテキスト
(e.g. "Convergence failed: T below dew point") は取得不可。

#### 未実装 Phase 7+

- place_unit / connect_stream — 新規装置/配線。COM の Operations.Add() 系を使うが API パスが複雑、需要が出てから対応
- DataTables / Workbook export — Excel 連携で数値表エクスポート
- Optimizer (HYOPT) — 最適化変数定義と実行

## Phase 7: ドキュメント・公開

- [ ] 使用例 README に追加
- [ ] GitHub に公開(brack101/AspenPlus-MCP-Server に並ぶ HYSYS版として)
- [ ] X / Zenn で発信

---

## Phase 8: 汎用 COM メソッド呼び出し (2026-05-15 実装完了)

### 実装ツール (合計 34 ツール)

- [x] `hysys_call_method(path, method, args, confirm)` — 任意 COM オブジェクトの汎用メソッド呼び出し
  - **動機**: 反応セットを Fluid Package にアタッチする `ReactionSet.AssociateFluidPackage(fp)` のような操作のために、専用ツールを毎回作るのではなく汎用ツール 1 個で対応
  - **実装**: `path` は introspect と同じ namespace で eval、`args` は `[{type:'literal'|'path', value:...}]` の順序付きリスト、`confirm=True` 必須
  - **代表用途**:
    - `ReactionSet.AssociateFluidPackage(fp)` — 反応セットを FP にアタッチ
    - `ColumnFlowsheet.Operations.Add(...)` — 塔反応段の追加 (引数シグネチャは試行錯誤要)
  - **セキュリティ**: introspect と同じ tokens ブラックリスト (`__`, `import `, `exec(` 等)、method 名にもメタ文字禁止

## Phase 7: ソルバー制御の即時対応 (2026-05-15 実装完了)

### 実装ツール (合計 33 ツール)

- [x] `hysys_set_solver_can_solve(value: bool, confirm)` — `case.Solver.CanSolve = value` で全体ソルバー停止/再開
  - **動機**: 塔ソルバー停滞検知時に GUI 操作なしで Stop Run したい
  - **実装**: before/after/is_solving を返す
- [x] `hysys_column_reset(column_name: str, confirm)` — 塔の Column Environment Reset 相当
  - **動機**: 停滞解を捨てて再収束したい。フローシート全体 reset より範囲限定
  - **実装**: `col.ColumnFlowsheet.Reset()` (introspect で実在確認済)
- [x] `hysys_column_run(column_name: str)` — 塔個別 Run
  - **動機**: メインソルバー Hold でも塔単独起動したい / set→Run の連鎖を短く
  - **実装**: `col.ColumnFlowsheet.Run()`。返値に CurrentIteration / CfsConverged / SolvingStatus / EquilibriumError

### 既存ツールの不具合

- [ ] `hysys_set_column_spec_active` / `hysys_set_column_solver_param` — 塔ソルバーが「Solved with non-convergence」状態だと set が silent revert される
  - **症状**: applied=true, after=新値 と返るのに、直後の list/introspect では旧値のまま
  - **対策案**: set 前に塔ソルバー停止 → set → set 後に値検証 → 検証失敗ならエラー返す。`hysys_column_reset` 実装と組み合わせる

---

## 即時使用例

HYSYS ケースを開いた状態で、こんな対話ができます:

> Claude: Feed の組成を確認して
→ `hysys_get_stream("Feed")` → `Methanol: 1346 kgmole/h, H2O: 10666 kgmole/h`

> Claude: Feed の Methanol を 1500 kgmole/h に変えて再収束させて
→ `hysys_set_stream("Feed", composition_changes={"Methanol": 1500})`
→ `hysys_run()`
→ `hysys_get_status()` → "Converged in 234 iterations"

> Claude: 吸収塔の段別温度プロファイルを CSV で出して
→ `hysys_get_column_profile("Absorber")`
→ tabular data returned
