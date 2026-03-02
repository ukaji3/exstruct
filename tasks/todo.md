## Plan

- [x] `tasks/feature_spec.md` のIssue 72仕様を最終確認
- [x] `src/exstruct/mcp/render_runner.py` を新規追加
- [x] `src/exstruct/mcp/tools.py` に `CaptureSheetImagesToolInput/Output` と `run_capture_sheet_images_tool` を追加
- [x] `src/exstruct/mcp/server.py` に `exstruct_capture_sheet_images` ツール登録を追加
- [x] `src/exstruct/mcp/shared/output_path.py` に画像出力ディレクトリ解決ヘルパを追加
- [x] `src/exstruct/mcp/shared/a1.py` にシート修飾A1範囲パーサ/正規化を追加
- [x] `src/exstruct/render/__init__.py` に `sheet` / `a1_range` 指定対応を追加
- [x] `src/exstruct/mcp/__init__.py` の公開エクスポートを更新
- [x] `docs/mcp.md` を更新（新ツール仕様、例、制約）
- [x] `README.md` と `README.ja.md` のMCPツール一覧を更新

## Test Cases

- [x] `tests/mcp/test_tool_models.py` に新規入力/出力モデル検証を追加
- [x] `tests/mcp/test_tools_handlers.py` にハンドラ request/result 変換テストを追加
- [x] `tests/mcp/test_server.py` にツール登録と引数伝播テストを追加
- [x] `tests/mcp/shared/test_output_path.py` に未指定 `out_dir` 一意化テストを追加
- [x] `tests/render/test_render_init.py` に `sheet` / `range` 指定出力テストを追加
- [x] 既存render/mcpテストの回帰確認

## Verification

- [x] `uv run pytest tests/mcp/test_tool_models.py tests/mcp/test_tools_handlers.py tests/mcp/test_server.py tests/mcp/shared/test_output_path.py tests/render/test_render_init.py`
- [x] `uv run task precommit-run`

## Review

- Summary:
  - Issue 72として `exstruct_capture_sheet_images` をMCPに追加し、`sheet` / `range` 指定、`out_dir` 未指定時の一意ディレクトリ解決、COM必須チェックを実装した。
  - `render.export_sheet_images` に `sheet` / `a1_range` を追加し、対象シート/範囲のみ出力できるようにした。
- Verification:
  - `uv run pytest tests/mcp/test_tool_models.py tests/mcp/test_tools_handlers.py tests/mcp/test_server.py tests/mcp/shared/test_output_path.py tests/render/test_render_init.py` -> 148 passed
  - `uv run task precommit-run` -> ruff / ruff-format / mypy passed
- Risks:
  - `sheet` 名一致は厳密一致（大文字小文字含む）で比較しているため、利用側が異なる表記を渡すと不一致エラーになる。
  - 範囲指定時のレンダリング結果ページ数はExcelの改ページ設定に依存する。
- Follow-ups:
  - 必要であれば `sheet` 名の大小無視比較オプションを検討する。
  - 将来的に `copyPicture` 以外の軽量レンダリング経路を比較評価する。

## Timeout Hardening (2026-02-28)

- [x] Add subprocess join timeout + terminate/kill fallback in `src/exstruct/render/__init__.py`.
- [x] Add MCP-side timeout guard for `exstruct_capture_sheet_images` in `src/exstruct/mcp/server.py`.
- [x] Update/extend tests in `tests/render/test_render_init.py` and `tests/mcp/test_server.py`.

## Timeout Hardening Review

- Summary:
  - Added bounded wait for render subprocess and forced shutdown when join timeout is exceeded.
  - Added tool-level timeout in MCP capture handler with explicit timeout error message.
- Verification:
  - `uv run pytest tests/render/test_render_init.py tests/mcp/test_server.py -q`
  - `uv run task precommit-run`

## Follow-up Fix (non-finite timeout env) 2026-02-28

- [x] Reject non-finite timeout env values in `src/exstruct/render/__init__.py`.
- [x] Reject non-finite timeout env values in `src/exstruct/mcp/server.py`.
- [x] Extend tests in `tests/render/test_render_init.py` and `tests/mcp/test_server.py` for `NaN/inf/-inf`.

## Follow-up Fix Review

- Summary:
  - Added `math.isfinite(...)` checks so non-finite timeout env values fallback to defaults.
  - Closed both review findings for render and MCP timeout readers.
- Verification:
  - `uv run pytest tests/render/test_render_init.py tests/mcp/test_server.py -q`

## Rollout Plan (capture_sheet_images limited release, 2026-02-28)

- [x] `tasks/feature_spec.md` に限定提供（experimental）方針・運用ポリシー・GA基準を確定反映
- [x] MCPサーバー起動経路で `EXSTRUCT_RENDER_SUBPROCESS=0` を既定化（MCP実行時のみ）
- [x] `docs/mcp.md` に Experimental 表記、推奨環境変数、既知制約を追記
- [x] `README.md` / `README.ja.md` に運用注意（限定提供・依存条件）を追記
- [x] サブプロセスハング切り分け用の計測ログポイントを `render` 側に追加（export/join/queue/write）
- [x] サブプロセス経路の再現テストを追加（MCPコンテキスト相当でのタイムアウト/エラー検証）
- [x] 成功率評価用の代表Workbookセットと計測手順を `tasks/` に定義

## Rollout Verification

- [x] `uv run pytest tests/mcp/test_server.py tests/render/test_render_init.py`
- [x] `uv run task precommit-run`
- [ ] 手動確認: `exstruct_capture_sheet_images` を最小範囲 `sheet=Sheet1, range=A1:A1` で実行し、120sタイムアウトが解消していること

## Rollout Review (template)

- Summary:
  - MCP runtime で `EXSTRUCT_RENDER_SUBPROCESS=0` を既定化し、`capture_sheet_images` の限定提供方針を docs/README に反映。
  - `render` に段階ログ（export/subprocess/worker）を追加し、MCPコンテキスト相当のエラー伝播テストを追加。
  - 成功率評価の代表Workbookセットと手順を `tasks/capture_sheet_images_eval.md` として定義。
- Verification:
  - `uv run pytest tests/mcp/test_server.py tests/render/test_render_init.py -q` -> 78 passed
  - `uv run task precommit-run` -> ruff / ruff-format / mypy passed
- Risks:
  - `EXSTRUCT_RENDER_SUBPROCESS=0` はクラッシュ分離を弱めるため、長時間運用時のメモリ観測が必要。
- Follow-ups:
  - 実環境での手動確認（最小範囲ケース）を実施し、失敗時はログをもとに subprocess 経路根治へ移行。

## Subprocess Stabilization Plan (2026-03-02)

- [ ] `tasks/feature_spec.md` の 2026-03-02 Addendum をベースに実装方針を確定
- [ ] `src/exstruct/render/` にサブプロセス worker 専用エントリポイントを追加（親の `stdin/-c` 実行に非依存）
- [ ] `src/exstruct/render/__init__.py` の `_render_pdf_pages_subprocess` を結果先行フローへ変更（join先行を廃止）
- [ ] startup/result/join の3段階タイムアウトと stage-aware エラー整形を実装
- [ ] 例外/タイムアウト時のプロセス終了手順を統一（terminate/kill/cleanup）
- [ ] `tests/render/test_render_init.py` に回帰テストを追加（worker結果返却済み + join遅延ケースで成功扱い）
- [ ] `tests/render/test_render_init.py` に回帰テストを追加（worker bootstrapping 失敗時に actionable メッセージ）
- [ ] `tests/render/test_render_init.py` に回帰テストを追加（result timeout / join timeout のエラーステージ区別）
- [ ] `tests/mcp/test_server.py` に `EXSTRUCT_RENDER_SUBPROCESS=1` 相当の伝播/失敗メッセージ検証を追加
- [ ] 代表Workbookセットで手動再検証（`EXSTRUCT_RENDER_SUBPROCESS=1`）
- [ ] docs 更新（`docs/mcp.md`, `README.md`, `README.ja.md`）: サブプロセス再有効化条件と既知制約

## Subprocess Stabilization Verification

- [ ] `uv run pytest tests/render/test_render_init.py tests/mcp/test_server.py -q`
- [ ] `uv run task precommit-run`

## Subprocess Stabilization Review (template)

- Summary:
  - 
- Verification:
  - 
- Risks:
  - 
- Follow-ups:
  - 
