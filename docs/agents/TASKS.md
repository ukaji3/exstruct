# Task List

未完了 [ ], 完了 [x]

## 数式取得機能追加

- [x] `SheetData`に`formulas_map`フィールドを追加し、シリアライズ対象に含める
- [x] `StructOptions`に`include_formulas_map: bool = False`を追加し、verbose時の既定挙動と整合させる
- [x] openpyxlで`data_only=False`の読み取りパスを追加し、`formulas_map`用の走査処理を実装する
- [x] `.xls`かつ数式取得ONの場合はCOM経由で`formulas_map`を取得し、遅延警告を出す
- [x] `formulas_map`の仕様（=付きの式文字列、空文字除外、=のみ許可、共有/配列は未展開）に沿った抽出ロジックを追加
- [x] openpyxlの配列数式（`ArrayFormula`）は`value.text`から式文字列を取得する分岐を追加
- [x] CLI/ドキュメント/READMEの出力モード説明に`formulas_map`の条件を追記する
- [x] テスト要件に`formulas_map`関連（ON/OFF、verbose既定、.xls COM分岐）を追加する

## PR #44 指摘対応

- [x] `src/exstruct/render/__init__.py` の `_page_index_from_suffix` を2桁固定ではなく可変桁の数値サフィックスに対応させ、`_rename_pages_for_print_area` の上書きリスクを解消する
- [x] `src/exstruct/render/__init__.py` の `_export_sheet_pdf` の `finally` 内 `return` を削除し、PrintArea 復元失敗はログに残して例外を握りつぶさない
- [x] `src/exstruct/core/pipeline.py` の `step_extract_formulas_map_*` の挙動を docstring に合わせる（失敗時にログしてスキップ）か、docstring を実装に合わせて修正する
- [x] `docs/README.ja.md` の `**verbose**` 説明行を日本語に統一する

## PR #44 コメント/Codecov 対応

- [x] Codecov パッチカバレッジ低下（60.53%）の指摘に対応し、対象ファイルの不足分テストを追加する（`src/exstruct/render/__init__.py`, `src/exstruct/core/cells.py`, `src/exstruct/core/backends/com_backend.py`, `src/exstruct/core/pipeline.py`, `src/exstruct/core/backends/openpyxl_backend.py`）
- [x] Codecov の「Files with missing lines」で具体的な未カバー行を確認し、テスト観点を整理する
- [x] Codacy 警告対応: `src/exstruct/render/__init__.py:274` の finally 内 return により例外が握りつぶされる可能性（`PyLintPython3_W0150`）を解消する

## PR #44 CodeRabbit 再レビュー対応

- [ ] `scripts/codacy_issues.py`: トークン未設定時の `sys.exit(1)` をモジュールトップから排除し、`get_token()` または `main()` で検証する
- [ ] `scripts/codacy_issues.py`: `format_for_ai` の `sys.exit` を `ValueError` に置換し、呼び出し側でバリデーションする
- [ ] `scripts/codacy_issues.py`: `urlopen` の非2xxチェック（到達不能）を削除または `HTTPError` 側へ寄せる
- [ ] `scripts/codacy_issues.py`: `status` の固定値バリデーションを廃止する（固定なら直代入／必要なら CLI 引数化）
- [ ] `tests/backends/test_print_areas_openpyxl.py`: `PrintAreaData` 型に合わせる＋関連テストに Google スタイル docstring を付与
- [ ] `tests/core/test_pipeline.py`: 無効な `MergedCellRange` を有効な非重複レンジに修正する
- [ ] `tests/backends/test_backends.py`: `sheets` のクラス属性共有を避け、インスタンス属性に変更する
- [ ] `tests/render/test_render_init.py` / `tests/utils.py` / `tests/models/test_models_export.py`: docstring/コメントの指摘を反映する
- [ ] `src/exstruct/render/__init__.py`: Protocol クラスに Google スタイル docstring を追加する
