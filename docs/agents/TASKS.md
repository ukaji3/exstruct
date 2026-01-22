# Task List

未完了 [ ], 完了 [x]

## 数式取得機能追加

- [ ] `SheetData`に`formulas_map`フィールドを追加し、シリアライズ対象に含める
- [ ] `StructOptions`に`include_formulas_map: bool = False`を追加し、verbose時の既定挙動と整合させる
- [ ] openpyxlで`data_only=False`の読み取りパスを追加し、`formulas_map`用の走査処理を実装する
- [ ] `.xls`かつ数式取得ONの場合はCOM経由で`formulas_map`を取得し、遅延警告を出す
- [ ] `formulas_map`の仕様（=付きの式文字列、空文字除外、=のみ許可、共有/配列は未展開）に沿った抽出ロジックを追加
- [ ] openpyxlの配列数式（`ArrayFormula`）は`value.text`から式文字列を取得する分岐を追加
- [ ] CLI/ドキュメント/READMEの出力モード説明に`formulas_map`の条件を追記する
- [ ] テスト要件に`formulas_map`関連（ON/OFF、verbose既定、.xls COM分岐）を追加する
