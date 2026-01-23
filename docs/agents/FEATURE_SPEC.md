# Feature Spec for AI Agent (Phase-by-Phase)

本ドキュメントは AI エージェント向けに、段階的に実装を進めるための仕様メモです。

---

## 数式取得機能追加

- 新たに数式文字列をそのまま取得する機能を追加
- `SheetData`モデルに`formulas_map`を新設予定`formulas_map: dict[str, list[tuple[int,int]]]`
- 数式の値は定義されている数式をそのまま取得する
- セル座標はcolors_mapと同じようにr,cの数値で表記
- デフォルトはverboseモード以上で出力、もしくはオプションからONにする
- 定義されている数式文字列をシンプルに取得する実装
- 数式の表記形式は「=A1」のようにユーザーが見るままの数式文字列にする
- 共有数式や配列数式は一旦は展開しない実装にする
- 空文字は除外、=だけのセルも数式文字として取得
- formulas_mapのキーは「式文字列（先頭=を含む）」で固定する
- 既存の値はSheetData.rowsにあり、数式はSheetData.formulas_mapにあることで共存する
- データ取得時はformulas_map が ON のときだけ data_only=False で再読込
- オプションは`StructOptions`にて`include_formulas_map: bool = False`で設定を受け付ける
- `.xls`形式かつ数式取得ONの時は処理が遅くなるという警告を出しつつ、COMで取得処理をする。
- cell.value が ArrayFormula の場合に value.text（実際の式文字列）を使う

---

## 今後のオプション検討メモ

- 表検知スコアリングの閾値を CLI/環境変数で調整可能にする
- 出力モード（light/standard/verbose）に応じてテーブル候補数を制限するオプション

---

## 実装方針

- 小さなステップごとにテスト追加、または既存フィクスチャで手動確認
- 短い関数・責務分割でスコアリング調整をしやすくする
- 外部公開前なので、破壊的変更はコメントや仕様に明示して段階的に移行する
