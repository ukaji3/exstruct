# COM vs OOXML 実装比較ドキュメント

ExStruct は Windows 環境では Excel COM API、Linux/macOS 環境では OOXML (Office Open XML) パーサーを使用して Excel ファイルから Shape と Chart を抽出します。本ドキュメントでは、両実装の詳細な機能比較と実装方法の違いを説明します。

## 目次

1. [アーキテクチャ概要](#アーキテクチャ概要)
2. [Shape 抽出機能](#shape-抽出機能)
3. [Chart 抽出機能](#chart-抽出機能)
4. [実装方法の比較](#実装方法の比較)
5. [機能対応表](#機能対応表)
6. [制限事項と注意点](#制限事項と注意点)

---

## アーキテクチャ概要

### COM 実装 (Windows)

```
Excel Application (COM)
    ↓
xlwings ライブラリ
    ↓
shapes.py / charts.py
    ↓
Pydantic モデル (Shape, Chart)
```

- **依存**: Microsoft Excel がインストールされた Windows 環境
- **利点**: Excel の全機能にアクセス可能、高精度
- **欠点**: Windows 限定、Excel ライセンス必要

### OOXML 実装 (クロスプラットフォーム)

```
.xlsx ファイル (ZIP アーカイブ)
    ↓
xml.etree.ElementTree
    ↓
drawing.py / chart.py
    ↓
Pydantic モデル (Shape, Chart)
```

- **依存**: Python 標準ライブラリのみ
- **利点**: 全プラットフォーム対応、Excel 不要
- **欠点**: 一部機能が制限される

---

## Shape 抽出機能

### 1. 位置情報 (l, t)

#### COM 実装
```python
# shapes.py
shape_obj = Shape(
    l=int(shp.left),
    t=int(shp.top),
    ...
)
```
- `xlwings.Shape.left` / `xlwings.Shape.top` プロパティを使用
- 単位: ポイント (1/72 インチ)

#### OOXML 実装
```python
# drawing.py
def _get_xfrm_position(elem: Element) -> tuple[int, int, int, int] | None:
    xfrm = elem.find(".//a:xfrm", NS)
    off = xfrm.find("a:off", NS)
    x = int(off.get("x", "0"))
    y = int(off.get("y", "0"))
    return (emu_to_pixels(x), emu_to_pixels(y), ...)
```
- DrawingML の `<a:xfrm><a:off>` 要素から取得
- 単位: EMU (English Metric Units) → ピクセルに変換
- 変換式: `pixels = emu / 914400 * 96`

---

### 2. サイズ情報 (w, h)

#### COM 実装
```python
w=int(shp.width) if mode == "verbose" or shape_type_str == "Group" else None,
h=int(shp.height) if mode == "verbose" or shape_type_str == "Group" else None,
```

#### OOXML 実装
```python
ext = xfrm.find("a:ext", NS)
cx = int(ext.get("cx", "0"))
cy = int(ext.get("cy", "0"))
return (..., emu_to_pixels(cx), emu_to_pixels(cy))
```
- `<a:ext cx="..." cy="...">` から取得
- verbose モードまたはグループ図形の場合のみ出力

---

### 3. テキスト抽出

#### COM 実装
```python
try:
    text = shp.text.strip() if shp.text else ""
except Exception:
    text = ""
```
- `xlwings.Shape.text` プロパティを使用

#### OOXML 実装
```python
def _get_text_from_element(elem: Element) -> str:
    texts: list[str] = []
    for t_elem in elem.findall(".//a:t", NS):
        if t_elem.text:
            texts.append(t_elem.text)
    return "".join(texts).strip()
```
- `<a:t>` 要素（テキストラン）を全て収集して結合
- リッチテキストの書式情報は無視

---

### 4. 図形種別 (type)

#### COM 実装
```python
type_num = shp.api.Type
shape_type_str = MSO_SHAPE_TYPE_MAP.get(type_num, f"Unknown({type_num})")
astype_num = shp.api.AutoShapeType
autoshape_type_str = MSO_AUTO_SHAPE_TYPE_MAP.get(astype_num, ...)
type_label = f"{shape_type_str}-{autoshape_type_str}"
```
- `MsoShapeType` と `MsoAutoShapeType` の組み合わせ
- 例: `"AutoShape-FlowchartProcess"`

#### OOXML 実装
```python
def _get_preset_geometry(elem: Element) -> str | None:
    prst_geom = elem.find(".//a:prstGeom", NS)
    return prst_geom.get("prst") if prst_geom is not None else None

# PRESET_GEOM_MAP で変換
PRESET_GEOM_MAP = {
    "flowChartProcess": "AutoShape-FlowchartProcess",
    "rect": "AutoShape-Rectangle",
    ...
}
```
- `<a:prstGeom prst="...">` 属性から取得
- マッピングテーブルで COM 互換の名前に変換

---

### 5. Shape ID 割り当て

#### COM 実装
```python
node_index = 0
for shp in iter_shapes_recursive(root):
    if not is_relationship_geom:
        node_index += 1
        shape_id = node_index
    if excel_name:
        if shape_id is not None:
            excel_names.append((excel_name, shape_id))
```
- 非コネクター図形に連番 ID を割り当て
- Excel 内部名と ID のマッピングを保持

#### OOXML 実装
```python
def _assign_shape_ids(parse_results: list[_ShapeParseResult]) -> None:
    excel_id_to_node_id: dict[str, int] = {}
    node_index = 0
    # First pass: assign node IDs to non-connector shapes
    for result in parse_results:
        if not result.is_connector and result.excel_id:
            node_index += 1
            result.shape.id = node_index
            excel_id_to_node_id[result.excel_id] = node_index
```
- `<xdr:cNvPr id="...">` から Excel ID を取得
- 非コネクター図形に連番 ID を割り当て
- Excel ID → Node ID のマッピングを構築

---

### 6. コネクター方向 (direction)

#### COM 実装
```python
def compute_line_angle_deg(w: float, h: float) -> float:
    return math.degrees(math.atan2(h, w)) % 360.0

def angle_to_compass(angle: float) -> str:
    dirs = ["E", "NE", "N", "NW", "W", "SW", "S", "SE"]
    idx = int(((angle + 22.5) % 360) // 45)
    return dirs[idx]

angle = compute_line_angle_deg(float(shp.width), float(shp.height))
shape_obj.direction = angle_to_compass(angle)
```

#### OOXML 実装
```python
def _compute_direction(width: int, height: int) -> str | None:
    angle = math.degrees(math.atan2(-height, width))
    if angle < 0:
        angle += 360
    # Map angle to compass direction
    if 337.5 <= angle or angle < 22.5:
        return "E"
    elif 22.5 <= angle < 67.5:
        return "NE"
    ...
```
- 両実装とも `atan2` で角度を計算し、8方位に変換
- 座標系の違いにより符号が異なる

---

### 7. 矢印スタイル (begin_arrow_style, end_arrow_style)

#### COM 実装
```python
begin_style = int(shp.api.Line.BeginArrowheadStyle)
end_style = int(shp.api.Line.EndArrowheadStyle)
shape_obj.begin_arrow_style = begin_style
shape_obj.end_arrow_style = end_style
```
- `MsoArrowheadStyle` 列挙値を直接使用

#### OOXML 実装
```python
ARROW_HEAD_MAP: dict[str, int] = {
    "none": 1,
    "triangle": 2,
    "stealth": 3,
    "diamond": 4,
    "oval": 5,
    "arrow": 2,
}

def _get_arrow_styles(elem: Element) -> tuple[int | None, int | None]:
    ln = elem.find(".//a:ln", NS)
    head_end = ln.find("a:headEnd", NS)
    tail_end = ln.find("a:tailEnd", NS)
    head_type = head_end.get("type", "none")
    begin_style = ARROW_HEAD_MAP.get(head_type, 1)
    ...
```
- `<a:ln><a:headEnd type="...">` から取得
- OOXML 名を COM 互換の数値に変換

---

### 8. コネクター接続情報 (begin_id, end_id)

#### COM 実装
```python
connector = shp.api.ConnectorFormat
begin_shape = connector.BeginConnectedShape
if begin_shape is not None:
    begin_name = getattr(begin_shape, "Name", None)
# 後で名前から ID に変換
shape_obj.begin_id = name_to_id.get(begin_name)
```
- `ConnectorFormat.BeginConnectedShape` / `EndConnectedShape` を使用
- 接続先図形の名前を取得し、ID マッピングで変換

#### OOXML 実装
```python
def _get_connector_endpoints(elem: Element) -> tuple[str | None, str | None]:
    cnv_cxn_sp_pr = elem.find("xdr:nvCxnSpPr/xdr:cNvCxnSpPr", NS)
    st_cxn = cnv_cxn_sp_pr.find("a:stCxn", NS)
    end_cxn = cnv_cxn_sp_pr.find("a:endCxn", NS)
    start_id = st_cxn.get("id") if st_cxn is not None else None
    end_id = end_cxn.get("id") if end_cxn is not None else None
    return (start_id, end_id)

# Second pass: resolve connector endpoints
for result in parse_results:
    if result.is_connector:
        if result.start_cxn_id and result.start_cxn_id in excel_id_to_node_id:
            result.shape.begin_id = excel_id_to_node_id[result.start_cxn_id]
```
- `<a:stCxn id="...">` / `<a:endCxn id="...">` から Excel ID を取得
- Excel ID → Node ID マッピングで変換

---

### 9. 回転 (rotation)

#### COM 実装
```python
rot = float(shp.api.Rotation)
if abs(rot) > 1e-6:
    shape_obj.rotation = rot
```

#### OOXML 実装
```python
def _get_rotation(elem: Element) -> float | None:
    xfrm = elem.find(".//a:xfrm", NS)
    rot_str = xfrm.get("rot")
    # OOXML rotation is in 1/60000 of a degree
    rot_emu = int(rot_str)
    rot_deg = rot_emu / 60000.0
    return rot_deg if abs(rot_deg) >= 1e-6 else None
```
- OOXML は 1/60000 度単位で格納
- 度に変換して返却

---

### 10. グループ図形のフラット化

#### COM 実装
```python
def iter_shapes_recursive(shp: xw.Shape) -> Iterator[xw.Shape]:
    yield shp
    if shp.api.Type == 6:  # Group
        items = shp.api.GroupItems
        for i in range(1, items.Count + 1):
            inner = items.Item(i)
            xl_shape = shp.parent.shapes[inner.Name]
            yield from iter_shapes_recursive(xl_shape)
```

#### OOXML 実装
```python
def _parse_group_shapes(grp_sp: Element, mode: str) -> list[_ShapeParseResult]:
    results: list[_ShapeParseResult] = []
    for sp in grp_sp.findall("xdr:sp", NS):
        result = _parse_shape_element(sp, mode, is_cxn_sp=False)
        if result is not None:
            results.append(result)
    for cxn_sp in grp_sp.findall("xdr:cxnSp", NS):
        result = _parse_shape_element(cxn_sp, mode, is_cxn_sp=True)
        if result is not None:
            results.append(result)
    # Recursively parse nested groups
    for nested_grp in grp_sp.findall("xdr:grpSp", NS):
        results.extend(_parse_group_shapes(nested_grp, mode))
    return results
```
- 両実装とも再帰的にグループ内の図形を展開
- ネストされたグループも対応

---

### 11. 出力モードフィルタリング

#### COM 実装
```python
def _should_include_shape(
    text: str, shape_type_num: int | None, ..., output_mode: str = "standard"
) -> bool:
    if output_mode == "light":
        return False
    is_relationship = (shape_type_num in (3, 9) or 
                       "Arrow" in autoshape_type_str or ...)
    if output_mode == "standard":
        return bool(text) or is_relationship
    return True  # verbose
```

#### OOXML 実装
```python
def _should_include_shape(text: str, type_label: str, is_connector: bool, mode: str) -> bool:
    if mode == "light":
        return False
    if mode == "verbose":
        return True
    # standard mode
    if text:
        return True
    if is_connector:
        return True
    if type_label and "Arrow" in type_label:
        return True
    return False
```

| モード | 出力内容 |
|--------|---------|
| light | Shape 抽出をスキップ |
| standard | テキストあり or コネクター/矢印 |
| verbose | 全ての Shape |

---

## Chart 抽出機能

### 1. チャート種別 (chart_type)

#### COM 実装
```python
chart_com = sheet.api.ChartObjects(ch.name).Chart
chart_type_num = chart_com.ChartType
chart_type_label = XL_CHART_TYPE_MAP.get(chart_type_num, f"unknown_{chart_type_num}")
```
- `XlChartType` 列挙値を使用

#### OOXML 実装
```python
CHART_TYPE_MAP = {
    "lineChart": "Line",
    "barChart": "Bar",
    "pieChart": "Pie",
    ...
}

def _get_chart_type(plot_area: Element) -> str:
    for tag, type_name in CHART_TYPE_MAP.items():
        if plot_area.find(f"c:{tag}", NS) is not None:
            return type_name
    return "unknown"
```
- `<c:plotArea>` 内のチャート要素タグで判定

---

### 2. タイトル (title)

#### COM 実装
```python
title = chart_com.ChartTitle.Text if chart_com.HasTitle else None
```

#### OOXML 実装
```python
def _get_chart_title(chart_elem: Element) -> str | None:
    title_elem = chart_elem.find(".//c:title", NS)
    # Try rich text first
    for t_elem in title_elem.findall(".//a:t", NS):
        if t_elem.text:
            return t_elem.text.strip()
    # Try string reference
    str_ref = title_elem.find(".//c:strRef/c:strCache/c:pt/c:v", NS)
    if str_ref is not None and str_ref.text:
        return str_ref.text.strip()
    return None
```
- リッチテキストまたはセル参照からタイトルを取得

---

### 3. Y軸情報 (y_axis_title, y_axis_range)

#### COM 実装
```python
y_axis = chart_com.Axes(2, 1)  # xlValue, xlPrimary
if y_axis.HasTitle:
    y_axis_title = y_axis.AxisTitle.Text
y_axis_range = [y_axis.MinimumScale, y_axis.MaximumScale]
```

#### OOXML 実装
```python
def _get_axis_title(plot_area: Element, axis_type: str) -> str:
    axis = plot_area.find(f"c:{axis_type}", NS)
    title = axis.find("c:title", NS)
    for t_elem in title.findall(".//a:t", NS):
        if t_elem.text:
            return t_elem.text.strip()
    return ""

def _get_axis_range(plot_area: Element, axis_type: str) -> list[float]:
    axis = plot_area.find(f"c:{axis_type}", NS)
    scaling = axis.find("c:scaling", NS)
    min_elem = scaling.find("c:min", NS)
    max_elem = scaling.find("c:max", NS)
    ...
```
- `<c:valAx>` 要素から軸情報を取得

---

### 4. 系列データ (series)

#### COM 実装
```python
for s in chart_com.SeriesCollection():
    parsed = parse_series_formula(getattr(s, "Formula", ""))
    # =SERIES(name, x_range, y_range, plot_order) をパース
    series_list.append(ChartSeries(
        name=s.Name,
        name_range=parsed["name_range"],
        x_range=parsed["x_range"],
        y_range=parsed["y_range"],
    ))
```
- `=SERIES(...)` 数式をパースして範囲を抽出

#### OOXML 実装
```python
def _get_series_data(ser_elem: Element) -> ChartSeries:
    name, name_range = _extract_series_name(ser_elem)
    x_range = _extract_range_from_ref(ser_elem.find("c:cat", NS), ["c:strRef", "c:numRef"])
    y_range = _extract_range_from_ref(ser_elem.find("c:val", NS), ["c:numRef"])
    return ChartSeries(name=name, name_range=name_range, x_range=x_range, y_range=y_range)
```
- `<c:ser>` 要素から直接範囲参照を取得

---

## 実装方法の比較

### データアクセス方式

| 項目 | COM | OOXML |
|------|-----|-------|
| アクセス方式 | オブジェクトモデル | XML パース |
| データソース | Excel プロセス | ZIP 内 XML ファイル |
| 座標単位 | ポイント | EMU |
| 型情報 | 列挙値 (int) | 文字列属性 |

### ファイル構造 (OOXML)

```
sample.xlsx (ZIP)
├── xl/
│   ├── workbook.xml           # ブック情報
│   ├── worksheets/
│   │   └── sheet1.xml         # シートデータ
│   ├── drawings/
│   │   └── drawing1.xml       # 図形定義 (DrawingML)
│   ├── charts/
│   │   └── chart1.xml         # チャート定義 (ChartML)
│   └── _rels/
│       └── workbook.xml.rels  # リレーションシップ
└── [Content_Types].xml
```

### XML 名前空間

```python
# DrawingML
NS = {
    "xdr": "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}

# ChartML
NS = {
    "c": "http://schemas.openxmlformats.org/drawingml/2006/chart",
    "a": "http://schemas.openxmlformats.org/drawingml/2006/main",
}
```

---

## 機能対応表

### Shape 機能

| 機能 | COM | OOXML | 備考 |
|------|:---:|:-----:|------|
| 位置 (l, t) | ✓ | ✓ | |
| サイズ (w, h) | ✓ | ✓ | verbose/Group のみ |
| テキスト | ✓ | ✓ | |
| 種別 (type) | ✓ | ✓ | マッピングテーブルで互換 |
| ID 割り当て | ✓ | ✓ | |
| 方向 (direction) | ✓ | ✓ | コネクター/線のみ |
| 矢印スタイル | ✓ | ✓ | |
| コネクター接続 (begin_id, end_id) | ✓ | ✓ | |
| 回転 (rotation) | ✓ | ✓ | |
| グループフラット化 | ✓ | ✓ | 再帰的に展開 |
| モードフィルタリング | ✓ | ✓ | light/standard/verbose |

### Chart 機能

| 機能 | COM | OOXML | 備考 |
|------|:---:|:-----:|------|
| 位置 (l, t) | ✓ | ✓ | |
| サイズ (w, h) | ✓ | ✓ | verbose のみ |
| チャート種別 | ✓ | ✓ | |
| タイトル | ✓ | ✓ | |
| Y軸タイトル | ✓ | ✓ | |
| Y軸範囲 | ✓ | ✓ | |
| 系列名 | ✓ | ✓ | |
| 系列範囲 (x_range, y_range) | ✓ | ✓ | |

---

## OOXML で実装できない/困難な機能

OOXML (Office Open XML) はファイルフォーマット仕様であり、Excel アプリケーションの動的な計算機能にはアクセスできません。以下に、OOXML で実装が不可能または困難な機能とその技術的理由を説明します。

### 1. 自動計算された Y 軸範囲

#### 機能概要
チャートの Y 軸範囲（最小値・最大値）は、Excel が系列データに基づいて自動計算する場合があります。

#### COM 実装
```python
y_axis = chart_com.Axes(2, 1)  # xlValue, xlPrimary
y_axis_range = [y_axis.MinimumScale, y_axis.MaximumScale]
```
- Excel プロセスが計算した実際の値を取得

#### OOXML の制限
```xml
<!-- 自動スケールの場合、min/max 要素が存在しない -->
<c:scaling>
  <c:orientation val="minMax"/>
  <!-- <c:min val="..."/> が省略される -->
  <!-- <c:max val="..."/> が省略される -->
</c:scaling>
```

#### 実装不可能な理由
- OOXML ファイルには「自動」という設定のみが保存される
- 実際の min/max 値は Excel が描画時に動的に計算する
- 計算には系列データの解析、適切な目盛り間隔の決定など複雑なロジックが必要
- Excel の自動スケーリングアルゴリズムは非公開

#### 現在の動作
- OOXML: `<c:min>` / `<c:max>` 要素が存在する場合のみ値を返す
- 自動スケールの場合は空リスト `[]` を返す

---

### 2. セル参照から解決されたタイトル/ラベル

#### 機能概要
チャートのタイトルや軸ラベルは、セル参照（例: `=Sheet1!$A$1`）で指定できます。

#### COM 実装
```python
title = chart_com.ChartTitle.Text  # 解決済みの文字列
```
- Excel が参照を解決し、実際のテキストを返す

#### OOXML の制限
```xml
<c:title>
  <c:tx>
    <c:strRef>
      <c:f>Sheet1!$A$1</c:f>
      <!-- キャッシュがない場合、値が不明 -->
    </c:strRef>
  </c:tx>
</c:title>
```

#### 実装困難な理由
- OOXML にはセル参照の数式のみが保存される
- 参照先のセル値を取得するには、シートデータ（`xl/worksheets/sheet*.xml`）をパースする必要がある
- 数式が複雑な場合（例: `=CONCATENATE(A1, B1)`）、数式エンジンが必要
- 現在の実装では `<c:strCache>` のキャッシュ値を使用するが、キャッシュが古い可能性がある

#### 現在の動作
- OOXML: キャッシュ値 `<c:strCache><c:pt><c:v>` を優先的に使用
- キャッシュがない場合は `None` を返す

---

### 3. 動的に計算されるフォント/スタイル情報

#### 機能概要
図形のテキストには、テーマやスタイルに基づいて動的に決定されるフォント情報があります。

#### COM 実装
```python
font = shp.api.TextFrame.Characters.Font
font_name = font.Name  # 解決済みのフォント名
font_size = font.Size  # 解決済みのサイズ
```

#### OOXML の制限
```xml
<a:rPr>
  <a:latin typeface="+mn-lt"/>  <!-- テーマフォント参照 -->
</a:rPr>
```

#### 実装困難な理由
- `+mn-lt` はテーマの「本文フォント（ラテン文字）」への参照
- 実際のフォント名を解決するには `xl/theme/theme1.xml` のパースが必要
- テーマの継承、オーバーライド、ロケール依存の処理が複雑
- ExStruct の現在のスコープ外（テキスト内容のみ抽出）

#### 現在の動作
- OOXML: フォント情報は抽出しない（テキスト内容のみ）

---

### 4. OLE オブジェクト / 埋め込みオブジェクト

#### 機能概要
Excel には他のアプリケーション（Word、PowerPoint、PDF など）のオブジェクトを埋め込むことができます。

#### COM 実装
```python
if shp.api.Type == 12:  # msoOLEControlObject
    ole = shp.api.OLEFormat
    # OLE オブジェクトにアクセス可能
```

#### OOXML の制限
```xml
<mc:AlternateContent>
  <mc:Choice Requires="v">
    <!-- VML 形式の OLE オブジェクト -->
  </mc:Choice>
</mc:AlternateContent>
```

#### 実装不可能な理由
- OLE オブジェクトはバイナリ形式で埋め込まれる
- オブジェクトの内容を解釈するには、各アプリケーション固有のパーサーが必要
- セキュリティ上の理由から、OLE オブジェクトの自動実行は危険
- ExStruct のスコープ外（Shape/Chart のみ対象）

#### 現在の動作
- OOXML: OLE オブジェクトはスキップ（抽出対象外）

---

### 5. マクロ / VBA コード

#### 機能概要
`.xlsm` ファイルには VBA マクロが含まれる場合があります。

#### COM 実装
```python
# VBA プロジェクトにアクセス可能（セキュリティ設定による）
vba = workbook.api.VBProject
```

#### OOXML の制限
- VBA コードは `xl/vbaProject.bin` にバイナリ形式で保存
- OOXML 標準の一部ではない（Microsoft 拡張）

#### 実装不可能な理由
- VBA バイナリ形式は非公開仕様
- マクロの実行には VBA ランタイムが必要
- セキュリティリスクが高い
- ExStruct のスコープ外

#### 現在の動作
- OOXML: マクロは無視（データ抽出のみ）

---

### 6. 条件付き書式の評価結果

#### 機能概要
セルや図形には条件付き書式が適用され、データに応じて色やスタイルが変化します。

#### COM 実装
```python
# 現在の表示状態を取得
interior_color = cell.api.DisplayFormat.Interior.Color
```

#### OOXML の制限
```xml
<conditionalFormatting>
  <cfRule type="cellIs" operator="greaterThan">
    <formula>100</formula>
  </cfRule>
</conditionalFormatting>
```

#### 実装不可能な理由
- OOXML にはルール定義のみが保存される
- 実際の評価結果は Excel が動的に計算
- 条件式の評価には数式エンジンが必要
- ExStruct のスコープ外（書式情報は抽出対象外）

#### 現在の動作
- OOXML: 条件付き書式は無視

---

### 7. 数式の計算結果（キャッシュなしの場合）

#### 機能概要
セルに数式が含まれる場合、計算結果が必要になることがあります。

#### COM 実装
```python
value = cell.value  # 計算済みの値
formula = cell.formula  # 数式
```

#### OOXML の制限
```xml
<c:v>123</c:v>  <!-- キャッシュ値 -->
<c:f>SUM(A1:A10)</c:f>  <!-- 数式 -->
```

#### 実装困難な理由
- OOXML にはキャッシュ値が保存されるが、ファイル保存後に元データが変更された場合は古い
- 数式を再計算するには完全な数式エンジンが必要
- Excel の数式は 400 以上の関数をサポート
- 外部参照、名前付き範囲、配列数式など複雑な機能がある

#### 現在の動作
- OOXML: キャッシュ値を使用（存在する場合）
- キャッシュがない場合は数式文字列を返す

---

### 8. 複合チャートの詳細な種別判定

#### 機能概要
Excel では複数のチャート種別を組み合わせた複合チャートを作成できます。

#### COM 実装
```python
chart_type = chart_com.ChartType  # 主要なチャート種別
# 各系列ごとの種別も取得可能
for series in chart_com.SeriesCollection():
    series_type = series.ChartType
```

#### OOXML の制限
```xml
<c:plotArea>
  <c:barChart>...</c:barChart>
  <c:lineChart>...</c:lineChart>  <!-- 複数のチャート要素 -->
</c:plotArea>
```

#### 実装困難な理由
- OOXML では複数のチャート要素が並列に存在
- どれが「主要」かの判定ロジックが不明確
- COM は Excel の内部判定結果を返すが、そのアルゴリズムは非公開

#### 現在の動作
- OOXML: 最初に見つかったチャート種別を返す
- 複合チャートの場合、実際の主要種別と異なる可能性がある

---

## 制限事項と注意点

### OOXML 実装の制限（その他）

1. **座標精度**
   - EMU → ピクセル変換で若干の誤差が生じる可能性
   - COM は Excel 内部の正確な値を取得

2. **図形種別マッピング**
   - 全ての `MsoAutoShapeType` に対応するマッピングが必要
   - 未知の種別は `AutoShape-{prst}` 形式で出力

3. **チャート種別**
   - 複合チャートの場合、最初に見つかった種別を返す
   - COM は主要なチャート種別を返す

### 使い分けの指針

| 環境 | 推奨実装 | 理由 |
|------|---------|------|
| Windows + Excel | COM | 高精度、全機能対応 |
| Windows (Excel なし) | OOXML | Excel 不要 |
| Linux / macOS | OOXML | 唯一の選択肢 |
| CI/CD 環境 | OOXML | ヘッドレス実行可能 |

### フォールバック動作

`integrate.py` では自動的にフォールバックが行われます：

```python
# COM が利用可能な場合
if _is_com_available():
    shapes = get_shapes_with_position(workbook, mode)
    charts = get_charts(sheet, mode)
# COM が利用不可の場合
else:
    shapes = get_shapes_ooxml(xlsx_path, mode)
    charts = get_charts_ooxml(xlsx_path, mode)
```

---

## 参考資料

- [ECMA-376 Office Open XML File Formats](https://www.ecma-international.org/publications-and-standards/standards/ecma-376/)
- [DrawingML Reference](https://docs.microsoft.com/en-us/openspecs/office_standards/ms-odrawxml/)
- [xlwings Documentation](https://docs.xlwings.org/)
- [MsoShapeType Enumeration](https://docs.microsoft.com/en-us/office/vba/api/office.msoshapetype)
