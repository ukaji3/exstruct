# API Design — ExStruct

ExStruct のトップレベル Python API 設計。

## 基本 API

```python
import exstruct as xs

data = xs.extract("file.xlsx")
```

### 戻り値

`WorkbookData`（Pydantic model）

---

## CLI

```bash
exstruct file.xlsx --format json > output.json
```

オプション：

- `--image` → PNG 出力
- `--pdf` → PDF 出力
- `--dpi` → PNG 解像度
- `--sheets-dir` → シートごとの分割出力
- `--multiple` → 複数ファイル処理

---

## JSON/YAML/TOON 出力

```python
data = xs.extract("file.xlsx")

xs.export("file.xlsx")
data.to_json(pretty=True)
data.to_yaml()
data.to_toon()
data.save("file.json")
data["Sheet1"]          # WorkbookData.__getitem__
for name, sheet in data:  # WorkbookData.__iter__
    print(name, len(sheet.rows))
```

---

## 今後拡張予定

- `xs.detect_tables()`
- `xs.semantic_shapes()`
- `xs.smartart()`
