# Excelサンプルと空ひな形

このディレクトリのExcelファイルは、`ExcelToJson`へ渡す入力ブックのサンプルです。製品仕様の正本は [`../docs/REQUIREMENTS.md`](../docs/REQUIREMENTS.md) です。

## `sample.xlsx`

要件定義14章の変換例に対応する、変換可能なサンプルです。

| シート | 内容 |
|---|---|
| `setting` | `rootType=object`、`emptyCell=omit` |
| `root` | `profile`へのobject参照、`items`へのarray参照、`nickNames`へのscalar-array参照 |
| `profile` | `age`のnumber変換と`birthday`のdate変換 |
| `items` | 同じIDを持つ2行のarray参照 |
| `nickNames` | 同じIDを持つ2行のscalar-array参照 |

`root`の`nickNames`列には`scalar-array:nickNames`の旧式Excelメモがあり、参照先は`ID`と`value`の2列です。`value`の型メモは省略されているため、既定の`text`として変換されます。

```powershell
ExcelToJson.exe .\samples\sample.xlsx
```

変換すると、`samples/sample.json`へ次の内容が出力されます。

```json
{
  "name": "Alice",
  "profile": {
    "age": 30,
    "birthday": "1996-05-10"
  },
  "items": [
    {
      "name": "Apple",
      "quantity": 2
    },
    {
      "name": "Orange",
      "quantity": 3
    }
  ],
  "nickNames": [
    "Allie",
    "Ali"
  ]
}
```

## `scalar-array.xlsx`

`rootType=scalar-array`で、`root`シートの`value`列をJSONルートの文字列配列へ変換するサンプルです。

```powershell
ExcelToJson.exe .\samples\scalar-array.xlsx
```

変換結果は次のとおりです。

```json
[
  "Allie",
  "Ali"
]
```

## `template.xlsx`

新しい入力ブックを作り始めるための空ひな形です。

- `setting`には`rootType=array`、`emptyCell=omit`と、空の日付書式設定があります。
- `root`は`ID`ヘッダーだけを持ち、データ行はありません。
- `_guide`は変換対象外で、列、ヘッダーの既定型、データセルの型上書き、参照、scalar-array、ID、空行終端の入力方法を説明します。
- 空のまま変換した場合は`[]`を出力します。
- `rootType=object`へ変更する場合は、`root`へデータ行をちょうど1行追加してください。

作業時は元ファイルを残すため、`template.xlsx`を別名でコピーして使用してください。
