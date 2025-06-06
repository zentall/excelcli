# excelcli

## English

A simple CLI tool for extracting data from Excel files (`.xlsx`) and exporting it to CSV.  
You can extract specific rows or columns from multiple Excel files using glob patterns.

### Features

- Extract rows or columns from Excel files
- Output as CSV
- Supports glob patterns for batch processing
- Optionally include the source file name in the CSV

### Usage

#### Extract columns

```sh
excelcli extract-col ./reports/*.xlsx --sheet Sheet1 --col-range A:C --rows 1,2,3 --headers "Name,Sales,Dept" --output output.csv --with-filename
```

#### Extract rows

```sh
excelcli extract-row ./reports/*.xlsx --sheet Sheet1 --range 2:5 --columns B,D,E --headers "Name,Sales,Dept" --output output.csv --with-filename
```

#### Options

- `--sheet`: Sheet name (default: Sheet1)
- `--col-range`: Column range (e.g., A:C)
- `--rows`: Row numbers (comma-separated, e.g., 1,3,5)
- `--range`: Row range (e.g., 2:10)
- `--columns`: Columns to extract (comma-separated, e.g., B,D,E)
- `--headers`: CSV header names (comma-separated)
- `--output`: Output CSV file
- `--with-filename`: Include the Excel file name in the CSV

### Test

Integration tests are in `tests/integration_test.rs`.  
Put your test Excel file in `tests/data/test.xlsx`.

---

## 日本語

Excelファイル（`.xlsx`）からデータを抽出し、CSVに出力するシンプルなCLIツールです。  
複数ファイルをglobパターンで一括処理できます。

### 特長

- Excelファイルから行または列を抽出
- CSV形式で出力
- globパターンで複数ファイル対応
- オプションでファイル名をCSVに含めることが可能

### 使い方

#### 列抽出

```sh
excelcli extract-col ./reports/*.xlsx --sheet Sheet1 --col-range A:C --rows 1,2,3 --headers "名前,売上,部署" --output output.csv --with-filename
```

#### 行抽出

```sh
excelcli extract-row ./reports/*.xlsx --sheet Sheet1 --range 2:5 --columns B,D,E --headers "名前,売上,部署" --output output.csv --with-filename
```

#### 主なオプション

- `--sheet`: シート名（デフォルト: Sheet1）
- `--col-range`: 列範囲（例: A:C）
- `--rows`: 行番号（カンマ区切り、例: 1,3,5）
- `--range`: 行範囲（例: 2:10）
- `--columns`: 抽出する列（カンマ区切り、例: B,D,E）
- `--headers`: CSVヘッダー名（カンマ区切り）
- `--output`: 出力CSVファイル
- `--with-filename`: CSVにExcelファイル名を含める

### テスト

統合テストは `tests/integration_test.rs` にあります。  
テスト用Excelファイルは `tests/data/test.xlsx` に配置してください。