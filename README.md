# Excel to CSV Converter

ExcelファイルをCSVファイルに変換するシンプルなPowerShellツールです。

## ファイル構成

```
ExcelToCSV/
├── ExcelToCsv-Simple.bat     # ドラッグ&ドロップ用バッチファイル
├── Convert-ExcelToCsv.ps1    # PowerShell変換スクリプト
└── README.md                 # このファイル
```

## 機能

- ✅ **ドラッグ&ドロップ対応** - バッチファイルにExcelファイルをドロップするだけ
- ✅ **相対パス対応** - カレントディレクトリや相対パスでファイル指定可能
- ✅ **複数シート対応** - 特定シートまたは全シートの変換
- ✅ **自動出力** - 元ファイルと同じフォルダにCSV出力
- ✅ **エラーハンドリング** - 適切なエラーメッセージとCOMオブジェクト解放

## 使用方法

### 1. ドラッグ&ドロップ（推奨）

1. `ExcelToCsv-Simple.bat` をデスクトップにショートカット作成
2. 変換したいExcelファイルをバッチファイルにドラッグ&ドロップ
3. 自動的に同じフォルダにCSVファイルが作成される

### 2. コマンドライン実行

```powershell
# 基本変換（最初のシートをCSVに変換）
powershell.exe -ExecutionPolicy Bypass -File "Convert-ExcelToCsv.ps1" -ExcelFilePath "data.xlsx"

# 出力先指定
powershell.exe -ExecutionPolicy Bypass -File "Convert-ExcelToCsv.ps1" -ExcelFilePath "data.xlsx" -OutputPath "C:\output\"

# 特定シート変換
powershell.exe -ExecutionPolicy Bypass -File "Convert-ExcelToCsv.ps1" -ExcelFilePath "data.xlsx" -WorksheetName "Sheet2"

# 全シート変換
powershell.exe -ExecutionPolicy Bypass -File "Convert-ExcelToCsv.ps1" -ExcelFilePath "data.xlsx" -AllWorksheets
```

## パラメータ

| パラメータ名 | 必須 | 説明 |
|-------------|------|------|
| `ExcelFilePath` | ○ | 変換するExcelファイルのパス |
| `OutputPath` | × | CSV出力先フォルダ（省略時は元ファイルと同じフォルダ） |
| `WorksheetName` | × | 変換するシート名（省略時は最初のシート） |
| `AllWorksheets` | × | 全シートを変換する場合に指定 |

## 動作環境

- **OS**: Windows 10/11
- **PowerShell**: Windows PowerShell 5.1以上
- **Excel**: Microsoft Excelがインストールされている環境
- **ファイル形式**: .xlsx, .xls

## セットアップ

1. フォルダを任意の場所に配置
2. PowerShellの実行ポリシーを設定（初回のみ）:
   ```powershell
   Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
   ```
3. `ExcelToCsv-Simple.bat`のショートカットをデスクトップに作成（推奨）

## 出力仕様

- **出力ファイル名**: `元ファイル名.csv`
- **全シート変換時**: `元ファイル名_シート名.csv`
- **出力場所**: 元ファイルと同じフォルダ（OutputPath未指定時）
- **文字コード**: CSV形式（Excel標準）

## エラー対処

### 実行ポリシーエラー
```
running scripts is disabled on this system
```
**対処法**: PowerShellを管理者権限で起動し、以下を実行
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### Excel COMエラー
```
Excel file not found / COM object creation failed
```
**対処法**:
- Microsoft Excelがインストールされているか確認
- ファイルパスが正しいか確認
- 32bit PowerShellを試す（64bit環境で32bit Excelの場合）

### ファイルアクセスエラー
**対処法**:
- Excelファイルが他のプロセスで開かれていないか確認
- ファイルの読み取り権限があるか確認

## 変換例

```
入力: data.xlsx (Sheet1, Sheet2, Sheet3)
実行: ドラッグ&ドロップ
出力: data.csv (Sheet1の内容)

入力: data.xlsx
実行: -AllWorksheets指定
出力: data_Sheet1.csv, data_Sheet2.csv, data_Sheet3.csv
```

