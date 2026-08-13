# ExcelToJson

Excelブックの表形式データを、シート間参照による階層構造を持つJSONへ変換するWindows向けCLIです。

製品仕様の正本は [`docs/REQUIREMENTS.md`](docs/REQUIREMENTS.md)、開発・変更時の作業規約は [`AGENTS.md`](AGENTS.md) です。

## 主な機能

- `.xlsx` の `setting` と `root` シートを読み取ってJSONへ変換
- `text` / `number` / `boolean` / `date` の明示的な型変換
- `object:シート名` / `array:シート名` による単段・多段参照
- 実レコード経路による循環参照検出
- UTF-8 BOMなし、インデント付きJSON
- 変換成功後だけ既存JSONを安全に置換
- Self-contained / Single-fileのWindows x64実行ファイル

入力ブックの詳細な規則は正本を参照してください。

## 対応環境

### 実行環境

- Windows 10 22H2以降のx64（Windows 11を含む）
- publish済みの `ExcelToJson.exe` を使う場合、.NET Runtime、Visual Studio、Microsoft Excelは不要

### 開発環境

- Visual Studio 2022 17.14の最新サービス版
- Visual Studio Installerの「.NET デスクトップ開発」ワークロード
- .NET SDK 9.0.3xx以降の9.0系

リポジトリの `global.json` はSDK 9.0.300を下限とし、同じ9.0系列の最新feature bandへ追従します。NuGet依存は中央管理され、`packages.lock.json` と一致しないrestoreは失敗します。

> [!IMPORTANT]
> .NET 9のサポートは2026年11月10日に終了します。Visual Studio 2022での正式なIDEビルドを維持するためV1では.NET 9を採用していますが、継続保守では.NET 10以降への移行が必要です。

## プロジェクト構成

```text
ExcelToJson.slnx
├─ src/ExcelToJson.Cli    CLI境界、終了コード、コンソール表示
├─ src/ExcelToJson.Core   Excel読取、解析、変換、参照解決、JSON出力
└─ tests/ExcelToJson.Tests
                          単体、ClosedXML統合、CLIテスト
```

依存方向は `Cli -> Core` です。CoreではClosedXMLの型をWorkbook読取境界の内側に閉じ込め、変換と参照解決には中立な内部モデルを渡します。System.Text.Jsonによる書き込みも変換規則から分離しています。

## ビルドとテスト

PowerShellでリポジトリルートから実行します。

```powershell
dotnet restore ExcelToJson.slnx --locked-mode
dotnet build ExcelToJson.slnx --configuration Release --no-restore
dotnet test ExcelToJson.slnx --configuration Release --no-restore
```

Visual Studio 2022では `ExcelToJson.slnx` を開き、ソリューションのビルドとTest Explorerから同じテストを実行できます。テスト基盤はxUnit v3とMicrosoft Testing Platformです。

依存関係を意図的に更新する場合だけ、一時的にlocked modeを無効化してlockファイルを再生成し、差分をレビューしてください。通常の開発ではlockファイルを変更しません。

任意で.NET CLIとテスト基盤のテレメトリを無効化できます。

```powershell
$env:DOTNET_CLI_TELEMETRY_OPTOUT = '1'
$env:TESTINGPLATFORM_TELEMETRY_OPTOUT = '1'
```

## 実行

```powershell
ExcelToJson.exe C:\data\sample.xlsx
```

成功すると `C:\data\sample.json` を生成し、その絶対パスを標準出力へ表示します。失敗時は日本語の診断を標準エラーへ表示し、終了コード1を返します。

フレームワーク依存のDebugビルドは次のようにも実行できます。

```powershell
dotnet run --project src/ExcelToJson.Cli -- C:\data\sample.xlsx
```

## Windows x64向けpublish

```powershell
dotnet publish src/ExcelToJson.Cli/ExcelToJson.Cli.csproj `
  --configuration Release `
  --no-restore `
  /p:PublishProfile=win-x64
```

成果物は次へ生成されます。

```text
src/ExcelToJson.Cli/bin/Release/net9.0/win-x64/publish/ExcelToJson.exe
```

publish profileはSelf-contained、Single-file、trimming無効です。成果物の確認では `ExcelToJson.exe` だけを別ディレクトリへコピーし、単独で変換できることを検証します。

## 数式セルの制約

数式セルはClosedXMLによる再計算結果を優先し、再計算できない場合はExcel保存済みのキャッシュ結果を使用します。ClosedXMLが式を評価できず、有効なキャッシュ結果もない場合は変換エラーです。配布先でExcelを起動して再計算することはありません。
