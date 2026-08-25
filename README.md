# ExcelToJson

Excelブックの表形式データを、シート間参照による階層構造を持つJSONへ変換するWindows向けCLIです。

製品仕様の正本は [`docs/REQUIREMENTS.md`](docs/REQUIREMENTS.md)、開発・変更時の作業規約は [`AGENTS.md`](AGENTS.md) です。

## 主な機能

- `.xlsx` の `setting` と `root` シートを読み取ってJSONへ変換
- `text` / `number` / `boolean` / `date` の明示的な型変換
- ヘッダーメモを既定型とし、データセルのメモでそのセルだけJSON型を上書き
- `object:シート名` / `array:シート名` による単段・多段参照
- 実レコード経路による循環参照検出
- UTF-8 BOMなし、インデント付きJSON
- 変換成功後だけ既存JSONを安全に置換
- Self-contained / Single-fileのWindows x64実行ファイル

入力ブックの詳細な規則は正本を参照してください。

## サンプルと空ひな形

リポジトリの [`samples`](samples/README.md) には、次のExcelファイルがあります。

- `sample.xlsx`: シート間参照と型変換を含む、要件定義14章に対応した実行可能なサンプル
- `template.xlsx`: `setting`、空の `root`、入力方法を説明する `_guide` を備えた作業開始用ひな形

どちらもリポジトリ内の参考ファイルであり、Windows向けのpublish成果物には含まれません。

## 対応環境

### 実行環境

- Windows 10 22H2以降のx64（Windows 11を含む）
- publish済みの `ExcelToJson.exe` を使う場合、.NET Runtime、Visual Studio、Microsoft Excelは不要

### 配布用EXEの作成環境

アプリケーションを開発せず、ソースコードから配布用の `ExcelToJson.exe` を作成するだけの場合は、次の環境が必要です。

- Windows x64
- .NET SDK 9.0.3xx以降の9.0系
- このリポジトリのソースコード
- 初回のNuGet restoreに必要なインターネット接続

Visual StudioとMicrosoft Excelは不要です。.NET Runtimeだけでは作成できないため、必ず.NET **SDK** をインストールしてください。

### 開発環境

- Visual Studio 2022 17.14の最新サービス版
- Visual Studio Installerの「.NET デスクトップ開発」ワークロード
- .NET SDK 9.0.3xx以降の9.0系

リポジトリの `global.json` はSDK 9.0.300を下限とし、同じ9.0系列の最新feature bandへ追従します。NuGet依存は中央管理され、`packages.lock.json` と一致しないrestoreは失敗します。

> [!IMPORTANT]
> .NET 9のサポートは2026年11月10日に終了します。Visual Studio 2022での正式なIDEビルドを維持するためV1では.NET 9を採用していますが、継続保守では.NET 10以降への移行が必要です。

## 配布用の単体EXEを作成する

ここでは、開発やテストコードの変更を行わず、ソースコードからWindows x64向けの配布用EXEを作成する手順を示します。

PowerShellでリポジトリルートへ移動し、使用されるSDKを確認します。

```powershell
dotnet --version
```

`9.0.3xx`以降の9.0系が表示されることを確認してください。続いて、依存関係を復元し、空の出力先を指定してpublishします。

```powershell
dotnet restore src/ExcelToJson.Cli/ExcelToJson.Cli.csproj `
  --locked-mode `
  --runtime win-x64

dotnet publish src/ExcelToJson.Cli/ExcelToJson.Cli.csproj `
  --configuration Release `
  --runtime win-x64 `
  --self-contained true `
  --no-restore `
  --output artifacts/ExcelToJson-win-x64 `
  /p:PublishSingleFile=true `
  /p:PublishTrimmed=false `
  /p:IncludeNativeLibrariesForSelfExtract=true
```

`dotnet publish`にはソースコードのReleaseビルドも含まれるため、単体EXEの作成だけが目的なら、事前の`dotnet build`は不要です。このコマンドはpublish条件をすべて明示しており、Visual Studioのpublish profileには依存しません。

成果物は次へ生成されます。

```text
artifacts/ExcelToJson-win-x64/ExcelToJson.exe
```

出力先のファイルを確認します。

```powershell
Get-ChildItem artifacts/ExcelToJson-win-x64 -File
```

空の出力先へpublishした場合、配布に必要なファイルは `ExcelToJson.exe` だけです。このEXEには.NETランタイムと依存ライブラリが含まれます。trimmingは、ClosedXMLなどが必要とするコードを誤って除去しないよう無効にしています。

`ExcelToJson.exe`だけを任意のディレクトリへコピーし、入力ブックを指定して実行できることを確認してください。

```powershell
artifacts/ExcelToJson-win-x64/ExcelToJson.exe C:\data\sample.xlsx
```

成功すると `C:\data\sample.json` が生成されます。実行先に.NET Runtime、Visual Studio、Microsoft Excelは不要です。

> [!NOTE]
> `dotnet build`の出力ディレクトリにある `ExcelToJson.exe` は、配布用の単体EXEとは限りません。配布には、必ず上記のSelf-contained / Single-file指定で生成したpublish出力を使用してください。

## 開発者向け: プロジェクト構成

```text
ExcelToJson.slnx
├─ src/ExcelToJson.Cli    CLI境界、終了コード、コンソール表示
├─ src/ExcelToJson.Core   Excel読取、解析、変換、参照解決、JSON出力
└─ tests/ExcelToJson.Tests
                          単体、ClosedXML統合、CLIテスト
```

依存方向は `Cli -> Core` です。CoreではClosedXMLの型をWorkbook読取境界の内側に閉じ込め、変換と参照解決には中立な内部モデルを渡します。System.Text.Jsonによる書き込みも変換規則から分離しています。

## 開発者向け: ビルドとテスト

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

## 数式セルの制約

数式セルはClosedXMLによる再計算結果を優先し、再計算できない場合はExcel保存済みのキャッシュ結果を使用します。ClosedXMLが式を評価できず、有効なキャッシュ結果もない場合は変換エラーです。配布先でExcelを起動して再計算することはありません。
