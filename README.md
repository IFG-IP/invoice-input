# 請求書入力ツール

請求書・領収書取込の最初の技術検証です。

## 検証できること

- ブラウザからフォルダ単位または複数ファイルを投入
- 複数回に分けて選択したファイルを同じ一覧へ追加
- ユーザーが請求書・領収書のPDF/画像を画面からアップロード
- 拡張子で画像ルートとPDF解析ルートに仕分け
- PDF.jsでPDF内部テキストを抽出
- 抽出文字数が20文字以上ならデジタルPDFとして判定
- 20文字未満なら画像PDFとして1ページ目をPNG相当のCanvasレンダリングに回し、画像ルートへ合流
- JPG / PNG / HEIC / HEIFをCanvasでリサイズし、JPEGとして軽量化

## 使い方

操作方法だけを簡単に確認したい場合は、納品用の簡易マニュアル [MANUAL.md](./MANUAL.md) を参照してください。

手動で起動する場合は、このフォルダで以下を実行してください。

```powershell
cd "invoice-input(アプリケーション本体）"
powershell -NoProfile -ExecutionPolicy Bypass -File .\serve.ps1
```
PDF解析とHEIC変換はCDNライブラリを使うため、初回表示時にインターネット接続が必要です。

## AI精度検証

1. 画像またはPDFを読み込みます。
2. `.env` にGeminiの接続設定を入れます。
3. 必要に応じてモデルを変更します。初期値は `gemini-3.1-flash-lite` です。
4. `次へ` を押すと、Excelフォーマットの項目に合わせて抽出し、目視確認画面へ進みます。

抽出したい項目は、VS Code上でアプリと同じフォルダに置いたExcel/CSVの1行目から自動で読み込みます。ユーザーが画面でExcelを選択する必要はありません。

起動時には、アプリと同じフォルダに置いたExcel/CSVを既定テンプレートとして自動で読み込みます。複数ある場合は更新日時が新しいものを使います。ローカルにExcel/CSVがない場合は `templates/extraction-fields.csv` を読み込みます。

例:

```text
日付 | 取引先 | 合計金額 | 税額 | 請求書番号 | 支払期日
```

既知の列名は `date` や `total_amount` などの標準キーに寄せます。未知の列名は `field_1` のようなキーで抽出し、画面上ではExcelの列名を表示します。

Generative Language APIを使う場合の `.env` は以下の形式です。
ローカルとAWS Lambdaの検証では、まずこちらを使うのがおすすめです。

```text
GEMINI_PROVIDER=generative
GEMINI_DEFAULT_MODEL=gemini-3.1-flash-lite
GEMINI_API_KEY=...
TEMPLATE_EXCEL_URL=https://...
```

Vertex AI APIを使う場合はAPIキーではなくOAuth認証が必要です。Google Cloud CLIで認証したうえで以下の形式にします。

```text
GEMINI_PROVIDER=vertex
GEMINI_DEFAULT_MODEL=gemini-3.1-flash-lite
GOOGLE_CLOUD_PROJECT=your-project-id
GOOGLE_CLOUD_LOCATION=asia-northeast1
TEMPLATE_EXCEL_URL=https://...
```

Google Cloud CLIを使う場合は、事前に以下を実行してください。

```powershell
gcloud auth application-default login
gcloud config set project your-project-id
```

一時的にアクセストークンを直接使う場合は、`.env` に `VERTEX_ACCESS_TOKEN=...` を追加できます。Vertex AIの `generateContent` は `VERTEX_API_KEY` では呼べないため、APIキーで動かす場合は `GEMINI_PROVIDER=generative` と `GEMINI_API_KEY` を使ってください。

期待値と比較したい場合は、`期待値JSON` に以下の形式で入力してください。

```json
{
  "sample.jpg": {
    "date": "2026-05-01",
    "vendor": "ABC商店",
    "total_amount": 12345,
    "tax_amount": 1122,
    "currency": "JPY"
  }
}
```

キーにはファイル名、フォルダ込みのファイル名、または拡張子なしの名前を使えます。

GeminiのAPIキーやVertex AIの認証情報は `serve.ps1` がサーバー側で読み込みます。ブラウザや `app.js` には渡しません。

## 画面とログ

ユーザー画面には、アップロード、次へ、目視確認、Excel出力だけを表示します。

PDF/画像の判定ルート、抽出JSON、期待値比較、エラー詳細などの開発者向け情報は、`serve.ps1` を起動しているPowerShellターミナルへ出力します。

## 目視確認

書類をアップロードすると `目視確認` のファイル一覧に表示されます。AI抽出後は、抽出済みの内容を同じ画面で確認・修正できます。

- 左側に請求書・領収書のプレビュー
- 右側に抽出データの編集フォーム
- 複数ファイルの場合は、目視確認上部のファイル一覧で切り替え
- 現在編集中のファイルは太い枠で表示
- 右側の抽出データ欄にも編集中のファイル名を表示
- 抽出データを手で修正すると自動で反映
- 必須項目が未入力の場合は赤枠で表示
- `JSONコピー` で修正後データをコピー
- `skyberry Excel出力` で、VS Code上に置いたExcelフォーマットの2行目以降へ抽出データを流し込んだExcelを出力

## skyberry Excel出力

アプリと同じフォルダに置いたExcelの1行目を列名として読み込み、AI抽出・目視修正済みのデータを2行目以降に書き込みます。

ブラウザからローカルの元Excelを直接上書きすることはできないため、出力時は `skyberry-import-YYYYMMDD-HHMM.xlsx` という新しいExcelファイルをダウンロードします。

## 次の検証候補

- サーバー側OCRの比較検証
- skyberry標準項目へのマッピングJSON設計
- 確認・修正画面のテーブル編集
- Excel出力後の必須項目チェック

# プログラム解説

## 1. システム概要

このプログラムは、請求書・領収書のPDFまたは画像をブラウザ上で取り込み、AIで必要項目を抽出し、目視確認・手修正を行ったうえで、skyberry取込用Excelを出力する検証ツールです。ローカル実行では `serve.ps1` をAPIプロキシとして使い、AWSデプロイ時は同等のAPIプロキシをAWS Lambda / API Gatewayで提供する構成を想定しています。

主な処理は以下です。

1. 請求書・領収書ファイルをブラウザから追加する。
2. PDF.jsやCanvasでPDF・画像を解析し、AIに渡しやすい形式へ整える。
3. ローカルサーバー `serve.ps1` 経由でGemini APIまたはVertex AIへ抽出依頼を送る。
4. Excelフォーマットの列名に合わせて抽出結果を作成する。
5. 目視確認画面でファイルごとに内容を確認・修正する。
6. 修正済みデータをskyberry取込用Excelとしてダウンロードする。

---

## 2. 技術スタック

### フロントエンド

- HTML / CSS / JavaScript
- PDF.js: PDFの読み込み、テキスト抽出、ページ画像化
- heic2any: HEIC / HEIF画像のブラウザ変換
- SheetJS: Excelテンプレートの読み込み、skyberry取込用Excelの生成
- Canvas API: 画像リサイズ、PDFページのJPEG化
- File API / Drag and Drop API: 複数ファイル・フォルダ単位の投入

### バックエンド

- PowerShell
- `serve.ps1`: ローカルHTTPサーバー兼APIプロキシ
- `System.Net.Sockets.TcpListener`: ローカルサーバー実装
- `.env`: Gemini / Vertex AI / テンプレートURLなどの設定値管理
- AWS Lambda: AWSデプロイ時のGemini APIプロキシ
- Amazon API Gateway: ブラウザからLambdaを呼び出すHTTP API
- AWS Amplify HostingまたはS3 + CloudFront: フロントエンド配信

### 外部API

- Google Gemini API: 請求書・領収書の項目抽出
- Vertex AI Gemini API: Google Cloud認証でGeminiを利用する場合の代替経路
- `TEMPLATE_EXCEL_URL`: SharePointなどに置いたExcelテンプレートを取得する場合に利用
- Whisper API、Dify API、Notion APIは現在の実装では使用していません。

### 開発・実行環境

- Windows
- PowerShell
- Webブラウザ
- ローカルURL: `http://127.0.0.1:5173/`
- AWS: Amplify Hosting / API Gateway / Lambda
- Google Cloud CLI: Vertex AIを使う場合のみ必要

---

## 3. システムアーキテクチャ

```text
ユーザー
  |
  | PDF / JPG / PNG / HEICを投入
  v
ブラウザ画面 index.html / app.js
  |
  | PDF.js / heic2any / Canvas / SheetJSで前処理
  |
  | POST /api/gemini/generate
  v
ローカルサーバー serve.ps1
  |
  | .envからAPIキー・認証情報を読み込み
  |
  +--> Gemini API または Vertex AI Gemini API
  |
  +--> Excelテンプレート取得
  |
  v
ブラウザへ抽出結果を返却
  |
  | 目視確認・手修正
  v
skyberry-import-YYYYMMDD-HHMM.xlsx をダウンロード
```

1. ブラウザはファイル投入、PDF・画像の前処理、目視確認、Excel出力を担当します。
2. `serve.ps1` はAPIキーをブラウザへ渡さず、Gemini APIへの中継だけを行います。
3. Excelテンプレートは、ローカルに置いたExcel/CSV、`TEMPLATE_EXCEL_URL`、またはフォールバックCSVから読み込みます。
4. 出力Excelは元テンプレートを直接上書きせず、新しいファイルとしてダウンロードします。

### AWSデプロイ時の構成

AWS上では、ブラウザで動くフロントエンドと、Gemini APIを呼び出すバックエンドを分けて配置します。APIキーをフロントエンドに含めないため、ブラウザは直接Gemini APIを呼ばず、API Gateway経由でLambdaを呼び出します。

```text
ユーザー
  |
  | ブラウザでアクセス
  v
AWS Amplify Hosting または S3 + CloudFront
  |
  | index.html / styles.css / app.js / config.js を配信
  |
  | POST {apiBaseUrl}/api/gemini/generate
  | GET  {apiBaseUrl}/api/template/extraction-fields
  | GET  {apiBaseUrl}/api/health
  v
Amazon API Gateway
  |
  v
AWS Lambda
  |
  | 環境変数またはSecrets ManagerからAPIキーを取得
  v
Gemini API / Vertex AI Gemini API
```

AWSデプロイ時の役割は以下です。

1. フロントエンドはAmplify Hosting、またはS3 + CloudFrontで静的ファイルとして配信します。
2. `config.js` の `apiBaseUrl` にAPI GatewayのURLを設定します。
3. Lambdaはローカル版 `serve.ps1` の `/api/gemini/generate`、`/api/template/extraction-fields`、`/api/health` 相当の処理を担当します。
4. Gemini APIキーやVertex AI認証情報はLambdaの環境変数、またはAWS Secrets Managerに保存します。
5. API Gatewayでは、フロントエンドのドメインから呼び出せるようにCORSを設定します。
6. 請求書PDFや画像はブラウザ内で前処理し、必要なデータのみAPIへ送ります。恒久保存は現在の実装では行いません。

---

## 4. API詳細と利用コスト

### Whisper API

- 現在は使用していません。
- 音声入力や音声文字起こし機能はありません。
- そのため、このプログラムにおけるWhisper API利用料は発生しません。

### Dify API

- 現在は使用していません。
- Difyのワークフロー、アプリID、APIキー、エンドポイントはコード上にありません。
- そのため、このプログラムにおけるDify API利用料は発生しません。

### Notion API（オプション）

- 現在は使用していません。
- 抽出結果をNotionデータベースへ保存する機能は未実装です。
- 将来追加する場合は、Notion連携用のAPIキー、データベースID、保存項目のマッピング設計が必要です。

### Gemini API / Vertex AI Gemini API

- 現在の主なAI抽出処理で使用します。
- Generative Language APIを使う場合は、`.env` に `GEMINI_PROVIDER=generative` と `GEMINI_API_KEY` を設定します。
- Vertex AIを使う場合は、`.env` に `GEMINI_PROVIDER=vertex`、`GOOGLE_CLOUD_PROJECT`、`GOOGLE_CLOUD_LOCATION` を設定し、Google Cloud CLIまたはアクセストークンで認証します。
- 既定モデルは `GEMINI_DEFAULT_MODEL` で指定します。未指定時は `gemini-3.1-flash-lite` を使います。

2026-05-19時点のGemini Developer API公式料金では、`gemini-3.1-flash-lite` Standardは以下です。

- 入力: $0.25 / 100万トークン（text / image / video）
- 音声入力: $0.50 / 100万トークン
- 出力: $1.50 / 100万トークン
- Batch / Flex利用時は、入力 $0.125 / 100万トークン、出力 $0.75 / 100万トークン

参考:

- Gemini Developer API pricing: https://ai.google.dev/gemini-api/docs/pricing
- Vertex AI / Agent Platform pricing: https://cloud.google.com/vertex-ai/docs/generative-ai/pricing

---

## 5. 月間予想コスト

以下は、`gemini-3.1-flash-lite` Standardを使い、1ファイルあたり入力8,000トークン、出力1,000トークン程度と仮定した概算です。実際の料金は、PDFページ数、画像枚数、抽出項目数、出力JSONの長さによって変わります。

| 使用量 | 想定トークン量 | Gemini入力料金 | Gemini出力込み概算 |
| --- | --- | --- | --- |
| 100件/月 | 入力80万 / 出力10万 | 約$0.20 | 約$0.35 |
| 1,000件/月 | 入力800万 / 出力100万 | 約$2.00 | 約$3.50 |
| 10,000件/月 | 入力8,000万 / 出力1,000万 | 約$20.00 | 約$35.00 |

Whisper API、Dify API、Notion APIは現在未使用のため、上記には含めていません。為替、税、Google Cloud側の追加条件、無料枠、将来の価格変更は別途確認が必要です。

---

## 6. セットアップと実行方法

### 前提条件

- Windows環境
- PowerShellが利用できること
- ブラウザでローカルURLを開けること
- Gemini APIキー、またはVertex AIを利用できるGoogle Cloud認証
- PDF.jsなどのCDNライブラリを初回読み込みできるインターネット接続

### インストール手順

1. このフォルダ一式を任意の場所に配置します。
2. `.env` を作成または更新し、Gemini APIの接続情報を設定します。
3. skyberry取込用のExcelフォーマットをアプリと同じフォルダに置きます。
4. 必要に応じて `config.js` の `apiBaseUrl` を設定します。ローカル実行では空のままで問題ありません。

### 実行方法

PowerShellでこのフォルダを開き、以下を実行します。

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\serve.ps1
```

起動後、ブラウザで以下を開きます。

```text
http://127.0.0.1:5173/
```

5173番ポートが使用中の場合は、別ポートで起動します。

```powershell
powershell -NoProfile -ExecutionPolicy Bypass -File .\serve.ps1 -Port 5174
```

### APIキー設定

Generative Language APIを使う場合:

```text
GEMINI_PROVIDER=generative
GEMINI_DEFAULT_MODEL=gemini-3.1-flash-lite
GEMINI_API_KEY=your-api-key
TEMPLATE_EXCEL_URL=https://example.com/template.xlsx
```

Vertex AIを使う場合:

```text
GEMINI_PROVIDER=vertex
GEMINI_DEFAULT_MODEL=gemini-3.1-flash-lite
GOOGLE_CLOUD_PROJECT=your-project-id
GOOGLE_CLOUD_LOCATION=asia-northeast1
TEMPLATE_EXCEL_URL=https://example.com/template.xlsx
```

Vertex AIでGoogle Cloud CLIを使う場合:

```powershell
gcloud auth application-default login
gcloud config set project your-project-id
```

### AWSデプロイ時の設定

AWSに配置する場合は、フロントエンド側とAPI側で以下を設定します。

フロントエンド側の `config.js`:

```javascript
window.APP_CONFIG = {
  apiBaseUrl: "https://xxxxxxxxxx.execute-api.ap-northeast-1.amazonaws.com"
};
```

ローカル実行では `apiBaseUrl` は空のままで問題ありません。AWSデプロイ時は、API GatewayのベースURLを入れます。末尾の `/` は付けても付けなくても動くようにしています。

Lambda側の環境変数例:

```text
GEMINI_PROVIDER=generative
GEMINI_DEFAULT_MODEL=gemini-3.1-flash-lite
GEMINI_API_KEY=your-api-key
TEMPLATE_EXCEL_URL=https://example.com/template.xlsx
```

Vertex AIをAWS Lambdaから使う場合は、Google Cloudのサービスアカウント認証情報やアクセストークンの扱いを別途設計する必要があります。検証用途では、まず `GEMINI_PROVIDER=generative` と `GEMINI_API_KEY` を使う構成がシンプルです。

API Gateway / Lambdaで必要なエンドポイント:

| メソッド | パス | 用途 |
| --- | --- | --- |
| GET | `/api/health` | API疎通確認、既定モデル確認 |
| GET | `/api/template/extraction-fields` | Excel/CSVテンプレート取得 |
| POST | `/api/gemini/generate` | Gemini APIへの抽出依頼 |
| POST | `/api/log` | ブラウザ側ログの受け取り |

AWSデプロイ後の確認手順:

1. AmplifyまたはCloudFrontのURLで画面を開きます。
2. ブラウザの開発者ツールで `/api/health` が200で返ることを確認します。
3. PDFまたは画像を1件だけ投入し、`次へ` でAI抽出まで進むことを確認します。
4. 目視確認画面で抽出データを修正し、`skyberry Excel出力` でExcelがダウンロードされることを確認します。
5. APIキーや認証情報が `app.js`、`config.js`、ブラウザの通信内容に露出していないことを確認します。

---

## 7. 注意事項と制限

1. ブラウザを更新すると、画面上の作業状態は消えます。
2. 元のExcelテンプレートは直接上書きされません。
3. 出力Excelは `skyberry-import-YYYYMMDD-HHMM.xlsx` として新規ダウンロードされます。
4. AI抽出結果は必ず目視確認してください。
5. 必須項目が未入力の場合は赤枠で表示されます。
6. 請求書に存在しない項目は、skyberry出力時の確認画面で確認してから出力できます。
7. PDF内部テキストは最大25ページ分、画像プレビューは最大12ページ分を処理対象にしています。
8. APIキーは `serve.ps1` 側で読み込み、ブラウザには渡しません。
9. CDNライブラリを使っているため、初回表示時やライブラリ未キャッシュ時はインターネット接続が必要です。

---

## 8. 将来の拡張可能性

- OCR専用APIとの比較検証
- skyberry項目へのマッピング設定画面
- 複数行明細の抽出・編集
- Notion、SharePoint、Google Driveなどへの保存連携
- 抽出精度レポートの自動生成
- 請求書種別ごとの抽出プロンプト切り替え
- Excel出力前の入力チェック強化
- サーバー常駐化、認証、操作ログ保存

---

## 9. トラブルシューティング

| 症状 | 原因候補 | 対応 |
| --- | --- | --- |
| `Start()` でポートエラーが出る | 5173番ポートが既に使用中 | 既存サーバーを使うか、`-Port 5174` など別ポートで起動する |
| 画面が開かない | `serve.ps1` が起動していない | PowerShellで起動コマンドを再実行する |
| AI抽出でエラーになる | `.env` のAPIキー不足、モデル名誤り、ネットワーク不通 | `GEMINI_API_KEY`、`GEMINI_DEFAULT_MODEL`、接続状況を確認する |
| AWSデプロイ画面でAPIに接続できない | `config.js` の `apiBaseUrl` 未設定、API Gateway URL誤り、CORS未設定 | API GatewayのURL、ステージ、CORS、`/api/health` の応答を確認する |
| LambdaでAI抽出が失敗する | Lambda環境変数不足、タイムアウト、Gemini APIキー誤り | `GEMINI_API_KEY`、`GEMINI_PROVIDER`、Lambdaタイムアウト、CloudWatch Logsを確認する |
| Vertex AIで認証エラーになる | OAuth認証またはプロジェクト設定不足 | `gcloud auth application-default login` と `GOOGLE_CLOUD_PROJECT` を確認する |
| PDFが正しく表示されない | PDF.jsのCDN読み込み失敗、特殊PDF | インターネット接続を確認し、再読み込みする |
| HEIC画像が読めない | heic2anyのCDN読み込み失敗 | インターネット接続を確認し、JPG/PNG変換済み画像で再投入する |
| 必須項目が赤枠になる | 抽出漏れ、または請求書に記載がない | 入力できるものは修正し、記載がないものは出力時の確認画面で判断する |
| Excelが出力されない | テンプレート未読込、ブラウザ制限、必須確認の未完了 | テンプレートファイル、確認画面、ブラウザのダウンロード設定を確認する |
