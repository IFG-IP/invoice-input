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

おすすめは `open-local.cmd` をダブルクリックして開く方法です。

起動するとローカルサーバーが立ち上がり、ブラウザで以下を開きます。

```text
http://127.0.0.1:5173/
```

手動で起動する場合は、このフォルダで以下を実行してください。

```powershell
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
- `編集内容を反映` で画面上のJSONに反映
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
