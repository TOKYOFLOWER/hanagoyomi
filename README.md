# hanagoyomi — 花暦 🌿

花市場の報告書PDF → お客様向け情報シートに自動変換するツール。

## 概要

PDFの市場報告書をアップロードすると、Claude APIが内容を読みやすく整形し、Google Sheetsに自動で書き込みます。

### 処理フロー

```
[PDF] → pdf.js でテキスト抽出（ブラウザ側）
     → Google Apps Script に送信
     → Claude API で文章を変換
     → Google Sheets に新規タブとして書き込み
```

## ファイル構成

| ファイル | 説明 |
|---------|------|
| `index.html` | GitHub Pages でホストされるフロントエンド UI |
| `gas/Code.gs` | Google Apps Script（Claude API 呼び出し + Sheets 書き込み） |

## セットアップ

### フロントエンド（GitHub Pages）

1. このリポジトリの Settings → Pages で `main` ブランチを公開
2. `index.html` 内の `gasUrl` を自分の GAS Web App URL に書き換える

### Google Apps Script

1. [Google Apps Script](https://script.google.com/) で新規プロジェクトを作成
2. `gas/Code.gs` の内容を貼り付ける
3. スクリプトプロパティに以下を設定：
   - `ANTHROPIC_API_KEY` — Claude API キー
   - `SPREADSHEET_ID` — 書き込み先スプレッドシートの ID
4. ウェブアプリとしてデプロイし、URL を `index.html` に設定

## 技術スタック

- **フロントエンド**: HTML / CSS / JavaScript（バニラ）
- **PDF解析**: [pdf.js](https://mozilla.github.io/pdf.js/) v3.11.174
- **バックエンド**: Google Apps Script
- **AI**: Claude API（Anthropic）
- **データ保存**: Google Sheets
