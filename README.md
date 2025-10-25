# 📄 請求書自動生成システム

Google Apps Script を使って、Google スプレッドシートのデータから請求書 PDF を自動生成し、メールで送信するシステムです。

## ✨ 主な機能

- **📊 スプレッドシート連携**: Google スプレッドシートから請求データを自動読み込み
- **📧 メール自動送信**: 作成した請求書 PDF のリンクを自動でメール送信
- **🔄 バッチ処理対応**: 複数の請求書を一括で作成・送信可能
- **🔗 Notion 連携**: Notion データベースから作業時間を自動取得して請求書に反映
- **📅 日付自動生成**: 請求日とファイル名を自動で生成
- **🛡️ エラーハンドリング**: エラーが発生しても他の請求書処理を継続
- **🎨 カスタマイズ可能**: テンプレートを自由にデザイン可能
- **💾 自動保存**: PDF を指定した Google Drive フォルダに自動保存
- **🧪 テストモード**: メール送信せずに動作確認できるテストモード搭載

## 🚀 クイックスタート（初心者向け）

### ステップ 1: 必要なファイルを準備

#### 1-1. データスプレッドシートを作成

1. Google スプレッドシートを新規作成
2. 以下の列を作成（`google-spread-sheets/invoice-data.csv` を参考）
   ```
   seller_name, seller_address, seller_email, item_1_name,
   item_1_number, item_1_price, seller_bank_name, seller_bank_type,
   seller_bank_number, seller_bank_holder_name, output_folder_id
   ```
3. データ行にサンプルデータを入力

**データ例:**
| seller_name | seller_address | seller_email | item_1_name | item_1_number | item_1_price |
|-------------|----------------|--------------|-------------|---------------|--------------|
| 株式会社サンプル A | 東京都千代田区千代田 1-1-1 | sample@example.com | Web システム開発費 | 1 | 150000 |

#### 1-2. テンプレートスプレッドシートを作成

1. Google スプレッドシートを新規作成
2. 請求書のレイアウトをデザイン
3. データを入れたい場所にプレースホルダーを配置
   - 例: `{{seller_name}}`, `{{invoice_date}}`, `{{item_1_price}}`
4. 詳細は `google-spread-sheets/invoice-template.csv` を参照

#### 1-3. Google Drive にフォルダを作成

1. Google Drive で請求書 PDF 保存用のフォルダを作成
2. フォルダの ID をコピー
   - フォルダを開いて、URL の `folders/` の後の文字列がフォルダ ID
   - 例: `https://drive.google.com/drive/folders/1abc...xyz` → `1abc...xyz`

### ステップ 2: スクリプトをセットアップ

#### 2-1. Google Apps Script プロジェクトを作成

1. Google スプレッドシート（どれでも可）を開く
2. メニューから `拡張機能` → `Apps Script` を選択
3. 新しいプロジェクトが開きます

#### 2-2. スクリプトをコピー

1. `scripts/invoice-generator.gs` の内容をすべてコピー
2. Apps Script エディタに貼り付け（既存のコードは削除）

#### 2-3. 環境変数を設定（重要！）

スクリプトの上部にある以下の 3 つの値を設定します：

```javascript
// データスプレッドシートのID
// スプレッドシートのURLから取得: https://docs.google.com/spreadsheets/d/【ここ】/edit
const DATA_SPREADSHEET_ID = "あなたのデータスプレッドシートID";

// テンプレートスプレッドシートのID
const TEMPLATE_SPREADSHEET_ID = "あなたのテンプレートスプレッドシートID";

// Notion APIキー（Notion連携を使う場合のみ）
const NOTION_API_KEY = "あなたのNotion APIキー";
```

**スプレッドシート ID の取得方法:**

- スプレッドシートを開く
- URL を確認: `https://docs.google.com/spreadsheets/d/【ここがID】/edit`
- 【ここが ID】の部分をコピー

#### 2-4. メール設定をカスタマイズ

```javascript
const EMAIL_ENABLED = true; // メール送信を使う場合は true
const TEST_MODE = true; // 初回は true にしてテスト
const SENDER_EMAIL = "your-email@example.com"; // あなたのメールアドレス
const SENDER_NAME = "あなたの会社名";
```

⚠️ **初めて使う場合は `TEST_MODE = true` にしてテストしましょう！**

### ステップ 3: 実行してみる

#### 3-1. 初回実行（権限の許可）

1. Apps Script エディタで関数を選択: `createInvoices`
2. 実行ボタン（▶）をクリック
3. 権限の確認画面が表示されます
   - 「権限を確認」をクリック
   - Google アカウントを選択
   - 「詳細」をクリック
   - 「（プロジェクト名）に移動」をクリック
   - 「許可」をクリック

#### 3-2. 実行ログを確認

1. Apps Script エディタの下部に「実行ログ」タブが表示されます
2. ログを確認して処理が成功したかチェック
   ```
   📊 [開始] Notionから作業時間を取得します
   ✅ PDF作成成功 - 20251018_株式会社DROX様_請求書.pdf
   ✅ メール送信成功 → sample@example.com
   🎉 すべての処理が完了しました！
   ```

#### 3-3. 結果を確認

1. Google Drive の指定フォルダに PDF が保存されているか確認
2. TEST_MODE の場合、実際にメールは送信されません
3. 問題なければ `TEST_MODE = false` にして本番運用

## 📋 詳細ドキュメント

より詳しい情報は以下のドキュメントを参照してください：

- 📖 **[詳細セットアップガイド](docs/SETUP_GUIDE.md)** - ステップバイステップの詳しい手順
- 📋 **[テンプレート説明](google-spread-sheets/README.md)** - テンプレートの構造とプレースホルダー一覧
- 📝 **[クイックリファレンス](docs/QUICK_REFERENCE.md)** - よく使う機能の早見表

## 📁 プロジェクト構成

```
invoice-automation/
├── README.md                           # このファイル（プロジェクト概要）
├── docs/
│   ├── SETUP_GUIDE.md                 # 詳細なセットアップ手順
│   └── QUICK_REFERENCE.md             # クイックリファレンス
├── google-spread-sheets/
│   ├── README.md                      # テンプレート詳細説明
│   ├── invoice-data.csv               # データスプレッドシートのサンプル
│   └── invoice-template.csv           # テンプレートのサンプル
└── scripts/
    └── invoice-generator.gs            # メインスクリプト（日本語コメント付き）
```

## 🔧 必要なもの

- Google アカウント
- Google スプレッドシート（無料）
- Google Drive（無料）
- Google Apps Script（Google アカウントに含まれています）
- Notion（オプション: 作業時間自動取得を使う場合）

## 💡 使い方の例

### 例 1: 月次請求書を一括作成

1. データスプレッドシートに今月の請求先を入力
2. `createInvoices()` を実行
3. 全ての請求書 PDF が自動生成され、メールで送信される

### 例 2: Notion の作業記録から自動請求

1. Notion で日々の作業時間を記録
2. 月末に `createInvoices()` を実行
3. Notion から作業時間が自動集計され、請求書に反映される

## ⚙️ 設定項目

### 基本設定

```javascript
// シート名（デフォルトは "Sheet1"）
const DATA_SHEET_NAME = "Sheet1";
const TEMPLATE_SHEET_NAME = "Sheet1";
```

### メール設定

```javascript
// メール送信を有効化
const EMAIL_ENABLED = true;

// テストモード（true: メール送信しない、false: 実際に送信）
const TEST_MODE = false;

// 送信者情報
const SENDER_EMAIL = "your-email@example.com";
const SENDER_NAME = "あなたの会社名";
```

### Notion 連携設定（オプション）

```javascript
// NotionデータベースID
const NOTION_DATABASE_ID = "あなたのNotionデータベースID";

// Notion APIバージョン（通常は変更不要）
const NOTION_API_VERSION = "2022-06-28";
```

**Notion 設定手順:**

1. https://www.notion.so/my-integrations でインテグレーション作成
2. 作成したトークンを `NOTION_API_KEY` に設定
3. Notion データベースをインテグレーションと共有
4. データベース ID を `NOTION_DATABASE_ID` に設定

**必要な Notion プロパティ:**

- `Start Time`: 日付型（開始時刻）
- `End Time`: 日付型（終了時刻）
- `hours`: 数値型またはフォーミュラ型（作業時間）

## 📊 データスプレッドシートの列構成

| 列名                    | 説明                          | 例                         |
| ----------------------- | ----------------------------- | -------------------------- |
| seller_name             | 請求先会社名                  | 株式会社サンプル A         |
| seller_address          | 請求先住所                    | 東京都千代田区千代田 1-1-1 |
| seller_email            | メール送信先                  | sample@example.com         |
| item_1_name             | 項目名                        | Web システム開発費         |
| item_1_number           | 数量                          | 1（または作業時間）        |
| item_1_price            | 単価                          | 150000                     |
| seller_bank_name        | 銀行名                        | 三井住友銀行 東京支店(111) |
| seller_bank_type        | 口座種別                      | 普通                       |
| seller_bank_number      | 口座番号                      | 1234567                    |
| seller_bank_holder_name | 口座名義                      | カ）サンプルエー           |
| output_folder_id        | PDF 保存先フォルダ ID（必須） | 1abc...xyz                 |

⚠️ **注意:**

- 1 行目は必ずヘッダー（列名）にしてください
- `output_folder_id` は必須項目です（空白だとエラーになります）
- 列の順番は自由です（ヘッダー名で判別します）

## 🎨 テンプレートのプレースホルダー

テンプレートスプレッドシートで使えるプレースホルダー：

### 請求書情報

- `{{invoice_date}}` - 請求日（自動生成）

### 請求先情報

- `{{seller_name}}` - 請求先会社名
- `{{seller_address}}` - 請求先住所

### 項目情報

- `{{item_1_name}}` - 項目名
- `{{item_1_number}}` - 数量
- `{{item_1_price}}` - 単価

### 銀行情報

- `{{seller_bank_name}}` - 銀行名
- `{{seller_bank_type}}` - 口座種別
- `{{seller_bank_number}}` - 口座番号
- `{{seller_bank_holder_name}}` - 口座名義

詳細は [テンプレート説明](google-spread-sheets/README.md) を参照してください。

## 🔄 処理の流れ

1. 📂 データスプレッドシートから請求データを読み込み
2. 📊 （オプション）Notion データベースから作業時間を取得
3. 📄 各請求データに対してテンプレートシートをコピー
4. 🔄 プレースホルダー（`{{xxx}}`）を実際のデータに置き換え
5. 📄 入力済みシートを PDF に変換
6. 💾 指定した Google Drive フォルダに PDF を保存
7. 🔗 PDF 共有リンクを生成
8. 📧 請求先にメールで請求書リンクを送信
9. 🗑️ 一時的な作業用シートを削除
10. 📊 処理結果のサマリーをログに出力

## 🛠️ トラブルシューティング

### よくある問題と解決方法

#### Q1: 「権限が不足しています」と表示される

**A:** 初回実行時に権限の許可が必要です。

1. 「権限を確認」をクリック
2. 「詳細」→「（プロジェクト名）に移動」をクリック
3. 「許可」をクリック

#### Q2: PDF が作成されない

**A:** 以下を確認してください：

- `DATA_SPREADSHEET_ID` が正しいか
- `TEMPLATE_SPREADSHEET_ID` が正しいか
- `output_folder_id` が正しいか
- フォルダへのアクセス権限があるか

#### Q3: メールが送信されない

**A:**

- `EMAIL_ENABLED = true` になっているか確認
- `TEST_MODE = false` になっているか確認（テストモードでは実際に送信されません）
- `seller_email` に有効なメールアドレスが入っているか確認

#### Q4: Notion からデータが取得できない

**A:**

- Notion API キーが正しいか確認
- Notion データベース ID が正しいか確認
- Notion データベースがインテグレーションと共有されているか確認
- 必須プロパティ（Start Time, End Time, hours）が存在するか確認

#### Q5: 「シートが見つかりません」エラー

**A:**

- `DATA_SHEET_NAME` がスプレッドシートのシート名と一致しているか確認
- `TEMPLATE_SHEET_NAME` がテンプレートのシート名と一致しているか確認

より詳しいトラブルシューティングは [セットアップガイド](docs/SETUP_GUIDE.md#troubleshooting) を参照してください。

## 📧 メール本文のカスタマイズ

メール本文は `createEmailBody()` 関数で定義されています。
HTML 形式で自由にカスタマイズ可能です。

デフォルトのメール内容：

```
【件名】
請求書送付のご案内 [YYYY年MM月分]

【本文】
〇〇様

いつもお世話になっております。
株式会社DROXです。

YYYY年MM月分の請求書をお送りいたします。
下記リンクよりご確認ください。

[請求書を確認する（ボタン）]

PDF URL: https://drive.google.com/...
```

## 🎯 活用例

### ケース 1: フリーランス・個人事業主

- 月末に全クライアントの請求書を一括作成
- 作業時間を Notion で管理して自動集計
- メールで自動送信して手間を削減

### ケース 2: 小規模企業・スタートアップ

- 複数案件の請求書を効率的に管理
- テンプレートを統一して見栄えを向上
- 請求業務の属人化を防ぐ

### ケース 3: 定期請求が発生する業務

- 月次の定期請求を自動化
- データを更新するだけで請求書作成
- 請求漏れを防止

## 📝 カスタマイズのヒント

### テンプレートをカスタマイズ

- Google スプレッドシートで自由にデザイン可能
- ロゴ画像の挿入も可能
- 複数の項目（item_2, item_3...）にも対応可能

### メール本文をカスタマイズ

- `createEmailBody()` 関数を編集
- HTML/CSS で見栄えを調整可能
- 添付ファイルの追加も可能

### 処理をスケジュール実行

1. Apps Script エディタで「トリガー」を設定
2. 「時間主導型」を選択
3. 実行タイミングを設定（例: 毎月 1 日午前 9 時）

## 🔒 セキュリティとプライバシー

- スクリプトはあなたの Google アカウント内で実行されます
- データは外部サーバーに送信されません（Notion 連携除く）
- API キーは適切に管理してください
- スプレッドシートの共有設定に注意してください

## 📄 ライセンス

このプロジェクトは教育目的および個人利用のために提供されています。
自由に改変・利用していただけます。

## 🤝 コントリビューション

改善案やバグ修正のプルリクエストを歓迎します！
フォークして自由にカスタマイズしてください。

## 💬 サポート

詳しい使い方については、以下をご覧ください：

- [詳細セットアップガイド](docs/SETUP_GUIDE.md)
- [テンプレート説明](google-spread-sheets/README.md)

---

❤️ 面倒な請求書作成作業を自動化して、本業に集中しましょう！
