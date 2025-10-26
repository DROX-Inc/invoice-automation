/**
 * ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 * 📄 Google Apps Script - 請求書PDF自動生成ツール
 * ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
 *
 * 【このスクリプトの機能】
 * - Googleスプレッドシートのデータから請求書PDFを自動生成
 * - Notionデータベースから作業時間を自動取得
 * - PDFをGoogle Driveに保存
 * - 請求書をメールで自動送信
 *
 * 【主な特徴】
 * ✓ ヘッダーベースのデータマッピング（列の順番は自由）
 * ✓ 日付の自動フォーマット
 * ✓ バッチ処理対応（複数の請求書を一括作成）
 * ✓ エラーハンドリング（エラーが発生しても処理を継続）
 * ✓ テストモード（メール送信なしで動作確認可能）
 *
 * 【セットアップ手順】初心者の方向け
 * 1. データスプレッドシートを作成（invoice-data.csv を参照）
 *    - 請求先情報、項目、金額などを入力します
 *
 * 2. テンプレートスプレッドシートを作成（invoice-template.csv を参照）
 *    - {{seller_name}}のようなプレースホルダーを使います
 *
 * 3. 下記の環境変数を設定（必須！）
 *    - DATA_SPREADSHEET_ID: データスプレッドシートのID
 *    - TEMPLATE_SPREADSHEET_ID: テンプレートスプレッドシートのID
 *    - NOTION_API_KEY: Notion APIキー
 *
 * 4. createInvoices() 関数を実行
 *    - Google Apps Scriptエディタで関数を選択して実行ボタンを押す
 *
 * 【初回実行時の注意】
 * - Google Apps Scriptが各種サービス（Drive、Gmail、Notion）への
 *   アクセス許可を求めます。必ず許可してください。
 */

// ============================================================================
// ⚙️ 環境変数 - 最初にここを設定してください！（必須）
// ============================================================================
// 【重要】以下の3つの変数は必ず設定する必要があります

// データスプレッドシートのID
// 取得方法: スプレッドシートのURLから取得
// 例: https://docs.google.com/spreadsheets/d/【ここがID】/edit
const DATA_SPREADSHEET_ID = "";

// テンプレートスプレッドシートのID
// 取得方法: スプレッドシートのURLから取得
// 例: https://docs.google.com/spreadsheets/d/【ここがID】/edit
const TEMPLATE_SPREADSHEET_ID = "";

// Notion APIキー
// 取得方法: https://www.notion.so/my-integrations でインテグレーションを作成
// 作成後に表示される「Internal Integration Token」をコピー
const NOTION_API_KEY = "";

// ============================================================================
// 📋 設定セクション
// ============================================================================

/**
 * 【データスプレッドシートの列構成】
 * データスプレッドシートには以下の列が必要です（順番は自由）：
 *
 * - seller_name: 請求先会社名（例: 株式会社ABC）
 * - invoice_needed: 請求書生成フラグ（TRUE/FALSE）※TRUEの場合のみ請求書を生成
 * - seller_address: 請求先住所
 * - seller_email: メール送信先アドレス
 * - item_1_name: 項目名（例: コンサルティング業務）
 * - item_1_price: 単価（例: 時間単価）
 * - seller_bank_name: 銀行名
 * - seller_bank_type: 口座種別（普通/当座）
 * - seller_bank_number: 口座番号
 * - seller_bank_holder_name: 口座名義
 * - output_folder_id: （必須）PDFを保存するGoogle DriveフォルダのID
 * - notion_user_id: NotionユーザーID（UUID形式）
 *
 * 【自動生成される項目】
 * - 請求日: スクリプト実行時の日付が自動設定されます
 * - PDFファイル名: YYYYMMDD_株式会社DROX様_請求書.pdf の形式で自動生成
 */

/**
 * データスプレッドシート内のシート名
 * 通常は 'Sheet1' ですが、異なる場合は変更してください
 */
const DATA_SHEET_NAME = "Sheet1";

/**
 * テンプレートスプレッドシート内のシート名
 * 通常は 'Sheet1' ですが、異なる場合は変更してください
 */
const TEMPLATE_SHEET_NAME = "Sheet1";

// ============================================================================
// 📧 メール送信設定
// ============================================================================

/**
 * 【メール送信の設定】
 *
 * EMAIL_ENABLED: メール送信機能の有効/無効
 *   - true: メール送信を有効にする
 *   - false: メール送信を完全に無効にする（PDFのみ作成）
 *
 * TEST_MODE: テストモード
 *   - true: メールを実際に送信せず、ログだけ出力（テスト用）
 *   - false: 実際にメールを送信（本番用）
 *   ⚠️ 初めて使う場合は true にしてテストすることをおすすめします
 *
 * SENDER_EMAIL: 送信者のメールアドレス
 *   - Google Apps Scriptを実行するGoogleアカウントのメールアドレス
 *
 * SENDER_NAME: 送信者の表示名
 *   - メール受信時に表示される送信者名
 */
const EMAIL_ENABLED = true; // メール送信を有効にする
const TEST_MODE = false; // テストモードを無効にする（本番運用時は false）
const SENDER_EMAIL = "koki-hata@drox-inc.com"; // 送信者メールアドレス
const SENDER_NAME = "株式会社DROX"; // 送信者名

// ============================================================================
// 🔗 Notion API 設定
// ============================================================================

/**
 * 【Notion データベース設定】
 *
 * Notionから作業時間を自動取得する場合に設定します
 *
 * セットアップ手順:
 * 1. https://www.notion.so/my-integrations でインテグレーションを作成
 * 2. 作成したインテグレーションのトークンを NOTION_API_KEY に設定（上部の環境変数セクション）
 * 3. Notionデータベースをインテグレーションと共有
 *    - データベースページの右上「...」→「コネクトの追加」→ インテグレーション名を選択
 * 4. データベースのIDを NOTION_DATABASE_ID に設定
 *    - データベースURLから取得: https://notion.so/【ここがID】?v=...
 *
 * 【Notionデータベースの必須プロパティ】
 * - Start Time: 日付型（開始時刻）
 * - End Time: 日付型（終了時刻）
 * - hours: 数値型またはフォーミュラ型（作業時間）
 */
const NOTION_DATABASE_ID = "25143e83afc180cfaa2bff97b74538eb"; // NotionデータベースのID
const NOTION_API_VERSION = "2022-06-28"; // Notion APIのバージョン（通常変更不要）

// ============================================================================
// 🔗 Notion API 関数群
// ============================================================================

/**
 * Notionデータベースのスキーマ（構造）を取得する関数
 * 【用途】デバッグ時にデータベースのプロパティ名や型を確認できます
 * 【初心者向け説明】
 * - この関数は、Notionデータベースにどんなプロパティ（列）があるかを調べます
 * - 例: "Start Time"という名前のプロパティが"date"型であることなどを確認できます
 *
 * @returns {Object} データベースのプロパティ情報
 */
function getNotionDatabaseSchema() {
  console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
  console.log("🔍 [開始] Notionデータベースのスキーマを取得します");
  console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");

  try {
    // --------------------------------------------------------------------------
    // ▼ STEP 1: APIリクエストの準備 ▼
    // --------------------------------------------------------------------------
    // Notion APIからデータベースの情報を取得するためのURLを組み立てます。
    // このURLは、どのデータベースの情報を取得したいかをNotionに伝えるためのものです。
    const url = `https://api.notion.com/v1/databases/${NOTION_DATABASE_ID}`;
    console.log(`構築されたURL: ${url}`);

    // APIリクエストに必要な設定（オプション）を定義します。
    const options = {
      method: "get", // 'get'は「データを取得する」という意味のHTTPメソッドです。
      headers: {
        // ヘッダー：リクエストに関する追加情報
        // Authorization: APIを利用するための「認証キー」です。これにより、Notionは誰からのリクエストかを識別します。
        Authorization: `Bearer ${NOTION_API_KEY}`,
        // Notion-Version: 使用するNotion APIのバージョンを指定します。
        "Notion-Version": NOTION_API_VERSION,
      },
      // muteHttpExceptions: trueに設定すると、APIがエラーを返した場合（例：404 Not Found）でも
      // スクリプトが停止せず、後続のコードでエラー処理を行えるようになります。
      muteHttpExceptions: true,
    };
    console.log("リクエストヘッダー:", options.headers);

    // --------------------------------------------------------------------------
    // ▼ STEP 2: Notion APIへのリクエスト送信 ▼
    // --------------------------------------------------------------------------
    console.log("📡 Notion APIにリクエストを送信中...");
    // UrlFetchApp.fetch() を使って、実際にNotion APIにリクエストを送信します。
    const response = UrlFetchApp.fetch(url, options);
    const responseCode = response.getResponseCode();
    const responseBody = response.getContentText();
    console.log(`📬 Notion APIからの応答ステータスコード: ${responseCode}`);

    // --------------------------------------------------------------------------
    // ▼ STEP 3: レスポンスの処理 ▼
    // --------------------------------------------------------------------------
    // レスポンスコードが200の場合、リクエストは成功です。
    if (responseCode === 200) {
      console.log("✅ データベース情報の取得に成功しました。");
      // 応答データ（JSON形式のテキスト）をオブジェクトに変換します。
      const data = JSON.parse(responseBody);

      console.log("=== データベースのプロパティ一覧 ===");
      // 取得したデータからプロパティ情報を取り出し、名前と型をログに出力します。
      // これにより、データベースにどのような列があるかを確認できます。
      for (const [propName, propConfig] of Object.entries(data.properties)) {
        console.log(`  - 📋 プロパティ名: "${propName}", 型: ${propConfig.type}`);
      }
      console.log("====================================");
      console.log("✅ [完了] getNotionDatabaseSchema() - 成功");
      console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n");
      return data.properties;
    } else {
      // 200以外のコードはエラーを示します。
      console.error(`❌ データベーススキーマの取得に失敗しました。ステータスコード: ${responseCode}`);
      console.error("--- Notionからのエラーレスポンス START ---");
      console.error(responseBody);
      console.error("--- Notionからのエラーレスポンス END ---");
      console.error("💡 ヒント: NOTION_API_KEY または NOTION_DATABASE_ID が間違っている可能性があります。");
      return null;
    }
  } catch (error) {
    console.error("❌ スクリプト実行中に予期せぬエラーが発生しました:", error.message);
    console.error("💡 ヒント: インターネット接続またはGoogle Apps Scriptのサービスに問題がある可能性があります。");
    return null;
  }
}

/**
 * Notionデータベースから指定期間・特定ユーザーの作業時間を取得する関数
 * 【機能】指定した日付範囲内かつ指定ユーザーの"hours"プロパティの値を合計します
 * 【初心者向け説明】
 * - Notionデータベースに記録されている作業時間を自動で集計します
 * - 例: 2025年10月1日〜31日の間に特定のメンバーが記録した作業時間をすべて足し算します
 * - 請求書の作業時間を手動で計算する必要がなくなります
 *
 * @param {string} startDate - 開始日（YYYY-MM-DD形式、例: "2025-10-01"）
 * @param {string} endDate - 終了日（YYYY-MM-DD形式、例: "2025-10-31"）
 * @param {string} notionUserId - NotionユーザーID（UUID形式）
 * @returns {number} 合計作業時間（時間単位）
 */
function fetchNotionHours(startDate, endDate, notionUserId) {
  console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
  console.log("📊 [開始] Notionから作業時間を取得します");
  console.log(`📅 対象期間: ${startDate} 〜 ${endDate}`);
  console.log(`👤 対象ユーザーID: ${notionUserId}`);
  console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");

  try {
    // --------------------------------------------------------------------------
    // ▼ STEP 1: 入力パラメータの検証 ▼
    // --------------------------------------------------------------------------
    if (!startDate || !endDate) {
      console.error(`❌ 日付パラメータが無効です: 開始日="${startDate}", 終了日="${endDate}"`);
      throw new Error(`日付パラメータが必須です。受信値: 開始日="${startDate}", 終了日="${endDate}"`);
    }
    if (!notionUserId) {
      console.error(`❌ NotionユーザーIDが無効です: "${notionUserId}"`);
      throw new Error(`NotionユーザーIDが必須です。受信値: "${notionUserId}"`);
    }
    console.log("✅ パラメータの検証完了");

    // --------------------------------------------------------------------------
    // ▼ STEP 2: APIリクエストの準備 ▼
    // --------------------------------------------------------------------------
    // Notionデータベースから情報を「問い合わせ（クエリ）」するためのURLです。
    const url = `https://api.notion.com/v1/databases/${NOTION_DATABASE_ID}/query`;
    console.log(`構築されたURL: ${url}`);

    // Notionに「どのようなデータが欲しいか」を伝えるための条件（ペイロード）を作成します。
    // これは、データベースから特定の条件に合うデータだけを絞り込むために使います。
    const payload = {
      // `filter`：データを絞り込むための条件を指定します。
      filter: {
        // `and`：ここに列挙されたすべての条件を満たすデータのみを取得します。
        and: [
          {
            property: "Start Time", // データベースの「Start Time」という名前のプロパティを対象にします。
            date: {
              on_or_after: startDate, // その日付が `startDate` 以降であること。
            },
          },
          {
            property: "End Time", // データベースの「End Time」という名前のプロパティを対象にします。
            date: {
              on_or_before: endDate, // その日付が `endDate` 以前であること。
            },
          },
          {
            property: "User", // データベースの「User」プロパティ（people型）を対象にします。
            people: {
              contains: notionUserId, // 指定されたユーザーIDを含むこと。
            },
          },
        ],
      },
      // `page_size`：1回のリクエストで取得するデータの上限数。Notion APIの最大値は100です。
      page_size: 100,
    };

    console.log("📤 Notionに送信するフィルター条件（ペイロード）:", JSON.stringify(payload, null, 2));

    // APIリクエストに必要な設定（オプション）を定義します。
    const options = {
      method: "post", // 'post'は「データを問い合わせる・作成する」などの操作で使います。今回はクエリなのでpostです。
      headers: {
        Authorization: `Bearer ${NOTION_API_KEY}`,
        "Notion-Version": NOTION_API_VERSION,
        "Content-Type": "application/json", // 送信するデータがJSON形式であることを示します。
      },
      // `payload`：上で作成したフィルター条件を、JSON形式の文字列に変換して設定します。
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    };
    console.log("リクエストヘッダー:", options.headers);

    // --------------------------------------------------------------------------
    // ▼ STEP 3: データ取得とページネーション処理 ▼
    // --------------------------------------------------------------------------
    let totalHours = 0;
    let hasMore = true; // 次のページがあるかどうかを示すフラグ
    let nextCursor = null; // 次のページの開始位置を示すID
    let pageCount = 0;

    console.log("🔄 データ取得ループを開始します（結果が100件を超える場合、複数回リクエストされます）");
    // `hasMore`がtrueである限り、ループを続けます。
    while (hasMore) {
      pageCount++;
      console.log(`\n📄 ページ ${pageCount} の取得処理を開始...`);

      // 2ページ目以降の場合、`start_cursor`をペイロードに追加して、前回の続きからデータを取得します。
      if (nextCursor) {
        console.log(`   ...前回の続きから取得します (cursor: ${nextCursor})`);
        payload.start_cursor = nextCursor;
        options.payload = JSON.stringify(payload); // ペイロードを更新
      }

      // APIリクエストを送信
      console.log(`   📡 Notion APIにページ ${pageCount} のデータをリクエスト...`);
      const response = UrlFetchApp.fetch(url, options);
      const responseCode = response.getResponseCode();
      const responseBody = response.getContentText();
      console.log(`   📬 Notion APIからの応答ステータスコード: ${responseCode}`);

      if (responseCode !== 200) {
        console.error(`❌ Notion APIがエラーを返しました。ステータスコード: ${responseCode}`);
        console.error("--- Notionからのエラーレスポンス START ---");
        console.error(responseBody);
        console.error("--- Notionからのエラーレスポンス END ---");
        throw new Error(`Notion APIがエラーコードを返しました: ${responseCode}`);
      }

      const data = JSON.parse(responseBody);
      console.log(`   ✅ ページ ${pageCount} のデータを取得しました（${data.results.length}件）`);

      // --------------------------------------------------------------------------
      // ▼ STEP 4: 取得したデータの処理 ▼
      // --------------------------------------------------------------------------
      if (data.results && data.results.length > 0) {
        // 最初のページの最初のアイテムの構造をログに出力
        if (pageCount === 1 && data.results.length > 0) {
          console.log("🔍 [デバッグ情報] 最初のページの最初のアイテムのデータ構造:");
          console.log(JSON.stringify(data.results[0], null, 2));
        }

        let pageHours = 0;

        // 取得した各データ（ページ）に対して処理を行います。
        data.results.forEach((page, index) => {
          const pageId = page.id;
          let hoursValue = 0;
          try {
            // `page.properties.hours` に作業時間のデータが格納されています。
            const hoursProperty = page.properties.hours;
            if (hoursProperty) {
              // プロパティの型（'number'または'formula'）に応じて値を取得します。
              if (hoursProperty.type === "number") {
                hoursValue = hoursProperty.number || 0;
              } else if (hoursProperty.type === "formula") {
                // フォーミュラ型の場合、計算結果が `formula.number` に入っています。
                hoursValue = hoursProperty.formula.number || 0;
              }
              // console.log(`     - item ${index + 1} (ID: ${pageId}): ${hoursValue.toFixed(2)} 時間`); // 個別の時間はデバッグ情報で確認できるため、コメントアウト
              pageHours += hoursValue;
            } else {
              console.warn(`     - item ${index + 1} (ID: ${pageId}): 'hours' プロパティが見つかりません。`);
            }
          } catch (e) {
            console.error(`     - item ${index + 1} (ID: ${pageId}): データ処理中にエラーが発生しました: ${e.message}`);
          }
        });

        totalHours += pageHours;
        console.log(`   👉 このページの合計時間: ${pageHours.toFixed(2)}時間`);
        console.log(`   累計: ${totalHours.toFixed(2)}時間`);
      }

      // --------------------------------------------------------------------------
      // ▼ STEP 5: 次のページの有無を確認 ▼
      // --------------------------------------------------------------------------
      // レスポンスに `has_more: true` が含まれている場合、まだ続きのデータがあります。
      hasMore = data.has_more || false;
      // `next_cursor` は、次のリクエストでどこからデータを取得すればよいかを示すIDです。
      nextCursor = data.next_cursor || null;

      if (hasMore) {
        console.log("   ...まだデータがあります。次のページを取得します。");
      } else {
        console.log("   ...すべてのデータを取得しました。");
      }
    }

    console.log("\n" + "─".repeat(50));
    console.log(`✅ [成功] 全ページの作業時間の合計が完了しました。`);
    console.log(`   - 取得ページ数: ${pageCount}ページ`);
    console.log(`   - 合計作業時間: ${totalHours.toFixed(2)}時間`);
    console.log("─".repeat(50) + "\n");
    console.log("✅ [完了] fetchNotionHours() - 成功");
    console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n");
    return totalHours;
  } catch (error) {
    console.error(`❌ Notionからのデータ取得中に致命的なエラーが発生しました: ${error.message}`);
    console.error(
      "💡 ヒント: Notion APIキー、データベースID、プロパティ名（hours, Start Time, End Time）が正しいか、またインテグレーションがデータベースに共有されているかを確認してください。"
    );
    console.error("❌ [完了] fetchNotionHours() - 失敗");
    console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n");
    throw error;
  }
}

// ============================================================================
// 📝 メイン関数（請求書作成の中心処理）
// ============================================================================

/**
 * 【メイン関数】請求書を一括作成する関数
 * 【初心者向け説明】
 * この関数がスクリプトの中心です。以下の処理を自動で行います：
 * 1. Googleスプレッドシートから請求書データを読み込み
 * 2. Notionから作業時間を取得
 * 3. テンプレートに情報を埋め込み
 * 4. PDFを作成してGoogle Driveに保存
 * 5. 請求書をメールで送信
 *
 * 【使い方】
 * Google Apps Scriptのエディタで「createInvoices」関数を選択して実行ボタンを押すだけ！
 */
function createInvoices() {
  console.log("╔════════════════════════════════════════════════════════════════╗");
  console.log("║                    請求書自動作成を開始します                    ║");
  console.log("╚════════════════════════════════════════════════════════════════╝\n");

  // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  // ステップ1: スプレッドシートとテンプレートを取得
  // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  console.log("📂 [ステップ1] スプレッドシートを開いています...");

  // データが入っているスプレッドシートを開く
  const dataSpreadsheet = SpreadsheetApp.openById(DATA_SPREADSHEET_ID);
  const dataSheet = dataSpreadsheet.getSheetByName(DATA_SHEET_NAME);

  // シートが見つからない場合はエラーを表示して終了
  if (!dataSheet) {
    console.error(`❌ データシート "${DATA_SHEET_NAME}" が見つかりません。`);
    console.error("💡 ヒント: DATA_SHEET_NAMEの設定を確認してください");
    return;
  }
  console.log(`✅ データシートを開きました: "${DATA_SHEET_NAME}"`);

  // 請求書のテンプレートスプレッドシートを開く
  const templateSpreadsheet = SpreadsheetApp.openById(TEMPLATE_SPREADSHEET_ID);
  const templateSheet = templateSpreadsheet.getSheetByName(TEMPLATE_SHEET_NAME);

  // テンプレートシートが見つからない場合はエラーを表示して終了
  if (!templateSheet) {
    console.error(`❌ テンプレートシート "${TEMPLATE_SHEET_NAME}" が見つかりません。`);
    console.error("💡 ヒント: TEMPLATE_SHEET_NAMEの設定を確認してください");
    return;
  }
  console.log(`✅ テンプレートシートを開きました: "${TEMPLATE_SHEET_NAME}"\n`);

  // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  // ステップ2: スプレッドシートからデータを読み込む
  // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  console.log("📊 [ステップ2] データを読み込んでいます...");

  // シート全体のデータを取得（2次元配列として取得される）
  const data = dataSheet.getDataRange().getValues();

  // 1行目はヘッダー（列名）
  const headers = data[0];
  console.log(`📋 ヘッダー: ${headers.join(", ")}`);

  // 2行目以降が実際のデータ
  const dataRows = data.slice(1);

  // 各行をオブジェクトに変換（ヘッダーをキーにする）
  // 例: { seller_name: "株式会社ABC", seller_email: "abc@example.com", ... }
  const invoiceDataRows = dataRows.map((row) => {
    const invoice = {};
    headers.forEach((header, index) => {
      invoice[header] = row[index];
    });
    return invoice;
  });

  console.log(`✅ ${invoiceDataRows.length}件の請求書データを読み込みました\n`);

  // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  // ステップ3: カウンター変数を初期化（処理結果を記録するため）
  // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  let successCount = 0; // PDF作成成功数
  let failureCount = 0; // PDF作成失敗数
  let skippedCount = 0; // スキップ数
  let emailsSentCount = 0; // メール送信成功数
  let emailsFailedCount = 0; // メール送信失敗数

  console.log("╔════════════════════════════════════════════════════════════════╗");
  console.log(`║              ${invoiceDataRows.length}件の請求書作成を開始します                      ║`);
  console.log("╚════════════════════════════════════════════════════════════════╝\n");

  // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  // ステップ4: 請求書データを1行ずつ処理
  // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

  invoiceDataRows.forEach((invoice, index) => {
    console.log("┌────────────────────────────────────────────────────────────┐");
    console.log(`│  請求書 ${index + 1}/${invoiceDataRows.length} の作成を開始                                    │`);
    console.log(`│  宛先: ${invoice.seller_name}                                  `);
    console.log("└────────────────────────────────────────────────────────────┘");

    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    // ステップ4-0: invoice_neededフラグのチェック
    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    console.log("  🔍 請求書生成フラグを確認中...");

    // invoice_neededが明示的にTRUEでない場合はスキップ
    const invoiceNeeded = String(invoice.invoice_needed).toUpperCase();
    if (invoiceNeeded !== "TRUE") {
      skippedCount++;
      console.log(`  ⊘ 請求書 ${index + 1}: invoice_needed=${invoice.invoice_needed} のためスキップします`);
      console.log(`  ℹ️ ${invoice.seller_name} の請求書生成は不要と設定されています\n`);
      return;
    }

    console.log(`  ✅ invoice_needed=TRUE - 請求書生成を続行します`);

    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    // ステップ4-1: テンプレートシートをコピーして作業用シートを作成
    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    console.log("  📄 テンプレートをコピーして作業用シートを作成中...");
    const copiedSheet = templateSheet.copyTo(templateSpreadsheet);
    const newSheetName = `temp_${index + 1}`; // 一時的なシート名
    copiedSheet.setName(newSheetName);
    console.log(`  ✅ 作業用シート作成完了: "${newSheetName}"`);

    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    // ステップ4-2: 現在の日付を取得してPDFファイル名を生成
    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    console.log("  📅 日付情報を生成中...");
    const currentDate = new Date();
    const formattedDate = Utilities.formatDate(currentDate, "JST", "yyyy/MM/dd");
    const pdfFileName = Utilities.formatDate(currentDate, "JST", "yyyyMMdd") + "_株式会社DROX様_請求書";
    console.log(`  ✅ 請求日: ${formattedDate}`);
    console.log(`  ✅ PDFファイル名: ${pdfFileName}.pdf`);

    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    // ステップ4-3: Notionから作業時間を取得（重要！）
    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    console.log("  ⏰ Notionから作業時間を取得中...");
    let itemNumber = 0; // デフォルトは0時間

    // notion_user_idの検証
    if (!invoice.notion_user_id || invoice.notion_user_id.trim() === "") {
      console.warn(`  ⚠️ notion_user_idが指定されていません。Notionからのデータ取得をスキップします`);
      console.log(`  ℹ️ 作業時間: 0時間（データなし）`);
    } else {
      try {
        // TODO: 将来的には動的に日付を設定できるようにする
        // 現在は固定値を使用
        const startDate = "2025-10-01";
        const endDate = "2025-10-31";

        console.log(`  📊 対象期間: ${startDate} 〜 ${endDate}`);
        console.log(`  👤 対象ユーザー: ${invoice.seller_name} (ID: ${invoice.notion_user_id})`);

        const notionHours = fetchNotionHours(startDate, endDate, invoice.notion_user_id);

        if (notionHours !== null && notionHours !== undefined) {
          itemNumber = notionHours;
          console.log(`  ✅ Notionから取得した作業時間を使用: ${notionHours}時間`);
        } else {
          console.log(`  ⚠️ Notionからデータが取得できませんでした。作業時間: 0時間`);
        }
      } catch (notionError) {
        console.error(`  ❌ Notionデータの取得に失敗しました。作業時間: 0時間`);
        console.error(`  エラー詳細: ${notionError.message}`);
        // エラーが発生しても0時間で続行
      }
    }

    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    // ステップ4-4: プレースホルダーを実際のデータに置き換え
    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    console.log("  🔄 テンプレートにデータを埋め込み中...");

    // 請求日を置き換え
    copiedSheet.createTextFinder("{{invoice_date}}").replaceAllWith(formattedDate);
    console.log("    ✓ 請求日");

    // 請求先情報を置き換え
    copiedSheet.createTextFinder("{{seller_name}}").replaceAllWith(invoice.seller_name);
    copiedSheet.createTextFinder("{{seller_address}}").replaceAllWith(invoice.seller_address);
    console.log("    ✓ 請求先情報");

    // 項目詳細を置き換え（Notionまたはスプレッドシートから取得した作業時間を使用）
    copiedSheet.createTextFinder("{{item_1_name}}").replaceAllWith(invoice.item_1_name);
    copiedSheet.createTextFinder("{{item_1_number}}").replaceAllWith(itemNumber);
    copiedSheet.createTextFinder("{{item_1_price}}").replaceAllWith(invoice.item_1_price);
    console.log("    ✓ 項目詳細");

    // 銀行情報を置き換え
    copiedSheet.createTextFinder("{{seller_bank_name}}").replaceAllWith(invoice.seller_bank_name);
    copiedSheet.createTextFinder("{{seller_bank_type}}").replaceAllWith(invoice.seller_bank_type);
    copiedSheet.createTextFinder("{{seller_bank_number}}").replaceAllWith(invoice.seller_bank_number);
    copiedSheet.createTextFinder("{{seller_bank_holder_name}}").replaceAllWith(invoice.seller_bank_holder_name);
    console.log("    ✓ 銀行情報");

    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    // ステップ4-5: 変更をスプレッドシートに即座に反映
    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    SpreadsheetApp.flush(); // バッファをフラッシュして確実に保存
    console.log("  ✅ データの埋め込みが完了しました");

    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    // ステップ4-6: 出力先フォルダを検証（必須項目）
    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    console.log("  📁 出力先フォルダを確認中...");

    if (!invoice.output_folder_id || invoice.output_folder_id.trim() === "") {
      failureCount++;
      console.error(`  ❌ 請求書 ${index + 1}: output_folder_id（出力先フォルダID）が指定されていません`);
      console.error("  💡 ヒント: スプレッドシートのoutput_folder_id列にGoogle DriveのフォルダIDを入力してください");
      templateSpreadsheet.deleteSheet(copiedSheet);
      return;
    }

    let targetFolder;
    try {
      targetFolder = DriveApp.getFolderById(invoice.output_folder_id);
      console.log(`  ✅ 出力先フォルダID: ${invoice.output_folder_id}`);
    } catch (e) {
      failureCount++;
      console.error(`  ❌ 請求書 ${index + 1}: 無効なoutput_folder_id: ${invoice.output_folder_id}`);
      console.error(`  エラー詳細: ${e.message}`);
      console.error("  💡 ヒント: フォルダIDが正しいか、アクセス権限があるか確認してください");
      templateSpreadsheet.deleteSheet(copiedSheet);
      return;
    }

    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    // ステップ4-7: PDFを作成してGoogle Driveに保存
    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    console.log("  📄 PDFを作成中...");

    let pdfFile = null;
    try {
      pdfFile = createPdfInDrive(templateSpreadsheet, copiedSheet.getSheetId(), targetFolder, pdfFileName);
      successCount++;
      console.log(`  ✅ 請求書 ${index + 1}: PDF作成成功 - ${pdfFileName}.pdf`);
    } catch (e) {
      failureCount++;
      console.error(`  ❌ 請求書 ${index + 1}: PDF作成失敗 - ${pdfFileName}.pdf`);
      console.error(`  エラー詳細: ${e.message}`);
      templateSpreadsheet.deleteSheet(copiedSheet);
      return;
    }

    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    // ステップ4-8: メール送信（有効化されている場合のみ）
    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

    if (EMAIL_ENABLED && pdfFile) {
      console.log("  📧 メール送信処理を開始...");

      const recipientEmail = invoice.seller_email;
      const recipientName = invoice.seller_name;

      if (recipientEmail && recipientEmail.includes("@")) {
        const emailSent = sendInvoiceEmail(recipientEmail, recipientName, pdfFile, formattedDate);

        if (emailSent) {
          emailsSentCount++;
          console.log(`  ✅ 請求書 ${index + 1}: メール送信成功 → ${recipientEmail}`);
        } else {
          emailsFailedCount++;
          console.error(`  ❌ 請求書 ${index + 1}: メール送信失敗 → ${recipientEmail}`);
        }
      } else {
        console.warn(`  ⚠️ 請求書 ${index + 1}: 有効なメールアドレスがないためメール送信をスキップします`);
      }
    } else if (!EMAIL_ENABLED) {
      console.log("  ℹ️ メール送信が無効化されています（EMAIL_ENABLED = false）");
    }

    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    // ステップ4-9: 作業用シートを削除（クリーンアップ）
    // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
    templateSpreadsheet.deleteSheet(copiedSheet);
    console.log(`  🗑️ 作業用シート "${newSheetName}" を削除しました\n`);
  });

  // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  // 処理完了サマリーを表示
  // ━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
  console.log("\n╔════════════════════════════════════════════════════════════════╗");
  console.log("║                    📊 処理結果サマリー                          ║");
  console.log("╚════════════════════════════════════════════════════════════════╝");
  console.log(`\n📝 処理した請求書の総数: ${invoiceDataRows.length}件`);
  console.log(`\n【PDF作成結果】`);
  console.log(`  ✅ 成功: ${successCount}件`);
  console.log(`  ❌ 失敗: ${failureCount}件`);
  console.log(`  ⊘ スキップ: ${skippedCount}件`);

  if (EMAIL_ENABLED) {
    console.log(`\n【メール送信結果】`);
    console.log(`  ✅ 送信成功: ${emailsSentCount}件`);
    console.log(`  ❌ 送信失敗: ${emailsFailedCount}件`);

    if (TEST_MODE) {
      console.log("\n⚠️ テストモードが有効でした - 実際にはメールは送信されていません");
      console.log("   本番環境で使用する場合は TEST_MODE を false に設定してください");
    }
  } else {
    console.log("\n📧 メール送信: 無効化されています（EMAIL_ENABLED = false）");
  }

  console.log("\n" + "=".repeat(64));
  console.log("🎉 すべての処理が完了しました！");
  console.log("=".repeat(64) + "\n");
}

// ============================================================================
// 📄 PDF作成関数
// ============================================================================

/**
 * スプレッドシートの指定シートをPDFとしてGoogle Driveに保存する関数
 * 【初心者向け説明】
 * - スプレッドシートのシートをPDFファイルに変換します
 * - Google DriveのエクスポートAPIを使用して高品質なPDFを生成します
 * - 用紙サイズ、余白、配置などを細かく設定できます
 *
 * @param {Spreadsheet} spreadsheet - PDF化するシートが含まれるスプレッドシート
 * @param {string} sheetId - PDF化するシートのID
 * @param {Folder} folder - PDFを保存するGoogle Driveフォルダ
 * @param {string} fileName - PDFのファイル名（拡張子なし）
 * @returns {File} 作成されたPDFファイルオブジェクト
 */
function createPdfInDrive(spreadsheet, sheetId, folder, fileName) {
  console.log("    📄 [開始] PDFを作成してGoogle Driveに保存します");

  try {
    // --------------------------------------------------------------------------
    // ▼ STEP 1: PDFエクスポート用のURLを構築 ▼
    // --------------------------------------------------------------------------
    // GoogleスプレッドシートをPDFとしてエクスポートするための特別なURLを組み立てます。
    // このURLにアクセスすると、指定した形式でPDFが生成されます。
    const spreadsheetId = spreadsheet.getId();
    const exportUrl =
      `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?` +
      "format=pdf" +              // フォーマット: PDF
      "&gid=" + sheetId +         // 対象シートのID
      "&size=A4" +                // 用紙サイズ: A4
      "&portrait=true" +          // 向き: 縦
      "&scale=3" +                // スケール: 高さに合わせる (推奨)
      "&top_margin=0.2" +         // 上余白 (インチ)
      "&bottom_margin=0.2" +      // 下余白 (インチ)
      "&left_margin=0.2" +        // 左余白 (インチ)
      "&right_margin=0.2" +       // 右余白 (インチ)
      "&sheetnames=false" +       // シート名を非表示
      "&printtitle=false" +       // スプレッドシート名を非表示
      "&gridlines=false" +        // グリッド線を非表示
      "&fzr=false" +              // 固定行を無視
      "&horizontal_alignment=CENTER" + // 水平配置: 中央
      "&vertical_alignment=TOP";     // 垂直配置: 上

    console.log(`      - PDFエクスポートURLを構築しました。`);
    // console.log(`      URL: ${exportUrl}`); // デバッグ時にURLを確認したい場合はこの行を有効化

    // --------------------------------------------------------------------------
    // ▼ STEP 2: APIリクエストの準備 ▼
    // --------------------------------------------------------------------------
    // Googleのサービスにアクセスするための「認証トークン」を取得します。
    // これにより、スクリプトがユーザーの代わりにGoogle DriveやSpreadsheetを操作する許可を得ます。
    const token = ScriptApp.getOAuthToken();

    // APIリクエストに必要な設定（オプション）を定義します。
    const options = {
      headers: {
        // Authorizationヘッダーに認証トークンを設定します。
        Authorization: "Bearer " + token,
      },
      muteHttpExceptions: true, // エラーが発生してもスクリプトを停止させない
    };
    console.log("      - APIリクエストのオプションを設定しました。");

    // --------------------------------------------------------------------------
    // ▼ STEP 3: PDFデータの取得と保存 ▼
    // --------------------------------------------------------------------------
    console.log("      - 📥 PDFデータをダウンロードしています...");
    // UrlFetchApp.fetch() を使って、エクスポートURLにアクセスし、PDFデータを取得します。
    const response = UrlFetchApp.fetch(exportUrl, options);
    const responseCode = response.getResponseCode();

    if (responseCode === 200) {
      console.log(`      - ✅ PDFデータのダウンロードに成功しました (ステータスコード: ${responseCode})`);
      // 取得したPDFデータを「Blob」というバイナリデータ形式に変換し、ファイル名を設定します。
      const blob = response.getBlob().setName(fileName + ".pdf");

      console.log(`      - 💾 PDFをGoogle Driveフォルダ「${folder.getName()}」に保存しています...`);
      // 指定されたGoogle DriveフォルダにPDFファイルを作成します。
      const pdfFile = folder.createFile(blob);
      console.log(`      - ✅ PDFの保存が完了しました: ${pdfFile.getName()}`);

      // 作成したPDFファイルを返す
      return pdfFile;
    } else {
      // エラー処理
      const errorResponse = response.getContentText();
      console.error(`      - ❌ PDFのダウンロードに失敗しました (ステータスコード: ${responseCode})`);
      console.error(`      - エラー内容: ${errorResponse}`);
      throw new Error(`PDFの生成に失敗しました。ステータスコード: ${responseCode}`);
    }
  } catch (error) {
    console.error(`    ❌ PDF作成処理中にエラーが発生しました: ${error.message}`);
    // エラーを呼び出し元に再スローして、メインの処理で捕捉できるようにします。
    throw error;
  }
}

// ============================================================================
// メール送信関数
// ============================================================================

/**
 * 請求書PDFのリンクを含むメールを送信する関数
 * 【初心者向け説明】
 * - 作成した請求書PDFへのリンクをメールで送信します
 * - 日本語のビジネスメール形式でメールを作成します
 * - TEST_MODE を true にすると実際には送信せずログだけ出力されます（テスト用）
 *
 * @param {string} recipientEmail - 受信者のメールアドレス
 * @param {string} recipientName - 受信者の名前
 * @param {File} pdfFile - Google DriveのPDFファイルオブジェクト
 * @param {string} invoiceDate - 請求日（YYYY/MM/DD形式）
 * @returns {boolean} メール送信成功時は true、失敗時は false
 */
function sendInvoiceEmail(recipientEmail, recipientName, pdfFile, invoiceDate) {
  console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
  console.log("📧 [開始] メール送信処理");
  console.log(`📬 宛先: ${recipientName} <${recipientEmail}>`);
  console.log(`📅 請求日: ${invoiceDate}`);
  console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");

  try {
    // ステップ1: メールアドレスの検証
    // @マークが含まれているか簡易チェック
    if (!recipientEmail || !recipientEmail.includes("@")) {
      console.error(`❌ 無効なメールアドレスです: ${recipientEmail}`);
      console.error("❌ [完了] sendInvoiceEmail() - 失敗（無効なメールアドレス）");
      console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n");
      return false;
    }
    console.log("✓ メールアドレスの検証完了");

    // ステップ2: PDFファイルの共有URLを取得
    console.log("🔗 PDF共有URLを生成中...");
    const pdfUrl = getPdfShareableUrl(pdfFile);

    // ステップ3: 請求日から年と月を抽出
    // getMonth()は0-11を返すので+1する必要があります
    const dateObj = new Date(invoiceDate);
    const year = dateObj.getFullYear();
    const month = dateObj.getMonth() + 1;
    console.log(`✓ 対象期間: ${year}年${month}月`);

    // ステップ4: メールの件名を作成
    const subject = `請求書送付のご案内 [${year}年${month}月分]`;
    console.log(`✓ 件名: ${subject}`);

    // ステップ5: メール本文を作成（HTML形式）
    console.log("✍️ メール本文を作成中...");
    const body = createEmailBody(recipientName, year, month, pdfUrl);

    // ステップ6: メールオプションを設定
    const options = {
      name: SENDER_NAME, // 送信者名
      htmlBody: body, // HTML形式の本文
      from: SENDER_EMAIL, // 送信者メールアドレス
    };

    // ステップ7: メールを送信（またはテストモードの場合はログ出力のみ）
    if (TEST_MODE) {
      // テストモード: 実際には送信せず、ログだけ出力
      console.log("⚠️ テストモード: メールは実際には送信されません");
      console.log(`   宛先: ${recipientEmail}`);
      console.log(`   件名: ${subject}`);
      console.log(`   PDF URL: ${pdfUrl}`);
      console.log("✅ [完了] sendInvoiceEmail() - 成功（テストモード）");
      console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n");
      return true;
    } else {
      // 本番モード: 実際にメールを送信
      console.log("📤 メールを送信中...");
      GmailApp.sendEmail(recipientEmail, subject, body, options);
      console.log(`✅ メール送信成功: ${recipientEmail}`);
      console.log("✅ [完了] sendInvoiceEmail() - 成功");
      console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n");
      return true;
    }
  } catch (error) {
    console.error(`❌ メール送信に失敗しました: ${recipientEmail}`);
    console.error(`エラー詳細: ${error.message}`);
    console.error("💡 ヒント: Gmail APIの権限、送信者メールアドレス、受信者メールアドレスを確認してください");
    console.error("❌ [完了] sendInvoiceEmail() - 失敗");
    console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n");
    return false;
  }
}

/**
 * Google DriveのPDFファイルの共有URLを生成する関数
 * 【初心者向け説明】
 * - PDFファイルを「リンクを知っている全員が閲覧可能」に設定します
 * - メールで送信できる共有リンクを生成します
 * - 受信者はこのリンクからPDFをダウンロードできます
 *
 * @param {File} pdfFile - PDFファイルオブジェクト
 * @returns {string} PDF共有URL
 */
function getPdfShareableUrl(pdfFile) {
  console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
  console.log("🔗 [開始] PDF共有URL生成");
  console.log(`📄 ファイル名: ${pdfFile.getName()}`);
  console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");

  try {
    // ステップ1: ファイルの共有設定を変更
    // 「リンクを知っている全員」が「閲覧」できるように設定
    console.log("🔓 共有設定を変更中...");
    pdfFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    console.log("✅ 共有設定完了: リンクを知っている全員が閲覧可能");

    // ステップ2: ファイルIDを取得して共有URLを生成
    const fileId = pdfFile.getId();
    const shareableUrl = `https://drive.google.com/file/d/${fileId}/view?usp=sharing`;

    console.log(`✅ 共有URL生成完了`);
    console.log(`   URL: ${shareableUrl}`);
    console.log("✅ [完了] getPdfShareableUrl() - 成功");
    console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n");
    return shareableUrl;
  } catch (error) {
    console.error(`❌ 共有URLの生成に失敗しました: ${error.message}`);
    console.error("💡 ヒント: Google Driveの権限を確認してください");
    console.error("❌ [完了] getPdfShareableUrl() - 失敗");
    console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n");
    throw error;
  }
}

/**
 * 日本語ビジネスメール形式のメール本文を作成する関数
 * 【初心者向け説明】
 * - HTML形式で見栄えの良いメール本文を作成します
 * - 日本語のビジネスメール定型文を使用します
 * - PDF閲覧用のボタンとリンクを含めます
 *
 * @param {string} recipientName - 受信者の名前（「〜様」を付けて表示）
 * @param {number} year - 請求年
 * @param {number} month - 請求月
 * @param {string} pdfUrl - PDFファイルへのURL
 * @returns {string} HTML形式のメール本文
 */
function createEmailBody(recipientName, year, month, pdfUrl) {
  console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");
  console.log("✉️ [開始] メール本文作成");
  console.log(`📝 宛先: ${recipientName}`);
  console.log(`📅 対象期間: ${year}年${month}月`);
  console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━");

  // HTML形式のメール本文を作成
  // スタイルを含めることで、メールクライアントで見栄え良く表示されます
  const htmlBody = `
    <div style="font-family: 'Hiragino Sans', 'Meiryo', sans-serif; line-height: 1.6; color: #333;">
      <!-- 宛名 -->
      <p>${recipientName} 様</p>

      <!-- 挨拶文 -->
      <p>いつもお世話になっております。<br>
      株式会社DROXです。</p>

      <!-- 本文 -->
      <p>${year}年${month}月分の請求書をお送りいたします。<br>
      下記リンクよりご確認ください。</p>

      <!-- PDFリンク（ボタン形式） -->
      <p style="margin: 20px 0;">
        <a href="${pdfUrl}" style="background-color: #2563EB; color: white; padding: 12px 24px; text-decoration: none; border-radius: 6px; display: inline-block; font-weight: 500;">
          請求書を確認する
        </a>
      </p>

      <!-- 確認依頼（目立つ形式） -->
      <div style="background-color: #EFF6FF; border: 2px solid #3B82F6; border-radius: 8px; padding: 20px; margin: 30px 0; text-align: center;">
        <p style="font-size: 14px; color: #1E40AF; margin: 0;">
          問題なければ、<strong style="font-size: 18px; color: #1D4ED8;">「OK DROX!」</strong>と返信お願いします。
        </p>
      </div>

      <!-- 締めの挨拶 -->
      <p style="margin-top: 30px;">
      ご不明な点がございましたら、お気軽にお問い合わせください。<br>
      今後ともよろしくお願いいたします。</p>

      <!-- 区切り線 -->
      <hr style="margin: 30px 0; border: none; border-top: 1px solid #ddd;">

      <!-- 署名 -->
      <p style="font-size: 12px; color: #666;">
      株式会社DROX<br>
      ${SENDER_EMAIL}<br>
      </p>
    </div>
  `;

  console.log("✅ メール本文作成完了");
  console.log(`   本文の長さ: ${htmlBody.length}文字`);
  console.log("✅ [完了] createEmailBody() - 成功");
  console.log("━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n");
  return htmlBody;
}
