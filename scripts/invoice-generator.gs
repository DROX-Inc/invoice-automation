/**
 * Google Apps Script - Invoice PDF Generator
 *
 * This script automatically generates invoice PDFs from spreadsheet data.
 *
 * Features:
 * - Header-based data mapping (column order independent)
 * - Automatic date formatting
 * - Batch processing
 * - Error handling
 *
 * Setup Instructions:
 * 1. Update the configuration section below with your IDs
 * 2. Ensure your data spreadsheet has headers in Row 1
 * 3. Use {{header_name}} placeholders in your template
 * 4. Run createInvoices() function
 */

// ============================================================================
// CONFIGURATION SECTION
// Update these values with your own spreadsheet and folder IDs
// ============================================================================

/**
 * Spreadsheet ID containing invoice data
 * Find it in the URL: https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit
 */
const DATA_SPREADSHEET_ID = "YOUR_DATA_SPREADSHEET_ID";

/**
 * Spreadsheet ID containing invoice template
 * Find it in the URL: https://docs.google.com/spreadsheets/d/SPREADSHEET_ID/edit
 */
const TEMPLATE_SPREADSHEET_ID = "YOUR_TEMPLATE_SPREADSHEET_ID";

/**
 * Google Drive Folder ID where PDFs will be saved
 * Find it in the URL: https://drive.google.com/drive/folders/FOLDER_ID
 */
const OUTPUT_FOLDER_ID = "YOUR_DRIVE_FOLDER_ID";

/**
 * Name of the sheet containing invoice data in the data spreadsheet
 * Default is usually 'Sheet1'
 */
const DATA_SHEET_NAME = "Sheet1";

/**
 * Name of the sheet containing the invoice template
 * Default is usually 'Sheet1'
 */
const TEMPLATE_SHEET_NAME = "Sheet1";

// ============================================================================
// MAIN FUNCTIONS
// ============================================================================

function createInvoices() {
  // --- スプレッドシートやフォルダを取得 ---

  const dataSpreadsheet = SpreadsheetApp.openById(DATA_SPREADSHEET_ID);

  const dataSheet = dataSpreadsheet.getSheetByName(DATA_SHEET_NAME);

  if (!dataSheet) {
    SpreadsheetApp.getUi().alert(`データシート「${DATA_SHEET_NAME}」が見つかりません。`);

    return;
  }

  const templateSpreadsheet = SpreadsheetApp.openById(TEMPLATE_SPREADSHEET_ID);

  const templateSheet = templateSpreadsheet.getSheetByName(TEMPLATE_SHEET_NAME);

  if (!templateSheet) {
    SpreadsheetApp.getUi().alert(`テンプレートシート「${TEMPLATE_SHEET_NAME}」が見つかりません。`);

    return;
  }

  const outputFolder = DriveApp.getFolderById(OUTPUT_FOLDER_ID); // --- データシートから請求書情報を取得 --- // getValues()でシートの全データを二次元配列として取得

  const data = dataSheet.getDataRange().getValues(); // ヘッダー行（1行目）を除外して、2行目から処理を開始

  const invoiceDataRows = data.slice(1); // --- 請求書データを1行ずつ処理 ---

  invoiceDataRows.forEach((row) => {
    // A列からF列のデータをそれぞれ変数に格納

    const [
      invoiceDate, // 請求日 (A列)
      invoiceNumber, // 請求書番号 (B列)
      companyName, // 会社名 (C列)
      subject, // 件名 (D列)
      amount, // 金額 (E列)
      pdfFileName, // PDFファイル名 (F列)
    ] = row; // 会社名が空欄の行はスキップ

    if (!companyName) {
      return;
    } // --- 請求書作成処理 --- // 1. テンプレートシートをコピーして、作業用シートを作成 // 作業用シートはデータシート側に作成される

    const copiedSheet = templateSheet.copyTo(dataSpreadsheet);

    const newSheetName = `temp_${invoiceNumber}`; // 一時的なシート名

    copiedSheet.setName(newSheetName); // 2. プレースホルダーを実際のデータに置換 // 日付のフォーマットを 'yyyy/MM/dd' 形式に整える

    const formattedDate = Utilities.formatDate(new Date(invoiceDate), "JST", "yyyy/MM/dd");

    copiedSheet.createTextFinder("{{請求日}}").replaceAllWith(formattedDate);

    copiedSheet.createTextFinder("{{請求書番号}}").replaceAllWith(invoiceNumber);

    copiedSheet.createTextFinder("{{会社名}}").replaceAllWith(companyName);

    copiedSheet.createTextFinder("{{件名}}").replaceAllWith(subject); // 金額は数値として転記（テンプレートの書式設定が適用される）

    copiedSheet.createTextFinder("{{金額}}").replaceAllWith(amount); // 3. スプレッドシートへの変更を即時反映させる

    SpreadsheetApp.flush(); // 4. PDFを作成してGoogle Driveに保存

    try {
      createPdfInDrive(dataSpreadsheet, copiedSheet.getSheetId(), outputFolder, pdfFileName);

      console.log(`${pdfFileName}.pdf を作成しました。`);
    } catch (e) {
      console.error(`${pdfFileName}.pdf の作成中にエラーが発生しました: ${e.message}`);
    } // 5. 不要になった作業用シートを削除

    dataSpreadsheet.deleteSheet(copiedSheet);
  });

  SpreadsheetApp.getUi().alert("すべての請求書PDFの作成が完了しました。");
}

/**

 * 指定されたシートをPDFとしてGoogle Driveに保存する関数

 * @param {Spreadsheet} spreadsheet - PDF化するシートが含まれるスプレッドシート

 * @param {string} sheetId - PDF化するシートのID

 * @param {Folder} folder - 保存先のGoogle Driveフォルダ

 * @param {string} fileName - 保存するPDFのファイル名（拡張子なし）

 */
function createPdfInDrive(spreadsheet, sheetId, folder, fileName) {
  const spreadsheetId = spreadsheet.getId();

  // PDF出力用のURLを作成
  const url =
    `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?` +
    "format=pdf" +
    "&gid=" +
    sheetId + // シートIDを指定
    "&size=A4" + // 用紙サイズ（A3/A4/A5/letter/legal等）
    "&portrait=true" + // 縦向き（false = 横向き）
    "&scale=3" + // スケール設定：
    // 1 = 通常 (100%)
    // 2 = 幅に合わせる
    // 3 = 高さに合わせる ★おすすめ
    // 4 = ページに合わせる
    "&top_margin=0.2" + // 上マージン（インチ）
    "&bottom_margin=0.2" + // 下マージン（インチ）
    "&left_margin=0.2" + // 左マージン（インチ）
    "&right_margin=0.2" + // 右マージン（インチ）
    "&sheetnames=false" + // シート名を非表示
    "&printtitle=false" + // スプレッドシート名を非表示
    "&gridlines=false" + // グリッド線を非表示
    "&fzr=false" + // 固定行を無視
    "&horizontal_alignment=CENTER" + // 水平方向の配置（LEFT/CENTER/RIGHT）
    "&vertical_alignment=TOP"; // 垂直方向の配置（TOP/MIDDLE/BOTTOM）

  const token = ScriptApp.getOAuthToken();
  const options = {
    headers: {
      Authorization: "Bearer " + token,
    },
    muteHttpExceptions: true,
  };

  // URLからPDFデータを取得
  const response = UrlFetchApp.fetch(url, options);
  const blob = response.getBlob().setName(fileName + ".pdf");

  // フォルダにPDFファイルを作成
  folder.createFile(blob);
}
