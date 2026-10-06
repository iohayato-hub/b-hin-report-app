/**
 * ==============================================================================
 * B品報告アプリ - バックエンドスクリプト (code.gs)
 * ==============================================================================
 * 
 * 【仕様概要】
 * 1. スプレッドシートへの直接追記（LockService・通し番号採番なし・シンプル高安定）
 * 2. シート自動生成（「B品報告実績」全11列、「商品マスタ」）
 * 3. 柔軟なスプレッドシート指定（URL指定 / ID指定 / バインド型自動検出）
 * 4. 商品マスタ全件取得 Web API (doGet: ?action=getMaster)
 * 5. B品・直し品移動データ登録 (doPost / google.script.run)
 * 
 * 【スプレッドシート列構成（全11列）】
 * 1. 登録日時 (yyyy/MM/dd HH:mm:ss)
 * 2. 作業者名
 * 3. 移動元倉庫コード
 * 4. 移動先倉庫コード (1008 または 1006)
 * 5. 発生理由
 * 6. 商品コード
 * 7. 商品名1
 * 8. 商品名2
 * 9. サイズ
 * 10. JANコード (文字列表示 @ 指定)
 * 11. 数量
 */

// ==============================================================================
// 【設定】スプレッドシートの指定（URL または ID）
// ==============================================================================
// ※スプレッドシート画面の「拡張機能 > Apps Script」から開いている場合は空欄（''）のままでOKです。
// ※スタンドアロンスクリプトの場合や、特定のスプレッドシートに書き込みたい場合は
//   ここにスプレッドシートのURL（またはID）を貼り付けてください。
//   例: const SPREADSHEET_URL_OR_ID = 'https://docs.google.com/spreadsheets/d/1xxxx/edit';
const SPREADSHEET_URL_OR_ID = '';

// シート名定義
const SHEET_RECORD_NAME = 'B品報告実績';
const SHEET_MASTER_NAME = '商品マスタ';

// ヘッダー行定義（全11列）
const RECORD_HEADERS = [
  '登録日時',        // 1. A列 (yyyy/MM/dd HH:mm:ss)
  '作業者名',        // 2. B列
  '移動元倉庫コード',// 3. C列
  '移動先倉庫コード',// 4. D列 (1008 または 1006)
  '発生理由',        // 5. E列
  '商品コード',      // 6. F列
  '商品名1',         // 7. G列
  '商品名2',         // 8. H列
  'サイズ',          // 9. I列
  'JANコード',       // 10. J列 (文字列として保持)
  '数量'             // 11. K列
];

// 商品マスタのヘッダー定義
const MASTER_HEADERS = [
  '商品コード',
  '商品名1',
  '商品名2',
  'サイズ',
  'JANコード'
];

/**
 * 対象スプレッドシートを取得（URL指定 / ID指定 / getActiveSpreadsheet）
 */
function getTargetSpreadsheet(optionalUrlOrId) {
  const target = (optionalUrlOrId && String(optionalUrlOrId).trim()) 
    || SPREADSHEET_URL_OR_ID.trim();

  if (target) {
    try {
      if (target.startsWith('http://') || target.startsWith('https://')) {
        return SpreadsheetApp.openByUrl(target);
      } else {
        return SpreadsheetApp.openById(target);
      }
    } catch (err) {
      throw new Error('指定されたスプレッドシートを開けませんでした。URL/IDおよびGoogleアカウントの共有権限を確認してください: ' + err.toString());
    }
  }

  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) {
    return active;
  }

  throw new Error('【スプレッドシート未指定エラー】\n対象のスプレッドシートが見つかりません。\n\n【解決策】\n① スプレッドシートを開き、上のメニューの「拡張機能 > Apps Script」を開いてこのコードをデプロイする。\nまたは\n② code.gs 先頭の SPREADSHEET_URL_OR_ID にスプレッドシートのURLを貼り付ける。\nまたは\n③ アプリ画面下の設定欄でスプレッドシートURLを入力してください。');
}

/**
 * Web App の GET リクエスト処理
 * - HTML画面の表示
 * - マスタデータ取得 API (?action=getMaster)
 * - 疎通確認 API (?action=ping)
 */
function doGet(e) {
  const action = (e && e.parameter && e.parameter.action) ? e.parameter.action : '';
  const spreadsheetUrl = (e && e.parameter && e.parameter.spreadsheetUrl) ? e.parameter.spreadsheetUrl : '';

  // 1. 商品マスタ取得 API
  if (action === 'getMaster') {
    try {
      const ss = getTargetSpreadsheet(spreadsheetUrl);
      const masterData = getProductMasterData(ss);
      return createJsonResponse({
        success: true,
        master: masterData,
        count: masterData.length,
        sheetName: ss.getName()
      });
    } catch (err) {
      return createJsonResponse({
        success: false,
        error: err.toString()
      });
    }
  }

  // 2. 疎通確認 API
  if (action === 'ping') {
    try {
      const ss = getTargetSpreadsheet(spreadsheetUrl);
      const sheet = getOrCreateRecordSheet(ss);
      return createJsonResponse({
        status: 'ok',
        success: true,
        sheetName: ss.getName(),
        rowCount: sheet.getLastRow(),
        timestamp: Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss')
      });
    } catch (err) {
      return createJsonResponse({
        status: 'error',
        success: false,
        error: err.toString()
      });
    }
  }

  // 3. GAS Web AppとしてHTMLを直接配信する場合
  try {
    return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('B品報告アプリ')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, maximum-scale=1, user-scalable=no, viewport-fit=cover')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  } catch (err) {
    return ContentService.createTextOutput('B品報告アプリ GAS Backend is running.');
  }
}

/**
 * Web App の POST リクエスト処理 (外部Webアプリ / モバイル端末からの送信)
 */
function doPost(e) {
  try {
    let payload = null;
    if (e && e.postData && e.postData.contents) {
      const parsed = JSON.parse(e.postData.contents);
      payload = parsed.payload ? parsed.payload : parsed;
    } else if (e && e.parameter) {
      payload = e.parameter;
    }

    if (!payload) {
      return createJsonResponse({
        success: false,
        error: '送信データが空です。'
      });
    }

    const result = sendInventoryData(payload);
    return createJsonResponse(result);
  } catch (err) {
    return createJsonResponse({
      success: false,
      error: 'POSTリクエスト処理中にエラーが発生しました: ' + err.toString()
    });
  }
}

/**
 * B品報告データの登録（通し番号・排他ロックなしのシンプル追加処理）
 * ※ google.script.run からも直接呼び出し可能
 * 
 * @param {Object} data 報告データ
 * @return {Object} 結果 { success: true, timestamp, sheetName, message }
 */
function sendInventoryData(data) {
  try {
    const ss = getTargetSpreadsheet(data && (data.spreadsheetUrl || data.spreadsheetId));
    const sheet = getOrCreateRecordSheet(ss);

    // 登録日時 (JST yyyy/MM/dd HH:mm:ss)
    const formattedTimestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');

    // 数量
    const quantity = Math.max(1, parseInt(data.quantity, 10) || 1);

    // JANコード文字列のフォーマット（先頭0落ち防止）
    const janString = data.janCode ? "'" + String(data.janCode).trim() : '';

    // 書込行データ（全11列）
    const row = [
      formattedTimestamp,           // 1. 登録日時 (A列)
      data.workerName || '',        // 2. 作業者名 (B列)
      data.sourceWarehouse || '',   // 3. 移動元倉庫コード (C列)
      data.destWarehouse || '',     // 4. 移動先倉庫コード (D列: 1008 または 1006)
      data.reason || '',            // 5. 発生理由 (E列)
      data.productCode || '',       // 6. 商品コード (F列)
      data.name1 || '',             // 7. 商品名1 (G列)
      data.name2 || '',             // 8. 商品名2 (H列)
      data.size || '',              // 9. サイズ (I列)
      janString,                    // 10. JANコード (J列)
      quantity                      // 11. 数量 (K列)
    ];

    // スプレッドシートへ追記
    sheet.appendRow(row);

    // JANコード（J列: 第10列）をテキスト書式に設定
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 10).setNumberFormat('@');

    SpreadsheetApp.flush(); // 即時反映

    return {
      success: true,
      timestamp: formattedTimestamp,
      sheetName: ss.getName(),
      productName: (data.name1 || '') + ' ' + (data.name2 || ''),
      quantity: quantity,
      message: 'スプレッドシートへ正常に登録されました。'
    };

  } catch (error) {
    return {
      success: false,
      error: 'データ保存エラー: ' + error.toString()
    };
  }
}

/**
 * 商品マスタシートから全件取得
 */
function getProductMasterData(optionalSs) {
  const ss = optionalSs || getTargetSpreadsheet();
  const sheet = getOrCreateMasterSheet(ss);
  const lastRow = sheet.getLastRow();

  if (lastRow <= 1) {
    return [];
  }

  // A2:E最終行を取得
  const range = sheet.getRange(2, 1, lastRow - 1, 5);
  const values = range.getValues();
  const master = [];

  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const productCode = String(row[0] || '').trim();
    const name1 = String(row[1] || '').trim();
    const name2 = String(row[2] || '').trim();
    const size = String(row[3] || '').trim();
    const janCode = String(row[4] || '').trim();

    if (productCode || janCode) {
      master.push({
        productCode: productCode,
        name1: name1,
        name2: name2,
        size: size,
        janCode: janCode
      });
    }
  }

  return master;
}

/**
 * 「B品報告実績」シートを取得（なければヘッダー付きで新規作成）
 */
function getOrCreateRecordSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_RECORD_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_RECORD_NAME, 0);
    // ヘッダー行書き込み（全11列）
    sheet.appendRow(RECORD_HEADERS);

    // ヘッダーデザイン
    const headerRange = sheet.getRange(1, 1, 1, RECORD_HEADERS.length);
    headerRange.setBackground('#1e293b'); // スレートダーク
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');
    sheet.setFrozenRows(1);

    // 列幅の調整（全11列）
    sheet.setColumnWidth(1, 160);  // 1. 登録日時
    sheet.setColumnWidth(2, 110);  // 2. 作業者名
    sheet.setColumnWidth(3, 120);  // 3. 移動元倉庫コード
    sheet.setColumnWidth(4, 120);  // 4. 移動先倉庫コード
    sheet.setColumnWidth(5, 130);  // 5. 発生理由
    sheet.setColumnWidth(6, 130);  // 6. 商品コード
    sheet.setColumnWidth(7, 200);  // 7. 商品名1
    sheet.setColumnWidth(8, 180);  // 8. 商品名2
    sheet.setColumnWidth(9, 80);   // 9. サイズ
    sheet.setColumnWidth(10, 140); // 10. JANコード
    sheet.setColumnWidth(11, 80);  // 11. 数量
  }
  return sheet;
}

/**
 * 「商品マスタ」シートを取得（なければサンプルデータ付きで新規作成）
 */
function getOrCreateMasterSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_MASTER_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_MASTER_NAME);
    sheet.appendRow(MASTER_HEADERS);

    const headerRange = sheet.getRange(1, 1, 1, MASTER_HEADERS.length);
    headerRange.setBackground('#0f766e'); // ティール
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');
    sheet.setFrozenRows(1);

    // サンプルマスタデータの投入
    const sampleRows = [
      ['KN-2024-001', 'プレミアムメリノウールセーター', 'クルーネック / ネイビー', 'M', "'4901234567890"],
      ['KN-2024-002', 'プレミアムメリノウールセーター', 'クルーネック / ネイビー', 'L', "'4901234567891"],
      ['KN-2024-003', 'プレミアムメリノウールセーター', 'クルーネック / チャコール', 'M', "'4901234567892"],
      ['CD-5010-WHT', 'ハイテンションコード編みパーカー', 'ストレッチドローコード / ホワイト', 'FREE', "'4560123456789"],
      ['CD-5011-BLK', 'ハイテンションコード編みパーカー', 'ストレッチドローコード / ブラック', 'FREE', "'4560123456796"],
      ['SW-8801-BEG', 'ローゲージケーブルニットカーディガン', 'ウッドボタン / ベージュ', 'S', "'4912345678012"],
      ['SW-8802-BEG', 'ローゲージケーブルニットカーディガン', 'ウッドボタン / ベージュ', 'M', "'4912345678029"],
      ['SW-8803-BEG', 'ローゲージケーブルニットカーディガン', 'ウッドボタン / ベージュ', 'L', "'4912345678036"],
      ['CS-1020-001', 'ヘビーウェイト長袖カットソー', '天竺編みバインダーネック / オフ白', 'M', "'4923456789013"],
      ['CS-1020-002', 'ヘビーウェイト長袖カットソー', '天竺編みバインダーネック / オフ白', 'L', "'4923456789020"],
      ['PT-3301-KHK', 'ストレッチリブイージーパンツ', 'スピンドルコード仕様 / カーキ', 'M', "'4934567890124"],
      ['PT-3302-KHK', 'ストレッチリブイージーパンツ', 'スピンドルコード仕様 / カーキ', 'L', "'4934567890131"]
    ];

    const target = sheet.getRange(2, 1, sampleRows.length, 5);
    target.setValues(sampleRows);
    sheet.getRange(2, 5, sampleRows.length, 1).setNumberFormat('@');

    sheet.setColumnWidth(1, 130);
    sheet.setColumnWidth(2, 220);
    sheet.setColumnWidth(3, 200);
    sheet.setColumnWidth(4, 90);
    sheet.setColumnWidth(5, 140);
  }
  return sheet;
}

/**
 * CORS対応 JSONレスポンスの生成
 */
function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}
