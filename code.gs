/**
 * ==============================================================================
 * B品報告アプリ - バックエンドスクリプト (code.gs)
 * ==============================================================================
 * 
 * 【主な機能】
 * 1. LockService による厳格な排他制御（同時書き込み時の通し番号重複防止）
 * 2. シート自動生成（「B品報告実績」「商品マスタ」）
 * 3. 柔軟なスプレッドシート指定（URL指定 / ID指定 / バインド型自動検出）
 * 4. 商品マスタ全件取得 Web API (doGet: ?action=getMaster)
 * 5. B品・直し品移動データ登録 & 自動採番 (doPost / google.script.run)
 * 6. 通し番号（No.〇〇）の返却
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

// ヘッダー行定義（全12列）
const RECORD_HEADERS = [
  '通し番号',        // A列 (1から連番)
  '登録日時',        // B列 (yyyy/MM/dd HH:mm:ss)
  '作業者名',        // C列
  '移動元倉庫コード',// D列
  '移動先倉庫コード',// E列 (1008 または 1006)
  '発生理由',        // F列
  '商品コード',      // G列
  '商品名1',         // H列
  '商品名2',         // I列
  'サイズ',          // J列
  'JANコード',       // K列 (文字列として保持)
  '数量'             // L列
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

  // 2. 疎通確認・現在番号確認 API
  if (action === 'ping') {
    try {
      const ss = getTargetSpreadsheet(spreadsheetUrl);
      const sheet = getOrCreateRecordSheet(ss);
      const lastRow = sheet.getLastRow();
      let currentLastNo = 0;
      if (lastRow > 1) {
        const val = sheet.getRange(lastRow, 1).getValue();
        currentLastNo = Number(val) || 0;
      }
      return createJsonResponse({
        status: 'ok',
        success: true,
        sheetName: ss.getName(),
        lastNumber: currentLastNo,
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
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
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
 * B品報告データの登録・自動採番（排他制御 LockService 適用）
 * ※ google.script.run からも直接呼び出し可能
 * 
 * @param {Object} data 報告データ
 * @return {Object} 採番結果 { success: true, assignedNumber: 42, assignedNumbers: [42], ... }
 */
function sendInventoryData(data) {
  // ★ 重要要件: LockService による厳格な排他制御
  const lock = LockService.getScriptLock();
  const hasLock = lock.waitLock(10000); // 10秒待機

  if (!hasLock) {
    return {
      success: false,
      error: '他ユーザーが書き込み中のため排他ロックを取得できませんでした。数秒後に再送してください。'
    };
  }

  try {
    const ss = getTargetSpreadsheet(data && (data.spreadsheetUrl || data.spreadsheetId));
    const sheet = getOrCreateRecordSheet(ss);

    const lastRow = sheet.getLastRow();
    let nextSeqNo = 1;

    if (lastRow > 1) {
      // 最終行のA列通し番号を取得して +1
      const lastVal = sheet.getRange(lastRow, 1).getValue();
      const parsedNo = Number(lastVal);
      if (!isNaN(parsedNo) && parsedNo >= 1) {
        nextSeqNo = parsedNo + 1;
      } else {
        // 安全策: A列の最大値を走査
        const aValues = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
        let maxNo = 0;
        for (let i = 0; i < aValues.length; i++) {
          const num = Number(aValues[i][0]);
          if (!isNaN(num) && num > maxNo) {
            maxNo = num;
          }
        }
        nextSeqNo = maxNo + 1;
      }
    }

    // 登録日時 (JST yyyy/MM/dd HH:mm:ss)
    const formattedTimestamp = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');

    // ★ 1商品1採番（個体管理）の自動展開要件:
    // 数量 N (例: 3) の場合、スプレッドシートへ N行 (例: 3行) 展開して書き込み、各行の数量列には 1 を記録
    const quantity = Math.max(1, parseInt(data.quantity, 10) || 1);

    // JANコード文字列のフォーマット（先頭0落ち防止）
    const janString = data.janCode ? "'" + String(data.janCode).trim() : '';

    const rows = [];
    const assignedNumbers = [];

    for (let i = 0; i < quantity; i++) {
      const currentSeqNo = nextSeqNo + i;
      assignedNumbers.push(currentSeqNo);

      // 書込行データの組み立て（全12列、数量は各行常に 1）
      rows.push([
        currentSeqNo,                                  // 1. 通し番号 (A列)
        formattedTimestamp,                            // 2. 登録日時 (B列)
        data.workerName || '',                         // 3. 作業者名
        data.sourceWarehouse || '',                    // 4. 移動元倉庫コード
        data.destWarehouse || '',                      // 5. 移動先倉庫コード (1008 または 1006)
        data.reason || '',                             // 6. 発生理由
        data.productCode || '',                        // 7. 商品コード
        data.name1 || '',                              // 8. 商品名1
        data.name2 || '',                              // 9. 商品名2
        data.size || '',                               // 10. サイズ
        janString,                                     // 11. JANコード
        1                                              // 12. 数量 (常に 1 で1行ずつ展開)
      ]);
    }

    // 新規行への一括書き込み
    const targetRowStart = lastRow + 1;
    sheet.getRange(targetRowStart, 1, quantity, 12).setValues(rows);

    // 書式設定（通し番号: 数値フォーマット、JANコード: テキスト表示フォーマット）
    sheet.getRange(targetRowStart, 1, quantity, 1).setNumberFormat('#,##0');
    sheet.getRange(targetRowStart, 11, quantity, 1).setNumberFormat('@');

    SpreadsheetApp.flush(); // 即時反映

    const startNumber = nextSeqNo;
    const endNumber = nextSeqNo + quantity - 1;

    return {
      success: true,
      assignedNumber: startNumber,
      assignedNumbers: assignedNumbers,
      startNumber: startNumber,
      endNumber: endNumber,
      count: quantity,
      timestamp: formattedTimestamp,
      sheetName: ss.getName(),
      productName: (data.name1 || '') + ' ' + (data.name2 || ''),
      message: quantity > 1
        ? '正常に記録され、' + quantity + '点分の通し番号（No.' + startNumber + ' ～ No.' + endNumber + '）が発番されました。'
        : '正常に記録され、通し番号（No.' + startNumber + '）が発番されました。'
    };

  } catch (error) {
    return {
      success: false,
      error: 'データ保存エラー: ' + error.toString()
    };
  } finally {
    // ロックを確実に解放
    lock.releaseLock();
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
    // ヘッダー行書き込み
    sheet.appendRow(RECORD_HEADERS);

    // ヘッダーデザイン
    const headerRange = sheet.getRange(1, 1, 1, RECORD_HEADERS.length);
    headerRange.setBackground('#1e293b'); // スレートダーク
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    headerRange.setHorizontalAlignment('center');
    sheet.setFrozenRows(1);

    // 列幅の自動調整
    sheet.setColumnWidth(1, 90);   // 通し番号
    sheet.setColumnWidth(2, 160);  // 登録日時
    sheet.setColumnWidth(3, 110);  // 作業者名
    sheet.setColumnWidth(4, 120);  // 移動元
    sheet.setColumnWidth(5, 120);  // 移動先
    sheet.setColumnWidth(6, 120);  // 発生理由
    sheet.setColumnWidth(7, 130);  // 商品コード
    sheet.setColumnWidth(8, 200);  // 商品名1
    sheet.setColumnWidth(9, 180);  // 商品名2
    sheet.setColumnWidth(10, 80);  // サイズ
    sheet.setColumnWidth(11, 140); // JANコード
    sheet.setColumnWidth(12, 80);  // 数量
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
