// 環 -wa- 予約フォーム — Google Apps Script
// デプロイ手順は README.md を参照してください

var SPREADSHEET_ID    = '15pwB9beyd2nmvM2_GMgmHqL72Sa9z5TajrDbF6hCyG4';
var SHEET_NAME        = '予約';
var NOTIFY_EMAIL      = 'wa.sasebo@gmail.com';

/**
 * フォームからの GET リクエスト（URLSearchParams）を受け取り
 * スプレッドシートへ記録 → 通知メール送信
 *
 * カラム構成:
 *   A: 予約受付日時  B: 体験日付  C: 時間帯  D: 体験メニュー
 *   E: 人数（大人）  F: 子供の人数  G: お名前  H: 電話番号
 *   I: メールアドレス  J: ご要望・備考  K: ステータス
 */
function doGet(e) {
  var output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);

  try {
    var p = e.parameter || {};

    // デモモード（パラメータなし）
    if (!p.date) {
      output.setContent(JSON.stringify({ status: 'ok', service: '環 -wa- 予約API' }));
      return output;
    }

    writeToSheet(p);
    sendNotificationEmail(p);

    output.setContent(JSON.stringify({ result: 'success' }));

  } catch (err) {
    Logger.log('Error: ' + err.toString());
    output.setContent(JSON.stringify({ result: 'error', message: err.toString() }));
  }

  return output;
}

function writeToSheet(p) {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    var headers = [
      '予約受付日時', '体験日付', '時間帯', '体験メニュー',
      '人数（大人）', '子供の人数', 'お名前', '電話番号',
      'メールアドレス', 'ご要望・備考', 'ステータス'
    ];
    sheet.appendRow(headers);
    var hRange = sheet.getRange(1, 1, 1, headers.length);
    hRange.setBackground('#5C3423');
    hRange.setFontColor('#FFFFFF');
    hRange.setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  var now = Utilities.formatDate(new Date(), 'Asia/Tokyo', 'yyyy/MM/dd HH:mm:ss');

  sheet.appendRow([
    now,
    p.date     || '',
    p.time     || '',
    p.menu     || '',
    p.adults   || '',
    p.children || '0',
    p.name     || '',
    p.phone    || '',
    p.email    || '',
    p.notes    || '',
    '未確認'
  ]);
}

function sendNotificationEmail(p) {
  var subject = '【環 -wa-】予約リクエストが届きました';

  var body = [
    '以下の内容で予約リクエストを受け付けました。',
    '',
    '─────────────────────────',
    '体験日付　：' + (p.date  || ''),
    '時間帯　　：' + (p.time  || ''),
    '体験メニュー：' + (p.menu || ''),
    '人数（大人）：' + (p.adults   || '') + '名',
    '子供の人数　：' + (p.children || '0') + '名',
    '─────────────────────────',
    'お名前　　：' + (p.name  || ''),
    '電話番号　：' + (p.phone || ''),
    'メール　　：' + (p.email || ''),
    'ご要望・備考：' + (p.notes || 'なし'),
    '─────────────────────────',
    '',
    'スプレッドシートでステータスをご確認ください。',
  ].join('\n');

  MailApp.sendEmail({
    to:      NOTIFY_EMAIL,
    subject: subject,
    body:    body
  });
}

/**
 * POST フォールバック（旧クライアント向け）
 */
function doPost(e) {
  var output = ContentService.createTextOutput();
  output.setMimeType(ContentService.MimeType.JSON);

  try {
    var raw  = (e && e.postData && e.postData.contents) ? e.postData.contents : '{}';
    var data = JSON.parse(raw);
    writeToSheet(data);
    sendNotificationEmail(data);
    output.setContent(JSON.stringify({ result: 'success' }));
  } catch (err) {
    Logger.log('Error: ' + err.toString());
    output.setContent(JSON.stringify({ result: 'error', message: err.toString() }));
  }

  return output;
}
