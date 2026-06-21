// 中央講義領取記錄器（所有演講共用同一張總表）
// 容器綁定試算表：13SmRlC3oEN7j7V17BWPWFJV3tqzNKuggob-cWi-dXRs
// 部署：執行身分=我、誰可存取=任何人
//
// 用法：各演講網站 POST（sendBeacon, text/plain）一段 JSON：
//   { email, handout: "講義名稱", source: "場次/網址", time: ISO字串 }
// 全部 append 到同一分頁，欄位：時間 / Email / 講義 / 來源

const SHEET_NAME = '講義領取總表';

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['時間', 'Email', '講義', '來源']);
    sheet.getRange(1, 1, 1, 4).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = getSheet_();
    sheet.appendRow([
      data.time ? new Date(data.time) : new Date(),
      data.email || '',
      data.handout || data.course || '',   // 相容舊欄位名 course
      data.source || '',
    ]);
    return ContentService
      .createTextOutput(JSON.stringify({ ok: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doGet(e) {
  // ?count=1 → 回傳目前資料列數，用來驗證寫入是否成功
  if (e && e.parameter && e.parameter.count) {
    try {
      const sheet = getSheet_();
      return ContentService
        .createTextOutput(JSON.stringify({ ok: true, rows: sheet.getLastRow() - 1 }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (err) {
      return ContentService
        .createTextOutput(JSON.stringify({ ok: false, error: err.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }
  return ContentService
    .createTextOutput(JSON.stringify({ status: '中央講義領取記錄器 active' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// 首次部署後在編輯器手動 Run 一次以授權 Sheets 權限
function authorize() {
  const sheet = getSheet_();
  Logger.log('Authorized. 目前列數: ' + (sheet.getLastRow() - 1));
}
