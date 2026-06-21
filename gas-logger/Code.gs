// TB 透析患者抗結核用藥 講義領取 — Email 記錄
// 容器綁定試算表：13SmRlC3oEN7j7V17BWPWFJV3tqzNKuggob-cWi-dXRs
// 部署為 Web App：執行身分=我、誰可存取=任何人

const SHEET_NAME = '講義領取名單';

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(['時間', 'Email', '課程']);
      sheet.getRange(1, 1, 1, 3).setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    sheet.appendRow([
      data.time ? new Date(data.time) : new Date(),
      data.email || '',
      data.course || '',
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

function doGet() {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'TB dialysis handout logger active' }))
    .setMimeType(ContentService.MimeType.JSON);
}

// 首次部署後執行一次以授權 Sheets 權限（在編輯器手動 Run）
function authorize() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Logger.log('Authorized for: ' + ss.getName());
}
