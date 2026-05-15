const STORE_SHEET_NAME = "_DASHBOARD_DATA";
const STORE_CELL = "A1";

function doGet(e) {
  return respond_({
    ok: true,
    data: readDashboardData_(),
  }, e);
}

function doPost(e) {
  const body = JSON.parse(e.postData.contents || "{}");

  if (body.action === "write" && body.data) {
    writeDashboardData_(body.data);
  }

  return respond_({ ok: true }, e);
}

function readDashboardData_() {
  const sheet = getStoreSheet_();
  const raw = sheet.getRange(STORE_CELL).getValue();
  if (!raw) return null;
  return JSON.parse(raw);
}

function writeDashboardData_(data) {
  const sheet = getStoreSheet_();
  sheet.getRange(STORE_CELL).setValue(JSON.stringify(data));
  sheet.hideSheet();
}

function getStoreSheet_() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(STORE_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(STORE_SHEET_NAME);
    sheet.hideSheet();
  }
  return sheet;
}

function respond_(payload, e) {
  const json = JSON.stringify(payload);
  const callback = e && e.parameter && e.parameter.callback;

  if (callback) {
    return ContentService
      .createTextOutput(`${callback}(${json});`)
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService
    .createTextOutput(json)
    .setMimeType(ContentService.MimeType.JSON);
}
