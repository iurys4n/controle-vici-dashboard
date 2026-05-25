// Backend único para os dashboards de controle de horas (VAESO e Vici).
// Multi-store: o mesmo deploy serve os dois via ?store=vaeso|vici.
// Mantém o estado atual + um snapshot por dia (histórico p/ rollback).
//
// Compatibilidade: store ausente => "vici", gravando no mesmo lugar de antes
// (_DASHBOARD_DATA!A1), pra não perder os dados já existentes da Vici.

const STORE_CELL = "A1";
const HISTORY_DAYS = 90;

const STORES = {
  vici:  { dataSheet: "_DASHBOARD_DATA",       histSheet: "_HIST_VICI" },
  vaeso: { dataSheet: "_DASHBOARD_DATA_VAESO", histSheet: "_HIST_VAESO" },
};

function doGet(e) {
  const store = resolveStore_(e, null);

  if (e && e.parameter && e.parameter.action === "history") {
    return respond_({ ok: true, history: listHistory_(store) }, e);
  }

  return respond_({ ok: true, data: readData_(store) }, e);
}

function doPost(e) {
  const body = JSON.parse((e && e.postData && e.postData.contents) || "{}");
  const store = resolveStore_(e, body);

  if (body.action === "write" && body.data) {
    const lock = LockService.getScriptLock();
    lock.waitLock(20000);
    try {
      writeData_(store, body.data);
      snapshotDaily_(store, body.data);
    } finally {
      lock.releaseLock();
    }
  }

  return respond_({ ok: true }, e);
}

function resolveStore_(e, body) {
  const key =
    (body && body.store) ||
    (e && e.parameter && e.parameter.store) ||
    "vici";
  return STORES[key] || STORES.vici;
}

function readData_(store) {
  const sheet = getSheet_(store.dataSheet);
  const raw = sheet.getRange(STORE_CELL).getValue();
  if (!raw) return null;
  return JSON.parse(raw);
}

function writeData_(store, data) {
  const sheet = getSheet_(store.dataSheet);
  sheet.getRange(STORE_CELL).setValue(JSON.stringify(data));
  sheet.hideSheet();
}

// Um snapshot por dia: atualiza a linha do dia atual ou cria uma nova,
// e mantém só os últimos HISTORY_DAYS dias.
function snapshotDaily_(store, data) {
  const sheet = getSheet_(store.histSheet);
  const tz = Session.getScriptTimeZone();
  const today = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd");
  const stamp = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss");
  const json = JSON.stringify(data);

  const last = sheet.getLastRow();
  if (last >= 1) {
    const keys = sheet.getRange(1, 1, last, 1).getValues();
    for (let i = 0; i < keys.length; i++) {
      if (keys[i][0] === today) {
        sheet.getRange(i + 1, 1, 1, 3).setValues([[today, stamp, json]]);
        return;
      }
    }
  }

  sheet.appendRow([today, stamp, json]);

  const rows = sheet.getLastRow();
  if (rows > HISTORY_DAYS) {
    sheet.deleteRows(1, rows - HISTORY_DAYS);
  }
}

function listHistory_(store) {
  const sheet = getSheet_(store.histSheet);
  const last = sheet.getLastRow();
  if (last < 1) return [];
  return sheet
    .getRange(1, 1, last, 2)
    .getValues()
    .map((r) => ({ date: r[0], savedAt: r[1] }));
}

function getSheet_(name) {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
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
