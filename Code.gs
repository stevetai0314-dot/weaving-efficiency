const SHEET_ID = '1cVAV8odDf457CJHGuV_Yy7JOB97vbiCWy2Y4x-bfE4M';
const RECORD_TAB = '登記紀錄';

// ── 前端讀取上次時數 ──────────────────────────
function doGet(e) {
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(RECORD_TAB);

    if (!sheet || sheet.getLastRow() <= 1) {
      return jsonOut({});
    }

    const data = sheet.getDataRange().getValues();
    const h = data[0];
    const C = {
      time:    h.indexOf('時間戳記'),
      machine: h.indexOf('機台號'),
      current: h.indexOf('本次時數'),
      stopped: h.indexOf('停機'),
    };

    // 每台機台取最新一筆非停機的本次時數
    const latest = {};
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const id  = String(row[C.machine]);
      if (!id) continue;
      if (row[C.stopped]) continue;
      if (row[C.current] === '' || row[C.current] === null) continue;

      const t = new Date(row[C.time]);
      if (!latest[id] || t > latest[id].t) {
        latest[id] = { t, hours: row[C.current] };
      }
    }

    const result = {};
    Object.keys(latest).forEach(k => { result[k] = latest[k].hours; });
    return jsonOut(result);

  } catch (err) {
    return jsonOut({ error: err.message });
  }
}

// ── 前端送出登記資料 ──────────────────────────
function doPost(e) {
  try {
    const payload = JSON.parse(e.parameter.data);
    const ss    = SpreadsheetApp.openById(SHEET_ID);
    let sheet   = ss.getSheetByName(RECORD_TAB);

    if (!sheet) {
      sheet = ss.insertSheet(RECORD_TAB);
    }

    // 建立表頭（第一次使用）
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['時間戳記','班別','組別','機台號','上次時數','本次時數','稼動率%','停機']);
      sheet.setFrozenRows(1);
    }

    const ts      = Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'yyyy-MM-dd HH:mm:ss');
    const shift   = payload.shift;
    const group   = payload.group;
    const records = payload.records;

    const rows = records.map(r => [
      ts,
      shift,
      group,
      r.machineId,
      r.lastHours   || '',
      r.stopped ? '' : r.currentHours,
      r.stopped ? '' : r.efficiency,
      r.stopped ? 1  : 0,
    ]);

    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    return jsonOut({ status: 'ok', written: rows.length });

  } catch (err) {
    return jsonOut({ status: 'error', message: err.message });
  }
}

function jsonOut(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
