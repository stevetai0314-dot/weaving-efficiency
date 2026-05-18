const SHEET_ID = '1cVAV8odDf457CJHGuV_Yy7JOB97vbiCWy2Y4x-bfE4M';
const RECORD_TAB = '登記紀錄';

// ── 前端讀取上次時數 ──────────────────────────
// 回傳格式：{ "廠別:機台號": 本次時數, ... }
// 舊資料無廠別欄位時向下相容，預設視為 織一
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
      factory: h.indexOf('廠別'),
      machine: h.indexOf('機台號'),
      current: h.indexOf('本次時數'),
      stopped: h.indexOf('停機'),
    };

    // Pass 1: find the latest date per machine key
    const latestDate = {};
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const id  = String(row[C.machine]);
      if (!id) continue;
      if (row[C.stopped]) continue;
      if (row[C.current] === '' || row[C.current] === null) continue;

      const factory = (C.factory >= 0 && row[C.factory]) ? String(row[C.factory]) : '織一';
      const key = factory + ':' + id;

      const t = new Date(row[C.time]);
      const dateStr = Utilities.formatDate(t, 'Asia/Ho_Chi_Minh', 'yyyy-MM-dd');
      if (!latestDate[key] || dateStr > latestDate[key]) {
        latestDate[key] = dateStr;
      }
    }

    // Pass 2: among rows on the latest date, keep the max hours
    const latest = {};
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const id  = String(row[C.machine]);
      if (!id) continue;
      if (row[C.stopped]) continue;
      if (row[C.current] === '' || row[C.current] === null) continue;

      const factory = (C.factory >= 0 && row[C.factory]) ? String(row[C.factory]) : '織一';
      const key = factory + ':' + id;

      const t = new Date(row[C.time]);
      const dateStr = Utilities.formatDate(t, 'Asia/Ho_Chi_Minh', 'yyyy-MM-dd');
      if (dateStr !== latestDate[key]) continue;

      if (!latest[key] || row[C.current] > latest[key].hours) {
        latest[key] = { hours: row[C.current] };
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
      sheet.appendRow(['時間戳記','廠別','班別','組別','機台號','上次時數','本次時數','稼動率%','停機']);
      sheet.setFrozenRows(1);
    }

    const ts      = Utilities.formatDate(new Date(), 'Asia/Ho_Chi_Minh', 'yyyy-MM-dd HH:mm:ss');
    const factory = payload.factory || '織一';
    const shift   = payload.shift;
    const group   = payload.group;
    const records = payload.records;

    const rows = records.map(r => [
      ts,
      factory,
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
