/**
 * ROTAX Air Log - Google Apps Script (GAS) クラウド同期バックエンド
 * 
 * 【機能】
 * - doGet(e) : スプレッドシートの全レコードをJSON形式で返却
 * - doPost(e): アプリ側から送信されたレコードをIDベースでUpsert（新規追加・更新・削除）し、
 *              マージ後の最新全レコードを返却（双方向同期対応）
 */

const SHEET_NAME = 'Logs';
const HEADERS = [
  'id',
  'date',
  'time',
  'circuit',
  'vehicle',
  'tireSet',
  'driver',
  'condition',
  'ambientTemp',
  'trackTemp',
  'laps',
  'coldFL',
  'coldFR',
  'coldRL',
  'coldRR',
  'warmFL',
  'warmFR',
  'warmRL',
  'warmRR',
  'deltaFL',
  'deltaFR',
  'deltaRL',
  'deltaRR',
  'memo',
  'updatedAt'
];

function getOrCreateSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#1c2230')
      .setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function readAllRecords_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(sheet.getLastColumn(), HEADERS.length);
  if (lastRow <= 1) return [];

  const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const rawHeaders = values[0].map(String);
  const records = [];

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    if (!row[0] && !row[1] && !row[3]) continue;
    const obj = {};
    for (let c = 0; c < rawHeaders.length; c++) {
      const key = rawHeaders[c];
      if (!key) continue;
      let val = row[c];
      if (val instanceof Date) {
        if (key === 'date') {
          val = Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        } else if (key === 'time') {
          val = Utilities.formatDate(val, Session.getScriptTimeZone(), 'HH:mm');
        } else {
          val = val.toISOString();
        }
      }
      obj[key] = val !== undefined && val !== null ? String(val) : '';
    }
    if (!obj.id) {
      obj.id = 'row-' + r + '-' + new Date().getTime();
    }
    records.push(obj);
  }
  return records;
}

function writeAllRecords_(sheet, records) {
  records.sort(function(a, b) {
    const dtA = (a.date || '') + 'T' + (a.time || '00:00');
    const dtB = (b.date || '') + 'T' + (b.time || '00:00');
    return dtB.localeCompare(dtA);
  });

  const lastRow = sheet.getLastRow();
  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, Math.max(sheet.getLastColumn(), HEADERS.length)).clearContent();
  }

  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);

  if (records.length === 0) return;

  const rows = records.map(function(rec) {
    return HEADERS.map(function(h) {
      return rec[h] !== undefined && rec[h] !== null ? rec[h] : '';
    });
  });

  sheet.getRange(2, 1, rows.length, HEADERS.length).setValues(rows);
}

function doGet(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(10000);
  try {
    const sheet = getOrCreateSheet_();
    const records = readAllRecords_(sheet);
    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'success',
        count: records.length,
        records: records,
        syncedAt: new Date().toISOString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'error',
        message: err.toString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.tryLock(15000);
  try {
    const sheet = getOrCreateSheet_();
    const rawBody = e.postData && e.postData.contents ? e.postData.contents : '{}';
    const payload = JSON.parse(rawBody);
    const action = payload.action || 'sync';

    const existingRecords = readAllRecords_(sheet);
    const recordMap = {};
    existingRecords.forEach(function(rec) {
      recordMap[String(rec.id)] = rec;
    });

    if (action === 'sync' && Array.isArray(payload.records)) {
      const deletedIds = Array.isArray(payload.deletedIds) ? payload.deletedIds.map(String) : [];
      deletedIds.forEach(function(delId) {
        delete recordMap[delId];
      });

      payload.records.forEach(function(incoming) {
        if (!incoming || !incoming.id) return;
        const id = String(incoming.id);
        if (deletedIds.indexOf(id) !== -1) return;

        const current = recordMap[id];
        if (!current) {
          recordMap[id] = incoming;
        } else {
          const inTime = incoming.updatedAt ? new Date(incoming.updatedAt).getTime() : 0;
          const curTime = current.updatedAt ? new Date(current.updatedAt).getTime() : 0;
          if (inTime >= curTime) {
            recordMap[id] = incoming;
          }
        }
      });
    } else if (action === 'overwrite' && Array.isArray(payload.records)) {
      const freshMap = {};
      payload.records.forEach(function(incoming) {
        if (incoming && incoming.id) {
          freshMap[String(incoming.id)] = incoming;
        }
      });
      writeAllRecords_(sheet, Object.keys(freshMap).map(function(k) { return freshMap[k]; }));
      const finalOverwrite = readAllRecords_(sheet);
      return ContentService
        .createTextOutput(JSON.stringify({
          status: 'success',
          action: 'overwrite',
          count: finalOverwrite.length,
          records: finalOverwrite,
          syncedAt: new Date().toISOString()
        }))
        .setMimeType(ContentService.MimeType.JSON);
    } else if (action === 'delete' && payload.id) {
      delete recordMap[String(payload.id)];
    }

    const mergedList = Object.keys(recordMap).map(function(k) {
      return recordMap[k];
    });
    writeAllRecords_(sheet, mergedList);

    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'success',
        action: action,
        count: mergedList.length,
        records: mergedList,
        syncedAt: new Date().toISOString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({
        status: 'error',
        message: err.toString()
      }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}
