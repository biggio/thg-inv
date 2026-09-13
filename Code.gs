/**
 * Google Apps Script 後端 API (Code.gs)
 * 負責讀取與寫入「貨架[填單]」A~G 欄位
 * 全面支援外部網頁以 CORS mode: no-cors 或 GET/JSONP 發送
 */

const SHEET_NAME = "貨架[填單]";

// 處理 GET 請求 (查詢存量、取得記錄、或以 GET 寫入避免跨域問題)
function doGet(e) {
  const params = (e && e.parameter) ? e.parameter : {};
  let result = {};

  try {
    if (params.action === 'submit') {
      // 支援前端以 GET 方式寫入（100% 避開瀏覽器 CORS 預檢限制）
      result = submitRecord(params);
    } else if (params.action === 'getRecords') {
      result = getRecentRecords();
    } else if (params.action === 'queryStock') {
      result = queryStock(params.sku);
    } else {
      result = { status: "online", message: "THG 庫存管理 API 運作中", timestamp: new Date() };
    }
  } catch (err) {
    result = { success: false, error: err.toString() };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// 處理 POST 請求
function doPost(e) {
  let result = {};
  try {
    let postData = {};
    if (e.postData && e.postData.contents) {
      try {
        postData = JSON.parse(e.postData.contents);
      } catch (jsonErr) {
        postData = e.parameter || {};
      }
    } else if (e.parameter) {
      postData = e.parameter;
    }
    result = submitRecord(postData);
  } catch (err) {
    result = { success: false, error: err.toString() };
  }

  return ContentService.createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

// 取得最近記錄 (最多 20 筆)
function getRecentRecords() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(["日期", "料號", "數量", "儲位", "進或出", "經辦人", "備註"]);
    }
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      return { success: true, records: [] };
    }
    
    const startRow = Math.max(2, lastRow - 19);
    const numRows = lastRow - startRow + 1;
    const values = sheet.getRange(startRow, 1, numRows, 7).getValues();
    
    const records = values.map((row, idx) => {
      let dateVal = row[0];
      if (dateVal instanceof Date) {
        dateVal = Utilities.formatDate(dateVal, Session.getScriptTimeZone() || "GMT+8", "yyyy-MM-dd HH:mm");
      }
      return {
        id: startRow + idx,
        date: String(dateVal || ''),
        sku: String(row[1] || '').trim(),
        qty: Number(row[2]) || 0,
        location: String(row[3] || '').trim(),
        type: String(row[4] || '').trim(),
        operator: String(row[5] || '').trim(),
        note: String(row[6] || '').trim()
      };
    }).reverse();
    
    return { success: true, records: records };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// 依料號實時計算各儲位庫存分佈與總結存
function queryStock(sku) {
  try {
    if (!sku) return { success: true, stockMap: {}, total: 0 };
    sku = String(sku).trim();
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return { success: true, stockMap: {}, total: 0 };
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: true, stockMap: {}, total: 0 };
    
    const values = sheet.getRange(2, 1, lastRow - 1, 7).getValues();
    const stockMap = {};
    let total = 0;
    
    values.forEach(row => {
      const rowSku = String(row[1] || '').trim();
      if (rowSku.toLowerCase() === sku.toLowerCase()) {
        const qty = Number(row[2]) || 0;
        const loc = String(row[3] || '未指定儲位').trim();
        const type = String(row[4] || '').trim();
        
        const delta = (type === '出' || type === '出庫') ? -qty : qty;
        stockMap[loc] = (stockMap[loc] || 0) + delta;
        total += delta;
      }
    });
    
    return {
      success: true,
      sku: sku,
      stockMap: stockMap,
      total: total
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  }
}

// 寫入一筆庫存紀錄 (A~G 欄，含併發鎖定)
function submitRecord(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // 最多等 10 秒
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(["日期", "料號", "數量", "儲位", "進或出", "經辦人", "備註"]);
    }
    
    const now = new Date();
    const formattedDate = data.date ? data.date : Utilities.formatDate(now, Session.getScriptTimeZone() || "GMT+8", "yyyy-MM-dd HH:mm:ss");
    
    const sku = String(data.sku || '').trim();
    const qty = Number(data.qty) || 0;
    const location = String(data.location || '').trim();
    const type = String(data.type || '進').trim();
    const operator = String(data.operator || '').trim();
    const note = String(data.note || '').trim();
    
    if (!sku) throw new Error("料號不能為空");
    if (qty <= 0) throw new Error("數量必須大於 0");
    if (!location) throw new Error("儲位不能為空");
    
    sheet.appendRow([
      formattedDate, // A: 日期
      sku,           // B: 料號
      qty,           // C: 數量
      location,      // D: 儲位
      type,          // E: 進或出
      operator,      // F: 經辦人
      note           // G: 備註
    ]);
    
    return {
      success: true,
      message: "寫入成功",
      record: {
        date: formattedDate,
        sku: sku,
        qty: qty,
        location: location,
        type: type,
        operator: operator,
        note: note
      }
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  } finally {
    lock.releaseLock();
  }
}
