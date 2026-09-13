/**
 * Google Apps Script 後端 API (Code.gs)
 * 負責讀取與寫入「貨架[填單]」A~G 欄位
 * 
 * 業務規則更新：
 * 1. A 欄日期格式：M/d/yyyy (例如: 9/11/2026)
 * 2. E 欄格式：進庫為 "1 In"，出庫為 "3 Out"
 * 3. 支援取得「Active_Item」分頁 B 欄的可用料號清單供前端 Auto-Complete (自動完成)
 * 4. 精確以 A 欄有無資料判定實際最後一行，徹底避免公式造成的空行跳格
 */

const SHEET_NAME = "貨架[填單]";
const ACTIVE_ITEM_SHEET_NAME = "Active_Item";

// 輔助函式：精確找出 A 欄最後一個有實質內容的行號
function getActualLastRow(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 1;

  const aValues = sheet.getRange(1, 1, lastRow, 1).getValues();
  for (let r = aValues.length - 1; r >= 0; r--) {
    const val = aValues[r][0];
    if (val !== "" && val !== null && val !== undefined) {
      return r + 1;
    }
  }
  return 1;
}

// 取得「Active_Item」分頁 B 欄的所有可用料號
function getActiveItems() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(ACTIVE_ITEM_SHEET_NAME);
    if (!sheet) {
      return { success: true, items: [] };
    }

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: true, items: [] };

    // 抓取 B 欄（從第 2 行到最後一行）
    const bValues = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
    const itemsSet = new Set();

    for (let i = 0; i < bValues.length; i++) {
      const val = String(bValues[i][0] || '').trim();
      if (val) {
        itemsSet.add(val);
      }
    }

    return {
      success: true,
      items: Array.from(itemsSet)
    };
  } catch (err) {
    return { success: false, error: err.toString(), items: [] };
  }
}

// 處理 GET 請求
function doGet(e) {
  const params = (e && e.parameter) ? e.parameter : {};
  let result = {};

  try {
    if (params.action === 'submit') {
      result = submitRecord(params);
    } else if (params.action === 'getRecords') {
      result = getRecentRecords();
    } else if (params.action === 'queryStock') {
      result = queryStock(params.sku);
    } else if (params.action === 'getActiveItems') {
      result = getActiveItems();
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
      sheet.getRange(1, 1, 1, 7).setValues([["日期", "料號", "數量", "儲位", "進或出", "經辦人", "備註"]]);
    }
    
    const actualLastRow = getActualLastRow(sheet);
    if (actualLastRow <= 1) {
      return { success: true, records: [] };
    }
    
    const startRow = Math.max(2, actualLastRow - 19);
    const numRows = actualLastRow - startRow + 1;
    const values = sheet.getRange(startRow, 1, numRows, 7).getValues();
    
    const records = [];
    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      let dateVal = row[0];
      if (dateVal === "" || dateVal === null || dateVal === undefined) continue;
      
      if (dateVal instanceof Date) {
        // 格式化為 M/d/yyyy
        dateVal = Utilities.formatDate(dateVal, Session.getScriptTimeZone() || "GMT+8", "M/d/yyyy");
      }
      records.push({
        id: startRow + i,
        date: String(dateVal),
        sku: String(row[1] || '').trim(),
        qty: Number(row[2]) || 0,
        location: String(row[3] || '').trim(),
        type: String(row[4] || '').trim(),
        operator: String(row[5] || '').trim(),
        note: String(row[6] || '').trim()
      });
    }
    
    return { success: true, records: records.reverse() };
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
    
    const actualLastRow = getActualLastRow(sheet);
    if (actualLastRow <= 1) return { success: true, stockMap: {}, total: 0 };
    
    const values = sheet.getRange(2, 1, actualLastRow - 1, 7).getValues();
    const stockMap = {};
    let total = 0;
    
    values.forEach(row => {
      if (!row[0]) return; // A 欄有值才算

      const rowSku = String(row[1] || '').trim();
      if (rowSku.toLowerCase() === sku.toLowerCase()) {
        const qty = Number(row[2]) || 0;
        const loc = String(row[3] || '未指定儲位').trim();
        const type = String(row[4] || '').trim();
        
        // 判定進出庫: "3 Out" 或 "出" 為扣帳，"1 In" 或 "進" 為入庫
        const delta = (type === '3 Out' || type === '出' || type === '出庫') ? -qty : qty;
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

// 寫入一筆庫存紀錄
function submitRecord(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.getRange(1, 1, 1, 7).setValues([["日期", "料號", "數量", "儲位", "進或出", "經辦人", "備註"]]);
    }
    
    // 規則 1：A 欄日期格式應為 9/11/2026 (即 M/d/yyyy)
    const now = new Date();
    const formattedDate = data.date ? data.date : Utilities.formatDate(now, Session.getScriptTimeZone() || "GMT+8", "M/d/yyyy");
    
    const sku = String(data.sku || '').trim();
    const qty = Number(data.qty) || 0;
    const location = String(data.location || '').trim();
    
    // 規則 2：E 欄進庫應為 "1 In"，出庫為 "3 Out"
    let rawType = String(data.type || '').trim();
    let typeFormatted = "1 In";
    if (rawType === '出' || rawType === '3 Out' || rawType === '出庫' || rawType === '3') {
      typeFormatted = "3 Out";
    } else {
      typeFormatted = "1 In";
    }
    
    const operator = String(data.operator || '').trim();
    const note = String(data.note || '').trim();
    
    if (!sku) throw new Error("料號不能為空");
    if (qty <= 0) throw new Error("數量必須大於 0");
    if (!location) throw new Error("儲位不能為空");
    
    // 以 A 欄最後有值的下一行定點寫入
    const actualLastRow = getActualLastRow(sheet);
    const targetRow = actualLastRow + 1;
    
    sheet.getRange(targetRow, 1, 1, 7).setValues([[
      formattedDate, // A: 日期 (M/d/yyyy)
      sku,           // B: 料號
      qty,           // C: 數量
      location,      // D: 儲位
      typeFormatted, // E: 進或出 ("1 In" 或 "3 Out")
      operator,      // F: 經辦人
      note           // G: 備註
    ]]);
    
    return {
      success: true,
      message: "寫入成功",
      targetRow: targetRow,
      record: {
        date: formattedDate,
        sku: sku,
        qty: qty,
        location: location,
        type: typeFormatted,
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
