/**
 * Google Apps Script 後端 API (Code.gs)
 * 負責讀取、寫入、修改、刪除「貨架[填單]」A~G 欄位
 */

const SHEET_NAME = "貨架[填單]";
const ACTIVE_ITEM_SHEET_NAME = "Active_Item";

// 高性能查找 A 欄最後有內容的行號
function getActualLastRowFast(sheet) {
  try {
    const finder = sheet.getRange("A:A").createTextFinder(".+").useRegularExpression(true);
    const results = finder.findAll();
    if (results && results.length > 0) {
      return results[results.length - 1].getRow();
    }
  } catch (e) {}
  return Math.max(1, sheet.getLastRow());
}

// 取得試算表中所有現存的分頁名稱（診斷用）
function getAllSheetNames() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheets().map(s => s.getName());
}

// 取得「Active_Item」分頁 B 欄可用料號
function getActiveItems(forceRefresh) {
  const cache = CacheService.getScriptCache();
  const cacheKey = "thg_active_items_list";

  if (!forceRefresh) {
    const cachedData = cache.get(cacheKey);
    if (cachedData) {
      try {
        return JSON.parse(cachedData);
      } catch (e) {}
    }
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(ACTIVE_ITEM_SHEET_NAME);
    if (!sheet) {
      const sheets = ss.getSheets();
      for (let s of sheets) {
        if (s.getName().trim().toLowerCase() === ACTIVE_ITEM_SHEET_NAME.toLowerCase()) {
          sheet = s;
          break;
        }
      }
    }
    if (!sheet) return { success: true, items: [] };

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: true, items: [] };

    const bValues = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
    const items = [];
    for (let i = 0; i < bValues.length; i++) {
      const val = String(bValues[i][0] || '').trim();
      if (val) items.push(val);
    }

    const uniqueItems = Array.from(new Set(items));
    const response = {
      success: true,
      items: uniqueItems,
      updatedAt: new Date().getTime()
    };

    try {
      cache.put(cacheKey, JSON.stringify(response), 21600);
    } catch (cacheErr) {}

    return response;
  } catch (err) {
    return { success: false, error: err.toString(), items: [] };
  }
}

// 取得 Who 人員清單（超寬容自動偵測：List、Lists、清單、人員等）
function getWhoList(forceRefresh) {
  const cache = CacheService.getScriptCache();
  const cacheKey = "thg_who_list";

  if (!forceRefresh) {
    const cachedData = cache.get(cacheKey);
    if (cachedData) {
      try {
        return JSON.parse(cachedData);
      } catch (e) {}
    }
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const allSheets = ss.getSheets();
    const sheetNames = allSheets.map(s => s.getName());

    // 依序嘗試多種可能的分頁名稱
    let targetSheet = null;
    const candidates = ['list', 'lists', '清單', '人員', '經辦人', '名單', 'who'];
    
    // 1. 完全相同或大小寫不拘
    for (let s of allSheets) {
      const sName = s.getName().trim().toLowerCase();
      if (sName === 'list' || candidates.includes(sName)) {
        targetSheet = s;
        break;
      }
    }

    // 2. 包含 'list' 字眼
    if (!targetSheet) {
      for (let s of allSheets) {
        if (s.getName().toLowerCase().includes('list')) {
          targetSheet = s;
          break;
        }
      }
    }

    if (!targetSheet) {
      return {
        success: false,
        error: "找不到 List 分頁。目前試算表所有分頁為：[" + sheetNames.join(', ') + "]"
      };
    }

    const lastRow = targetSheet.getLastRow();
    if (lastRow <= 1) {
      return { success: true, list: [], sheetFound: targetSheet.getName() };
    }

    // 讀取 C 欄 (第 3 欄)
    const cValues = targetSheet.getRange(2, 3, lastRow - 1, 1).getValues();
    const whoList = [];

    for (let i = 0; i < cValues.length; i++) {
      const val = String(cValues[i][0] || '').trim();
      if (val) {
        whoList.push(val);
      }
    }

    const uniqueWho = Array.from(new Set(whoList));
    const response = {
      success: true,
      list: uniqueWho,
      sheetFound: targetSheet.getName(),
      count: uniqueWho.length,
      updatedAt: new Date().getTime()
    };

    try {
      cache.put(cacheKey, JSON.stringify(response), 21600);
    } catch (cacheErr) {}

    return response;
  } catch (err) {
    return { success: false, error: err.toString(), list: [] };
  }
}

// 處理 GET 請求
function doGet(e) {
  const params = (e && e.parameter) ? e.parameter : {};
  let result = {};

  try {
    if (params.action === 'submit') {
      result = submitRecord(params);
    } else if (params.action === 'update') {
      result = updateRecord(params);
    } else if (params.action === 'delete') {
      result = deleteRecord(params.rowIndex);
    } else if (params.action === 'getRecords') {
      result = getRecentRecords();
    } else if (params.action === 'queryStock') {
      result = queryStock(params.sku);
    } else if (params.action === 'getActiveItems') {
      result = getActiveItems(params.force === '1');
    } else if (params.action === 'getWhoList') {
      result = getWhoList(params.force === '1');
    } else if (params.action === 'getSheets') {
      result = { success: true, sheets: getAllSheetNames() };
    } else {
      result = { status: "online", message: "THG 庫存管理 API 運作中" };
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

    if (postData.action === 'update') {
      result = updateRecord(postData);
    } else if (postData.action === 'delete') {
      result = deleteRecord(postData.rowIndex);
    } else {
      result = submitRecord(postData);
    }
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
    if (!sheet) return { success: true, records: [] };
    
    const actualLastRow = getActualLastRowFast(sheet);
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
      if (!dateVal) continue;
      
      if (dateVal instanceof Date) {
        dateVal = Utilities.formatDate(dateVal, Session.getScriptTimeZone() || "GMT+8", "M/d/yyyy");
      }
      records.push({
        rowIndex: startRow + i,
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

// 修改特定行號的記錄
function updateRecord(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    const rowIndex = parseInt(data.rowIndex);
    if (!rowIndex || rowIndex < 2) throw new Error("無效的資料行號");

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error("找不到工作表");

    const sku = String(data.sku || '').trim();
    const qty = Number(data.qty) || 0;
    const location = String(data.location || '').trim();
    
    let rawType = String(data.type || '').trim();
    let typeFormatted = (rawType === '出' || rawType === '3 Out' || rawType === '出庫' || rawType === '3') ? "3 Out" : "1 In";
    
    const operator = String(data.operator || '').trim();
    const note = String(data.note || '').trim();
    
    let formattedDate = data.date;
    if (!formattedDate) {
      const existingDate = sheet.getRange(rowIndex, 1).getValue();
      if (existingDate instanceof Date) {
        formattedDate = Utilities.formatDate(existingDate, Session.getScriptTimeZone() || "GMT+8", "M/d/yyyy");
      } else {
        formattedDate = String(existingDate || '');
      }
    }

    sheet.getRange(rowIndex, 1, 1, 7).setValues([[
      formattedDate,
      sku,
      qty,
      location,
      typeFormatted,
      operator,
      note
    ]]);

    return { success: true, message: "修改成功", rowIndex: rowIndex };
  } catch (err) {
    return { success: false, error: err.toString() };
  } finally {
    lock.releaseLock();
  }
}

// 刪除特定行號的記錄
function deleteRecord(rowIndex) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(5000);
    rowIndex = parseInt(rowIndex);
    if (!rowIndex || rowIndex < 2) throw new Error("無效的資料行號");

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) throw new Error("找不到工作表");

    sheet.deleteRow(rowIndex);

    return { success: true, message: `已成功刪除第 ${rowIndex} 行記錄` };
  } catch (err) {
    return { success: false, error: err.toString() };
  } finally {
    lock.releaseLock();
  }
}

// 依料號實時計算存量
function queryStock(sku) {
  try {
    if (!sku) return { success: true, stockMap: {}, total: 0 };
    sku = String(sku).trim();
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) return { success: true, stockMap: {}, total: 0 };
    
    const textFinder = sheet.getRange("B:B").createTextFinder(sku).matchEntireCell(true);
    const foundCells = textFinder.findAll();
    
    const stockMap = {};
    let total = 0;
    
    if (foundCells && foundCells.length > 0) {
      foundCells.forEach(cell => {
        const row = cell.getRow();
        if (row === 1) return;
        const rowData = sheet.getRange(row, 1, 1, 5).getValues()[0];
        if (!rowData[0]) return;
        
        const qty = Number(rowData[2]) || 0;
        const loc = String(rowData[3] || '未指定儲位').trim();
        const type = String(rowData[4] || '').trim();
        
        const delta = (type === '3 Out' || type === '出' || type === '出庫') ? -qty : qty;
        stockMap[loc] = (stockMap[loc] || 0) + delta;
        total += delta;
      });
    }
    
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
    lock.waitLock(5000);
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.getRange(1, 1, 1, 7).setValues([["日期", "料號", "數量", "儲位", "進或出", "Who", "備註"]]);
    }
    
    const now = new Date();
    const formattedDate = data.date ? data.date : Utilities.formatDate(now, Session.getScriptTimeZone() || "GMT+8", "M/d/yyyy");
    
    const sku = String(data.sku || '').trim();
    const qty = Number(data.qty) || 0;
    const location = String(data.location || '').trim();
    
    let rawType = String(data.type || '').trim();
    let typeFormatted = (rawType === '出' || rawType === '3 Out' || rawType === '出庫' || rawType === '3') ? "3 Out" : "1 In";
    
    const operator = String(data.operator || '').trim();
    const note = String(data.note || '').trim();
    
    if (!sku) throw new Error("料號不能為空");
    if (qty <= 0) throw new Error("數量必須大於 0");
    if (!location) throw new Error("儲位不能為空");
    
    const actualLastRow = getActualLastRowFast(sheet);
    const targetRow = actualLastRow + 1;
    
    sheet.getRange(targetRow, 1, 1, 7).setValues([[
      formattedDate,
      sku,
      qty,
      location,
      typeFormatted,
      operator,
      note
    ]]);
    
    return {
      success: true,
      message: "寫入成功",
      targetRow: targetRow
    };
  } catch (err) {
    return { success: false, error: err.toString() };
  } finally {
    lock.releaseLock();
  }
}
