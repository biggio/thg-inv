/**
 * Google Apps Script 後端 API (Code.gs)
 * 負責讀取與寫入「貨架[填單]」A~G 欄位
 * 
 * 核心優化：
 * 針對有陣列公式 (ARRAYFORMULA 等) 的試算表，避免 sheet.getLastRow() 誤判空白行，
 * 精確以「A 欄 (日期) 是否有實質資料」為基準判定最後真實行與追加寫入位置。
 */

const SHEET_NAME = "貨架[填單]";

// 輔助函式：精確找出 A 欄最後一個有實質內容的行號 (Row Index)
function getActualLastRow(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 1; // 只有表頭或空表

  // 抓取 A 欄全部資料
  const aValues = sheet.getRange(1, 1, lastRow, 1).getValues();
  
  // 從最後一行往回找第一個 A 欄不是空的行
  for (let r = aValues.length - 1; r >= 0; r--) {
    const val = aValues[r][0];
    if (val !== "" && val !== null && val !== undefined) {
      return r + 1; // 轉為 1-indexed 行號
    }
  }
  return 1;
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

// 取得最近記錄 (以 A 欄有值為準)
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
    
    // 只取最後 20 筆真實存在的資料
    const startRow = Math.max(2, actualLastRow - 19);
    const numRows = actualLastRow - startRow + 1;
    const values = sheet.getRange(startRow, 1, numRows, 7).getValues();
    
    const records = [];
    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      let dateVal = row[0];
      if (dateVal === "" || dateVal === null || dateVal === undefined) continue;
      
      if (dateVal instanceof Date) {
        dateVal = Utilities.formatDate(dateVal, Session.getScriptTimeZone() || "GMT+8", "yyyy-MM-dd HH:mm");
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

// 依料號實時計算各儲位庫存分佈與總結存 (以 A 欄有值為準)
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
      // 確保該行不是被公式撐開的空行 (A 欄必須有日期)
      if (!row[0]) return;

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

// 寫入一筆庫存紀錄：精準寫在「A 欄最後一筆有資料」的下一行
function submitRecord(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000); // 併發鎖定保護
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.getRange(1, 1, 1, 7).setValues([["日期", "料號", "數量", "儲位", "進或出", "經辦人", "備註"]]);
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
    
    // 【核心修復】：不用 appendRow()，而是精確計算 A 欄最後有值的下一行
    const actualLastRow = getActualLastRow(sheet);
    const targetRow = actualLastRow + 1;
    
    // 將 A~G 寫入該行
    sheet.getRange(targetRow, 1, 1, 7).setValues([[
      formattedDate, // A: 日期
      sku,           // B: 料號
      qty,           // C: 數量
      location,      // D: 儲位
      type,          // E: 進或出
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
