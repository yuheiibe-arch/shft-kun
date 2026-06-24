/**
 * ====================================================================
 * 04_SheetUtils.gs
 * スプレッドシートからのデータ取得および書き込み処理（★白紙化防止・安全上書き版）
 * ====================================================================
 */

function writeWithProtectionObj(sheetName, newRowObjs, defaultHeaders, checkColName, isTargetFunc, getPrimaryKeyFunc, multiCheckCols = null) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.getRange(1, 1, 1, defaultHeaders.length).setValues([defaultHeaders]);
  }
  
  const lastCol = sheet.getLastColumn();
  const lastRow = sheet.getLastRow();
  
  let actualHeaders = [];
  if (lastCol > 0) {
    actualHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());
  } else {
    actualHeaders = [...defaultHeaders];
    sheet.getRange(1, 1, 1, actualHeaders.length).setValues([actualHeaders]);
  }
  
  let existingDataObjs = [];
  if (lastRow > 1) {
    const rawValues = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
    existingDataObjs = rawValues.map(rowArr => {
      let obj = {};
      actualHeaders.forEach((h, i) => { obj[h] = rowArr[i]; });
      return obj;
    });
  }

  let finalDataObjs = [];
  const protectedKeys = new Set();
  const existingMap = new Map(); 

  existingDataObjs.forEach(rowObj => {
    if (!rowObj[actualHeaders[0]] && !rowObj[actualHeaders[1]]) return;
    const isTarget = isTargetFunc(rowObj);
    const isChecked = (rowObj[checkColName] === true || String(rowObj[checkColName]).toUpperCase() === "TRUE");
    const key = getPrimaryKeyFunc(rowObj);
    
    if (!existingMap.has(key)) existingMap.set(key, rowObj);

    if (!isTarget) {
      finalDataObjs.push(rowObj);
    } else if (isChecked) {
      finalDataObjs.push(rowObj);
      protectedKeys.add(key);
    }
  });

  newRowObjs.forEach(rowObj => {
    const key = getPrimaryKeyFunc(rowObj);
    if (!protectedKeys.has(key)) {
      if (existingMap.has(key)) {
        const oldRowObj = existingMap.get(key);
        let checkColIdx = actualHeaders.indexOf(checkColName);
        if (checkColIdx !== -1) {
           for (let i = checkColIdx + 1; i < actualHeaders.length; i++) {
             let h = actualHeaders[i];
             if (oldRowObj[h] !== undefined && oldRowObj[h] !== "") {
                rowObj[h] = oldRowObj[h];
             }
           }
        }
      }
      finalDataObjs.push(rowObj);
    }
  });

  const areaOrder = {"東京": 1, "神奈川": 2, "千葉": 3, "埼玉": 4, "関西": 5, "その他": 6};
  const getAreaW = (area) => areaOrder[area] || 99;

  const dowOrder = {"月": 1, "火": 2, "水": 3, "木": 4, "金": 5, "土": 6, "日": 7};
  const getDowW = (str) => {
    for (let k in dowOrder) {
      if (String(str).includes(k)) return dowOrder[k];
    }
    return 99;
  };

  finalDataObjs.sort((a, b) => {
    if (sheetName === "定期募集") {
      const timeA = String(a["開始時間"] || "").padStart(5, '0');
      const timeB = String(b["開始時間"] || "").padStart(5, '0');
      return getAreaW(a["エリア"]) - getAreaW(b["エリア"]) || 
             String(a["シフトタイトル"]).localeCompare(String(b["シフトタイトル"]), 'ja') || 
             getDowW(a["繰り返し曜日"]) - getDowW(b["繰り返し曜日"]) || 
             timeA.localeCompare(timeB);
    }
    if (sheetName === "単独募集") {
      const timeA = String(a["開始時間"] || "").padStart(5, '0');
      const timeB = String(b["開始時間"] || "").padStart(5, '0');
      return getAreaW(a["エリア"]) - getAreaW(b["エリア"]) || 
             String(a["拠点名"]).localeCompare(String(b["拠点名"]), 'ja') || 
             String(a["該当日"]).localeCompare(String(b["該当日"]), 'ja') ||
             timeA.localeCompare(timeB);
    }
    
    if (sheetName === "確定シフト作成") {
      const getTypeW = (t) => {
        const s = String(t);
        if (s.includes("定期非常勤")) return 2;
        if (s.includes("常勤")) return 1;
        if (s.includes("スポット")) return 3;
        return 99;
      };
      
      return getAreaW(a["エリア"]) - getAreaW(b["エリア"]) || 
             (getTypeW(a["種別"]) - getTypeW(b["種別"])) ||
             String(a["医師名"]).localeCompare(String(b["医師名"]), 'ja') ||
             String(a["拠点名"]).localeCompare(String(b["拠点名"]), 'ja') ||
             String(a["シフトタイトル"]).localeCompare(String(b["シフトタイトル"]), 'ja');
    }
    if (sheetName === "欠勤・シフトキャンセル作成") {
      return String(a["医師名"]).localeCompare(String(b["医師名"]), 'ja') || 
             String(a["該当日"]).localeCompare(String(b["該当日"]), 'ja') ||
             String(a["対象拠点"]).localeCompare(String(b["対象拠点"]), 'ja');
    }
    return 0;
  });

  // ==============================================================================
  // ★重要変更箇所：事前の clearContent を廃止しました。
  // ==============================================================================
  
  if (finalDataObjs.length === 0) {
    // 新しいデータが空の場合のみ、既存のデータをクリアして終了
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, sheet.getMaxColumns()).clearContent().clearDataValidations();
    }
    const maxR = sheet.getMaxRows();
    if (maxR > 2) sheet.deleteRows(3, maxR - 2);
    return;
  }

  const numRows = finalDataObjs.length;
  const numCols = Math.max(actualHeaders.length, Object.keys(finalDataObjs[0]).length);

  if (sheet.getMaxRows() < 1 + numRows) sheet.insertRowsAfter(sheet.getMaxRows(), (1 + numRows) - sheet.getMaxRows());
  if (sheet.getMaxColumns() < numCols) sheet.insertColumnsAfter(sheet.getMaxColumns(), numCols - sheet.getMaxColumns());

  // システムが知らない列にはnullをセット
  const writeData = finalDataObjs.map(obj => {
    return actualHeaders.map(h => {
      if (obj[h] !== undefined) return obj[h]; 
      if (!defaultHeaders.includes(h)) return null; 
      return ""; 
    });
  });

  // ★安全な上書き実行（ここでタイムアウトしても、古いデータが残ったままなので白紙になりません）
  const targetRange = sheet.getRange(2, 1, numRows, actualHeaders.length);
  targetRange.setValues(writeData).setHorizontalAlignment("left").setVerticalAlignment("top");

  let checkColNames = multiCheckCols || [checkColName];
  checkColNames.forEach(cName => {
    let idx = actualHeaders.indexOf(cName);
    if (idx !== -1) {
      const cbRange = sheet.getRange(2, idx + 1, numRows, 1); 
      cbRange.setDataValidation(SpreadsheetApp.newDataValidation().requireCheckbox().build());
      cbRange.setHorizontalAlignment("center").setVerticalAlignment("top");
    }
  });

  // ★書き込み成功後、もし古いデータの方が長くて下に残骸がはみ出している場合、その残骸だけを消去する
  if (lastRow - 1 > numRows) {
    const excessRows = (lastRow - 1) - numRows;
    sheet.getRange(2 + numRows, 1, excessRows, sheet.getMaxColumns()).clearContent().clearDataValidations();
  }

  // シート下部の余分な空白行を削除して整える
  const finalMaxRows = sheet.getMaxRows();
  const neededRows = Math.max(2, 1 + numRows);
  if (finalMaxRows > neededRows) {
    sheet.deleteRows(neededRows + 1, finalMaxRows - neededRows);
  }
}

function _readFlatSheet(ss, sheetName, isAbsence) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return [];

  let headers = data[0];
  let daishinIdx = headers.findIndex(h => String(h).includes("代診"));
  let reqIdx = headers.findIndex(h => String(h).includes("依頼者"));

  let results = [];
  for (let i = 1; i < data.length; i++) {
    let dVal = data[i][0];
    if (!dVal) continue;
    let d = new Date(dVal);
    if (isNaN(d.getTime())) continue;
    
    let sH = parseInt(String(data[i][2] instanceof Date ? Utilities.formatDate(data[i][2], "JST", "HH:mm") : data[i][2]).split(":")[0], 10);
    let eH = parseInt(String(data[i][3] instanceof Date ? Utilities.formatDate(data[i][3], "JST", "HH:mm") : data[i][3]).split(":")[0], 10);
    let loc = String(data[i][4]).trim();
    let reason = String(data[i][5]).trim();
    
    let docRaw = isAbsence ? (String(data[i][6]).trim() || String(data[i][5]).trim()) : String(data[i][5]).trim();
    let doc = docRaw.replace(/先生$/, "").trim();
    
    let isDaishin = daishinIdx !== -1 ? data[i][daishinIdx] : false;
    let requester = reqIdx !== -1 ? String(data[i][reqIdx]).trim() : "";
    
    results.push({ 
      dateObj: d, sH: sH, eH: eH, loc: loc, reason: reason, docName: doc,
      isDaishin: isDaishin, requester: requester
    });
  }
  return results;
}

function _readKyukanFlat() {
  const map = [];
  try {
    // ★ここを safeOpenByUrl に変更
    const kyuSs = safeOpenByUrl("https://docs.google.com/spreadsheets/d/1cbeXWojsxNMhQUo1c6VflF5hLUJUyfuOXCFbGP5jJEA/edit");
    const kyuSheet = kyuSs.getSheetByName("休館日");
    if (kyuSheet) {
      const data = kyuSheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        let d = new Date(data[i][0]);
        let loc = String(data[i][3]).trim();
        if (!isNaN(d.getTime()) && loc) {
          map.push({ dateObj: d, loc: loc });
        }
      }
    }
  } catch(e) {}
  return map;
}