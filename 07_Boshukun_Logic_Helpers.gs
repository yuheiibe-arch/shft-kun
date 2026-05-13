/**
 * ====================================================================
 * 07_Boshukun_Logic_Helpers.gs
 * 祝日・他シートからのデータマップ・開院日マップ構築
 * ====================================================================
 */

let _holidayMap = null;

function _debug_initHolidayMap(yearStr) {
  if (_holidayMap) return;
  _holidayMap = {};
  try {
    const url = 'https://docs.google.com/spreadsheets/d/1WlmirSDOPnIcV2cwY5ClXkMWBFMM-4zrAw4XFDNUPGw/edit';
    const ss = SpreadsheetApp.openByUrl(url);
    let sheet = ss.getSheetByName(yearStr) || ss.getSheetByName(`${yearStr}年度`);
    if (!sheet) return;

    const values = sheet.getDataRange().getValues();
    for (let i = 2; i < values.length; i++) {
      const row = values[i];
      const dateVal = row[0];
      if (!dateVal) continue;
      
      let dStr = "";
      if (dateVal instanceof Date) {
        dStr = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), "yyyy/MM/dd");
      } else {
        const parsed = new Date(dateVal);
        if (!isNaN(parsed.getTime())) dStr = Utilities.formatDate(parsed, Session.getScriptTimeZone(), "yyyy/MM/dd");
        else dStr = String(dateVal).trim().replace(/-/g, "/");
      }
      
      if (dStr) {
        _holidayMap[dStr] = {
          isWeekday: String(row[2] || "").trim() !== "",
          isHoliday: String(row[3] || "").trim() !== "",
          isNewYear: String(row[5] || "").trim() !== ""
        };
      }
    }
  } catch (e) {}
}

function _debug_isHoliday(dateStr, dow) {
  if (_holidayMap && _holidayMap[dateStr]) {
    const hData = _holidayMap[dateStr];
    if (hData.isHoliday || hData.isNewYear) return true;
    if (hData.isWeekday) return false;
  }
  return (dow === "土" || dow === "日");
}

function _debug_isTrueHoliday(dateStr) {
  if (_holidayMap && _holidayMap[dateStr]) {
    const hData = _holidayMap[dateStr];
    if (hData.isHoliday || hData.isNewYear) return true;
  }
  return false;
}

function _debug_buildAdvanceMap(locNames, term, yearStr) {
  const map = { byDow: {}, byDate: {} }; 
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const advanceSheet = ss.getSheetByName('先行応募');
  if (!advanceSheet) return map;
  
  const year = parseInt(yearStr, 10);
  const startDate = term === "上期" ? new Date(year, 3, 1) : new Date(year, 9, 1);
  const endDate = term === "上期" ? new Date(year, 8, 30) : new Date(year + 1, 2, 31);
  const dowNames = ["日", "月", "火", "水", "木", "金", "土"];
  
  const data = advanceSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const dateVal = row[0];
    if (!dateVal) continue;
    
    const d = new Date(dateVal);
    if (d < startDate || d > endDate) continue;
    
    const locFull = String(row[4]).trim();
    if (!locNames.some(l => locFull.includes(l))) continue;
    
    const dStr = Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy/MM/dd");
    const dow = dowNames[d.getDay()];
    const loc = locFull.split('_')[0].replace(/クリニック|診療所/g, "").trim();
    
    const startStr = String(row[2] instanceof Date ? Utilities.formatDate(row[2], Session.getScriptTimeZone(), "HH:mm") : row[2]);
    const endStr = String(row[3] instanceof Date ? Utilities.formatDate(row[3], Session.getScriptTimeZone(), "HH:mm") : row[3]);
    
    const sH = parseInt(startStr.split(":")[0], 10);
    const eH = parseInt(endStr.split(":")[0], 10);
    if (isNaN(sH) || isNaN(eH)) continue;
    
    const keyDow = `${loc}_${dow}`;
    if (!map.byDow[keyDow]) map.byDow[keyDow] = [];
    map.byDow[keyDow].push({ startH: sH, endH: eH });
    
    const keyDate = `${dStr}_${loc}`;
    if (!map.byDate[keyDate]) map.byDate[keyDate] = [];
    map.byDate[keyDate].push({ startH: sH, endH: eH });
  }
  return map;
}

function _debug_buildAbsencePatchMap() {
  const map = {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('お休み情報');
  if (!sheet) return map;
  
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const dateVal = row[0];
    if (!dateVal) continue;
    
    const d = new Date(dateVal);
    if (isNaN(d.getTime())) continue;
    
    const dStr = Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy/MM/dd");
    const loc = String(row[4]).split('_')[0].replace(/クリニック|診療所/g, "").trim();
    
    let rawDocName = String(row[6]).trim() || String(row[5]).trim();
    const docName = rawDocName.replace(/先生$/, "").trim();
    
    const startStr = String(row[2] instanceof Date ? Utilities.formatDate(row[2], Session.getScriptTimeZone(), "HH:mm") : row[2]);
    const endStr = String(row[3] instanceof Date ? Utilities.formatDate(row[3], Session.getScriptTimeZone(), "HH:mm") : row[3]);
    
    const sH = parseInt(startStr.split(":")[0], 10);
    const eH = parseInt(endStr.split(":")[0], 10);
    if (isNaN(sH) || isNaN(eH)) continue;
    
    const key = `${dStr}_${loc}`;
    if (!map[key]) map[key] = [];
    map[key].push({ doc: docName, startH: sH, endH: eH });
  }
  return map;
}

// ★ 新規追加: 各拠点の開院日を取得する関数
function _buildOpenDateMap(locNames) {
  const map = {};
  try {
    const masterUrl = 'https://docs.google.com/spreadsheets/d/14RbsDcv0nXfEwweki8-9cK3lQUg1XUuhozLNF9u2qAs/edit';
    const ss = SpreadsheetApp.openByUrl(masterUrl);
    const sheet = ss.getSheetByName('拠点名');
    const data = sheet.getDataRange().getValues();
    for(let i = 1; i < data.length; i++) {
      const name = String(data[i][0]).trim();
      const openDateVal = data[i][7]; // H列
      if (locNames.includes(name) && openDateVal) {
        let openDate = openDateVal instanceof Date ? openDateVal : new Date(openDateVal);
        if (!isNaN(openDate.getTime())) {
          map[name] = openDate;
        }
      }
    }
  } catch(e) {}
  return map;
}