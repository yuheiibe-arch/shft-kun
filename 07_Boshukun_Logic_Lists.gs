/**
 * ====================================================================
 * 07_Boshukun_Logic_Lists.gs
 * リスト成形・二重募集完全防止＆単独圧縮対応版
 * ====================================================================
 */

function _getAreaHelper(locName, map) {
  if (!locName) return "その他";
  if (locName === "欠勤" || locName === "有給" || locName === "休館日") return locName; 
  if (map[locName]) return map[locName];
  
  const cleanLoc = locName.replace(/[（\(\)）]/g, "").replace(/内科|小児科/g, "").trim();
  for (const key in map) {
    const cleanKey = key.replace(/[（\(\)）]/g, "").replace(/内科|小児科/g, "").trim();
    if (cleanLoc === cleanKey) return map[key];
    if (cleanLoc.includes(cleanKey) || cleanKey.includes(cleanLoc)) return map[key];
  }
  if (locName.includes("北葛西") || locName.includes("亀有") || locName.includes("西葛西") || locName.includes("代官山") || locName.includes("武蔵小山")) return "東京";
  return "その他";
}

function _debug_getKyukanMap() {
  const map = {};
  try {
    // ★ここを safeOpenByUrl に変更
    const kyuSs = safeOpenByUrl("https://docs.google.com/spreadsheets/d/1cbeXWojsxNMhQUo1c6VflF5hLUJUyfuOXCFbGP5jJEA/edit");
    const kyuSheet = kyuSs.getSheetByName("休館日");
    if (kyuSheet) {
      const kyuData = kyuSheet.getDataRange().getValues();
      for (let i = 1; i < kyuData.length; i++) {
        let dVal = kyuData[i][0];
        let kLoc = String(kyuData[i][3]).trim();
        if (!dVal || !kLoc) continue;
        
        let d = new Date(dVal);
        if (!isNaN(d.getTime())) {
          let dStr = Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy/MM/dd");
          const cleanLoc = kLoc.replace(/[（\(\)）]/g, "").replace(/内科|小児科/g, "").trim();
          map[`${dStr}_${cleanLoc}`] = true;
        }
      }
    }
  } catch(e) {}
  return map;
}

function _debug_generateRecruitmentLists(rawBoshuData, term, yearStr, areaMap) {
  const baseGroups = {};
  const allSegments = [];

  rawBoshuData.forEach(row => {
    if (row.groupType === 'BASE') {
      if (row.isTrueHoliday) {
        row.groupType = 'HOLIDAY';
        allSegments.push({ seg: [row], groupType: 'HOLIDAY' });
      } else {
        const key = `${row.loc}_${row.category}_${row.start}-${row.end}_${row.dow}`;
        if (!baseGroups[key]) baseGroups[key] = [];
        baseGroups[key].push(row);
      }
    } else {
      allSegments.push({ seg: [row], groupType: row.groupType });
    }
  });

  for (const key in baseGroups) {
    const group = baseGroups[key];
    group.sort((a, b) => new Date(a.date).getTime() - new Date(b.date).getTime());

    let isPerfectlyConsecutive = true;
    for (let i = 1; i < group.length; i++) {
      const diffDays = Math.round((new Date(group[i].date) - new Date(group[i-1].date)) / 86400000);
      if (diffDays !== 7) {
        isPerfectlyConsecutive = false; 
        break;
      }
    }

    if (isPerfectlyConsecutive && group.length >= 4) {
      // ※注意：現在の06_Boshukun_Mainでは、定期募集(BASE)はマスターから直接作っているため、
      // ここで BASE に振り分けられたものは最終出力には使われません。
      allSegments.push({ seg: group, groupType: 'BASE' });
    } else {
      group.forEach(row => row.groupType = 'SINGLE_SHORT');
      allSegments.push({ seg: group, groupType: 'SINGLE_SHORT' });
    }
  }

  const groupedSingles = {};
  const finalAllSegments = [];

  allSegments.forEach(item => {
    if (item.groupType === 'BASE') {
      finalAllSegments.push(item);
    } else {
      const row = item.seg[0];
      const key = `${row.loc}_${row.category}_${row.start}-${row.end}_${item.groupType}`;
      if (!groupedSingles[key]) groupedSingles[key] = [];
      groupedSingles[key].push(...item.seg);
    }
  });

  for (const key in groupedSingles) {
    const grp = groupedSingles[key];
    const uniqueGrp = [];
    const seenDates = new Set();
    grp.forEach(g => {
      if (!seenDates.has(g.date)) {
        seenDates.add(g.date);
        uniqueGrp.push(g);
      }
    });
    uniqueGrp.sort((a, b) => new Date(a.date).getTime() - new Date(b.date).getTime());
    finalAllSegments.push({ seg: uniqueGrp, groupType: uniqueGrp[0].groupType });
  }

  finalAllSegments.sort((a, b) => {
    const locA = a.seg[0].loc;
    const locB = b.seg[0].loc;
    if (locA !== locB) return locA.localeCompare(locB, 'ja');
    const dateA = new Date(a.seg[0].date).getTime();
    const dateB = new Date(b.seg[0].date).getTime();
    if (dateA !== dateB) return dateA - dateB;
    return a.seg[0].startHour - b.seg[0].startHour;
  });

  const regularList = [];
  const singleList = [];

  finalAllSegments.forEach(item => {
    if (item.groupType === 'BASE') {
      regularList.push(_debug_formatRegularRow(item.seg, term, yearStr, areaMap));
    } else {
      singleList.push(_debug_formatSingleRow(item.seg, item.groupType, areaMap));
    }
  });

  return { regularList, singleList };
}

function _debug_formatRegularRow(segment, term, yearStr, areaMap) {
  const first = segment[0];
  const last = segment[segment.length - 1]; 
  const totalHours = segment.reduce((sum, g) => sum + g.hours, 0);
  const totalCost = segment.reduce((sum, g) => sum + (g.wage * g.hours), 0);

  let displayLoc = first.loc;
  if (displayLoc === "亀有" || displayLoc === "北葛西") displayLoc = `${displayLoc}（${first.category}）`;

  const shiftTitle = `${displayLoc}／${first.start}-${first.end}／毎週${first.dow}曜`;
  
  const actualStartStr = Utilities.formatDate(new Date(first.date), Session.getScriptTimeZone(), "yyyy/MM/dd");
  const actualEndStr = Utilities.formatDate(new Date(last.date), Session.getScriptTimeZone(), "yyyy/MM/dd");
  const period = `${actualStartStr}～${actualEndStr}`;

  const repeatDow = `毎週${first.dow}曜日`;
  const area = _getAreaHelper(displayLoc, areaMap);

  return [area, shiftTitle, period, first.start, first.end, repeatDow, "", first.wage, totalHours, totalCost, false, ""];
}

function _debug_formatSingleRow(segment, groupType, areaMap) {
  const first = segment[0];
  const last = segment[segment.length - 1];
  const formatMMDD = (dStr) => dStr.split('/').slice(1).join('/');
  
  const totalHours = segment.reduce((sum, g) => sum + g.hours, 0);
  const totalCost = segment.reduce((sum, g) => sum + (g.wage * g.hours), 0);

  let reason = "";
  if (groupType === 'ABSENCE') reason = "欠勤等による追加募集";
  else if (groupType === 'HOLIDAY') reason = "祝日・年末年始";
  else if (groupType === 'SINGLE_SHORT') reason = "一部日程のみの空き枠"; 

  let displayLoc = first.loc;
  if (displayLoc === "亀有" || displayLoc === "北葛西") displayLoc = `${displayLoc}（${first.category}）`;

  let period = `${formatMMDD(first.date)}（${first.dow}）`;
  let repeatDow = "単独"; 
  
  if (segment.length > 1) {
    period = `${formatMMDD(first.date)}～${formatMMDD(last.date)}`;
    
    let isConsecutive = true;
    let baseDow = first.dow;
    for(let i=1; i<segment.length; i++) {
      if(segment[i].dow !== baseDow) { isConsecutive = false; break; }
      let diffDays = Math.round((new Date(segment[i].date) - new Date(segment[i-1].date))/86400000);
      if (diffDays !== 7) { isConsecutive = false; break; }
    }
    
    if (isConsecutive && segment.length >= 3) {
      repeatDow = `毎週${first.dow}曜日`;
    } else {
      repeatDow = "複数日程";
    }
  }

  const targetDates = segment.map(g => `${formatMMDD(g.date)}(${g.dow})`).join(', ');
  const area = _getAreaHelper(displayLoc, areaMap);

  return [area, displayLoc, reason, period, first.start, first.end, repeatDow, targetDates, first.wage, totalHours, totalCost, false, "", ""];
}

function _debug_generateConfirmedShiftForSheet(shiftData, targetMonths, fullYearMonths, yearStr, term, areaMap) {
  const outputData = [];
  const dowNames = ["日", "月", "火", "水", "木", "金", "土"];
  const summary = {}; 
  
  const openDateMap = typeof _buildOpenDateMap === "function" ? _buildOpenDateMap(Object.keys(shiftData).map(k=>k.split('_')[0])) : {};
  if (typeof _initDoctorSpecialWageMap === "function") _initDoctorSpecialWageMap();

  fullYearMonths.forEach(monthStr => {
    for (const locKey in shiftData) {
      const monthData = shiftData[locKey];
      if (!monthData) continue;
      
      const parts = locKey.split('_');
      const cleanLocName = parts[0];
      const category = parts.length > 1 ? parts[1] : "小児科";

      for (const dateStr in monthData) {
        if (!dateStr.startsWith(monthStr)) continue;
        
        const dateObj = new Date(dateStr);
        const dateTime = dateObj.getTime();
        
        if (openDateMap[cleanLocName] && dateObj < openDateMap[cleanLocName]) continue; 

        const dailyShifts = monthData[dateStr];
        const dow = dowNames[dateObj.getDay()];
        const weekNum = Math.ceil(dateObj.getDate() / 7);
        
        dailyShifts.forEach(shift => {
          if (!shift.doctorName || shift.doctorName === "休" || shift.doctorName === "募集") return;
          
          const docName = shift.doctorName;
          const cleanDocName = docName.replace(/先生$/, "").trim();
          let docType = shift.type || "定期非常勤";
          if (docType !== "常勤" && docType !== "定期非常勤") docType = "定期非常勤";

          if (docType === "定期非常勤" && !targetMonths.includes(monthStr)) return;

          if (_doctorSpecialWageMap && _doctorSpecialWageMap[cleanDocName]) {
            const rule = _doctorSpecialWageMap[cleanDocName];
            if (rule.validPeriods && rule.validPeriods.length > 0) {
              const periodsForLoc = rule.validPeriods.filter(cp => cp.locName.includes(cleanLocName) || cleanLocName.includes(cp.locName));
              if (periodsForLoc.length > 0) {
                const isWithinPeriod = periodsForLoc.some(cp => dateTime >= cp.start && dateTime <= cp.end);
                if (!isWithinPeriod) return; 
              }
            }
          }

          let sHour = 9, eHour = 18;
          let sMin = "00", eMin = "00";
          
          if (shift.startTime) {
             const timeParts = String(shift.startTime).split(":");
             sHour = parseInt(timeParts[0], 10);
             if(timeParts[1]) sMin = timeParts[1];
          } else if (shift.startHour !== undefined) {
             sHour = shift.startHour;
          }
          
          if (shift.endTime) {
             const timeParts = String(shift.endTime).split(":");
             eHour = parseInt(timeParts[0], 10);
             if(timeParts[1]) eMin = timeParts[1];
          } else if (shift.hours !== undefined) {
             eHour = sHour + shift.hours;
          }
          const timeStr = `${('0'+sHour).slice(-2)}:${sMin}-${('0'+eHour).slice(-2)}:${eMin}`;

          const key = `${cleanLocName}_${docName}_${timeStr}`;
          if (!summary[key]) {
            summary[key] = {
              loc: cleanLocName,
              category: category,
              doctorName: docName,
              type: docType,
              timeStr: timeStr, 
              sHour: sHour, 
              eHour: eHour, 
              wage: "", 
              weeks: { 1: new Set(), 2: new Set(), 3: new Set(), 4: new Set(), 5: new Set() },
              minDate: dateTime, 
              maxDate: dateTime
            };
          } else {
            if (dateTime < summary[key].minDate) summary[key].minDate = dateTime;
            if (dateTime > summary[key].maxDate) summary[key].maxDate = dateTime;
          }
          
          summary[key].weeks[weekNum].add(dow);

          if (docType === "定期非常勤" && !summary[key].wage) {
            summary[key].wage = _debug_getMultiZoneWageString(dateStr, dow, sHour, eHour, cleanLocName, category, cleanDocName);
          }
        });
      }
    }
  });
  
  const dowOrder = ["月", "火", "水", "木", "金", "土", "日"];
  
  for (const key in summary) {
    const data = summary[key];
    let formattedWeeks = [];
    
    for (let w = 1; w <= 5; w++) {
      if (data.weeks[w].size > 0) {
        const sortedDows = Array.from(data.weeks[w]).sort((a, b) => dowOrder.indexOf(a) - dowOrder.indexOf(b));
        formattedWeeks.push(`第${w} (${sortedDows.join('')})`);
      }
    }
    
    let isEveryWeekSame = true;
    let baseDows = Array.from(data.weeks[1]).sort((a, b) => dowOrder.indexOf(a) - dowOrder.indexOf(b)).join('');
    
    if (baseDows === "") isEveryWeekSame = false;
    else {
      for (let w = 2; w <= 4; w++) {
        let dows = Array.from(data.weeks[w]).sort((a, b) => dowOrder.indexOf(a) - dowOrder.indexOf(b)).join('');
        if (dows !== baseDows) isEveryWeekSame = false;
      }
    }
    
    let finalDowStr = "";
    if (isEveryWeekSame) {
      let w5Dows = Array.from(data.weeks[5]).sort((a, b) => dowOrder.indexOf(a) - dowOrder.indexOf(b)).join('');
      if (w5Dows === baseDows || w5Dows === "") finalDowStr = `毎週 (${baseDows})`;
      else finalDowStr = formattedWeeks.join('、');
    } else {
      finalDowStr = formattedWeeks.join('、');
    }
    
    const minD = new Date(data.minDate);
    const maxD = new Date(data.maxDate);
    const startDStr = Utilities.formatDate(new Date(minD.getFullYear(), minD.getMonth(), 1), Session.getScriptTimeZone(), "yyyy年MM月dd日");
    const endDStr = Utilities.formatDate(new Date(maxD.getFullYear(), maxD.getMonth() + 1, 0), Session.getScriptTimeZone(), "yyyy年MM月dd日");
    const customPeriodStr = `${startDStr}～${endDStr}`;

    const isJoukin = (data.type === "常勤");
    const outTitle = isJoukin ? `${data.doctorName}先生／通年` : `${data.doctorName}先生／${term}`;
    const outPeriod = customPeriodStr; 
    const outWage = isJoukin ? "" : data.wage; 

    let displayLoc = data.loc;
    if (displayLoc === "亀有" || displayLoc === "北葛西") {
      displayLoc = `${displayLoc}（${data.category}）`;
    }
    const area = _getAreaHelper(displayLoc, areaMap);

    outputData.push([
      area, displayLoc, data.type, `${data.doctorName}先生`, outTitle, outPeriod, 
      data.timeStr, finalDowStr, outWage, false, ""
    ]);
  }
  return outputData;
}

function _debug_getSingleDayShiftsFromSheets(yearStr, term, locNames, shiftData, areaMap) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const advanceSheet = ss.getSheetByName('先行応募');
  const substituteSheet = ss.getSheetByName('振替勤務');
  const absenceSheet = ss.getSheetByName('お休み情報');
  
  const year = parseInt(yearStr, 10);
  const startDate = new Date(year, 3, 1); 
  const endDate = new Date(year + 1, 2, 31); 
  const summary = {};

  const docTypeMap = {};
  if (shiftData) {
    for (const locKey in shiftData) {
      for (const dateStr in shiftData[locKey]) {
        shiftData[locKey][dateStr].forEach(shift => {
          if (shift.doctorName && shift.doctorName !== "休" && shift.doctorName !== "募集") {
            let dName = shift.doctorName.replace(/先生$/, "").trim();
            docTypeMap[dName] = shift.type || "定期非常勤";
          }
        });
      }
    }
  }

  const formatTimeStr = (val) => {
    if (val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), "HH:mm");
    return String(val).substring(0, 5); 
  };

  const processSheet = (sheet, typeName, isAbsence = false) => {
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    const dowNames = ["日", "月", "火", "水", "木", "金", "土"];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const dateVal = row[0];
      if (!dateVal) continue;
      
      const d = new Date(dateVal);
      if (d < startDate || d > endDate) continue;
      
      const locFull = String(row[4]).trim();
      if (!locNames.some(l => locFull.includes(l))) continue;
      
      const startTime = formatTimeStr(row[2]);
      const endTime = formatTimeStr(row[3]);
      
      let rawDocName = isAbsence ? (String(row[6]).trim() || String(row[5]).trim()) : String(row[5]).trim();
      const docName = rawDocName.replace(/先生$/, "").trim();
      const dow = dowNames[d.getDay()];
      
      let finalType = typeName;
      if (isAbsence) {
        finalType = String(row[5]).trim(); 
        if (finalType !== "欠勤" && finalType !== "有給") continue;
      }
      
      const docType = docTypeMap[docName] || "定期非常勤";
      let isJoukin = (docType === "常勤");
      
      if (typeName === "振替勤務") {
        if (isJoukin) finalType = "振替勤務（常勤）";
        else finalType = "振替勤務（定期非常勤）";
      }
      
      const key = `${docName}_${locFull}_${finalType}_${startTime}_${endTime}_${dow}`;
      if (!summary[key]) {
        summary[key] = {
          docName, loc: locFull, type: finalType, startTime, endTime, dow, isJoukin,
          dObjs: []
        };
      }
      summary[key].dObjs.push(d);
    }
  };

  processSheet(advanceSheet, "先行応募");
  processSheet(substituteSheet, "振替勤務"); 
  processSheet(absenceSheet, "お休み情報", true);

  const results = [];
  for (const key in summary) {
    const item = summary[key];
    item.dObjs.sort((a,b) => a - b);
    const cleanLoc = item.loc.split('_')[0].replace(/クリニック|診療所/g, "").trim();
    
    let segments = [];
    let currentSegment = [item.dObjs[0]];
    
    for (let i = 1; i < item.dObjs.length; i++) {
      const diffDays = Math.round((item.dObjs[i] - item.dObjs[i-1]) / 86400000);
      if (diffDays === 7) {
        currentSegment.push(item.dObjs[i]);
      } else {
        if (currentSegment.length >= 4) {
          segments.push({ dates: currentSegment, type: 'weekly' });
        } else {
          currentSegment.forEach(d => segments.push({ dates: [d], type: 'single' }));
        }
        currentSegment = [item.dObjs[i]];
      }
    }
    if (currentSegment.length >= 4) {
      segments.push({ dates: currentSegment, type: 'weekly' });
    } else {
      currentSegment.forEach(d => segments.push({ dates: [d], type: 'single' }));
    }
    
    segments.forEach(seg => {
      const sDates = seg.dates;
      const startD = sDates[0];
      const endD = sDates[sDates.length - 1];
      
      let periodStr = "";
      let dowStr = "";
      if (seg.type === 'single') {
        periodStr = Utilities.formatDate(startD, Session.getScriptTimeZone(), "yyyy年MM月dd日");
        dowStr = "単独";
      } else {
        periodStr = `${Utilities.formatDate(startD, Session.getScriptTimeZone(), "yyyy年MM月dd日")}～${Utilities.formatDate(endD, Session.getScriptTimeZone(), "yyyy年MM月dd日")}`;
        dowStr = `毎週${item.dow}曜日`;
      }
      
      const category = item.loc.includes("内科") ? "内科" : "小児科"; 
      let startHour = parseInt(item.startTime.split(":")[0], 10);
      let endHour = parseInt(item.endTime.split(":")[0], 10);
      
      let wage = "";
      // ★修正：欠勤・有給・休館日の場合は時給を計算せず、空欄にする
      if (!item.isJoukin && item.type !== "欠勤" && item.type !== "有給" && item.type !== "休館日") {
        wage = _debug_getMultiZoneWageString(Utilities.formatDate(startD, Session.getScriptTimeZone(), "yyyy/MM/dd"), item.dow, startHour, endHour, cleanLoc, category, item.docName);
      }

      let displayLoc = cleanLoc;
      if (displayLoc === "亀有" || displayLoc === "北葛西") {
        displayLoc = `${displayLoc}（${category}）`;
      }
      
      let finalDisplayLoc = displayLoc;
      let areaSearchKey = displayLoc;

      if (item.type === "有給" || item.type === "欠勤") {
        finalDisplayLoc = item.type; 
        areaSearchKey = item.type; 
      }
      
      const area = _getAreaHelper(areaSearchKey, areaMap);

      results.push([
        area, finalDisplayLoc, item.type, `${item.docName}先生`, "", periodStr, 
        `${item.startTime}-${item.endTime}`, dowStr, wage, false, ""
      ]);
    });
  }
  return results;
}

// ★修正：重複排除（seenKeys）と名前のスペース除去処理を追加
function _debug_getCancelShiftsFromSheets(yearStr, locNames, rawBoshuData, shiftData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const absenceSheet = ss.getSheetByName('お休み情報');
  const results = [];
  const year = parseInt(yearStr, 10);
  const startDate = new Date(year, 3, 1); 
  const endDate = new Date(year + 1, 2, 31); 
  
  // 重複チェック用のSet
  const seenKeys = new Set();
  
  const formatTimeStr = (val) => {
    if (val instanceof Date) return Utilities.formatDate(val, Session.getScriptTimeZone(), "HH:mm");
    return String(val).substring(0, 5); 
  };
  
  if (absenceSheet) {
    const data = absenceSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const dateVal = row[0];
      if (!dateVal) continue;
      
      const d = new Date(dateVal);
      if (d < startDate || d > endDate) continue;
      
      const loc = String(row[4]).trim();
      if (!locNames.some(l => loc.includes(l))) continue;
      
      const reason = String(row[5]).trim();
      // 「時短」も元の出力に含まれていたため許容リストに追加しています
      if (!["欠勤", "有給", "時短", "祝日", "年末年始"].includes(reason)) continue;
      
      const startTime = formatTimeStr(row[2]);
      const endTime = formatTimeStr(row[3]);
      
      // 医師名から全角・半角スペースを完全に除去
      let rawDocName = String(row[6]).trim() || String(row[5]).trim(); 
      let docName = rawDocName.replace(/[\s ]+/g, "").replace(/先生$/, "").trim();
      const doctor = `${docName}先生`;
      
      const dateStrFormatted = Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy年MM月dd日");
      const dStrForMatch = Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy/MM/dd");
      
      const cleanLoc = loc.split('_')[0].replace(/クリニック|診療所/g, "").trim();
      const category = loc.includes("内科") ? "内科" : "小児科";
      
      let displayLoc = cleanLoc;
      if (displayLoc === "亀有" || displayLoc === "北葛西") {
        displayLoc = `${displayLoc}（${category}）`;
      }

      const timeStr = `${startTime}〜${endTime}`;
      
      // 重複チェック：同じ医師・日・理由・拠点・時間がすでに登録されていればスキップ
      const uniqueKey = `${doctor}_${dateStrFormatted}_${reason}_${displayLoc}_${timeStr}`;
      if (seenKeys.has(uniqueKey)) {
        continue;
      }
      seenKeys.add(uniqueKey);
      
      const cancelStartStr = String(startTime).split(":")[0];
      const cancelEndStr = String(endTime).split(":")[0];
      let cancelStart = 9;
      let cancelEnd = 13;
      if (!isNaN(parseInt(cancelStartStr, 10))) cancelStart = parseInt(cancelStartStr, 10);
      if (!isNaN(parseInt(cancelEndStr, 10))) cancelEnd = parseInt(cancelEndStr, 10);
      
      const hasBoshu = rawBoshuData.some(b => {
        if (b.date !== dStrForMatch || b.loc !== cleanLoc) return false;
        const bStart = b.startHour;
        const bEnd = bStart + b.hours;
        return (cancelStart < bEnd && cancelEnd > bStart);
      });
      
      const boshuStatus = hasBoshu ? "済" : "";
      
      results.push([doctor, dateStrFormatted, reason, displayLoc, timeStr, boshuStatus, false, "", false]);
    }
  }

  const kyukanMap = _debug_getKyukanMap();
  for (const key in kyukanMap) {
    const [dateStr, cleanLoc] = key.split('_');
    const d = new Date(dateStr);
    if (d < startDate || d > endDate) continue;
    if (!locNames.some(l => cleanLoc.includes(l) || l.includes(cleanLoc))) continue;
    
    if (shiftData && shiftData[cleanLoc] && shiftData[cleanLoc][dateStr]) {
      shiftData[cleanLoc][dateStr].forEach(shift => {
        if (shift.doctorName && shift.doctorName !== "休" && shift.doctorName !== "募集") {
           // 休館日側でもスペース除去
           let docName = shift.doctorName.replace(/[\s ]+/g, "").replace(/先生$/, "").trim();
           let doctor = `${docName}先生`;
           let dateStrFormatted = Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy年MM月dd日");
           let category = shift.rawShift.includes("内科") ? "内科" : "小児科";
           
           let displayLoc = cleanLoc;
           if (displayLoc === "亀有" || displayLoc === "北葛西") {
             displayLoc = `${displayLoc}（${category}）`;
           }
           let timeStr = `${shift.startTime}〜${shift.endTime}`;
           
           // 休館日側でも重複チェック
           const uniqueKey = `${doctor}_${dateStrFormatted}_休館日_${displayLoc}_${timeStr}`;
           if (seenKeys.has(uniqueKey)) {
             return; 
           }
           seenKeys.add(uniqueKey);
           
           results.push([doctor, dateStrFormatted, "休館日", displayLoc, timeStr, "", false, "", false]);
        }
      });
    }
  }
  
  return results;
}