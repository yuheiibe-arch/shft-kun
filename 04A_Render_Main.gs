/**
 * ==========================================
 * 04A_Render_Main.gs
 * テンプレートの複製・メイン描画・パッチデータ適用
 * ★ 曜日・週番号の書き込み爆速化（API通信削減）パッチ適用版
 * ★【先行応募・振替等】マスタ判定維持 ＆ 不明時の忖度（両方出力）完全排除版
 * ★【原点回帰】初期キャンバス1200行拡張 ＆ 1マスずつの通信を完全廃止した一括バッチ上書き版
 * ==========================================
 */

var globalWageCache = {};
var globalLocationDict = null;
var globalOpenDatesCache = null;
var globalOverrideRawData = null;
var globalSpecialtyMap = null; 

function getShiftOverrides(ss, originalLocName, cleanLocName, yearMonthStr) {
  let overrides = { advance: {}, absence: {}, substitute: {}, kyukan: {}, senkouDocs: new Set(), substituteDocs: new Set() };
  
  const [yStr, mStr] = yearMonthStr.split('/');

  if (!globalSpecialtyMap) {
    globalSpecialtyMap = typeof buildDoctorSpecialtyMap === 'function' ? buildDoctorSpecialtyMap(yStr) : {};
  }

  if (!globalOverrideRawData) {
    globalOverrideRawData = { advance: [], absence: [], substitute: [], kyukan: [] };
    let sAdv = ss.getSheetByName('先行応募'); if(sAdv) globalOverrideRawData.advance = sAdv.getDataRange().getValues();
    let sAbs = ss.getSheetByName('お休み情報'); if(sAbs) globalOverrideRawData.absence = sAbs.getDataRange().getValues();
    let sSub = ss.getSheetByName('振替勤務'); if(sSub) globalOverrideRawData.substitute = sSub.getDataRange().getValues();
    try {
      const kyuSs = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/1cbeXWojsxNMhQUo1c6VflF5hLUJUyfuOXCFbGP5jJEA/edit");
      const kyuSheet = kyuSs.getSheetByName("休館日");
      if (kyuSheet) globalOverrideRawData.kyukan = kyuSheet.getDataRange().getValues();
    } catch(e) {}
  }

  const processOverrideData = (data, targetDict, type) => {
    if (!data || data.length === 0) return;
    for(let i=1; i<data.length; i++) {
      let dateVal = data[i][0];
      if(!dateVal) continue;
      let d = new Date(dateVal);
      if (isNaN(d.getTime())) continue;
      let dStr = Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy/MM/dd");
      
      if (!dStr.startsWith(yearMonthStr)) continue;
      let locFull = String(data[i][4]).trim();
      if(!locFull.includes(cleanLocName)) continue;
      
      let doc = type === 'absence' ? (String(data[i][6]).trim() || String(data[i][5]).trim()) : String(data[i][5]).trim();
      doc = doc.replace(/[\s ]+/g, "").replace(/先生$/, "").trim();

      if (originalLocName !== cleanLocName) {
        let targetCat = originalLocName.includes("内科") ? "内科" : "小児科";
        let otherCat = targetCat === "内科" ? "小児科" : "内科";
        
        let docSpec = globalSpecialtyMap[doc] || "不明";
        
        let isBoxMatch = locFull.includes(targetCat);
        let isOtherBoxMatch = locFull.includes(otherCat);
        let isSubjectMatch = docSpec.includes(targetCat);
        let isOtherSubjectMatch = docSpec.includes(otherCat);

        if (isOtherBoxMatch && !isBoxMatch) {
          continue; 
        } else if (isBoxMatch) {
          // 通す
        } else {
          if (targetCat === "内科") {
            if (!isSubjectMatch) {
              continue; 
            }
          } else {
            if (isOtherSubjectMatch) {
              continue; 
            }
          }
        }
      }
      
      let sTime = String(data[i][2] instanceof Date ? Utilities.formatDate(data[i][2], Session.getScriptTimeZone(), "HH:mm") : data[i][2]);
      let eTime = String(data[i][3] instanceof Date ? Utilities.formatDate(data[i][3], Session.getScriptTimeZone(), "HH:mm") : data[i][3]);
      
      let sH = parseInt(sTime.split(":")[0], 10);
      let eH = parseInt(eTime.split(":")[0], 10);
      if(isNaN(sH) || isNaN(eH)) continue;
      
      if(!targetDict[dStr]) targetDict[dStr] = [];
      targetDict[dStr].push({startH: sH, endH: eH, doc: doc});
      
      if(type === 'advance') overrides.senkouDocs.add(doc);
      if(type === 'substitute') overrides.substituteDocs.add(doc);
    }
  };

  processOverrideData(globalOverrideRawData.advance, overrides.advance, 'advance');
  processOverrideData(globalOverrideRawData.absence, overrides.absence, 'absence'); 
  processOverrideData(globalOverrideRawData.substitute, overrides.substitute, 'substitute');

  let kyuData = globalOverrideRawData.kyukan;
  if (kyuData && kyuData.length > 0) {
    for (let i = 1; i < kyuData.length; i++) {
      let dVal = kyuData[i][0];
      if (!dVal) continue;
      let d = new Date(dVal);
      if (isNaN(d.getTime())) continue;
      let dStr = Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy/MM/dd");
      
      if (!dStr.startsWith(yearMonthStr)) continue;
      let kLoc = String(kyuData[i][3]).trim();
      if (!kLoc.includes(cleanLocName)) continue;
      
      if (originalLocName !== cleanLocName) {
        let targetCat = originalLocName.includes("内科") ? "内科" : "小児科";
        let otherCat = targetCat === "内科" ? "小児科" : "内科";
        if (kLoc.includes(otherCat)) continue;
      }
      
      overrides.kyukan[dStr] = true;
    }
  }

  return overrides;
}

function renderShiftBlock(ss, originalLocName, finalSheetName, yearMonthStr, monthData) {
  const cleanLocName = originalLocName.replace(/（.*?）/, '');
  if (!globalLocationDict) globalLocationDict = typeof getLocationDictionary === 'function' ? getLocationDictionary() : {};
  if (!globalOpenDatesCache) globalOpenDatesCache = typeof getLocationOpenDates === 'function' ? getLocationOpenDates() : {};
  
  const locOpenDate = globalOpenDatesCache[cleanLocName];
  const [yStr, mStr] = yearMonthStr.split('/');
  const monthEndDate = new Date(parseInt(yStr, 10), parseInt(mStr, 10), 0);
  if (locOpenDate && monthEndDate < locOpenDate) {
    return false; 
  }

  if (!globalWageCache[cleanLocName]) {
    globalWageCache[cleanLocName] = typeof getClinicWages === 'function' ? getClinicWages(cleanLocName) : [];
  }
  const wageDataList = globalWageCache[cleanLocName];
  let overrides = getShiftOverrides(ss, originalLocName, cleanLocName, yearMonthStr);
  
  let sheet = ss.getSheetByName(finalSheetName);
  let isNewBlock = false;
  
  if (!sheet) {
    sheet = ss.insertSheet(finalSheetName);
    isNewBlock = true;

    // =========================================================
    // ★ 合理化パッチ①：初期キャンバス拡張
    // 1000行で足りなくなることを見越し、最初から1200行に拡張しておく
    // （これにより、ループ処理中の行追加による激重処理がゼロになります）
    // =========================================================
    let initialMaxCols = sheet.getMaxColumns();
    if (initialMaxCols > 16) {
      sheet.deleteColumns(17, initialMaxCols - 16);
    }
    
    let currentMaxRows = sheet.getMaxRows();
    if (currentMaxRows < 1200) {
      sheet.insertRowsAfter(currentMaxRows, 1200 - currentMaxRows);
    }
    // =========================================================
  }
  
  const templateSheet = ss.getSheetByName(typeof CONFIG !== 'undefined' && CONFIG.TEMPLATE_SHEET_NAME ? CONFIG.TEMPLATE_SHEET_NAME : "テンプレート");
  if (!templateSheet) throw new Error("テンプレートシートが見つかりません。");

  const edges = typeof getMonthEdges === 'function' ? getMonthEdges(yearMonthStr) : {};
  const tempRange = templateSheet.getDataRange();
  let lastRow = sheet.getLastRow();
  let startRow = 0;
  
  if (lastRow > 0) {
    let searchData = sheet.getRange(1, 1, lastRow, 4).getDisplayValues();
    for (let i = 0; i < searchData.length; i++) {
      if (searchData[i][0] === "適用開始" && String(searchData[i][3]).startsWith(yearMonthStr)) {
        startRow = (i + 1) - 15;
        if (startRow < 1) startRow = 1;
        break; 
      }
    }
  }
  
  if (startRow === 0) {
    startRow = lastRow === 0 ? 1 : lastRow + 4;
    isNewBlock = true;
  }
  
  let blockEndRow = startRow + tempRange.getNumRows() - 1;
  const requiredRows = blockEndRow + 10;
  
  // 1200行に拡張済みなので、ここはほぼ通過します（無駄な追加が起きない）
  if (sheet.getMaxRows() < requiredRows) {
    sheet.insertRowsAfter(sheet.getMaxRows(), Math.max(requiredRows - sheet.getMaxRows(), 100));
  }
  
  if (isNewBlock) {
    tempRange.copyTo(sheet.getRange(startRow, 1));
    const fullBlockRange = sheet.getRange(startRow, 1, tempRange.getNumRows(), 16);
    fullBlockRange.setBorder(true, true, true, true, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_THICK);
    sheet.getRange(blockEndRow, 1, 1, 16).setBorder(null, null, true, null, null, null, '#000000', SpreadsheetApp.BorderStyle.SOLID_THICK);
  }

  let stats = typeof initStats === 'function' ? initStats() : {}; 
  let doctorCosts = {};

  if (typeof _debug_initHolidayMap === "function") _debug_initHolidayMap(yStr);

  let allDocsThisMonth = new Set(["募集", "休館日", "未開院"]);
  let teikiDocs = new Set(); 

  Object.values(monthData).forEach(dayShifts => {
    dayShifts.forEach(s => {
      let cleanDocName = s.doctorName.replace(/[\s ]+/g, "").replace(/先生$/, "").trim();
      if (cleanDocName && cleanDocName !== "休") {
        allDocsThisMonth.add(cleanDocName);
        teikiDocs.add(cleanDocName); 
      }
    });
  });

  let filteredSenkouDocs = new Set();
  overrides.senkouDocs.forEach(doc => {
    let cleanDoc = doc.replace(/[\s ]+/g, "").replace(/先生$/, "").trim();
    if (!teikiDocs.has(cleanDoc)) {
      filteredSenkouDocs.add(doc); 
    }
  });
  overrides.senkouDocs = filteredSenkouDocs;

  overrides.senkouDocs.forEach(doc => allDocsThisMonth.add(doc));
  overrides.substituteDocs.forEach(doc => allDocsThisMonth.add(doc));
  
  const dynamicRule = SpreadsheetApp.newDataValidation().requireValueInList(Array.from(allDocsThisMonth), true).build();
  const groupedDays = typeof getMonthDaysGroupedByDOW === 'function' ? getMonthDaysGroupedByDOW(yearMonthStr) : [];
  
  let currentRow = startRow + 19;
  
  const targetRange = sheet.getRange(startRow + 19, 4, groupedDays.length * 2, 12);
  let currentSheetValues = targetRange.getValues();

  let writeValues = [];
  let writeBgs = [];
  let writeFonts = [];
  let writeRules = [];
  
  let actualShiftHasDoctorCount = 0; 

  const leftHeaderRange = sheet.getRange(startRow + 19, 1, groupedDays.length * 2, 2);
  let leftHeaderValues = leftHeaderRange.getValues();
  let leftHeaderModified = false;

  groupedDays.forEach((dayInfo, dIdx) => {
    let rIdx = dIdx * 2;
    if (!leftHeaderValues[rIdx][0]) {
      leftHeaderValues[rIdx][0] = dayInfo.dayOfWeek;
      leftHeaderModified = true;
    }
    if (!leftHeaderValues[rIdx][1]) {
      leftHeaderValues[rIdx][1] = dayInfo.weekNum;
      leftHeaderModified = true;
    }
    
    if (!dayInfo.isValid) {
      let emptyRow = new Array(12).fill("");
      let grayRow = new Array(12).fill("#cccccc");
      let nullRules = new Array(12).fill(null);
      writeValues.push(emptyRow, emptyRow);
      writeBgs.push(grayRow, grayRow);
      writeFonts.push(grayRow, grayRow);
      writeRules.push(nullRules, nullRules);
      currentRow += 2;
      return; 
    }
 
    const dailyShifts = monthData[dayInfo.dateStr] || [];
    const tetrisResult = typeof calculateTetrisAllocation === 'function' ? calculateTetrisAllocation(dailyShifts, cleanLocName) : {line1:[], line2:[], docTypes:{}};
    
    let currentDate = new Date(dayInfo.dateStr);
    let isBeforeOpen = false;
    if (!locOpenDate || currentDate < locOpenDate) isBeforeOpen = true;

    let isTrueHoliday = false;
    if (typeof _debug_isTrueHoliday === "function") isTrueHoliday = _debug_isTrueHoliday(dayInfo.dateStr);
    let dayType = dayInfo.dayOfWeek === "土" ? "土曜" : dayInfo.dayOfWeek === "日" ? "日曜" : "平日";
    if (isTrueHoliday) dayType = "日曜";
    let isHoliday = (dayType !== "平日");
    
    const dateKey = dayInfo.dateStr;
    const isKyukan = overrides.kyukan[dateKey];
    const dayAbsences = overrides.absence[dateKey] || [];
    const dayAdvances = overrides.advance[dateKey] || [];
    const daySubstitutes = overrides.substitute[dateKey] || [];

    if (isBeforeOpen) {
      for(let i = 0; i < 12; i++) {
        if(i !== 4 && i !== 5) {
          tetrisResult.line1[i] = "未開院";
          tetrisResult.line2[i] = "未開院";
        }
      }
    } else if (isKyukan) {
      for(let i = 0; i < 12; i++) {
        if(i !== 4 && i !== 5) {
          tetrisResult.line1[i] = "休館日";
          tetrisResult.line2[i] = "休館日";
        }
      }
    } else {
      dayAbsences.forEach(abs => {
        for(let h = abs.startH; h < abs.endH; h++) {
          let idx = h - 9;
          if(idx >= 0 && idx < 12 && idx !== 4 && idx !== 5) {
            let cl1 = tetrisResult.line1[idx] ? String(tetrisResult.line1[idx]).replace(/[\s ]+/g, "") : null;
            let cl2 = tetrisResult.line2[idx] ? String(tetrisResult.line2[idx]).replace(/[\s ]+/g, "") : null;
            if(cl1 === abs.doc) tetrisResult.line1[idx] = "募集";
            if(cl2 === abs.doc) tetrisResult.line2[idx] = "募集";
          }
        }
      });
      
      dayAdvances.forEach(adv => {
        for(let h = adv.startH; h < adv.endH; h++) {
          let idx = h - 9;
          if(idx >= 0 && idx < 12 && idx !== 4 && idx !== 5) {
            let cl1 = tetrisResult.line1[idx] ? String(tetrisResult.line1[idx]).replace(/[\s ]+/g, "") : null;
            let cl2 = tetrisResult.line2[idx] ? String(tetrisResult.line2[idx]).replace(/[\s ]+/g, "") : null;
            
            if (cl1 === adv.doc || cl2 === adv.doc) continue;

            if(tetrisResult.line1[idx] === "募集" || tetrisResult.line1[idx] === null) {
              tetrisResult.line1[idx] = adv.doc;
              tetrisResult.docTypes[adv.doc] = "先行応募";
            } else if(tetrisResult.line2[idx] === "募集" || tetrisResult.line2[idx] === null) {
              tetrisResult.line2[idx] = adv.doc;
              tetrisResult.docTypes[adv.doc] = "先行応募";
            }
          }
        }
      });
      
      daySubstitutes.forEach(sub => {
        for(let h = sub.startH; h < sub.endH; h++) {
          let idx = h - 9;
          if(idx >= 0 && idx < 12 && idx !== 4 && idx !== 5) {
            let cl1 = tetrisResult.line1[idx] ? String(tetrisResult.line1[idx]).replace(/[\s ]+/g, "") : null;
            let cl2 = tetrisResult.line2[idx] ? String(tetrisResult.line2[idx]).replace(/[\s ]+/g, "") : null;
            
            if (cl1 === sub.doc || cl2 === sub.doc) continue;

            if(tetrisResult.line1[idx] === "募集" || tetrisResult.line1[idx] === null) {
              tetrisResult.line1[idx] = sub.doc;
              tetrisResult.docTypes[sub.doc] = "振替勤務";
            } else if(tetrisResult.line2[idx] === "募集" || tetrisResult.line2[idx] === null) {
              tetrisResult.line2[idx] = sub.doc;
              tetrisResult.docTypes[sub.doc] = "振替勤務";
            }
          }
        }
      });
    }

    let line1V = [], line2V = [], line1Bg = [], line2Bg = [], line1Fc = [], line2Fc = [], line1Rl = [], line2Rl = [];
    
    for (let i = 0; i < 12; i++) {
      let res1 = processCell(tetrisResult.line1[i], tetrisResult.docTypes, dynamicRule, dayType, i, stats, true);
      line1V.push(res1.val); line1Bg.push(res1.bg); line1Fc.push(res1.fc); line1Rl.push(res1.rule);
      if (res1.val && res1.val !== "募集" && res1.val !== "休館日" && res1.val !== "未開院") actualShiftHasDoctorCount++;
      
      let res2 = processCell(tetrisResult.line2[i], tetrisResult.docTypes, dynamicRule, dayType, i, stats, false);
      line2V.push(res2.val); line2Bg.push(res2.bg); line2Fc.push(res2.fc); line2Rl.push(res2.rule);
      if (res2.val && res2.val !== "募集" && res2.val !== "休館日" && res2.val !== "未開院") actualShiftHasDoctorCount++;
    }
    
    if (!isBeforeOpen) {
      dailyShifts.forEach(shift => {
        if (shift.type === "定期非常勤") {
          let cleanShiftName = shift.doctorName.replace(/[\s ]+/g, "").replace(/先生$/, "").trim();
          if (!doctorCosts[cleanShiftName]) {
            doctorCosts[cleanShiftName] = { hours: 0, cost: 0, rawShifts: [], appliedRates: new Set() };
          }
          
          doctorCosts[cleanShiftName].rawShifts.push({ start: shift.startTime, end: shift.endTime, dow: dayInfo.dayOfWeek, week: dayInfo.weekNum });
          let costInfo = typeof calculateDailyCost === 'function' ? calculateDailyCost(shift, isHoliday, wageDataList) : {hours:0, cost:0, appliedRates:new Set()};
          doctorCosts[cleanShiftName].hours += costInfo.hours;
          doctorCosts[cleanShiftName].cost += costInfo.cost;
          costInfo.appliedRates.forEach(r => doctorCosts[cleanShiftName].appliedRates.add(r));
        }
      });
    }

    writeValues.push(line1V, line2V);
    writeBgs.push(line1Bg, line2Bg);
    writeFonts.push(line1Fc, line2Fc);
    writeRules.push(line1Rl, line2Rl);
    
    currentRow += 2;
  });

  if (leftHeaderModified) {
    leftHeaderRange.setValues(leftHeaderValues).setHorizontalAlignment("left");
  }

  if (!isNewBlock && actualShiftHasDoctorCount === 0) {
    let currentSheetHasDoctor = currentSheetValues.some(row => row.some(val => val && val !== "募集" && val !== "休" && val !== "休館日" && val !== "未開院" && val !== ""));
    if (currentSheetHasDoctor) {
      console.error(`🛑 [重大な危機を検知] 拠点「${finalSheetName}」のデータが空（0件）として計算されましたが、既存シートには医師のシフトが存在します。白紙化（誤上書き）を防ぐため、この拠点の同期処理を完全拒否（スキップ）しました。`);
      return false;
    }
  }

  // =========================================================
  // ★ 合理化パッチ②：一括バッチ書込（完全上書き方式）
  // 1マスごとの丁寧な通信（差分更新）を完全に廃止し、
  // 変更の有無に関わらず、メモリ上で完成させた1ヶ月分のデータを
  // 1回の通信で「ガバッ」と上書きする爆速モードで統一します。
  // =========================================================
  targetRange.setValues(writeValues)
             .setBackgrounds(writeBgs)
             .setFontColors(writeFonts)
             .setDataValidations(writeRules)
             .setHorizontalAlignment("left");
  
  console.log(`🏥 拠点 [${finalSheetName}] (${yearMonthStr}): 描画完了（一括バッチ書込）`);
  // =========================================================

  if (typeof writeDashboardInfo === 'function') {
    writeDashboardInfo(sheet, startRow, blockEndRow, edges, stats, doctorCosts, wageDataList, Array.from(overrides.senkouDocs));
  }
  
  // SpreadsheetApp.flush(); // 強制セーブによるファイル全体への負荷をなくすため削除済み
  return true; 
}

function processCell(docName, docTypes, baseRule, dayType, hourIdx, stats, isShin1) {
  let val = "", bg = "#ffffff", fc = "#000000", rule = baseRule;
  if (docName) docName = String(docName).replace(/[\s ]+/g, "");

  if (docName === "休") {
    bg = "#cccccc"; fc = "#cccccc";
    rule = null; 
  } else if (docName === "休館日") {
    val = "休館日"; bg = "#cccccc";
    fc = "#666666"; rule = baseRule;
  } else if (docName === "未開院") {
    val = "未開院";
    bg = "#e0e0e0"; fc = "#999999"; rule = baseRule; 
  } else if (docName === "募集") {
    val = "募集";
  } else if (docName) {
    val = docName;
    if (docTypes[docName] === "常勤") { if(stats.uniqueJoukin) stats.uniqueJoukin.add(docName); }
    else if (docTypes[docName] === "先行応募") { if(stats.uniqueSenkou) stats.uniqueSenkou.add(docName); }
    else if (docTypes[docName] === "振替勤務") { if(stats.uniqueSubstitute) stats.uniqueSubstitute.add(docName); }
    else { if(stats.uniqueTeiki) stats.uniqueTeiki.add(docName); }
  }

  if (docName !== "休" && docName !== "休館日" && docName !== "未開院" && stats.grid) {
    let timeZone = "";
    if (hourIdx >= 0 && hourIdx <= 3) timeZone = "9-13";
    else if (hourIdx >= 6 && hourIdx <= 8) timeZone = "15-18";
    else if (hourIdx >= 9 && hourIdx <= 11) timeZone = "18-21";
    
    if (isShin1) {
      stats.total1shin++;
      if (val === "募集") stats.boshu++;
      else if (docTypes[val] === "常勤") stats.joukin1shin++;
      else if (docTypes[val] === "先行応募") stats.senkou1shin++;
      else if (docTypes[val] === "定期非常勤") stats.teiki1shin++;
    } else {
      if (val && val !== "募集") stats.total2shinFilled++;
    }

    if (timeZone && stats.grid[dayType] && stats.grid[dayType][timeZone]) {
      stats.grid[dayType][timeZone].total++;
      if (val === "募集") stats.grid[dayType][timeZone].boshu++;
      else if (val) stats.grid[dayType][timeZone].filled++;
    }
  }

  return { val: val, bg: bg, fc: fc, rule: rule };
}