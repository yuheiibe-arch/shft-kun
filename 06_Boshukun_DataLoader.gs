/**
 * ====================================================================
 * 07_Boshukun_DataLoader.gs
 * カレンダー生成・マスタ読み込み・契約備考の解析エンジン
 * ★【完全版】退職日マスタ連携＆最強パーサー（全角記号網羅版）
 * ====================================================================
 */

function _loadBoshukunData(ctx) {
  let initialLocNames = [];
  if (ctx.mode === 'area') {
    initialLocNames = getTargetLocationsFromMaster(ctx.year, ctx.term, ctx.areas);
  } else {
    if (ctx.singleLoc) initialLocNames = [ctx.singleLoc];
  }

  if (initialLocNames.length === 0) {
    throw new Error("選択されたエリア・条件に合致する拠点がありません。");
  }

  // ★亀有・北葛西の強制展開 ＋ MQCの常時追加
  let expandedLocNames = [];
  initialLocNames.forEach(loc => {
    const clean = loc.split('_')[0].replace(/[（\(\)）]/g, "").replace(/内科|小児科/g, "").trim();
    if (clean === "亀有" || clean === "北葛西") {
      expandedLocNames.push(`${clean}_内科`);
      expandedLocNames.push(`${clean}_小児科`);
    } else {
      expandedLocNames.push(loc);
    }
  });
  if (!expandedLocNames.includes("MQC")) expandedLocNames.push("MQC");
  ctx.locNames = [...new Set(expandedLocNames)];

  ctx.targetDisplayLocs = ctx.locNames.map(locName => {
    const clean = locName.split('_')[0].replace(/[（\(\)）]/g, "").replace(/内科|小児科/g, "").trim();
    const cat = locName.includes("内科") ? "内科" : "小児科";
    if (clean === "亀有" || clean === "北葛西") return `${clean}（${cat}）`;
    return clean;
  });

  if (typeof initializeWageData === "function") initializeWageData();
  _debug_initHolidayMap(ctx.year);

  const locMaster = _buildLocMasterDict();
  ctx._areaMap = locMaster.areaMap;
  ctx._openDateMap = locMaster.openDateMap;

  const y = parseInt(ctx.year, 10);
  ctx.startDate = ctx.term === "上期" ? new Date(y, 3, 1) : new Date(y, 9, 1);
  ctx.endDate = ctx.term === "上期" ? new Date(y, 8, 30) : new Date(y + 1, 2, 31);
  ctx.startDStr = Utilities.formatDate(ctx.startDate, Session.getScriptTimeZone(), "yyyy/MM/dd");

  if (ctx.term === "上期") {
    ctx.p1Start = new Date(y, 3, 1);
    ctx.p1End = new Date(y, 5, 30);  ctx.p1TitleStr = "4~6";
    ctx.p2Start = new Date(y, 6, 1);
    ctx.p2End = new Date(y, 8, 30);  ctx.p2TitleStr = "7~9";
  } else {
    ctx.p1Start = new Date(y, 9, 1);
    ctx.p1End = new Date(y, 11, 31); ctx.p1TitleStr = "10~12";
    ctx.p2Start = new Date(y + 1, 0, 1);
    ctx.p2End = new Date(y + 1, 2, 31); ctx.p2TitleStr = "1~3";
  }

  // カレンダーキャッシュ生成
  let curDay = new Date(ctx.startDate.getTime());
  while (curDay <= ctx.endDate) {
    let dStr = Utilities.formatDate(curDay, "JST", "yyyy/MM/dd");
    let dN = ctx.dowNames[curDay.getDay()];
    let m = curDay.getMonth();
    let d = curDay.getDate();
    let isNY = (m === 11 && d >= 29) || (m === 0 && d <= 3);
    
    ctx.calendarCache.push({
      dateObj: new Date(curDay.getTime()),
      dStr: dStr,
      dN: dN,
      isHol: _debug_isTrueHoliday(dStr),
      isNY: isNY,
      wNum: Math.ceil(d / 7),
      getTime: curDay.getTime()
    });
    curDay.setDate(d + 1);
  }

  ctx.absences = _readFlatSheet(ctx.ss, "お休み情報", true);
  ctx.advances = _readFlatSheet(ctx.ss, "先行応募", false);
  ctx.substitutes = _readFlatSheet(ctx.ss, "振替勤務", false);
  ctx.kyukans = _readKyukanFlat();

  const masterUrl = 'https://docs.google.com/spreadsheets/d/1aEjphEv_63SeWQmwiOy9sx7IrMfawU01sHbKd_Ki4iA/edit';
  let masterSs;
  try { 
    masterSs = safeOpenByUrl(masterUrl);
  } catch (e) { 
    throw new Error("マスタへのアクセス権限がありません"); 
  }

  let masterDataCache = {};
  ["常勤", "定期非常勤"].forEach(type => {
    const sheet = masterSs.getSheetByName(`${type}${ctx.year}年度`);
    if (sheet) masterDataCache[type] = sheet.getDataRange().getValues();
  });

  ctx.locNames.forEach(locName => ctx.contractsByLoc[locName] = []);
  const locDict = typeof getLocationDictionary === "function" ? getLocationDictionary() : {};

  ["常勤", "定期非常勤"].forEach(type => {
    const data = masterDataCache[type];
    if (!data) return;
    const nameIdx = data[0].indexOf("医師名");
    const bikouIdx = data[0].indexOf("勤務備考");
    const subjIdx = data[0].findIndex(h => String(h).includes("専門") || String(h).includes("科目") || String(h).includes("診療科"));
    const contractWageIdx = data[0].indexOf("契約時給");
    const specialWageIdx = data[0].indexOf("特別時給の内訳");
    
    // ★ 退職日列のインデックスを取得
    const retireIdx = data[0].findIndex(h => String(h).includes("退職"));

    for (let r = 1; r < data.length; r++) {
      const docName = String(data[r][nameIdx]).replace(/先生$/, "").trim();
      const bikou = data[r][bikouIdx];
      const docSubject = subjIdx !== -1 ? String(data[r][subjIdx]).trim() : "";

      if (!docName || !bikou) continue;

      const contractType = contractWageIdx !== -1 ? String(data[r][contractWageIdx]) : "";
      const specialWageDetail = specialWageIdx !== -1 ? String(data[r][specialWageIdx]) : "";
      
      const parsedSlots = _parseComplexShiftText(bikou);

      // 1. 備考欄から期間を抽出する（★あらゆるハイフン・波ダッシュに対応）
      let periods = [];
      const periodRegex = /(\d{4})\/(\d{1,2})\/(\d{1,2})\s*[～~〜\-ー−]\s*(?:(\d{4})\/(\d{1,2})\/(\d{1,2}))?/g;
      let pMatch;
      while ((pMatch = periodRegex.exec(bikou)) !== null) {
        let fromDate = new Date(parseInt(pMatch[1]), parseInt(pMatch[2]) - 1, parseInt(pMatch[3]));
        let toDate = null;

        if (pMatch[4] && pMatch[5] && pMatch[6]) {
          toDate = new Date(parseInt(pMatch[4]), parseInt(pMatch[5]) - 1, parseInt(pMatch[6]), 23, 59, 59);
        }
        periods.push({ index: pMatch.index, from: fromDate, to: toDate });
      }

      // 2. マスタの「退職日」列をDateオブジェクトとして取得
      const retireVal = retireIdx !== -1 ? data[r][retireIdx] : "";
      let retireDateObj = null;
      if (retireVal instanceof Date && !isNaN(retireVal.getTime())) {
        retireDateObj = new Date(retireVal.getTime());
        retireDateObj.setHours(23, 59, 59, 999); 
      } else if (String(retireVal).trim() !== "") {
        let parsed = new Date(String(retireVal).trim());
        if (!isNaN(parsed.getTime())) {
          retireDateObj = parsed;
          retireDateObj.setHours(23, 59, 59, 999);
        }
      }

      // 3. 退職日がない場合のみ、最後の期間を「自動延長（無期限）」にする
      if (!retireDateObj && periods.length > 0) {
        periods[periods.length - 1].to = null;
      }

      for (let locKey in ctx.contractsByLoc) {
        let cleanLocName = locKey.split('_')[0].replace(/[（\(\)）]/g, "").replace(/内科|小児科/g, "").trim();
        let category = locKey.includes("内科") ? "内科" : "小児科";

        parsedSlots.forEach(pSlot => {
          let extractedLoc = pSlot.loc;
          let normalizedExtractedLoc = typeof normalizeLocationName === "function" ? normalizeLocationName(extractedLoc, locDict) : extractedLoc;
          
          if (!normalizedExtractedLoc.includes(cleanLocName) && !cleanLocName.includes(normalizedExtractedLoc)) return;

          if (cleanLocName === "亀有" || cleanLocName === "北葛西") {
            let otherCategory = category === "内科" ? "小児科" : "内科";
            let isSubjectMatch = docSubject.includes(category);
            let isOtherSubjectMatch = docSubject.includes(otherCategory);
            
            let isBoxMatch = extractedLoc.includes(category) || normalizedExtractedLoc.includes(category);
            let isOtherBoxMatch = extractedLoc.includes(otherCategory) || normalizedExtractedLoc.includes(otherCategory);
            
            if (isOtherBoxMatch && !isBoxMatch) return;
            if (!isBoxMatch && (!isSubjectMatch && (isOtherSubjectMatch || bikou.includes(otherCategory)))) return;
          }

          let validFrom = null;
          let validTo = null;
          let applicablePeriod = periods.slice().reverse().find(p => p.index < pSlot.originalIndex);
          if (applicablePeriod) {
            validFrom = applicablePeriod.from;
            validTo = applicablePeriod.to;
          }

          // 4. 【絶対ルール】退職日があれば、すべての期間の終了日を退職日で頭打ちにする
          if (retireDateObj) {
            if (validTo) {
              validTo = validTo < retireDateObj ? validTo : retireDateObj;
            } else {
              validTo = retireDateObj;
            }
          }
          
          ctx.contractsByLoc[locKey].push({ 
            docName: docName, 
            type: type, 
            freq: pSlot.freqStr, 
            dow: pSlot.dow, 
            sH: pSlot.sH, 
            eH: pSlot.eH, 
            isHolidayWork: pSlot.isHolidayWork, 
            isNewYearWork: pSlot.isNewYearWork, 
            weeks: pSlot.weeks, 
            bikou: bikou, 
            docSubject: docSubject,
            contractType: contractType, 
            specialWageDetail: specialWageDetail,
            validFrom: validFrom, 
            validTo: validTo 
          });
        });
      }
    }
  });
}

/**
 * どんなに汚い・複雑なテキストでも、正しく「拠点・週・曜日・時間・ルール」に分解する究極パーサー
 */
function _parseComplexShiftText(rawText) {
  let results = [];
  
  let isHolidayWork = true; 
  let isNewYearWork = true;
  let holidayMatch = rawText.match(/祝日[:：]?\s*勤務([あな])り/);
  if (holidayMatch) isHolidayWork = (holidayMatch[1] === "あ");
  let nyMatch = rawText.match(/年末年始[:：]?\s*勤務([あな])り/);
  if (nyMatch) isNewYearWork = (nyMatch[1] === "あ");

  let lines = rawText.split('\n');
  let cumulativeIndex = 0;
  
  lines.forEach(line => {
    let lineIndex = rawText.indexOf(line, cumulativeIndex);
    cumulativeIndex = lineIndex + line.length;
    
    let t = line.trim();
    if (!t || !t.includes("【")) return; 
    
    let locMatch = t.match(/【(.*?)】/);
    if (!locMatch) return;
    let location = locMatch[1];
    
    let weekMatch = t.match(/(毎週|[第1-5１-５・、,，\s]+)週?(月|火|水|木|金|土|日)曜?日?/);
    if (!weekMatch) return;
    
    let weekStr = weekMatch[1];
    let dayOfWeek = weekMatch[2];
    let targetWeeks = [];
    
    if (weekStr.includes("毎週")) {
      targetWeeks = [1, 2, 3, 4, 5];
    } else {
      let digits = weekStr.match(/\d/g);
      if (digits) {
        targetWeeks = digits.map(Number);
      }
    }
    
    // ★ 時間の抽出：波ダッシュ「〜」、全角マイナス「−」などを完全網羅
    let timeMatch = t.match(/(\d{1,2})[:：]?(\d{2})?\s*[～~〜\-ー−]\s*(\d{1,2})[:：]?(\d{2})?/);
    if (!timeMatch) return;
    
    let sH = parseInt(timeMatch[1], 10);
    let eH = parseInt(timeMatch[3], 10);
    if (isNaN(sH) || isNaN(eH)) return;
    
    results.push({
      loc: location,
      dow: dayOfWeek,
      weeks: targetWeeks,
      freqStr: targetWeeks.length === 5 ? "毎週" : `第${targetWeeks.join("・")}`,
      sH: sH,
      eH: eH,
      isHolidayWork: isHolidayWork,
      isNewYearWork: isNewYearWork,
      originalIndex: lineIndex
    });
  });
  
  return results;
}