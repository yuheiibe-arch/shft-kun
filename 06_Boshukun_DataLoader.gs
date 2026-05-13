/**
 * ====================================================================
 * 07_Boshukun_DataLoader.gs
 * カレンダー生成・マスタ読み込み・契約備考の解析エンジン
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
    ctx.p1Start = new Date(y, 3, 1);  ctx.p1End = new Date(y, 5, 30);  ctx.p1TitleStr = "4~6";
    ctx.p2Start = new Date(y, 6, 1);  ctx.p2End = new Date(y, 8, 30);  ctx.p2TitleStr = "7~9";
  } else {
    ctx.p1Start = new Date(y, 9, 1);  ctx.p1End = new Date(y, 11, 31); ctx.p1TitleStr = "10~12";
    ctx.p2Start = new Date(y + 1, 0, 1); ctx.p2End = new Date(y + 1, 2, 31); ctx.p2TitleStr = "1~3";
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
    masterSs = SpreadsheetApp.openByUrl(masterUrl); 
  } catch (e) { 
    throw new Error("マスタへのアクセス権限がありません"); 
  }

  let masterDataCache = {};
  ["常勤", "定期非常勤"].forEach(type => {
    const sheet = masterSs.getSheetByName(`${type}${ctx.year}年度`);
    if (sheet) masterDataCache[type] = sheet.getDataRange().getValues();
  });

  ctx.locNames.forEach(locName => ctx.contractsByLoc[locName] = []);

  // ★追加：表記ブレ吸収用の辞書を取得
  const locDict = typeof getLocationDictionary === "function" ? getLocationDictionary() : {};

  ["常勤", "定期非常勤"].forEach(type => {
    const data = masterDataCache[type];
    if (!data) return;
    const nameIdx = data[0].indexOf("医師名");
    const bikouIdx = data[0].indexOf("勤務備考");
    const subjIdx = data[0].findIndex(h => String(h).includes("専門") || String(h).includes("科目"));
    const contractWageIdx = data[0].indexOf("契約時給");
    const specialWageIdx = data[0].indexOf("特別時給の内訳");

    for (let r = 1; r < data.length; r++) {
      const docName = String(data[r][nameIdx]).replace(/先生$/, "").trim();
      const bikou = data[r][bikouIdx];
      const docSubject = subjIdx !== -1 ? String(data[r][subjIdx]).trim() : "";

      if (!docName || !bikou) continue;

      const contractType = contractWageIdx !== -1 ? String(data[r][contractWageIdx]) : "";
      const specialWageDetail = specialWageIdx !== -1 ? String(data[r][specialWageIdx]) : "";
      let isHolidayWork = !String(bikou).match(/祝日[:：]?勤務なし/);
      let isNewYearWork = !String(bikou).match(/年末年始[:：]?勤務なし/);

      // ★備考欄から「期間（YYYY/MM/DD～YYYY/MM/DD）」をすべて抽出し、その出現位置（index）を記録する
      let periods = [];
      const periodRegex = /(\d{4})\/(\d{1,2})\/(\d{1,2})\s*[～~-]\s*(\d{4})\/(\d{1,2})\/(\d{1,2})/g;
      let pMatch;
      while ((pMatch = periodRegex.exec(bikou)) !== null) {
        periods.push({
          index: pMatch.index,
          from: new Date(parseInt(pMatch[1]), parseInt(pMatch[2]) - 1, parseInt(pMatch[3])),
          to: new Date(parseInt(pMatch[4]), parseInt(pMatch[5]) - 1, parseInt(pMatch[6]), 23, 59, 59)
        });
      }

      for (let locKey in ctx.contractsByLoc) {
        let cleanLocName = locKey.split('_')[0].replace(/[（\(\)）]/g, "").replace(/内科|小児科/g, "").trim();
        let category = locKey.includes("内科") ? "内科" : "小児科";

        // ★修正：以前の足切り（if (!String(bikou).includes(cleanLocName)) continue;）は、表記ブレで弾かれる原因になるため削除しました。
        
        const regex = /【(.*?)】\s*(毎週|[第1-5１-５・、,，\s]+)週?(月|火|水|木|金|土|日)曜(?:日)?[:：]?\s*(\d{1,2}[:：]\d{2})\s*[～~-]\s*(\d{1,2}[:：]\d{2})/g;
        let match;
        
        while ((match = regex.exec(bikou)) !== null) {
          
          // ★修正：抽出した【 】の中身を正規化して判定
          let extractedLoc = match[1].trim();
          let normalizedExtractedLoc = typeof normalizeLocationName === "function" ? normalizeLocationName(extractedLoc, locDict) : extractedLoc;
          
          // ★修正：正規化された名前同士で比較
          if (!normalizedExtractedLoc.includes(cleanLocName) && !cleanLocName.includes(normalizedExtractedLoc)) continue;
          
          if (cleanLocName === "亀有" || cleanLocName === "北葛西") {
            let otherCategory = category === "内科" ? "小児科" : "内科";
            let isSubjectMatch = docSubject.includes(category);
            let isOtherSubjectMatch = docSubject.includes(otherCategory);
            // ★修正：正規化された名前も加味してカテゴリを判定
            let isBoxMatch = extractedLoc.includes(category) || normalizedExtractedLoc.includes(category);
            let isOtherBoxMatch = extractedLoc.includes(otherCategory) || normalizedExtractedLoc.includes(otherCategory);
            
            if (isOtherBoxMatch && !isBoxMatch) continue;
            if (!isBoxMatch && (!isSubjectMatch && (isOtherSubjectMatch || bikou.includes(otherCategory)))) continue;
          }

          // ★このシフトの記述位置より「前」にある直近の日付期間を探す（竹下先生対応）
          let validFrom = null;
          let validTo = null;
          let applicablePeriod = periods.slice().reverse().find(p => p.index < match.index);
          if (applicablePeriod) {
            validFrom = applicablePeriod.from;
            validTo = applicablePeriod.to;
          }

          const freqStr = match[2];
          const dow = match[3];
          const sH = parseInt(match[4].split(/[:：]/)[0], 10);
          const eH = parseInt(match[5].split(/[:：]/)[0], 10);

          let contractWeeks = [];
          if (freqStr.includes("毎週")) {
            contractWeeks = [1, 2, 3, 4, 5];
          } else {
            let m = freqStr.match(/[1-5１-５]/g);
            if (m) contractWeeks = m.map(v => parseInt(v.replace(/[１-５]/g, function(s) { return String.fromCharCode(s.charCodeAt(0) - 0xFEE0); }), 10));
          }
          
          ctx.contractsByLoc[locKey].push({ 
            docName: docName, type: type, freq: freqStr.trim(), dow: dow, 
            sH: sH, eH: eH, isHolidayWork: isHolidayWork, isNewYearWork: isNewYearWork, 
            weeks: contractWeeks, bikou: bikou, docSubject: docSubject,
            contractType: contractType, specialWageDetail: specialWageDetail,
            validFrom: validFrom, validTo: validTo // ★抽出した期間を保持
          });
        }
      }
    }
  });
}