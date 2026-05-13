/**
 * ====================================================================
 * 08B_3_LocConfirm.gs
 * 確定シフトへの書き出し、2診判定、お休み情報の適用
 * ★定期シフト・先行応募の二重登録防止ロック機能つき
 * ====================================================================
 */

function _processLocConfirm(ctx, locCtx) {
  let { cleanLocName, category, displayLoc, area, actualStartDate, activeContracts, locDicts, dailyBusyMap } = locCtx;

  // 1. 定期契約のマージと出力
  let mergedMap = new Map();
  activeContracts.forEach(c => {
    let vFrom = c.validFrom ? c.validFrom.getTime() : 0;
    let vTo = c.validTo ? c.validTo.getTime() : 0;
    let key = `${c.docName}_${c.type}_${vFrom}_${vTo}`;
    if (mergedMap.has(key)) mergedMap.get(key).shifts.push(c);
    else mergedMap.set(key, { docName: c.docName, type: c.type, validFrom: c.validFrom, validTo: c.validTo, bikou: c.bikou, isHolidayWork: c.isHolidayWork, isNewYearWork: c.isNewYearWork, shifts: [c] });
  });

  let mergedContracts = Array.from(mergedMap.values());
  let dowOrder = ["月", "火", "水", "木", "金", "土", "日"];
  mergedContracts.forEach(mc => mc.shifts.sort((a, b) => dowOrder.indexOf(a.dow) - dowOrder.indexOf(b.dow)));

  mergedContracts.forEach(mc => {
    let shiftsByDow = {};
    mc.shifts.forEach(sh => {
      if (!shiftsByDow[sh.dow]) shiftsByDow[sh.dow] = [];
      shiftsByDow[sh.dow].push(sh);
    });

    for (let dowKey in shiftsByDow) {
      let dowShifts = shiftsByDow[dowKey];

      const generateCautionStr = (sDate, eDate) => {
        let workCautionArr = [];
        let cancelCautionArr = [];
        let absenceCautionArr = []; 
        
        ctx.calendarCache.filter(day => day.getTime >= sDate.getTime() && day.getTime <= eDate.getTime()).forEach(cDay => {
          let activeShifts = dowShifts.filter(sh => _isContractActiveOnCache(sh, cDay));
          
          activeShifts.forEach(sh => {
            if (cDay.isHol) {
              let isWork = cDay.isNY ? sh.isNewYearWork : sh.isHolidayWork;
              let dStrClean = cDay.dStr || Utilities.formatDate(cDay.dateObj, "JST", "yyyy/MM/dd");
              
              let dummySh = Object.assign({}, sh);
              dummySh.specialWageDetail = ""; dummySh.contractType = "";

              if (isWork) {
                if (mc.type !== "常勤") {
                  let hWageStr = cleanLocName !== "MQC" ? _getFinalContractWage(dStrClean, cleanLocName, category, dummySh) : `${('0'+sh.sH).slice(-2)}:00-${('0'+sh.eH).slice(-2)}:00`;
                  let formattedWage = String(hWageStr).split('\n').map(line => line.replace(/-/g, '~').replace(/\//g, '：').replace(/¥/g, '￥')).join('\n');
                  workCautionArr.push(`${Utilities.formatDate(cDay.dateObj, "JST", "yyyy年MM月dd日")}（${cDay.dN}）祝日時給\n${formattedWage}`);
                }
              } else {
                let cWageStr = "";
                if (cleanLocName === "MQC") {
                  cWageStr = `${('0'+sh.sH).slice(-2)}:00-${('0'+sh.eH).slice(-2)}:00`;
                } else if (mc.type === "常勤") {
                  cWageStr = _getHolidayWageWrapper(dStrClean, cleanLocName, category, sh.sH, sh.eH);
                  if (!cWageStr) cWageStr = _getWageWrapper(dStrClean, cleanLocName, category, cDay.dN, sh.sH, sh.eH);
                } else {
                  cWageStr = _getFinalContractWage(dStrClean, cleanLocName, category, dummySh);
                }
                
                let formattedWage = cWageStr ? String(cWageStr).split('\n').map(line => line.replace(/-/g, '~').replace(/\//g, '：').replace(/¥/g, '￥')).join('\n') : "";
                cancelCautionArr.push(`${Utilities.formatDate(cDay.dateObj, "JST", "yyyy年MM月dd日")}（${cDay.dN}）祝日時給\n${formattedWage}`);
              }
            }
          });

          if (activeShifts.length > 0) {
            let dStrClean = cDay.dStr || Utilities.formatDate(cDay.dateObj, "JST", "yyyy/MM/dd");
            let dayAbsences = locDicts.absences[cDay.dStr] || locDicts.absences[dStrClean] || [];
            
            dayAbsences.filter(a => a.docName === mc.docName).forEach(abs => {
              let absWageStr = "";
              if (cleanLocName === "MQC") {
                absWageStr = `${('0'+abs.sH).slice(-2)}:00-${('0'+abs.eH).slice(-2)}:00`;
              } else if (mc.type === "常勤") {
                if (cDay.isHol) {
                  absWageStr = _getHolidayWageWrapper(dStrClean, cleanLocName, category, abs.sH, abs.eH);
                  if (!absWageStr) absWageStr = _getWageWrapper(dStrClean, cleanLocName, category, cDay.dN, abs.sH, abs.eH);
                } else {
                  absWageStr = _getWageWrapper(dStrClean, cleanLocName, category, cDay.dN, abs.sH, abs.eH);
                }
              } else {
                let dummySh = { sH: abs.sH, eH: abs.eH, dow: cDay.dN, specialWageDetail: "", contractType: "" };
                absWageStr = _getFinalContractWage(dStrClean, cleanLocName, category, dummySh);
              }
              
              let formattedAbsWage = absWageStr ? String(absWageStr).split('\n').map(line => line.replace(/-/g, '~').replace(/\//g, '：').replace(/¥/g, '￥')).join('\n') : "";
              let absStr = `${Utilities.formatDate(cDay.dateObj, "JST", "yyyy年MM月dd日")}（${cDay.dN}）：${abs.reason}\n${formattedAbsWage}`;
              
              if (!absenceCautionArr.includes(absStr)) absenceCautionArr.push(absStr);
            });
          }
        });
        
        let res = [];
        if (workCautionArr.length > 0) res.push(workCautionArr.join("\n\n"));
        if (cancelCautionArr.length > 0) res.push("【シフトキャンセル該当日（祝日・年末年始）】\n" + cancelCautionArr.join("\n\n")); 
        if (absenceCautionArr.length > 0) res.push("【お休み情報】\n" + absenceCautionArr.join("\n\n")); 
        return res.join("\n\n");
      };

      const pushTeikiConfirm = (sDate, eDate, tSuffix) => {
        let contractStart = mc.validFrom && mc.validFrom > sDate ? mc.validFrom : sDate;
        let contractEnd = mc.validTo && mc.validTo < eDate ? mc.validTo : eDate;
        if (contractEnd < actualStartDate || contractStart > contractEnd) return; 

        let chunkStart = contractStart < actualStartDate ? actualStartDate : contractStart;
        let pStr = `${Utilities.formatDate(chunkStart, "JST", "yyyy/MM/dd")}～${Utilities.formatDate(contractEnd, "JST", "yyyy/MM/dd")}`;
        let title = mc.type === "常勤" ? `${mc.docName}先生／${ctx.term}` : `${mc.docName}先生／${tSuffix}`;
        let cautionStr = generateCautionStr(chunkStart, contractEnd);

        let dowStr = dowShifts.map(sh => `${sh.freq}${sh.dow}曜日`).join('\n');
        let timeStr = dowShifts.map(sh => `${('0'+sh.sH).slice(-2)}:00-${('0'+sh.eH).slice(-2)}:00`).join('\n');
        
        let wageStr = "";
        if (mc.type !== "常勤") {
          let wageArr = dowShifts.map(sh => {
            let firstWorkingDay = ctx.calendarCache.find(day => day.getTime >= chunkStart.getTime() && day.getTime <= contractEnd.getTime() && _isContractActiveOnCache(sh, day));
            let representativeDateStr = firstWorkingDay ? firstWorkingDay.dStr : Utilities.formatDate(chunkStart, "JST", "yyyy/MM/dd");
            return _getFinalContractWage(representativeDateStr, cleanLocName, category, sh);
          });
          wageStr = [...new Set(wageArr)].join('\n');
        }
          
        ctx.confirmList.push({
          "エリア": area, "拠点名": displayLoc, "種別": mc.type, "医師名": `${mc.docName}先生`, "シフトタイトル": title, 
          "設定期間": pStr, "設定時間": timeStr, "設定曜日": dowStr, "時給": wageStr, "契約内容": mc.bikou, "注意箇所": cautionStr, "対応済": false,
          _sortDate: chunkStart.getTime(), _dow: dowKey 
        });
      };

      if (mc.type === "常勤") pushTeikiConfirm(actualStartDate, ctx.endDate, ctx.term);
      else { pushTeikiConfirm(ctx.p1Start, ctx.p1End, ctx.p1TitleStr); pushTeikiConfirm(ctx.p2Start, ctx.p2End, ctx.p2TitleStr); }
    }
  });

  const processSpotItem = (a, isBundle) => {
    let repDateObj = isBundle ? a.firstDate : a.dateObj;
    let dStrA = a.dStr || Utilities.formatDate(repDateObj, "JST", "yyyy/MM/dd");
    let dN = a.dow || ctx.dowNames[repDateObj.getDay()];
    
    let docNameClean = String(a.docName).replace(/先生$/, "").replace(/\s+/g, "").trim();
    let isJokin = ctx.jokinDocNames.has(docNameClean);
    let isHijokin = ctx.hijokinDocNames.has(docNameClean);
    let docMaster = ctx.docMasterData[docNameClean] || { specialWageDetail: "", contractType: "" };

    let typeName = "スポット";
    let shiftTitle = isBundle ? `【${a.typeName}】${displayLoc}／${('0'+a.sH).slice(-2)}:00-${('0'+a.eH).slice(-2)}:00／毎週${a.dow}曜` : `${a.docName}先生／${a.typeName}`;
    let cautionArr = [];

    if (a.typeName.includes("先行")) {
       typeName = isJokin ? "常勤（先行）" : (isHijokin ? "定期非常勤（先行）" : "スポット");
       cautionArr.push(isJokin ? "■業務内容：所定休出／先行応募\n■スタッフコメント：所定休出／先行応募" : "■業務内容：先行応募\n■スタッフコメント：先行応募");
    } else if (a.typeName.includes("振替")) {
       typeName = isJokin ? "常勤（振替）" : (isHijokin ? "定期非常勤（振替）" : "スポット");
       if ((a.isDaishin === true || String(a.isDaishin).toUpperCase() === "TRUE") && a.requester) {
         let reqName = String(a.requester).replace(/先生$/, "").trim() + "先生";
         shiftTitle = isBundle ? `【${reqName}の代診勤務】${displayLoc}／${('0'+a.sH).slice(-2)}:00-${('0'+a.eH).slice(-2)}:00／毎週${a.dow}曜` : `${a.docName}先生／${reqName}の代診勤務`;
         cautionArr.push(`■業務内容：${reqName}／代診勤務\n■スタッフコメント：${reqName}／代診勤務`);
       }
    }

    let absenceCautionArr = []; 
    let twoDrCautionArr = [];
    
    let targetDates = isBundle ? a.grp.map(d => d.dStr || Utilities.formatDate(d.dateObj, "JST", "yyyy/MM/dd")) : [dStrA];
    
    targetDates.forEach(targetDateStr => {
       let dailyOverlap = [];
       let busyList = dailyBusyMap[targetDateStr] || [];
       
       busyList.forEach(busy => {
         if (busy.docName !== a.docName) {
           let oSH = Math.max(a.sH, busy.sH);
           let oEH = Math.min(a.eH, busy.eH);
           if (oSH < oEH) dailyOverlap.push(`${('0'+oSH).slice(-2)}:00-${('0'+oEH).slice(-2)}:00`);
         }
       });

       if (dailyOverlap.length > 0) {
         let uniqueBlocks = [...new Set(dailyOverlap)];
         if (isBundle) twoDrCautionArr.push(`・${targetDateStr}: ${uniqueBlocks.join(", ")}`);
         else cautionArr.push(`【２診時間あり】${uniqueBlocks.join(", ")}`);
       }

       let isHol = _debug_isTrueHoliday(targetDateStr);

       if (isHol) { 
         let hWageStr = "";
         if (isJokin) {
           hWageStr = _getHolidayWageWrapper(targetDateStr, cleanLocName, category, a.sH, a.eH);
           if (!hWageStr) hWageStr = _getWageWrapper(targetDateStr, cleanLocName, category, dN, a.sH, a.eH);
         } else {
           let dummyShForHol = { docName: a.docName, dow: dN, sH: a.sH, eH: a.eH, specialWageDetail: docMaster.specialWageDetail, contractType: docMaster.contractType };
           hWageStr = _getFinalContractWage(targetDateStr, cleanLocName, category, dummyShForHol);
         }
         let formattedWage = hWageStr ? String(hWageStr).split('\n').map(line => line.replace(/-/g, '~').replace(/\//g, '：').replace(/¥/g, '￥')).join('\n') : "";
         cautionArr.push(`${targetDateStr}（${dN}）祝日時給\n${formattedWage}`);
       }

       let dayAbsences = locDicts.absences[targetDateStr] || [];
       dayAbsences.filter(abs => abs.docName === a.docName).forEach(abs => {
         let absWageStr = "";
         if (cleanLocName === "MQC") {
           absWageStr = `${('0'+abs.sH).slice(-2)}:00-${('0'+abs.eH).slice(-2)}:00`;
         } else if (isJokin) {
           if (isHol) {
             absWageStr = _getHolidayWageWrapper(targetDateStr, cleanLocName, category, abs.sH, abs.eH);
             if (!absWageStr) absWageStr = _getWageWrapper(targetDateStr, cleanLocName, category, dN, abs.sH, abs.eH);
           } else {
             absWageStr = _getWageWrapper(targetDateStr, cleanLocName, category, dN, abs.sH, abs.eH);
           }
         } else {
           let dummySh = { sH: abs.sH, eH: abs.eH, dow: dN, specialWageDetail: docMaster.specialWageDetail, contractType: docMaster.contractType, docName: a.docName };
           absWageStr = _getFinalContractWage(targetDateStr, cleanLocName, category, dummySh);
         }
         let formattedAbsWage = absWageStr ? String(absWageStr).split('\n').map(line => line.replace(/-/g, '~').replace(/\//g, '：').replace(/¥/g, '￥')).join('\n') : "";
         let aStr = `${targetDateStr}（${dN}）：${abs.reason}\n${formattedAbsWage}`;
         isBundle ? absenceCautionArr.push(aStr) : cautionArr.push("【お休み情報】\n" + aStr);
       });
    });

    if (isBundle) {
      if (twoDrCautionArr.length > 0) cautionArr.push("【２診時間あり】\n" + twoDrCautionArr.join("\n")); 
      if (absenceCautionArr.length > 0) cautionArr.push("【お休み情報】\n" + absenceCautionArr.join("\n\n")); 
    }

    let spotWageStr = !isJokin ? _getFinalContractWage(dStrA, cleanLocName, category, { docName: a.docName, dow: a.dow || dN, sH: a.sH, eH: a.eH, specialWageDetail: docMaster.specialWageDetail, contractType: docMaster.contractType }) : "";
    let periodStr = isBundle ? `${Utilities.formatDate(a.firstDate, "JST", "yyyy/MM/dd")}～${Utilities.formatDate(a.lastDate, "JST", "yyyy/MM/dd")}` : dStrA;

    ctx.confirmList.push({
      "エリア": area, "拠点名": displayLoc, "種別": typeName, "医師名": `${a.docName}先生`, "シフトタイトル": shiftTitle, 
      "設定期間": periodStr, "設定時間": `${('0'+a.sH).slice(-2)}:00-${('0'+a.eH).slice(-2)}:00`, 
      "設定曜日": isBundle ? `毎週${a.dow}曜日` : dN, "時給": spotWageStr, "契約内容": "", "注意箇所": cautionArr.join("\n\n"), "対応済": false,
      _sortDate: repDateObj.getTime()
    });
  };

  // =========================================================
  // ★追加：重複ロック機能（定期シフトと被っている先行・振替を自動除外）
  // =========================================================
  const getUniqueSpotShifts = (dictData) => {
    let results = [];
    Object.values(dictData).forEach(arr => {
      arr.forEach(spot => {
        let cDay = ctx.calendarCache.find(day => day.getTime === spot.dateObj.getTime());
        // 定期契約(activeContracts)に、同じ医師で同じ日のシフトが存在し、かつ時間が被っているかをチェック
        let isDuplicated = cDay && activeContracts.some(c => 
          c.docName === spot.docName && 
          _isContractActiveOnCache(c, cDay) && 
          Math.max(spot.sH, c.sH) < Math.min(spot.eH, c.eH) // 時間の重なり判定
        );
        
        // 被っていなければリストに加える（被っていればここで破棄される＝定期が優先される）
        if (!isDuplicated) {
          results.push(spot);
        }
      });
    });
    return results;
  };

  // 先行応募リストの生成（ロック機能付き）
  let flatAdvances = getUniqueSpotShifts(locDicts.advances);
  let advProcessed = typeof _processIrregularShifts === 'function' ? _processIrregularShifts(flatAdvances, "先行応募") : { singles: flatAdvances, bundled: [] };
  if(advProcessed.singles) advProcessed.singles.forEach(a => processSpotItem(a, false));
  if(advProcessed.bundled) advProcessed.bundled.forEach(b => processSpotItem(b, true));

  // 振替勤務リストの生成（ロック機能付き）
  let flatSubstitutes = getUniqueSpotShifts(locDicts.substitutes);
  let subProcessed = typeof _processIrregularShifts === 'function' ? _processIrregularShifts(flatSubstitutes, "振替勤務") : { singles: flatSubstitutes, bundled: [] };
  if(subProcessed.singles) subProcessed.singles.forEach(a => processSpotItem(a, false));
  if(subProcessed.bundled) subProcessed.bundled.forEach(b => processSpotItem(b, true));
}