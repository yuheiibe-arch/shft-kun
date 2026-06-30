/**
 * ====================================================================
 * 08B_4_LocCancel.gs
 * 欠勤、休館日のキャンセル処理と、追加の単独募集枠の生成
 * ====================================================================
 */

function _processLocCancel(ctx, locCtx) {
  let { cleanLocName, category, displayLoc, area, actualStartDate, locCalendar, activeContracts, locDicts } = locCtx;
  let locAbsences = [];
  Object.values(locDicts.absences).forEach(arr => locAbsences.push(...arr));
  
  locAbsences.forEach(a => {
    let overlappingBlocks = [];
    let foundContract = false;

    activeContracts.forEach(c => {
      let cDay = ctx.calendarCache.find(day => day.getTime === a.dateObj.getTime());
      if (cDay && c.docName === a.docName && _isContractActiveOnCache(c, cDay)) {
        foundContract = true;
        let oSH = Math.max(a.sH, c.sH);
        let oEH = Math.min(a.eH, c.eH);
        if (oSH < oEH) overlappingBlocks.push({sH: oSH, eH: oEH});
      }
    });

    if (!foundContract) overlappingBlocks.push({sH: a.sH, eH: a.eH});

    overlappingBlocks.forEach(blk => {
      ctx.cancelList.push({
        "医師名": `${a.docName}先生`, "該当日": Utilities.formatDate(a.dateObj, "JST", "yyyy年MM月dd日"), "理由": a.reason, 
        "対象拠点": displayLoc, "対象勤務時間": `${('0'+blk.sH).slice(-2)}:00〜${('0'+blk.eH).slice(-2)}:00`, "募集シフト作成指示": "済", "対応済": false, "対応者": "", "GASチェック": false
      });
      
      if (cleanLocName !== "MQC") {
        // ★ 修正：分割ルールに cleanLocName を渡して拠点情報を引き継ぐ
        _splitTimeIntoBlocks(blk.sH, blk.eH, cleanLocName).forEach(b => {
           const sig = `${cleanLocName}_${Utilities.formatDate(a.dateObj, "JST", "yyyy/MM/dd")}_${b.sH}_${b.eH}`;
           if (!ctx.pushedSingles.has(sig)) {
             ctx.tempSingles.push({ area: area, loc: displayLoc, cleanLoc: cleanLocName, cat: category, reason: "欠勤等による追加募集", dateObj: a.dateObj, dow: ctx.dowNames[a.dateObj.getDay()], sH: b.sH, eH: b.eH });
             ctx.pushedSingles.add(sig);
           }
        });
      }
    });
  });

  let locKyukans = [];
  Object.values(locDicts.kyukans).forEach(arr => locKyukans.push(...arr));
  
  locKyukans.forEach(k => {
    activeContracts.forEach(c => {
      let cDay = ctx.calendarCache.find(day => day.getTime === k.dateObj.getTime());
      if (cDay && _isContractActiveOnCache(c, cDay)) {
        ctx.cancelList.push({
          "医師名": `${c.docName}先生`, "该当日": Utilities.formatDate(k.dateObj, "JST", "yyyy年MM月dd日"), "理由": "休館日", 
          "対象拠点": displayLoc, "対象勤務時間": `${('0'+c.sH).slice(-2)}:00〜${('0'+c.eH).slice(-2)}:00`, "募集シフト作成指示": "", "対応済": false, "対応者": "", "GASチェック": false
        });
      }
    });
  });

  activeContracts.forEach(c => {
    locCalendar.forEach(cDay => {
      if (cDay.isHol && _isContractActiveOnCache(c, cDay)) {
        let isWork = cDay.isNY ? c.isNewYearWork : c.isHolidayWork;
        if (!isWork && cleanLocName !== "MQC") { 
          // ★ 修正：分割ルールに cleanLocName を渡して拠点情報を引き継ぐ
          _splitTimeIntoBlocks(c.sH, c.eH, cleanLocName).forEach(b => {
            const sig = `${cleanLocName}_${cDay.dStr}_${b.sH}_${b.eH}`;
            if (!ctx.pushedSingles.has(sig)) {
              ctx.tempSingles.push({ area: area, loc: displayLoc, cleanLoc: cleanLocName, cat: category, reason: "祝日・年末年始", dateObj: new Date(cDay.getTime), dow: cDay.dN, sH: b.sH, eH: b.eH });
              ctx.pushedSingles.add(sig);
            }
          });
        }
      }
      
      let isMatchDowAndWeek = false;
      if (c.dow === cDay.dN && (c.weeks.length === 5 || c.weeks.includes(cDay.wNum))) {
        isMatchDowAndWeek = true;
      }
      
      if (isMatchDowAndWeek) {
        let isOut = false;
        if (c.validFrom && cDay.getTime < c.validFrom.getTime()) isOut = true;
        if (c.validTo && cDay.getTime > c.validTo.getTime()) isOut = true;

        if (isOut && cleanLocName !== "MQC") {
           // ★ 修正：分割ルールに cleanLocName を渡して拠点情報を引き継ぐ
           _splitTimeIntoBlocks(c.sH, c.eH, cleanLocName).forEach(b => {
              let isCovered = activeContracts.some(otherC => {
                 if (_isContractActiveOnCache(otherC, cDay)) {
                    if (otherC.sH <= b.sH && otherC.eH >= b.eH) return true;
                 }
                 return false;
              });
              
              if (!isCovered) {
                const sig = `${cleanLocName}_${cDay.dStr}_${b.sH}_${b.eH}`;
                if (!ctx.pushedSingles.has(sig)) {
                  ctx.tempSingles.push({ area: area, loc: displayLoc, cleanLoc: cleanLocName, cat: category, reason: "契約期間外による追加募集", dateObj: new Date(cDay.getTime), dow: cDay.dN, sH: b.sH, eH: b.eH });
                  ctx.pushedSingles.add(sig);
                }
              }
           });
        }
      }
    });
  });
}