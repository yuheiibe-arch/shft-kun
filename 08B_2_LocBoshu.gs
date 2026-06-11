
/**
 * ====================================================================
 * 08B_2_LocBoshu.gs
 * 定期枠の空き計算と、不足分の募集（定期募集・単発募集）への出力
 * ====================================================================
 */

function _processLocBoshu(ctx, locCtx) {
  let { cleanLocName, category, displayLoc, area, actualStartDate, activeContracts, locDicts } = locCtx;

  if (cleanLocName === "MQC") return;

  let weeklyGrid = {};
  ctx.dowNames.forEach(dow => {
    weeklyGrid[dow] = new Array(12).fill(null).map(() => new Set([1, 2, 3, 4, 5]));
    weeklyGrid[dow][4].clear(); 
    weeklyGrid[dow][5].clear(); 
    if (cleanLocName === "北葛西") weeklyGrid[dow][11].clear();
  });

  activeContracts.forEach(c => {
    for (let h = c.sH; h < c.eH; h++) {
      let idx = h - 9;
      if (idx >= 0 && idx < 12 && idx !== 4 && idx !== 5) {
        c.weeks.forEach(w => weeklyGrid[c.dow][idx].delete(w));
      }
    }
  });

  ctx.dowNames.forEach(dow => {
    _extractWeeklyBlocks(weeklyGrid[dow]).forEach(b => {
      const wage = _getWageWrapper(ctx.startDStr, cleanLocName, category, dow, b.sH, b.eH);
      const dailyCost = _getDailyCost(ctx.startDStr, cleanLocName, category, dow, b.sH, b.eH);
      
      const pushBoshuChunk = (sDate, eDate, tSuffix) => {
        if (eDate < actualStartDate) return; 
        let chunkStart = (sDate < actualStartDate) ? actualStartDate : sDate;
        let pStr = `${Utilities.formatDate(chunkStart, "JST", "yyyy/MM/dd")}～${Utilities.formatDate(eDate, "JST", "yyyy/MM/dd")}`;
        
        let targetDaysCount = 0;
        let coveredDaysCount = 0; 
        let emptyDays = []; 
        let holArr = [];
        let irregArr = []; 
        
        ctx.calendarCache.filter(c => c.getTime >= chunkStart.getTime() && c.getTime <= eDate.getTime()).forEach(cDay => {
          if (cDay.dN === dow && b.weeksArr.includes(cDay.wNum)) {
            targetDaysCount++;
            let isCoveredToday = false; 
            
            if (cDay.isHol) holArr.push(`${cDay.dateObj.getMonth() + 1}/${cDay.dateObj.getDate()}(${dow})`);
            
            // 重いfilterループを、生成済みの辞書参照に置き換え（ロジックは完全同一）
            let advs = (locDicts.advances[cDay.dStr] || []).filter(adv => adv.sH < b.eH && adv.eH > b.sH);
            advs.forEach(adv => {
              isCoveredToday = true; 
              let baseTimeStr = `${('0'+adv.sH).slice(-2)}:00-${('0'+adv.eH).slice(-2)}:00`;
              irregArr.push(`${adv.docName}先生|||先行|||${baseTimeStr}|||${cDay.dateObj.getMonth() + 1}/${cDay.dateObj.getDate()}(${dow})`);
              if (cDay.isHol) {
                irregArr.push(`${adv.docName}先生|||先行|||（期間内の祝日）|||${cDay.dateObj.getMonth() + 1}/${cDay.dateObj.getDate()}(${dow})`);
              }
            });

            let subs = (locDicts.substitutes[cDay.dStr] || []).filter(sub => sub.sH < b.eH && sub.eH > b.sH);
            subs.forEach(sub => {
              isCoveredToday = true; 
              let baseTimeStr = `${('0'+sub.sH).slice(-2)}:00-${('0'+sub.eH).slice(-2)}:00`;
              irregArr.push(`${sub.docName}先生|||振替|||${baseTimeStr}|||${cDay.dateObj.getMonth() + 1}/${cDay.dateObj.getDate()}(${dow})`);
              if (cDay.isHol) {
                irregArr.push(`${sub.docName}先生|||振替|||（期間内の祝日）|||${cDay.dateObj.getMonth() + 1}/${cDay.dateObj.getDate()}(${dow})`);
              }
            });

            if (isCoveredToday) {
              coveredDaysCount++; 
            } else {
              emptyDays.push(cDay);
            }
          }
        });
        
        if (targetDaysCount === 0 || targetDaysCount === coveredDaysCount) {
          return; 
        }

        let coveredRate = coveredDaysCount / targetDaysCount;
        if (coveredRate >= 0.8) {
          emptyDays.forEach(cDay => {
            _splitTimeIntoBlocks(b.sH, b.eH).forEach(blk => {
              const sig = `${cleanLocName}_${cDay.dStr}_${blk.sH}_${blk.eH}`;
              if (!ctx.pushedSingles.has(sig)) {
                ctx.tempSingles.push({ 
                  area: area, loc: displayLoc, cleanLoc: cleanLocName, cat: category, 
                  reason: "定期枠（8割以上埋まり）の残り分募集", dateObj: new Date(cDay.getTime), 
                  dow: cDay.dN, sH: blk.sH, eH: blk.eH 
                });
                ctx.pushedSingles.add(sig);
              }
            });
          });
          return; 
        }

        const totalHours = b.hours * targetDaysCount;
        let holidaysStr = (dow !== "土" && dow !== "日" && holArr.length > 0) ? holArr.join(",") : "";
        let holWage = (holidaysStr !== "") ? _getHolidayWageWrapper(ctx.startDStr, cleanLocName, category, b.sH, b.eH) : "";
        let irregStr = irregArr.length > 0 ? irregArr.join("###") : ""; 

        ctx.masterRegularList.push({
          "エリア": area, "期間": pStr, "開始時間": b.sT, "終了時間": b.eT, 
          "時給": wage, "祝日時給": holWage, "祝日該当日": holidaysStr, 
          "募集時間": totalHours, 
          "コスト": dailyCost > 0 ? dailyCost * targetDaysCount : "", 
          "対応済": false, "先行・振替": irregStr,
          _loc: displayLoc, _sH: b.sH, _eH: b.eH, _freqStr: b.freqStr, _dow: dow, _tSuffix: tSuffix
        });

        if (cleanLocName === "北葛西" && b.sH === 18 && b.eH === 20) {
          let extra_sH = 17;
          let extra_eH = 20;
          let extra_wage = _getWageWrapper(ctx.startDStr, cleanLocName, category, dow, extra_sH, extra_eH);
          let extra_dailyCost = _getDailyCost(ctx.startDStr, cleanLocName, category, dow, extra_sH, extra_eH);
          let extra_holWage = (holidaysStr !== "") ? _getHolidayWageWrapper(ctx.startDStr, cleanLocName, category, extra_sH, extra_eH) : "";
          let extra_totalHours = 3 * targetDaysCount;
          
          ctx.masterRegularList.push({
            "エリア": area, "期間": pStr, "開始時間": "17:00", "終了時間": "20:00", 
            "時給": extra_wage, "祝日時給": extra_holWage, "祝日該当日": holidaysStr, 
            "募集時間": extra_totalHours, 
            "コスト": extra_dailyCost > 0 ? extra_dailyCost * targetDaysCount : "", 
            "対応済": false, "先行・振替": irregStr,
            _loc: displayLoc, _sH: extra_sH, _eH: extra_eH, _freqStr: b.freqStr, _dow: dow, _tSuffix: tSuffix
          });
        }
      };

      pushBoshuChunk(ctx.p1Start, ctx.p1End, ctx.p1TitleStr);
      pushBoshuChunk(ctx.p2Start, ctx.p2End, ctx.p2TitleStr);
    });
  });
}