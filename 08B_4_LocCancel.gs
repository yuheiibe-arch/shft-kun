/**
 * ====================================================================
 * 08C_SinglesProcessor.gs
 * 契約外・祝日・お休みなどで発生した追加募集（単発）のグループ化と登録
 * ★ 厳格な「毎週」判定対応版
 * ====================================================================
 */

function _processGroupedSingles(ctx) {
  let groupedSingles = {};
  ctx.tempSingles.forEach(s => {
    const key = `${s.loc}_${s.reason}_${s.sH}-${s.eH}_${s.dow}`;
    if (!groupedSingles[key]) groupedSingles[key] = [];
    groupedSingles[key].push(s);
  });
  
  for (const key in groupedSingles) {
    const grp = groupedSingles[key];
    grp.sort((a,b) => a.dateObj - b.dateObj);
    const first = grp[0];
    const last = grp[grp.length - 1];
    
    // ★ 修正: 確実に「毎週」かどうかを判定するため、1ヶ月あたりの出現回数を計算
    let monthMapCounter = new Set();
    grp.forEach(g => monthMapCounter.add(g.dateObj.getFullYear() + "-" + g.dateObj.getMonth()));
    let avgPerMonth = grp.length / monthMapCounter.size;
    
    // 月平均3.5回以上出現し、かつすべて同じ曜日なら「毎週」とみなす（それ以外は弾く）
    let isWeekly = (grp.length >= 3 && avgPerMonth >= 3.5 && grp.every(g => g.dow === first.dow));
    let repeatDow = isWeekly ? `毎週${first.dow}曜日` : (grp.length > 1 ? "複数日程" : "単独");
    
    let monthMap = {};
    grp.forEach(g => {
      let m = g.dateObj.getMonth() + 1;
      let d = g.dateObj.getDate();
      if (!monthMap[m]) monthMap[m] = [];
      monthMap[m].push(`${d}(${g.dow})`);
    });
    
    let datesStrArr = [];
    Object.keys(monthMap).map(Number).sort((a, b) => a - b).forEach(m => {
      datesStrArr.push(`${m}月：${monthMap[m].join(', ')}`);
    });
    const datesStr = datesStrArr.join('\n');
    
    const wage = _getWageWrapper(Utilities.formatDate(first.dateObj, "JST", "yyyy/MM/dd"), first.cleanLoc, first.cat, first.dow, first.sH, first.eH);
    if (!wage || wage === "") continue;

    const dailyCost = _getDailyCost(Utilities.formatDate(first.dateObj, "JST", "yyyy/MM/dd"), first.cleanLoc, first.cat, first.dow, first.sH, first.eH);
    const hours = (first.eH - first.sH) * grp.length;
    
    if (first.reason !== "定期枠（8割以上埋まり）の残り分募集" && first.reason === "契約期間外による追加募集" && repeatDow.includes("毎週")) {
      let m1 = first.dateObj.getMonth() + 1;
      let m2 = last.dateObj.getMonth() + 1;
      let tSuffix = (m1 === m2) ? `${m1}` : `${m1}~${m2}`;
      
      let pStart = new Date(first.dateObj.getFullYear(), first.dateObj.getMonth(), 1);
      let pEnd = new Date(last.dateObj.getFullYear(), last.dateObj.getMonth() + 1, 0);
      let roundedPeriod = `${Utilities.formatDate(pStart, "JST", "yyyy/MM/dd")}～${Utilities.formatDate(pEnd, "JST", "yyyy/MM/dd")}`;
      
      let holArr = [];
      let irregArr = [];
      
      grp.forEach(g => {
          let cDay = ctx.calendarCache.find(c => c.getTime === g.dateObj.getTime());
          if (cDay && cDay.isHol) holArr.push(`${cDay.dateObj.getMonth() + 1}/${cDay.dateObj.getDate()}(${g.dow})`);
          
          let advs = ctx.advances.filter(adv => {
            return _isTargetLocMatch(adv.loc, first.cleanLoc, first.cat) &&
                   adv.dateObj.getTime() === g.dateObj.getTime() &&
                   (adv.sH < first.eH && adv.eH > first.sH);
          });
          advs.forEach(adv => {
            let baseTimeStr = `${('0'+adv.sH).slice(-2)}:00-${('0'+adv.eH).slice(-2)}:00`;
            irregArr.push(`${adv.docName}先生|||先行|||${baseTimeStr}|||${g.dateObj.getMonth() + 1}/${g.dateObj.getDate()}(${g.dow})`);
            if (cDay && cDay.isHol) {
              irregArr.push(`${adv.docName}先生|||先行|||（期間内の祝日）|||${g.dateObj.getMonth() + 1}/${g.dateObj.getDate()}(${g.dow})`);
            }
          });

          let subs = ctx.substitutes.filter(sub => {
            return _isTargetLocMatch(sub.loc, first.cleanLoc, first.cat) &&
                   sub.dateObj.getTime() === g.dateObj.getTime() &&
                   (sub.sH < first.eH && sub.eH > first.sH);
          });
          subs.forEach(sub => {
            let baseTimeStr = `${('0'+sub.sH).slice(-2)}:00-${('0'+sub.eH).slice(-2)}:00`;
            irregArr.push(`${sub.docName}先生|||振替|||${baseTimeStr}|||${g.dateObj.getMonth() + 1}/${g.dateObj.getDate()}(${g.dow})`);
            if (cDay && cDay.isHol) {
              irregArr.push(`${sub.docName}先生|||振替|||（期間内の祝日）|||${g.dateObj.getMonth() + 1}/${g.dateObj.getDate()}(${g.dow})`);
            }
          });
      });
      let holidaysStr = (first.dow !== "土" && first.dow !== "日" && holArr.length > 0) ? holArr.join(",") : "";
      let holWage = (holidaysStr !== "") ? _getHolidayWageWrapper(Utilities.formatDate(first.dateObj, "JST", "yyyy/MM/dd"), first.cleanLoc, first.cat, first.sH, first.eH) : "";
      let irregStr = irregArr.length > 0 ? irregArr.join("###") : "";
      
      ctx.masterRegularList.push({
        "エリア": first.area, "期間": roundedPeriod, 
        "開始時間": `${('0'+first.sH).slice(-2)}:00`, "終了時間": `${('0'+first.eH).slice(-2)}:00`, 
        "時給": wage, "祝日時給": holWage, "祝日該当日": holidaysStr, 
        "募集時間": hours, 
        "コスト": dailyCost > 0 ? dailyCost * grp.length : "",
        "対応済": false, "先行・振替": irregStr,
        _loc: first.loc, _sH: first.sH, _eH: first.eH, _freqStr: "毎週", _dow: first.dow, _tSuffix: tSuffix
      });
    } else {
      const period = grp.length > 1 ? `${Utilities.formatDate(first.dateObj, "JST", "MM/dd")}～${Utilities.formatDate(last.dateObj, "JST", "MM/dd")}` : `${Utilities.formatDate(first.dateObj, "JST", "MM/dd")}（${first.dow}）`;
      
      ctx.singleList.push({
        "エリア": first.area, "拠点名": first.loc, "理由": first.reason, "期間": period, 
        "開始時間": `${('0'+first.sH).slice(-2)}:00`, "終了時間": `${('0'+first.eH).slice(-2)}:00`, 
        "繰り返し曜日": repeatDow, "該当日": datesStr, "時給": wage, 
        "募集時間": hours, 
        "コスト": dailyCost > 0 ? dailyCost * grp.length : "", 
        "対応済": false, "注意": ""
      });
    }
  }
}