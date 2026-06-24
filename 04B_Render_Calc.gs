/**
 * ==========================================
 * 04B_Render_Calc.gs
 * コスト計算・ダッシュボード書き込み・統計初期化
 * ★マスタ不在時の無限フリーズ・安全防衛版
 * ==========================================
 */

function calculateDailyCost(shift, isHoliday, wageDataList) {
  let sHour = parseInt(shift.startTime.split(':')[0], 10);
  let eHour = parseInt(shift.endTime.split(':')[0], 10);
  let hours = 0;
  let dailyCost = 0;
  let appliedRates = [];
  
  let prefix = isHoliday ? "hol" : "wd";
  let useYear = (shift.specialWageDetail && shift.specialWageDetail.includes("2025年度")) ? "y2025" : "y2026";
  
  // ====================================================================
  // ★ 安全防衛策1：そもそも時給マスタデータが空、または壊れている場合の完全ガード
  // ====================================================================
  if (!wageDataList || wageDataList.length === 0 || !wageDataList[0]) {
    console.warn(`[⚠️時給エラー] マスタデータ全体が空です。医師: ${shift.doctorName} (拠点: ${shift.rawShift || "不明"}) の計算をスキップします。`);
    return { hours: 0, cost: 0, appliedRates: [] };
  }
  
  let targetDept = wageDataList[0]; 
  if (wageDataList.length > 1) {
    if (shift.rawShift.includes("内科")) targetDept = wageDataList.find(w => w.department === "内科") || targetDept;
    else if (shift.rawShift.includes("小児科")) targetDept = wageDataList.find(w => w.department === "小児科") || targetDept;
  }
  
  // ====================================================================
  // ★ 安全防衛策2：該当する診療科のオブジェクト（または内部のrates属性）が存在しない場合のガード
  // ====================================================================
  if (!targetDept || !targetDept.rates) {
    console.warn(`[⚠️時給エラー] 診療科マスタの構造が不正です。医師: ${shift.doctorName}`);
    return { hours: 0, cost: 0, appliedRates: [] };
  }
  
  let rates = targetDept.rates[useYear] || targetDept.rates["y2026"];
  
  // ====================================================================
  // ★ 安全防衛策3：時給（rates）が undefined になった場合のフリーズガード
  // ====================================================================
  if (!rates) {
    console.warn(`[⚠️時給設定なし] 医師: ${shift.doctorName} 先生の時給データ(rates)がありません。時給0円として計算を安全に続行します。`);
    return { hours: 0, cost: 0, appliedRates: [] };
  }

  for (let h = sHour; h < eHour; h++) {
    if (h === 13 || h === 14) continue; 
    hours++;
    
    let r = 0;
    if (h >= 9 && h < 13) r = Number(rates[`${prefix}_am`]) || 0;
    else if (h >= 15 && h < 18) r = Number(rates[`${prefix}_pm`]) || 0;
    else if (h >= 18 && h < 21) r = Number(rates[`${prefix}_nt`]) || 0;
    
    if (shift.specialWageDetail && shift.specialWageDetail.includes("円")) {
       let match = shift.specialWageDetail.match(/(\d{1,2})[：:]\d{2}-(\d{1,2})[：:]\d{2}.*?(\d{1,2}(?:,\d{3})*)円/);
       if (match) {
         let spStart = parseInt(match[1], 10);
         let spEnd = parseInt(match[2], 10);
         if (h >= spStart && h < spEnd) r = parseInt(match[3].replace(/,/g, ''), 10);
       }
    }
    
    dailyCost += r;
    if (r > 0) appliedRates.push(r);
  }
  return { hours: hours, cost: dailyCost, appliedRates: appliedRates };
}

function generateShiftString(rawShifts) {
  let grouped = {};
  rawShifts.forEach(s => {
    let key = `${s.start}-${s.end}`;
    if (!grouped[key]) grouped[key] = {};
    if (!grouped[key][s.dow]) grouped[key][s.dow] = new Set();
    grouped[key][s.dow].add(s.week);
  });

  let results = [];
  for (let timeKey in grouped) {
    for (let dow in grouped[timeKey]) {
      let weeks = Array.from(grouped[timeKey][dow]).sort();
      let weekStr = weeks.length >= 4 ? "毎週" : "第" + weeks.join("・");
      results.push(`${timeKey}(${weekStr}${dow})`);
    }
  }
  return results.join(", ");
}

/**
 * 高速化：ダッシュボードのデータをメモリ上で作成し、1回で一括書き込みする
 */
function writeDashboardInfo(sheet, startRow, endRow, edges, stats, doctorCosts, wageDataList, senkouDocsArray) {
  const searchRange = sheet.getRange(startRow, 1, endRow - startRow + 1, 16);
  const values = searchRange.getValues();
  const calcRate = (part, total) => total > 0 ? Math.round((part / total) * 100) + "%" : "0%";

  const writeColMap = {
    "常)1診目コマ数": 4, "常)1診目割合": 4, "平日充足率(9-13)": 4, "平日超募集(9-13)": 4, 
    "土曜充足率(9-13)": 4, "土曜超募集(9-13)": 4, "日曜充足率(9-13)": 4, "日曜超募集(9-13)": 4, 
    "土日充足率(9-13)": 4, "土日超募集(9-13)": 4, "土日祝充足率(9-13)": 4, "土日祝超募集(9-13)": 4, "1診コマ数": 4, "適用開始": 4,
    
    "非)1診目コマ数": 7, "非)1診目割合": 7, "平日充足率(15-18)": 7, "平日超募集(15-18)": 7, 
    "土曜充足率(15-18)": 7, "土曜超募集(15-18)": 7, "日曜充足率(15-18)": 7, "日曜超募集(15-18)": 7, 
    "土日充足率(15-18)": 7, "土日超募集(15-18)": 7, "土日祝充足率(15-18)": 7, "土日祝超募集(15-18)": 7, "2診確保コマ数": 7, "有効期間": 7,
    
    "超募集コマ数": 10, "全体超募集率": 10, "平日充足率(18-21)": 10, "平日超募集(18-21)": 10, 
    "土曜充足率(18-21)": 10, "土曜超募集(18-21)": 10, "日曜充足率(18-21)": 10, "日曜超募集(18-21)": 10, 
    "土日充足率(18-21)": 10, "土日超募集(18-21)": 10, "土日祝充足率(18-21)": 10, "土日祝超募集(18-21)": 10, "2診割合": 10
  };

  const joukinColor = "#fce5cd"; 
  const teikiColor  = "#d9d2e9"; 
  const senkouColor = "#d9ead3"; 
  const emptyColor  = "#ffffff"; 

  let docBgUpdates = []; 
  let alignUpdates = []; 

  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      let cellText = String(values[r][c]).trim().replace(/（/g, "(").replace(/）/g, ")");
      if (!cellText) continue;
      
      let writeVal = null;

      if (cellText === "適用開始") writeVal = edges.start;
      else if (cellText === "有効期間") writeVal = edges.end;
      else if (cellText === "常)1診目コマ数") writeVal = stats.joukin1shin;
      else if (cellText === "非)1診目コマ数") writeVal = stats.teiki1shin;
      else if (cellText === "超募集コマ数") writeVal = stats.boshu;
      
      else if (cellText === "常)1診目割合") writeVal = calcRate(stats.joukin1shin, stats.total1shin);
      else if (cellText === "非)1診目割合") writeVal = calcRate(stats.teiki1shin, stats.total1shin);
      else if (cellText === "全体超募集率") writeVal = calcRate(stats.boshu, stats.total1shin);
      else if (cellText === "1診コマ数") writeVal = stats.total1shin;
      else if (cellText === "2診確保コマ数") writeVal = stats.total2shinFilled;
      else if (cellText === "2診割合") writeVal = calcRate(stats.total2shinFilled, stats.total1shin);
      else if (cellText === "募集枠合計") writeVal = stats.boshu + " 時間";
      else if (cellText === "常勤比率") {
        writeVal = stats.total1shin > 0 ? Math.round((stats.joukin1shin / stats.total1shin) * 100) + "%" : "0%";
      }

      let grids = ["平日", "土曜", "日曜"];
      let times = ["9-13", "15-18", "18-21"];
      grids.forEach(gk => {
        times.forEach(tk => {
          if (cellText === `${gk}充足率(${tk})`) writeVal = calcRate(stats.grid[gk][tk].filled, stats.grid[gk][tk].total);
          if (cellText === `${gk}超募集(${tk})`) writeVal = calcRate(stats.grid[gk][tk].boshu, stats.grid[gk][tk].total);
        });
      });

      times.forEach(tk => {
        let weTotal = stats.grid["土曜"][tk].total + stats.grid["日曜"][tk].total;
        let weFilled = stats.grid["土曜"][tk].filled + stats.grid["日曜"][tk].filled;
        let weBoshu = stats.grid["土曜"][tk].boshu + stats.grid["日曜"][tk].boshu;
        if (cellText === `土日充足率(${tk})` || cellText === `土日祝充足率(${tk})`) writeVal = calcRate(weFilled, weTotal);
        if (cellText === `土日超募集(${tk})` || cellText === `土日祝超募集(${tk})`) writeVal = calcRate(weBoshu, weTotal);
      });

      if (writeVal !== null) {
        let targetCol = writeColMap[cellText] ? writeColMap[cellText] : c + 2;
        values[r][targetCol - 1] = writeVal; 
      }

      if (cellText === "常勤医師") {
        let docs = Array.from(stats.uniqueJoukin);
        let bgs = new Array(7).fill(emptyColor);
        for (let i = 0; i < docs.length && i < 7; i++) { values[r][3 + i] = docs[i]; bgs[i] = joukinColor; }
        docBgUpdates.push({row: startRow + r, bgs: [bgs]});
      }
      if (cellText === "非常勤医師") {
        let docs = Array.from(stats.uniqueTeiki);
        let bgs = new Array(7).fill(emptyColor);
        for (let i = 0; i < docs.length && i < 7; i++) { values[r][3 + i] = docs[i]; bgs[i] = teikiColor; }
        docBgUpdates.push({row: startRow + r, bgs: [bgs]});
      }
      if (cellText === "先行応募" || cellText === "先行応募医師") {
        let docs = senkouDocsArray;
        let bgs = new Array(7).fill(emptyColor);
        for (let i = 0; i < docs.length && i < 7; i++) { values[r][3 + i] = docs[i]; bgs[i] = senkouColor; }
        docBgUpdates.push({row: startRow + r, bgs: [bgs]});
      }

      if (cellText === "医師名" && String(values[r][c+1]).trim() === "稼働時間") {
        let totalCost = 0;
        let idx = 0;
        
        for (let doc in doctorCosts) {
          let d = doctorCosts[doc];
          let contractDetails = generateShiftString(d.rawShifts);
          let sortedRates = Array.from(d.appliedRates).sort((a,b) => a - b);
          let rateStr = sortedRates.length > 0 ? sortedRates.map(r => r.toLocaleString()).join("/") : "時給表どおり";
          
          if (r + 1 + idx < values.length) {
            values[r + 1 + idx][c] = `${doc}（定非）`;
            values[r + 1 + idx][c + 1] = d.hours;
            values[r + 1 + idx][c + 2] = contractDetails;
            values[r + 1 + idx][c + 3] = rateStr;
            values[r + 1 + idx][c + 4] = d.cost.toLocaleString();
          }
          
          totalCost += d.cost;
          idx++;
        }
        
        if (idx > 0) {
          alignUpdates.push({ row: startRow + r + 1, col: c + 1, numRows: idx, numCols: 5 });
        }
        
        let defaultRate = wageDataList && wageDataList[0] && wageDataList[0].rates.y2026 ? wageDataList[0].rates.y2026 : {wd_am:0, wd_pm:0, wd_nt:0};
        const avgHourlyBoshu = (Number(defaultRate.wd_am) + Number(defaultRate.wd_pm) + Number(defaultRate.wd_nt)) / 3 || 12000;
        stats.totalMonthlyCost = totalCost + (stats.boshu * avgHourlyBoshu);
      }
      
      if (cellText === "日別拠点維持コスト" && stats.totalMonthlyCost) {
        let daily = Math.round(stats.totalMonthlyCost / edges.daysInMonth);
        values[r][c + 1] = daily.toLocaleString() + " 円";
      }
    }
  }
  
  searchRange.setValues(values);
  
  docBgUpdates.forEach(update => {
     sheet.getRange(update.row, 4, 1, 7).setBackgrounds(update.bgs);
  });
  
  alignUpdates.forEach(update => {
     sheet.getRange(update.row, update.col, update.numRows, update.numCols).setHorizontalAlignment("left");
  });
}

function initStats() {
  return {
    joukin1shin: 0, teiki1shin: 0, senkou1shin: 0, boshu: 0, total1shin: 0, total2shinFilled: 0,
    uniqueJoukin: new Set(), uniqueTeiki: new Set(), uniqueSenkou: new Set(), uniqueSubstitute: new Set(),
    grid: {
      "平日": { "9-13": { total:0, filled:0, boshu:0 }, "15-18": { total:0, filled:0, boshu:0 }, "18-21": { total:0, filled:0, boshu:0 } },
      "土曜": { "9-13": { total:0, filled:0, boshu:0 }, "15-18": { total:0, filled:0, boshu:0 }, "18-21": { total:0, filled:0, boshu:0 } },
      "日曜": { "9-13": { total:0, filled:0, boshu:0 }, "15-18": { total:0, filled:0, boshu:0 }, "18-21": { total:0, filled:0, boshu:0 } }
    }
  };
}

function getMonthDaysGroupedByDOW(yearMonthStr) {
  const [year, month] = yearMonthStr.split('/').map(Number);
  const daysInMonth = new Date(year, month, 0).getDate();
  const grouped = [];
  const dowMap = [1, 2, 3, 4, 5, 6, 0]; 
  const dowNames = {1:"月", 2:"火", 3:"水", 4:"木", 5:"金", 6:"土", 0:"日"};

  dowMap.forEach(targetDow => {
    let count = 0;
    for (let d = 1; d <= daysInMonth; d++) {
      let date = new Date(year, month - 1, d);
      if (date.getDay() === targetDow) {
        count++;
        grouped.push({ dateStr: `${year}/${('0'+month).slice(-2)}/${('0'+d).slice(-2)}`, dayOfWeek: dowNames[targetDow], weekNum: count, isValid: true });
      }
    }
    while (count < 5) {
      count++;
      grouped.push({ dateStr: null, dayOfWeek: dowNames[targetDow], weekNum: count, isValid: false });
    }
  });
  return grouped;
}