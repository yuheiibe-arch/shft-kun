/**
 * ====================================================================
 * 08B_1_LocHelpers.gs
 * 拠点処理用ヘルパーと、高速化のための辞書生成ロジック
 * ====================================================================
 */

function _isContractActiveOnCache(contract, cDay) {
  if (contract.dows) {
    if (!contract.dows.includes(cDay.dN)) return false;
  } else {
    if (cDay.dN !== contract.dow) return false;
  }
  if (contract.validFrom && cDay.getTime < contract.validFrom.getTime()) return false;
  if (contract.validTo && cDay.getTime > contract.validTo.getTime()) return false;
  if (contract.weeks.length === 5) return true;
  return contract.weeks.includes(cDay.wNum);
}

function _isTargetLocMatch(dataLoc, cleanLocName, category) {
  let match = dataLoc.includes(cleanLocName) || cleanLocName.includes(dataLoc);
  if (!match) return false;
  if (cleanLocName === "亀有" || cleanLocName === "北葛西") {
    let hasNaika = dataLoc.includes("内科");
    let hasShouni = dataLoc.includes("小児科") || dataLoc.includes("小児");
    if (category === "内科" && hasShouni && !hasNaika) return false;
    if (category === "小児科" && hasNaika && !hasShouni) return false;
  }
  return true;
}

// 🚀 タイムアウト回避のための辞書生成（元の判定ロジックを維持）
function _buildDailyDictsForLoc(ctx, cleanLocName, category, actualStartDate) {
  const buildDict = (dataArray) => {
    let dict = {};
    if (!dataArray) return dict;
    for (let i = 0; i < dataArray.length; i++) {
      let item = dataArray[i];
      if (_isTargetLocMatch(item.loc, cleanLocName, category) && item.dateObj >= actualStartDate && item.dateObj <= ctx.endDate) {
        let dStr = item.dStr || Utilities.formatDate(item.dateObj, "JST", "yyyy/MM/dd");
        item.dStr = dStr; 
        if (!dict[dStr]) dict[dStr] = [];
        dict[dStr].push(item);
      }
    }
    return dict;
  };

  return {
    advances: buildDict(ctx.advances),
    substitutes: buildDict(ctx.substitutes),
    absences: buildDict(ctx.absences),
    kyukans: buildDict(ctx.kyukans)
  };
}

// 🚀 2診アラート検知用の混雑マップ
function _buildDailyBusyMap(activeContracts, locCalendar, locDicts) {
  let map = {};
  activeContracts.forEach(c => {
    locCalendar.forEach(cDay => {
      if (_isContractActiveOnCache(c, cDay)) {
        if (!map[cDay.dStr]) map[cDay.dStr] = [];
        map[cDay.dStr].push({sH: c.sH, eH: c.eH, docName: c.docName});
      }
    });
  });
  ["advances", "substitutes"].forEach(key => {
    for(let dStr in locDicts[key]) {
      if(!map[dStr]) map[dStr] = [];
      locDicts[key][dStr].forEach(irr => {
        map[dStr].push({sH: irr.sH, eH: irr.eH, docName: irr.docName});
      });
    }
  });
  return map;
}