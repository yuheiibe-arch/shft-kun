/**
 * ====================================================================
 * 03_DataFormatter.gs
 * データの文字整形、時間のブロック化、辞書の生成など
 * ====================================================================
 */

function _debug_normalizeStr_ForFinal(str) {
  if (!str) return "";
  return str.toString().normalize('NFKC')
    .replace(/[（）()【】\[\]\s]/g, "")
    .replace(/(内科|小児科|クリニック|病院|モール|診療所)$/g, "")
    .replace(/ヶ/g, "ケ").trim();
}

function _processIrregularShifts(irregData, typeName) {
  let groups = {};
  irregData.forEach(d => {
    d.typeName = typeName; 
    let dN = GLOBAL_DOW_NAMES[d.dateObj.getDay()];
    let key = `${d.docName}_${d.sH}_${d.eH}_${dN}`;
    if (!groups[key]) groups[key] = [];
    groups[key].push(d);
  });

  let results = { bundled: [], singles: [] };
  for (let key in groups) {
    let grp = groups[key];
    grp.sort((a,b) => a.dateObj - b.dateObj);
    
    const firstDate = grp[0].dateObj;
    const lastDate = grp[grp.length - 1].dateObj;
    const daysDiff = (lastDate - firstDate) / 86400000;
    
    let totalDowCount = 0;
    let cur = new Date(firstDate);
    while(cur <= lastDate) { 
      totalDowCount++; 
      cur.setDate(cur.getDate() + 7); 
    }
    const attendanceRate = grp.length / totalDowCount;

    if (daysDiff >= 28 && attendanceRate >= 0.8) {
      results.bundled.push({
        grp: grp, 
        docName: grp[0].docName, 
        sH: grp[0].sH, 
        eH: grp[0].eH,
        dow: GLOBAL_DOW_NAMES[firstDate.getDay()], 
        firstDate: firstDate, 
        lastDate: lastDate, 
        typeName: typeName,
        isDaishin: grp[0].isDaishin,
        requester: grp[0].requester 
      });
    } else {
      grp.forEach(d => results.singles.push(d));
    }
  }
  return results;
}

function _extractContractInfo(bikou, targetLoc) {
  if (!bikou) return "";
  const lines = bikou.split(/\r\n|\n|\r/);
  let periodStr = "";
  let holidayStr = "";
  let locLines = [];

  for (let line of lines) {
    let t = line.trim();
    if (!t) continue;
    if (t.match(/\d{4}\/\d{1,2}\/\d{1,2}～\d{4}\/\d{1,2}\/\d{1,2}/)) {
      if (!periodStr) periodStr = t;
    } else if (t.includes("祝日") || t.includes("年末年始")) {
      if (!holidayStr) holidayStr = t;
    } else if (t.includes("【") && t.includes(targetLoc)) {
      locLines.push(t);
    }
  }
  
  let result = [];
  if (periodStr) result.push(periodStr);
  result = result.concat(locLines);
  if (holidayStr) result.push(holidayStr);
  
  return result.join('\n');
}

function _buildLocMasterDict() {
  let areaMap = {};
  let openDateMap = {};
  try {
    // ★ここを safeOpenByUrl に変更
    const ss = safeOpenByUrl('https://docs.google.com/spreadsheets/d/14RbsDcv0nXfEwweki8-9cK3lQUg1XUuhozLNF9u2qAs/edit');
    const data = ss.getSheetByName('拠点名').getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const n = String(data[i][0]).trim();
      const g = String(data[i][5]).trim();
      const a = String(data[i][6]).trim();
      
      if (!n || n === "MQC") continue;
      
      areaMap[n] = a.includes("東京") || g.includes("関東") ? "東京" : 
                   a.includes("神奈川") || g.includes("神奈川") ? "神奈川" : 
                   a.includes("千葉") || g.includes("千葉") ? "千葉" : 
                   a.includes("埼玉") || g.includes("埼玉") ? "埼玉" : 
                   a.includes("関西") || g.includes("大阪") ? "関西" : "その他";
      
      let d = data[i][7];
      d = (!d || String(d).trim() === "") ? new Date(2099, 11, 31) : (d instanceof Date ? d : new Date(d));
      if (!isNaN(d.getTime())) { 
        openDateMap[n] = d; 
        for(let j = 1; j <= 4; j++) {
          if(String(data[i][j]).trim()) {
            openDateMap[String(data[i][j]).trim()] = d; 
          }
        }
      }
    }
  } catch(e) { }
  return { areaMap, openDateMap };
}

function _getOpenDateHelperMaster(loc, map) {
  let c = loc.replace(/[（\(\)）]/g, "").replace(/内科|小児科/g, "").trim();
  if (map[c]) return map[c];
  c = c.replace(/アトレ|イーアス|セブンパーク|フルル/g, "").trim();
  if (map[c]) return map[c];
  for (let k in map) {
    if (c === k || c.includes(k) || k.includes(c)) return map[k];
  }
  return null;
}

function _getAreaHelper(loc, map) {
  if (!loc) return "その他";
  if (["欠勤", "有給", "休館日"].includes(loc)) return loc; 
  if (map[loc]) return map[loc];
  let c = loc.replace(/[（\(\)）]/g, "").replace(/内科|小児科/g, "").trim();
  for (let k in map) {
    if (c === k || c.includes(k) || k.includes(c)) return map[k];
  }
  return "その他";
}

function _splitTimeIntoBlocks(sH, eH) {
  let chunks = [];
  if (sH === 17) {
    chunks.push({sH: 17, eH: eH});
    return chunks;
  }
  if (sH < 13 && eH > 9) chunks.push({sH: 9, eH: 13});
  if (sH < 18 && eH > 15) chunks.push({sH: 15, eH: 18});
  if (sH < 22 && eH > 18) chunks.push({sH: 18, eH: 21});
  return chunks;
}

function _extractWeeklyBlocks(arr) {
  let blocks = [];
  let cur = null;
  let is17Start = false; 
  
  const tz = (i) => {
    if (is17Start && i >= 8) return "PM_NT_MERGED"; 
    return i < 4 ? "AM" : i < 6 ? "R" : i < 9 ? "PM" : "NT";
  };
  
  for (let i = 0; i <= 12; i++) {
    const setObj = i < 12 ? arr[i] : null;
    let val = null;
    if (setObj && setObj.size > 0) {
      const arrSorted = Array.from(setObj).sort();
      val = (arrSorted.length === 5) ? "毎週" : `第${arrSorted.join('・')}`;
    }

    if (val) {
      if (!cur) {
        is17Start = (i === 8); 
        cur = { s: i, e: i, tz: tz(i), freq: val, weeksArr: Array.from(setObj).sort() };
      } else if (cur.tz !== tz(i) || cur.freq !== val) { 
        blocks.push(cur); 
        is17Start = (i === 8); 
        cur = { s: i, e: i, tz: tz(i), freq: val, weeksArr: Array.from(setObj).sort() }; 
      } else {
        cur.e = i;
      }
    } else if (cur) { 
      blocks.push(cur); 
      cur = null; 
      is17Start = false;
    }
  }
  return blocks.map(b => ({ 
    sH: b.s + 9, eH: b.e + 10, 
    sT: `${('0'+(b.s+9)).slice(-2)}:00`, eT: `${('0'+(b.e+10)).slice(-2)}:00`, 
    hours: b.e - b.s + 1, freqStr: b.freq, weeksArr: b.weeksArr 
  }));
}

// =========================================================
// マージ関数
// =========================================================

function _mergeConfirmList(list) {
  if (!list || list.length === 0) return [];
  list.sort((a, b) => (a._sortDate || 0) - (b._sortDate || 0));
  return list.map(item => {
    let cleanItem = {};
    for (let key in item) {
      if (!key.startsWith('_')) cleanItem[key] = item[key];
    }
    return cleanItem;
  });
}

function _mergeMasterRegularList(list) {
  if (!list || list.length === 0) return [];
  const dowOrder = {"月":1, "火":2, "水":3, "木":4, "金":5, "土":6, "日":7};

  let grouped = {};
  list.forEach(item => {
    let key = `${item["エリア"]}_${item._loc}_${item._sH}_${item._eH}_${item._tSuffix}_${item._freqStr}`;
    if (!grouped[key]) grouped[key] = [];
    grouped[key].push(item);
  });

  let mergedList = [];
  
  for (let key in grouped) {
    let items = grouped[key];
    items.sort((a, b) => (dowOrder[a._dow] || 99) - (dowOrder[b._dow] || 99));

    let base = items[0];
    let cleanItem = {};

    cleanItem["エリア"] = base["エリア"] || "";

    // ★修正箇所：「毎週」の場合は曜日のみ、それ以外（「第2・4」など）は残す
    let uniqueDows = [...new Set(items.map(i => i._dow))];
    let dowStrForTitle = "";
    if (items.length === 1) {
      if (base._freqStr === "毎週") {
        dowStrForTitle = uniqueDows[0]; // 例: "日"
      } else {
        dowStrForTitle = `${base._freqStr}(${uniqueDows[0]})`; // 例: "第2・4(日)"
      }
    } else {
      // 複数曜日の場合
      dowStrForTitle = uniqueDows.join('');
    }

    if (base._loc && base._sH && base._eH && base._tSuffix) {
      cleanItem["シフトタイトル"] = `${base._loc}／${base._sH}-${base._eH}／${dowStrForTitle}／${base._tSuffix}`;
    } else {
      cleanItem["シフトタイトル"] = "";
    }

    cleanItem["期間"] = base["期間"] || "";
    cleanItem["開始時間"] = base["開始時間"] || "";
    cleanItem["終了時間"] = base["終了時間"] || "";

    // 繰り返し曜日
    let uniqueFreqDows = [...new Set(items.map(i => `${i._freqStr}${i._dow}曜日`))];
    cleanItem["繰り返し曜日"] = uniqueFreqDows.join('\n');

    // 時給
    let wages = items.map(i => i["時給"]);
    let allSameWage = wages.every(w => w === wages[0]);
    if (allSameWage) {
      cleanItem["時給"] = wages[0] || "";
    } else {
      let wageByDow = {};
      items.forEach(i => { if (!wageByDow[i._dow]) wageByDow[i._dow] = i["時給"]; });
      cleanItem["時給"] = Object.keys(wageByDow).map(d => `${d}曜:\n${wageByDow[d]}`).join('\n\n');
    }

    // 祝日時給
    let hWages = [...new Set(items.map(i => i["祝日時給"]).filter(w => w))];
    cleanItem["祝日時給"] = hWages.join('\n');

    // 祝日該当日
    let allHols = [];
    items.forEach(i => {
       if (i["祝日該当日"]) {
         i["祝日該当日"].split(',').forEach(p => {
            let m = p.match(/(\d+)\/(\d+\(.+\))/);
            if (m) allHols.push(`${m[1]}月：${m[2]}`);
            else allHols.push(p);
         });
       }
    });
    cleanItem["祝日該当日"] = [...new Set(allHols)].join('\n');

    // 募集時間とコストの合算
    cleanItem["募集時間"] = items.reduce((sum, i) => sum + (Number(i["募集時間"]) || 0), 0) || "";
    cleanItem["コスト"] = items.reduce((sum, i) => sum + (Number(i["コスト"]) || 0), 0) || "";

    // 先行・振替の改行フォーマット
    let irregGroups = {};
    items.forEach(i => {
        let rawIrreg = i["先行・振替"];
        if (rawIrreg && rawIrreg.includes("|||")) {
          rawIrreg.split("###").forEach(str => {
            let p = str.split("|||");
            if (p.length >= 4) {
              let k = `【${p[0]}（${p[1]}）】\n${p[2]}`;
              let dMatch = p[3].match(/(\d+)\/(\d+\(.+\))/);
              let fd = dMatch ? `${dMatch[1]}月：${dMatch[2]}` : p[3];
              if (!irregGroups[k]) irregGroups[k] = [];
              irregGroups[k].push(fd);
            }
          });
        }
    });
    let res = [];
    for (let k in irregGroups) {
      res.push(k + "\n" + [...new Set(irregGroups[k])].join("\n"));
    }
    cleanItem["先行・振替"] = res.join("\n\n");

    cleanItem["対応済"] = false;

    cleanItem._loc = base._loc;
    cleanItem._sH = base._sH;
    cleanItem._dowArr = uniqueDows;

    mergedList.push(cleanItem);
  }

  mergedList.sort((a, b) => {
     let locA = a._loc || "";
     let locB = b._loc || "";
     if (locA !== locB) return locA.localeCompare(locB, 'ja');
     if (a._sH !== b._sH) return a._sH - b._sH;
     let dowA = dowOrder[a._dowArr[0]] || 99;
     let dowB = dowOrder[b._dowArr[0]] || 99;
     return dowA - dowB;
  });

  return mergedList.map(item => {
     delete item._loc;
     delete item._sH;
     delete item._dowArr;
     return item;
  });
}