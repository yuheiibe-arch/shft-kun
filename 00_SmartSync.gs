/**
 * ==========================================
 * 00_SmartSync.gs
 * 差分チェックとスマート更新（順不同対応・超絶爆速版）
 * ★外部3ファイル分散出力の監視対応版
 * ==========================================
 */

function checkAndStartSmartSync(year, term) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  createProgressMonitor(ss, "計算中");
  let monitor = ss.getSheetByName("🗂️進行中モニター");
  if (!monitor) return;
  monitor.getRange("A1:B1").setValue("⚡ 変更箇所（差分）を爆速チェックしています...\nそのままお待ちください...");

  let existingLocs = [];
  
  // ★修正：マスタではなく、外部の3ファイルから出力済みシートをかき集める
  const TARGET_IDS = [
    "1rScroDlMNiRxThbaxGEuvhyCH2b9RoM6BadNSCAWsvI", // 東京埼玉
    "19Q-xVsMX0thz_rvmdo_GERJvhjFLQgw2pUkWaTYxQhE", // 関東
    "1fIFvTck_g9-Hp8MSpY7hIWwE5L2buJjBHIY4GBEf_MY"  // 関西
  ];
  
  TARGET_IDS.forEach(id => {
    try {
      let extSs = SpreadsheetApp.openById(id);
      extSs.getSheets().forEach(s => {
        let name = s.getName();
        if (name.startsWith(year)) {
          existingLocs.push(name.replace(year, ''));
        }
      });
    } catch(e) {
      console.warn("外部ファイルの読み込みに失敗しました: " + id);
    }
  });

  if (existingLocs.length === 0) {
    monitor.getRange("A1:B1").setValue("⚠️ 対象年度のシートが外部ファイルに見つかりませんでした。").setBackground("#ea4335");
    return;
  }

  let rawCache = getMasterRawData(year);
  let props = PropertiesService.getScriptProperties();
  
  let hasAnyHash = false;
  existingLocs.forEach(subLoc => {
    if (props.getProperty('HASH_' + year + '_' + subLoc)) hasAnyHash = true;
  });

  if (!hasAnyHash) {
    existingLocs.forEach(subLoc => {
      let newHash = computeStableHash(year, subLoc, rawCache);
      props.setProperty('HASH_' + year + '_' + subLoc, newHash);
    });
    monitor.getRange("A1:B1").setValue("✨ 初期セットアップ完了\n\n現在の状態を『基準』としてシステムに記憶しました！\n次からマスターを変更して実行すると、差分だけが爆速で更新されます。").setBackground("#34a853");
    return;
  }

  let targetLocs = [];
  existingLocs.forEach(subLoc => {
    let newHash = computeStableHash(year, subLoc, rawCache);
    let oldHash = props.getProperty('HASH_' + year + '_' + subLoc);

    if (newHash !== oldHash) {
      targetLocs.push(subLoc);
    }
  });

  if (targetLocs.length === 0) {
    monitor.getRange("A1:B1").setValue("✨ 変更はありませんでした！\n(すべてのシートが最新の状態です)").setBackground("#34a853");
    Utilities.sleep(3000);
    try { ss.deleteSheet(monitor); } catch(e) {}
    return;
  }

  startBackgroundBatch({
    year: year,
    term: term,
    locations: targetLocs
  });
}

function getMasterRawData(year) {
  const cache = {};
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  ['先行応募', 'お休み情報', '振替勤務'].forEach(sName => {
    let sheet = ss.getSheetByName(sName);
    if (sheet) cache[sName] = sheet.getDataRange().getValues();
  });

  try {
    const kyuSs = safeOpenByUrl("https://docs.google.com/spreadsheets/d/1cbeXWojsxNMhQUo1c6VflF5hLUJUyfuOXCFbGP5jJEA/edit");
    const kyuSheet = kyuSs.getSheetByName("休館日");
    if (kyuSheet) cache['休館日'] = kyuSheet.getDataRange().getValues();
  } catch(e) {}

  try {
    const masterSs = safeOpenByUrl('https://docs.google.com/spreadsheets/d/1aEjphEv_63SeWQmwiOy9sx7IrMfawU01sHbKd_Ki4iA/edit');
    ['常勤', '定期非常勤'].forEach(type => {
      let sheet = masterSs.getSheetByName(`${type}${year}年度`);
      if (sheet) cache[type] = sheet.getDataRange().getValues();
    });
  } catch(e) {}
  
  return cache;
}

function computeStableHash(year, subLocName, rawCache) {
  let baseLoc = subLocName.replace(/（.*?）/, '');
  let validRows = [];
  
  Object.keys(rawCache).forEach(key => {
    let rows = rawCache[key];
    if (rows) {
      rows.forEach(row => {
        let rowStr = row.map(cell => String(cell).trim()).join('|').replace(/\|+$/, '');
        if (rowStr === "") return;
        if (rowStr.includes(subLocName) || rowStr.includes(baseLoc)) {
          validRows.push(key + "::" + rowStr);
        }
      });
    }
  });

  validRows.sort();
  let str = validRows.join('\n');

  const rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, str);
  let txtHash = '';
  for (let i = 0; i < rawHash.length; i++) {
    let hashVal = rawHash[i];
    if (hashVal < 0) hashVal += 256;
    if (hashVal.toString(16).length == 1) txtHash += '0';
    txtHash += hashVal.toString(16);
  }
  return txtHash;
}

function doGet(e) {
  try {
    let targetYear = e.parameter.year;
    let targetsStr = e.parameter.targets; 

    if (!targetYear) {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let maxYear = 0;
      // doGet時の年度取得もマスタ基準に変更
      ss.getSheets().forEach(s => {
        const match = s.getName().match(/^(?:常勤|定期非常勤|)(\d{4})/);
        if (match) {
          const y = parseInt(match[1], 10);
          if (y > maxYear) maxYear = y;
        }
      });
      targetYear = String(maxYear);
    }

    if (targetYear !== "0") {
      let props = PropertiesService.getScriptProperties();
      props.setProperty('PENDING_DO_GET_YEAR', targetYear);
      props.setProperty('PENDING_DO_GET_TARGETS', targetsStr || "");

      const triggers = ScriptApp.getProjectTriggers();
      triggers.forEach(t => {
        if (t.getHandlerFunction() === 'processDoGetBackground') {
          ScriptApp.deleteTrigger(t);
        }
      });

      ScriptApp.newTrigger('processDoGetBackground').timeBased().after(1000).create();
    }
    
    return ContentService.createTextOutput("Success: Received and starting background process.");
  } catch (error) {
    return ContentService.createTextOutput("Error: " + error.message);
  }
}

function processDoGetBackground() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'processDoGetBackground') {
      ScriptApp.deleteTrigger(t);
    }
  });

  let props = PropertiesService.getScriptProperties();
  let targetYear = props.getProperty('PENDING_DO_GET_YEAR');
  let targetsStr = props.getProperty('PENDING_DO_GET_TARGETS');

  if (!targetYear) return;

  if (targetsStr) {
    let rawTargets = targetsStr.split(",");
    let finalTargets = new Set();
    
    let existingLocs = [];
    const TARGET_IDS = [
      "1rScroDlMNiRxThbaxGEuvhyCH2b9RoM6BadNSCAWsvI", 
      "19Q-xVsMX0thz_rvmdo_GERJvhjFLQgw2pUkWaTYxQhE", 
      "1fIFvTck_g9-Hp8MSpY7hIWwE5L2buJjBHIY4GBEf_MY"  
    ];
    TARGET_IDS.forEach(id => {
      try {
        SpreadsheetApp.openById(id).getSheets().forEach(s => {
          if (s.getName().startsWith(targetYear)) existingLocs.push(s.getName().replace(targetYear, ''));
        });
      } catch(e) {}
    });

    rawTargets.forEach(rt => {
      let clean = rt.replace('内科', '').replace('小児科', '').trim();
      existingLocs.forEach(el => {
        if (el.includes(clean)) finalTargets.add(el);
      });
    });

    if (finalTargets.size > 0) {
      startBackgroundBatch({
        year: targetYear,
        term: "通年",
        locations: Array.from(finalTargets)
      });
      return;
    }
  }

  checkAndStartSmartSync(targetYear, "通年");
}

function triggerNightlySmartSync() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let maxYear = 0;
  
  ss.getSheets().forEach(s => {
    const match = s.getName().match(/^(?:常勤|定期非常勤|)(\d{4})/);
    if (match) {
      const y = parseInt(match[1], 10);
      if (y > maxYear) maxYear = y;
    }
  });
  
  if (maxYear === 0) return;
  
  console.log(`🌙 夜間スマート更新を開始します: ${maxYear}年度 / 通年`);
  checkAndStartSmartSync(String(maxYear), "通年");
}