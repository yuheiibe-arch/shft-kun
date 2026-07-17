/**
 * ==========================================
 * 00_Batch_Engine.gs
 * バックグラウンドのバッチ処理・キュー管理
 * ★【完全体】3ファイル自動振り分け＆直接書き出しモデル
 * ==========================================
 */

function startBackgroundBatch(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();
  
  if (!payload || !payload.locations || payload.locations.length === 0) {
    return "Error: Locations Empty";
  }

  payload.locations = payload.locations.filter(loc => {
    let invalidWords = [
      "その他", "キャッシュ", "作業表", "お休み", "確定シフト", 
      "欠勤", "キャンセル", "手順書", "原本", "振替勤務", 
      "先行応募", "単独募集", "定期募集", "目次", "設定", 
      "テンプレート", "ダッシュボード"
    ];
    return !invalidWords.some(word => loc.includes(word));
  });

  if (payload.locations.length === 0) return "Error: 有効な出力対象拠点がありませんでした。";

  // ★ 拠点グループ判定（3ファイルへの振り分けマップ作成）
  const TARGET_IDS = {
    "TOKYO_SAITAMA": "1rScroDlMNiRxThbaxGEuvhyCH2b9RoM6BadNSCAWsvI",
    "KANTO": "19Q-xVsMX0thz_rvmdo_GERJvhjFLQgw2pUkWaTYxQhE",
    "KANSAI": "1fIFvTck_g9-Hp8MSpY7hIWwE5L2buJjBHIY4GBEf_MY"
  };
  const masterUrl = 'https://docs.google.com/spreadsheets/d/14RbsDcv0nXfEwweki8-9cK3lQUg1XUuhozLNF9u2qAs/edit';
  
  let destMap = {};
  try {
    let masterExternalSs = SpreadsheetApp.openByUrl(masterUrl);
    let masterSheet = masterExternalSs.getSheetByName("拠点名") || masterExternalSs.getSheets()[0];
    let data = masterSheet.getDataRange().getValues();
    let headers = data[0];
    let nameIdx = headers.indexOf("正規記載");
    let groupIdx = headers.indexOf("拠点グループ");
    let areaIdx = headers.indexOf("エリア");

    for (let r = 1; r < data.length; r++) {
      let clinicName = String(data[r][nameIdx]).trim();
      let groupName = String(data[r][groupIdx]).trim();
      let areaName = String(data[r][areaIdx]).trim();

      if (!clinicName || clinicName.toUpperCase() === "MQC") continue;

      if (groupName === "関東第一" || groupName === "関東第二" || groupName === "埼玉") {
        destMap[clinicName] = TARGET_IDS.TOKYO_SAITAMA;
      } else if (areaName === "関西") {
        destMap[clinicName] = TARGET_IDS.KANSAI;
      } else if (areaName === "関東") {
        destMap[clinicName] = TARGET_IDS.KANTO;
      }
    }
  } catch(e) {
    console.error("エリア分類マスタの取得に失敗: " + e.message);
  }
  
  payload.targetMap = destMap; // 判定結果をキューに保存

  // ★ テンプレートを各外部ファイルに自動同期（最新化）
  console.log("⏳ 各出力ファイルへテンプレートを同期中...");
  let templateSheet = ss.getSheetByName("テンプレート");
  let uniqueIds = [...new Set(Object.values(TARGET_IDS))];
  
  uniqueIds.forEach(id => {
    try {
      let tSs = SpreadsheetApp.openById(id);
      if (templateSheet) {
        let oldTemp = tSs.getSheetByName("テンプレート");
        if (oldTemp) tSs.deleteSheet(oldTemp); // 古いものを消して最新化
        let copied = templateSheet.copyTo(tSs);
        copied.setName("テンプレート");
        copied.hideSheet();
      }
    } catch(e) {
      console.warn("外部ファイルへの同期に失敗: " + e.message);
    }
  });

  payload.totalCount = payload.locations.length;
  payload.completedCount = 0;
  payload.currentMonthIndex = 0; 
  
  props.setProperty('BOSHUKUN_BATCH_QUEUE', JSON.stringify(payload));
  deleteTriggers();

  let oldCache = ss.getSheetByName("⚙️通信キャッシュ");
  if (oldCache) ss.deleteSheet(oldCache);

  if (typeof createProgressMonitor === 'function') createProgressMonitor(ss, payload.totalCount);
  if (typeof updateProgressMonitor === 'function') {
    updateProgressMonitor(ss, 0, payload.totalCount, payload.totalCount * 2, "準備中...（3ファイルへの分散出力を開始します）");
  }

  console.log(`🚀 startBackgroundBatch: ${payload.locations.join(", ")} の処理を開始。`);
  ScriptApp.newTrigger('processBatchQueue').timeBased().after(1000).create();
  return "Success";
}

function processBatchQueue() {
  const BATCH_START_TIME = Date.now();
  const TIME_LIMIT = 240000; 

  console.log("▶️ processBatchQueue 起動 (3ファイル分散出力モード)");
  deleteTriggers(); 
  
  const props = PropertiesService.getScriptProperties();
  const queueStr = props.getProperty('BOSHUKUN_BATCH_QUEUE');
  if (!queueStr) return;
  
  let queue = JSON.parse(queueStr);
  const masterSs = SpreadsheetApp.getActiveSpreadsheet(); // ★元のマスタファイル

  if (!queue.locations || queue.locations.length === 0) {
    props.deleteProperty('BOSHUKUN_BATCH_QUEUE');
    return;
  }

  try {
    let monitorSheet = masterSs.getSheetByName("🗂️進行中モニター");
    if (monitorSheet && monitorSheet.getRange("A2").getValue() === true) {
      monitorSheet.getRange("A1:B1").setValue("🛑 処理を緊急停止しました").setBackground("#ea4335");
      props.deleteProperty('BOSHUKUN_BATCH_QUEUE');
      return;
    }
  } catch(e) {}

  const targetYear = parseInt(queue.year, 10);
  const targetTerm = queue.term;
  const subLocName = queue.locations[0]; 
  const baseLocName = subLocName.replace(/（.*?）/, '');
  const finalSheetName = `${targetYear}${subLocName}`;

  if (queue.currentMonthIndex === undefined) queue.currentMonthIndex = 0;
  if (queue.retryCount === undefined) queue.retryCount = 0;

  // ★ 出力先ファイル（targetSs）の判定と取得
  const destMap = queue.targetMap || {};
  let targetId = destMap[baseLocName];
  
  if (!targetId) {
    console.warn(`⚠️ ${baseLocName} の出力先が見つかりません。デフォルトで関東用のファイルに出力します。`);
    targetId = "19Q-xVsMX0thz_rvmdo_GERJvhjFLQgw2pUkWaTYxQhE"; // 未分類は安全のため関東へ
  }
  const targetSs = SpreadsheetApp.openById(targetId);

  // マスタ情報（開院日）の取得
  const openDatesMap = getClinicOpeningDates(masterSs);
  const clinicOpenDate = openDatesMap[baseLocName];

  let baseMonths = [];
  if (targetTerm === "上期" || targetTerm === "通年") {
     for (let m = 4; m <= 9; m++) baseMonths.push(`${targetYear}/${('0' + m).slice(-2)}`);
  }
  if (targetTerm === "下期" || targetTerm === "通年") {
     let nextYear = targetYear + 1;
     for (let m = 10; m <= 12; m++) baseMonths.push(`${targetYear}/${('0' + m).slice(-2)}`);
     for (let m = 1; m <= 3; m++) baseMonths.push(`${nextYear}/${('0' + m).slice(-2)}`);
  }

  let targetMonths = baseMonths;
  if (clinicOpenDate) {
    let openYearMonthVal = (clinicOpenDate.getFullYear() * 100) + (clinicOpenDate.getMonth() + 1);
    targetMonths = targetMonths.filter(monthStr => {
      let parts = monthStr.split('/');
      let targetYearMonthVal = (parseInt(parts[0], 10) * 100) + parseInt(parts[1], 10);
      return targetYearMonthVal >= openYearMonthVal;
    });
  }

  const currentDate = new Date();
  const currentMonthNum = currentDate.getMonth() + 1;
  const currentYearNum = currentDate.getFullYear();
  const sheetExists = targetSs.getSheetByName(finalSheetName) !== null; // ★ 外部ファイル側に存在するかチェック

  if (sheetExists) {
    let currentYearMonthVal = (currentYearNum * 100) + currentMonthNum;
    targetMonths = targetMonths.filter(monthStr => {
      let parts = monthStr.split('/');
      let targetYearMonthVal = (parseInt(parts[0], 10) * 100) + parseInt(parts[1], 10);
      return targetYearMonthVal >= currentYearMonthVal;
    });
  }

  if (queue.currentMonthIndex === 0 && typeof updateProgressMonitor === 'function') {
    let remainMin = Math.ceil((queue.locations.length * Math.max(1, targetMonths.length)) * 0.1); 
    updateProgressMonitor(masterSs, queue.completedCount, queue.totalCount, remainMin, `[ ${subLocName} ] を外部ファイルへ展開中...`);
  }

  console.log(`⏳ データをロード中...`);
  let extractedData = typeof fetchAndOrganizeData === 'function' ? fetchAndOrganizeData(targetYear, targetTerm, [baseLocName]) : {};
  let specialtyMap = typeof buildDoctorSpecialtyMap === 'function' ? buildDoctorSpecialtyMap(targetYear) : {};
  let shiftData = extractedData.shifts || {};
  let originalDataForMonth = shiftData[baseLocName] || {};
  let isSplitTarget = (baseLocName === "亀有" || baseLocName === "北葛西");
  let targetCat = subLocName.includes("内科") ? "内科" : (subLocName.includes("小児科") ? "小児科" : "");

  while (queue.currentMonthIndex < targetMonths.length) {
    let yearMonthStr = targetMonths[queue.currentMonthIndex];
    console.log(`⏳ [RENDER] ${subLocName} (${yearMonthStr}) を処理中...`);

    try {
      let filteredDataForMonth = {};
      for (let dStr in originalDataForMonth) {
        let shifts = originalDataForMonth[dStr];
        if (isSplitTarget && targetCat) {
          let otherCat = targetCat === "内科" ? "小児科" : "内科";
          shifts = shifts.filter(s => {
            let docSpec = specialtyMap[s.doctorName] || "不明";
            let text = s.rawShift || "";
            let isBoxMatch = text.includes(targetCat);
            let isOtherBoxMatch = text.includes(otherCat);
            let isSubjectMatch = docSpec.includes(targetCat);
            let isOtherSubjectMatch = docSpec.includes(otherCat);
            if (isOtherBoxMatch && !isBoxMatch) return false; 
            else if (isBoxMatch) return true; 
            else {
              if (targetCat === "内科") return isSubjectMatch; 
              else return !isOtherSubjectMatch;
            }
          });
        }
        if (shifts.length > 0) filteredDataForMonth[dStr] = shifts;
      }

      // ★ 書き出し先（targetSs）を渡して描画を実行！
      renderShiftBlock(targetSs, subLocName, finalSheetName, yearMonthStr, filteredDataForMonth, targetTerm);
      console.log(`✅ [SUCCESS] ${yearMonthStr} 完了`);
      queue.retryCount = 0; 
      
    } catch(e) {
      console.error(`❌ [エラー] ${subLocName} (${yearMonthStr}): ${e.message}`);
      if (e.message.includes("タイムアウト") || e.message.includes("timeout") || e.message.includes("Timeout")) {
        queue.retryCount = (queue.retryCount || 0) + 1;
        if (queue.retryCount > 2) {
          queue.locations.shift();
          queue.currentMonthIndex = 0;
          queue.retryCount = 0;
        }
        props.setProperty('BOSHUKUN_BATCH_QUEUE', JSON.stringify(queue));
        ScriptApp.newTrigger('processBatchQueue').timeBased().after(10000).create();
        return;
      }
    }

    queue.currentMonthIndex++;

    if (Date.now() - BATCH_START_TIME > TIME_LIMIT && queue.currentMonthIndex < targetMonths.length) {
      console.log(`⏱️ 4分経過。安全のため新しいトリガーにバトンタッチします`);
      props.setProperty('BOSHUKUN_BATCH_QUEUE', JSON.stringify(queue));
      ScriptApp.newTrigger('processBatchQueue').timeBased().after(2000).create();
      return;
    }
  }

  console.log(`🧹 [CLEANUP] ${subLocName} の全月が展開完了。色塗りとトリミングを行います。`);
  let finalSheet = targetSs.getSheetByName(finalSheetName); // ★ targetSs から取得
  if (finalSheet) {
    if (finalSheet.getMaxColumns() > 16) finalSheet.deleteColumns(17, finalSheet.getMaxColumns() - 16);
    let lastDataRow = finalSheet.getLastRow();
    let maxSheetRows = finalSheet.getMaxRows();
    let targetMaxRows = lastDataRow + 3;
    if (maxSheetRows > targetMaxRows) {
      finalSheet.deleteRows(targetMaxRows + 1, maxSheetRows - targetMaxRows);
    }
    safeExecute(() => syncSheetIndependent(finalSheet), 3, "書式・プルダウン同期");
    SpreadsheetApp.flush(); 
    
    if (typeof computeStableHash === 'function') {
       let rawCache = typeof getMasterRawData === 'function' ? getMasterRawData(targetYear) : {};
       let newHash = computeStableHash(targetYear, subLocName, rawCache);
       props.setProperty('HASH_' + targetYear + '_' + subLocName, newHash);
    }
  }

  queue.locations.shift(); 
  queue.currentMonthIndex = 0;
  queue.retryCount = 0;
  queue.completedCount++;
  
  if (queue.locations.length > 0) {
    props.setProperty('BOSHUKUN_BATCH_QUEUE', JSON.stringify(queue));
    console.log(`➡️ 拠点が完了しました。30秒間待機した後に次の拠点へ進みます。`);
    ScriptApp.newTrigger('processBatchQueue').timeBased().after(30000).create();
  } else {
    if (typeof updateProgressMonitor === 'function') {
      updateProgressMonitor(masterSs, queue.totalCount, queue.totalCount, 0, "🎉 すべての外部出力が正常に完了しました！");
    }
    props.deleteProperty('BOSHUKUN_BATCH_QUEUE');
    console.log(`🎉 全拠点のバッチ処理が完了しました。`);
  }
}

function getClinicOpeningDates(ss) {
  let map = {};
  const masterSheet = ss.getSheetByName("初期設定");
  if (!masterSheet) return map;
  const data = masterSheet.getDataRange().getValues();
  if (data.length < 2) return map;
  let dateColIdx = -1, nameColIdx = -1;
  for (let r = 0; r < Math.min(5, data.length); r++) {
    for (let c = 0; c < data[r].length; c++) {
      let val = String(data[r][c]);
      if (val.includes("開院")) dateColIdx = c;
      if (val.includes("拠点") || val.includes("クリニック") || val.includes("施設")) nameColIdx = c;
    }
    if (dateColIdx !== -1 && nameColIdx !== -1) break;
  }
  if (dateColIdx !== -1 && nameColIdx !== -1) {
    for (let r = 1; r < data.length; r++) {
      let locName = String(data[r][nameColIdx]).trim();
      let rawDate = data[r][dateColIdx];
      if (locName) {
        if (rawDate instanceof Date) map[locName] = rawDate;
        else if (rawDate && String(rawDate).trim() !== "") map[locName] = new Date(rawDate);
        else map[locName] = null;
      }
    }
  }
  return map;
}

function buildDoctorSpecialtyMap(year) {
  const masterUrl = 'https://docs.google.com/spreadsheets/d/1aEjphEv_63SeWQmwiOy9sx7IrMfawU01sHbKd_Ki4iA/edit';
  let map = {};
  try {
    const masterSs = SpreadsheetApp.openByUrl(masterUrl);
    ["常勤", "定期非常勤"].forEach(type => {
      const sheet = masterSs.getSheetByName(`${type}${year}年度`);
      if (sheet) {
        const data = masterSs.getSheetByName(`${type}${year}年度`).getDataRange().getValues();
        if (data.length < 2) return;
        const nameIdx = data[0].indexOf("医師名");
        const subjIdx = data[0].findIndex(h => String(h).includes("専門") || String(h).includes("科目"));
        if (nameIdx !== -1 && subjIdx !== -1) {
          for (let r = 1; r < data.length; r++) {
            let docName = String(data[r][nameIdx]).replace(/先生$/, "").replace(/[\s ]+/g, "").trim();
            let spec = String(data[r][subjIdx]).trim();
            if (docName) {
              if (spec.includes("内科")) map[docName] = "内科";
              else if (spec.includes("小児科")) map[docName] = "小児科";
              else map[docName] = "その他";
            }
          }
        }
      }
    });
  } catch(e) {}
  return map;
}

function deleteTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'processBatchQueue') {
      ScriptApp.deleteTrigger(t);
    }
  });
}