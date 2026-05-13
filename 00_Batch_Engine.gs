/**
 * ==========================================
 * 00_Batch_Engine.gs
 * バックグラウンドのバッチ処理・キュー管理
 * ★既存シート完全防衛版（タイムアウト時の白紙化防止）
 * ==========================================
 */

function startBackgroundBatch(payload) {
  const props = PropertiesService.getScriptProperties();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  payload.totalCount = payload.locations.length;
  payload.completedCount = 0;
  payload.currentLocation = "";
  payload.globalStartTime = Date.now(); 
  
  props.setProperty('BOSHUKUN_BATCH_QUEUE', JSON.stringify(payload));
  deleteTriggers(); 
  
  // モニター画面を立ち上げて状況を可視化
  createProgressMonitor(ss, payload.totalCount);
  updateProgressMonitor(ss, 0, payload.totalCount, "計算中", "準備中...");
  
  ScriptApp.newTrigger('processBatchQueue').timeBased().after(1000).create();
  return "Success";
}

function processBatchQueue() {
  deleteTriggers(); 

  const props = PropertiesService.getScriptProperties();
  const queueStr = props.getProperty('BOSHUKUN_BATCH_QUEUE');
  if (!queueStr) return;
  
  let queue = JSON.parse(queueStr);
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!queue.locations || queue.locations.length === 0) {
    safeGenerateAreaIndexSheets(ss);
    props.deleteProperty('BOSHUKUN_BATCH_QUEUE');
    return;
  }

  const startTime = Date.now();
  const targetYear = queue.year;
  const targetTerm = queue.term;

  // 過去の月は無視し、今月以降の月だけを処理対象にする
  let targetMonths = [];
  const currentDate = new Date();
  const currentMonthNum = currentDate.getMonth() + 1;
  const currentYearNum = currentDate.getFullYear();
  
  if (targetTerm === "上期" || targetTerm === "通年") {
    for (let m = 4; m <= 9; m++) {
      let calcYear = parseInt(targetYear, 10);
      if (calcYear > currentYearNum || (calcYear === currentYearNum && m >= currentMonthNum)) {
        targetMonths.push(`${calcYear}/${('0' + m).slice(-2)}`);
      }
    }
  }
  if (targetTerm === "下期" || targetTerm === "通年") {
    let nextYear = parseInt(targetYear, 10) + 1;
    for (let m = 10; m <= 12; m++) {
      let calcYear = parseInt(targetYear, 10);
      if (calcYear > currentYearNum || (calcYear === currentYearNum && m >= currentMonthNum)) {
        targetMonths.push(`${calcYear}/${('0' + m).slice(-2)}`);
      }
    }
    for (let m = 1; m <= 3; m++) {
      if (nextYear > currentYearNum || (nextYear === currentYearNum && m >= currentMonthNum)) {
        targetMonths.push(`${nextYear}/${('0' + m).slice(-2)}`);
      }
    }
  }
  targetMonths.sort((a, b) => new Date(a + "/01") - new Date(b + "/01"));
  
  if (targetMonths.length === 0) {
    if (targetTerm === "上期" || targetTerm === "通年") {
      for (let m = 4; m <= 9; m++) targetMonths.push(`${targetYear}/${('0' + m).slice(-2)}`);
    }
    if (targetTerm === "下期" || targetTerm === "通年") {
      let nextYear = parseInt(targetYear, 10) + 1;
      for (let m = 10; m <= 12; m++) targetMonths.push(`${targetYear}/${('0' + m).slice(-2)}`);
      for (let m = 1; m <= 3; m++) targetMonths.push(`${nextYear}/${('0' + m).slice(-2)}`);
    }
    targetMonths.sort((a, b) => new Date(a + "/01") - new Date(b + "/01"));
  }

  // ★大改修：新規作成の負荷でスプレッドシートがタイムアウトしないよう、
  // 1回の処理拠点数を「1」に絞り、裏側で確実にバトンタッチしていく設定
  const CHUNK_SIZE = 1;
  let currentBatchLocs = queue.locations.splice(0, CHUNK_SIZE);
  
  let baseLocsForFetch = currentBatchLocs.map(l => l.replace(/（.*?）/, ''));
  baseLocsForFetch = [...new Set(baseLocsForFetch)];
  
  // =======================================================================
  // ★修正箇所1：エラーを握りつぶさず、失敗フラグを立ててログに記録する
  // =======================================================================
  let extractedData = {};
  let isFetchFailed = false;

  try {
    extractedData = typeof fetchAndOrganizeData === 'function' ? fetchAndOrganizeData(targetYear, targetTerm, baseLocsForFetch) : {};
  } catch (e) {
    console.error(`[🚨CRITICAL] ${baseLocsForFetch.join(', ')}のデータ取得中にタイムアウト等のエラーが発生しました: ${e.message}`);
    isFetchFailed = true; 
  }
  const shiftData = extractedData.shifts || {};
  // =======================================================================
  
  let rawCache = typeof getMasterRawData === 'function' ? getMasterRawData(targetYear) : {};
  let specialtyMap = typeof buildDoctorSpecialtyMap === 'function' ? buildDoctorSpecialtyMap(targetYear) : {};

  while (currentBatchLocs.length > 0) {
    
    // 緊急停止の監視
    try {
      let monitorSheet = ss.getSheetByName("🗂️進行中モニター");
      if (monitorSheet && monitorSheet.getMaxRows() >= 2 && monitorSheet.getMaxColumns() >= 1) {
        let isStopped = monitorSheet.getRange("A2").getValue();
        if (isStopped === true) {
          monitorSheet.getRange("A1:B1").setValue("🛑 処理を緊急停止しました").setBackground("#ea4335");
          props.deleteProperty('BOSHUKUN_BATCH_QUEUE');
          return; 
        }
      }
    } catch(e) {}

    // ★大改修：タイムアウトエラー(約250〜300秒)の前に安全に撤退するよう、
    // 180秒（3分）経過したら強制的に次のバッチに引き継ぐ
    if (Date.now() - startTime > 180000) {
      queue.locations = currentBatchLocs.concat(queue.locations);
      break;
    }

    const subLocName = currentBatchLocs.shift(); 
    const baseLocName = subLocName.replace(/（.*?）/, ''); 
    
    let elapsedMs = Date.now() - queue.globalStartTime;
    let timePerItem = queue.completedCount > 0 ? (elapsedMs / queue.completedCount) : 15000;
    let remainingItems = queue.totalCount - queue.completedCount;
    let remainMin = Math.ceil((timePerItem * remainingItems) / 60000);
    
    updateProgressMonitor(ss, queue.completedCount, queue.totalCount, remainMin, subLocName);
    
    queue.currentLocation = subLocName;
    props.setProperty('BOSHUKUN_BATCH_QUEUE', JSON.stringify(queue));

    try {
      const finalSheetName = `${targetYear}${subLocName}`;

      // =======================================================================
      // ★修正箇所2：最強の防衛線。取得失敗時や空データの時は絶対にシートに触れずスキップ
      // =======================================================================
      if (isFetchFailed || !shiftData[baseLocName] || Object.keys(shiftData[baseLocName]).length === 0) {
        console.warn(`[🛡️PROTECT 発動] ${subLocName} : 取得データが空です（タイムアウト等）。既存のシフトデータを守るため、シートの更新を完全スキップしました。`);
        // モニター画面にもスキップしたことを通知
        try {
          let monitorSheet = ss.getSheetByName("🗂️進行中モニター");
          if (monitorSheet) monitorSheet.getRange("A1:B1").setValue(`⚠️ 取得エラー発生\n[ ${subLocName} ] の既存データを保護してスキップしました...`);
          Utilities.sleep(1500); // ユーザーにスキップを視認させるための深呼吸
        } catch(e) {}
        
        throw new Error("FETCH_FAILED_PROTECTION"); // 意図的に下のcatchへ飛ばして描画処理を回避
      }
      // =======================================================================

      let hasValidMonth = false;

      targetMonths.forEach(yearMonthStr => {
        let originalDataForMonth = shiftData[baseLocName] ? shiftData[baseLocName] : {};
        let filteredDataForMonth = {};

        let isSplitTarget = (baseLocName === "亀有" || baseLocName === "北葛西");
        let targetCat = subLocName.includes("内科") ? "内科" : (subLocName.includes("小児科") ? "小児科" : "");

        for (let dStr in originalDataForMonth) {
          let shifts = originalDataForMonth[dStr];
          
          if (isSplitTarget && targetCat) {
            shifts = shifts.filter(s => {
              let docSpec = specialtyMap[s.doctorName] || "不明";
              if (docSpec === "内科" && targetCat === "内科") return true;
              if (docSpec === "小児科" && targetCat === "小児科") return true;
              if (docSpec === "内科" && targetCat === "小児科") return false;
              if (docSpec === "小児科" && targetCat === "内科") return false;
              let text = s.rawShift || "";
              if (text.includes(targetCat)) return true;
              let otherCat = targetCat === "内科" ? "小児科" : "内科";
              if (text.includes(otherCat)) return false;
              return true; 
            });
          }
          if (shifts.length > 0) filteredDataForMonth[dStr] = shifts;
        }

        const isRendered = renderShiftBlock(ss, subLocName, finalSheetName, yearMonthStr, filteredDataForMonth);
        if (isRendered) hasValidMonth = true;
      });

      let finalSheet = ss.getSheetByName(finalSheetName);
      if (finalSheet) {
        if (hasValidMonth) {
          if (finalSheet.getMaxColumns() > 16) finalSheet.deleteColumns(17, finalSheet.getMaxColumns() - 16);
          
          let lastRow = finalSheet.getLastRow();
          let rules = finalSheet.getConditionalFormatRules();
          if (rules.length === 0 || lastRow < 50) {
            syncSheetIndependent(finalSheet);
          }
          
          if (typeof computeStableHash === 'function') {
             let newHash = computeStableHash(targetYear, subLocName, rawCache);
             props.setProperty('HASH_' + targetYear + '_' + subLocName, newHash);
          }
        }
      }
    } catch(e) {
      // =======================================================================
      // ★修正箇所3：保護スキップ時はエラーログとして吐かず、成功ログとして残す
      // =======================================================================
      if (e.message === "FETCH_FAILED_PROTECTION") {
        console.log(`✅ [PROTECT 完了] ${subLocName} の既存データは白紙化されず、安全に維持されました。`);
      } else {
        console.error(`Error processing ${subLocName}: ${e.message}`);
      }
    }

    queue.completedCount++;
    queue.currentLocation = ""; 
    props.setProperty('BOSHUKUN_BATCH_QUEUE', JSON.stringify(queue));
  }

  if (queue.locations.length > 0) {
    props.setProperty('BOSHUKUN_BATCH_QUEUE', JSON.stringify(queue));
    ScriptApp.newTrigger('processBatchQueue').timeBased().after(60000).create();
  } else {
    queue.status = "COMPLETED";
    props.setProperty('BOSHUKUN_BATCH_QUEUE', JSON.stringify(queue));
    safeGenerateAreaIndexSheets(ss);
  }
}

function buildDoctorSpecialtyMap(year) {
  const masterUrl = 'https://docs.google.com/spreadsheets/d/1aEjphEv_63SeWQmwiOy9sx7IrMfawU01sHbKd_Ki4iA/edit';
  let map = {};
  try {
    const masterSs = SpreadsheetApp.openByUrl(masterUrl);
    ["常勤", "定期非常勤"].forEach(type => {
      const sheet = masterSs.getSheetByName(`${type}${year}年度`);
      if (sheet) {
        const data = sheet.getDataRange().getValues();
        if (data.length < 2) return;
        const nameIdx = data[0].indexOf("医師名");
        const subjIdx = data[0].findIndex(h => String(h).includes("専門") || String(h).includes("科目"));
        if (nameIdx !== -1 && subjIdx !== -1) {
          for (let r = 1; r < data.length; r++) {
            let docName = String(data[r][nameIdx]).replace(/先生$/, "").trim();
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