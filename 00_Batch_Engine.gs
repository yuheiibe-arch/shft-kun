/**
 * ==========================================
 * 00_Batch_Engine.gs
 * バックグラウンドのバッチ処理・キュー管理
 * ★【完全版】時間軸フィルター＆開院日未設定スキップ対応
 * ==========================================
 */

function startBackgroundBatch(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();
  
  if (!payload || !payload.locations || payload.locations.length === 0) {
    return "Error: Locations Empty";
  }

  payload.totalCount = payload.locations.length;
  payload.completedCount = 0;
  payload.globalStartTime = Date.now(); 
  payload.currentMonthIndex = 0; 
  
  props.setProperty('BOSHUKUN_BATCH_QUEUE', JSON.stringify(payload));
  deleteTriggers(); 
  
  if (typeof createProgressMonitor === 'function') createProgressMonitor(ss, payload.totalCount);
  if (typeof updateProgressMonitor === 'function') {
    updateProgressMonitor(ss, 0, payload.totalCount, payload.totalCount * 2, "準備中...（開院日マスタと時間軸を確認中）");
  }
  
  ScriptApp.newTrigger('processBatchQueue').timeBased().after(1000).create();
  return "Success";
}

function processBatchQueue() {
  const BATCH_START_TIME = Date.now();
  const SAFE_TIME_LIMIT = 270000; // 4.5分限界まで連続処理

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

  // 緊急停止チェック
  try {
    let monitorSheet = ss.getSheetByName("🗂️進行中モニター");
    if (monitorSheet && monitorSheet.getRange("A2").getValue() === true) {
      monitorSheet.getRange("A1:B1").setValue("🛑 処理を緊急停止しました").setBackground("#ea4335");
      props.deleteProperty('BOSHUKUN_BATCH_QUEUE');
      return; 
    }
  } catch(e) {}

  const targetYear = parseInt(queue.year, 10);
  const targetTerm = queue.term;
  const currentDate = new Date();
  const currentMonthNum = currentDate.getMonth() + 1;
  const currentYearNum = currentDate.getFullYear();

  const subLocName = queue.locations[0]; 
  const baseLocName = subLocName.replace(/（.*?）/, ''); 
  const finalSheetName = `${targetYear}${subLocName}`;
  const sheetExists = ss.getSheetByName(finalSheetName) !== null;

  if (queue.currentMonthIndex === undefined) queue.currentMonthIndex = 0;

  // =========================================================
  // ★ マスタから開院日を取得し、時間軸フィルターをかける
  // =========================================================
  const openDatesMap = getClinicOpeningDates(ss);
  const clinicOpenDate = openDatesMap[baseLocName]; // Dateオブジェクト または null

  let targetMonths = [];
  
  if (!clinicOpenDate) {
    // 開院日が未設定の場合は完全スキップ（出力不要）
    console.log(`[スキップ] ${subLocName} は開院日が未設定のため出力対象外です。`);
    targetMonths = []; 
  } else {
    // 1. 選択された期間（上期/下期/通年）のベースとなる月を全生成
    let baseMonths = [];
    if (targetTerm === "上期" || targetTerm === "通年") {
       for (let m = 4; m <= 9; m++) baseMonths.push(`${targetYear}/${('0' + m).slice(-2)}`);
    }
    if (targetTerm === "下期" || targetTerm === "通年") {
       let nextYear = targetYear + 1;
       for (let m = 10; m <= 12; m++) baseMonths.push(`${targetYear}/${('0' + m).slice(-2)}`);
       for (let m = 1; m <= 3; m++) baseMonths.push(`${nextYear}/${('0' + m).slice(-2)}`);
    }

    // 2. 開院月以降の月だけに絞り込む（時間軸フィルター）
    let openYearMonthVal = (clinicOpenDate.getFullYear() * 100) + (clinicOpenDate.getMonth() + 1);
    targetMonths = baseMonths.filter(monthStr => {
      let parts = monthStr.split('/');
      let targetYearMonthVal = (parseInt(parts[0], 10) * 100) + parseInt(parts[1], 10);
      return targetYearMonthVal >= openYearMonthVal;
    });

    // 3. 既にシートが存在する場合は「過去月」を上書きから保護する
    if (sheetExists) {
      let currentYearMonthVal = (currentYearNum * 100) + currentMonthNum;
      targetMonths = targetMonths.filter(monthStr => {
        let parts = monthStr.split('/');
        let targetYearMonthVal = (parseInt(parts[0], 10) * 100) + parseInt(parts[1], 10);
        return targetYearMonthVal >= currentYearMonthVal;
      });
    }
  }

  // データ取得
  let extractedData = {};
  try {
    extractedData = typeof fetchAndOrganizeData === 'function' ? fetchAndOrganizeData(targetYear, targetTerm, [baseLocName]) : {};
  } catch (e) {}
  const shiftData = extractedData.shifts || {};
  let rawCache = typeof getMasterRawData === 'function' ? getMasterRawData(targetYear) : {};
  let specialtyMap = typeof buildDoctorSpecialtyMap === 'function' ? buildDoctorSpecialtyMap(targetYear) : {};

  let originalDataForMonth = shiftData[baseLocName] || {};
  let isSplitTarget = (baseLocName === "亀有" || baseLocName === "北葛西");
  let targetCat = subLocName.includes("内科") ? "内科" : (subLocName.includes("小児科") ? "小児科" : "");

  let isRenderedAny = false;

  // モニター更新
  if (queue.currentMonthIndex === 0 && typeof updateProgressMonitor === 'function') {
    let msg = targetMonths.length > 0 ? `[ ${subLocName} ] を展開中... (${targetMonths.length}ヶ月分)` : `[ ${subLocName} ] は出力不要のためスキップします`;
    let remainMin = Math.ceil((queue.locations.length * Math.max(1, targetMonths.length)) * 0.15); 
    updateProgressMonitor(ss, queue.completedCount, queue.totalCount, remainMin, msg);
  }

  // =========================================================
  // ★ 連続描画ループ
  // =========================================================
  while (queue.currentMonthIndex < targetMonths.length) {
    if (Date.now() - BATCH_START_TIME > SAFE_TIME_LIMIT) {
      props.setProperty('BOSHUKUN_BATCH_QUEUE', JSON.stringify(queue));
      ScriptApp.newTrigger('processBatchQueue').timeBased().after(1000).create();
      return; 
    }

    let yearMonthStr = targetMonths[queue.currentMonthIndex];

    try {
      let filteredDataForMonth = {};
      for (let dStr in originalDataForMonth) {
        let shifts = originalDataForMonth[dStr];
        if (isSplitTarget && targetCat) {
          shifts = shifts.filter(s => {
            let docSpec = specialtyMap[s.doctorName] || "不明";
            if (docSpec === targetCat) return true;
            let text = s.rawShift || "";
            if (text.includes(targetCat)) return true;
            let otherCat = targetCat === "内科" ? "小児科" : "内科";
            if (text.includes(otherCat)) return false;
            return true; 
          });
        }
        if (shifts.length > 0) filteredDataForMonth[dStr] = shifts;
      }

      // フィルターを生き残った月は、データが0件でも確実に生成する
      const isRendered = renderShiftBlock(ss, subLocName, finalSheetName, yearMonthStr, filteredDataForMonth, targetTerm);
      if (isRendered) isRenderedAny = true;
      
    } catch(e) {
      console.error(`[エラー] ${subLocName} (${yearMonthStr}): ${e.message}`);
    }

    queue.currentMonthIndex++;
  }

  // =========================================================
  // ★ 1拠点完了時の色塗りと不要列削除
  // =========================================================
  if (queue.currentMonthIndex >= targetMonths.length) {
    let finalSheet = ss.getSheetByName(finalSheetName);
    if (finalSheet && isRenderedAny) {
      if (finalSheet.getMaxColumns() > 16) finalSheet.deleteColumns(17, finalSheet.getMaxColumns() - 16);
      safeExecute(() => syncSheetIndependent(finalSheet), 3, "書式・プルダウン同期");
      
      if (typeof computeStableHash === 'function') {
         let newHash = computeStableHash(targetYear, subLocName, rawCache);
         props.setProperty('HASH_' + targetYear + '_' + subLocName, newHash);
      }
    }

    queue.locations.shift(); 
    queue.currentMonthIndex = 0;
    queue.completedCount++;
  }

  if (queue.locations.length > 0) {
    props.setProperty('BOSHUKUN_BATCH_QUEUE', JSON.stringify(queue));
    ScriptApp.newTrigger('processBatchQueue').timeBased().after(1000).create();
  } else {
    if (typeof updateProgressMonitor === 'function') {
      updateProgressMonitor(ss, queue.totalCount, queue.totalCount, 0, "🎉 すべての出力が正常に完了しました！");
    }
    queue.status = "COMPLETED";
    props.setProperty('BOSHUKUN_BATCH_QUEUE', JSON.stringify(queue));
    safeGenerateAreaIndexSheets(ss);
  }
}

// =========================================================
// ★ 初期設定マスタから全開院日をマップ化（未設定はnull）
// =========================================================
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
        if (rawDate instanceof Date) {
          map[locName] = rawDate;
        } else if (rawDate && String(rawDate).trim() !== "") {
          map[locName] = new Date(rawDate);
        } else {
          map[locName] = null; // 未設定は明確に弾く
        }
      }
    }
  }
  return map;
}

function buildDoctorSpecialtyMap(year) {
  // 既存の処理
  const masterUrl = 'https://docs.google.com/spreadsheets/d/1aEjphEv_63SeWQmwiOy9sx7IrMfawU01sHbKd_Ki4iA/edit';
  let map = {};
  try {
    const masterSs = safeOpenByUrl(masterUrl);
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