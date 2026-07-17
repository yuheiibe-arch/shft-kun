/**
 * ==========================================
 * 00_Batch_Engine.gs
 * バックグラウンドのバッチ処理・キュー管理
 * ★【究極のタイムアウト対策＆白紙化防止版】
 * 強制セーブ(Flush)・フライング防止・開院日スキップ解除パッチ適用版
 * ★ 誤爆絶対防衛ガード適用（管理用・機能用シートの混入を完全ブロック）
 * ★【究極のV8メモリパンク対策】限界時間を1.5分に短縮しINTERNALエラーを根絶
 * ★【単一ターゲット・ステルス化】全シート走査・行削除・強制セーブの完全排除
 * ★【定期シフト】内科・小児科の絶対分類ルール完全適用版
 * ★【ゴミシート完全排除】タイムアウト連鎖遮断＆自動削除パッチ適用版
 * ★【APIデッドロック回避】トリガー間隔を10秒に延長しGoogleサーバーのロックを解除
 * ★【スマート・トリミング】描画完了時に余白3行を残して不要な空白行を自動削除
 * ★【二次災害防止＆リトライ】タイムアウト後のクラッシュ防衛と自動再挑戦機構
 * ==========================================
 */

function startBackgroundBatch(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();
  
  if (!payload || !payload.locations || payload.locations.length === 0) {
    return "Error: Locations Empty";
  }

  // =========================================================
  // 🛡️【誤爆絶対防衛ガード・強化版】
  // =========================================================
  payload.locations = payload.locations.filter(loc => {
    let invalidWords = [
      "その他", "キャッシュ", "作業表", "お休み", "確定シフト", 
      "欠勤", "キャンセル", "手順書", "原本", "振替勤務", 
      "先行応募", "単独募集", "定期募集", "目次", "設定", 
      "テンプレート", "ダッシュボード"
    ];
    return !invalidWords.some(word => loc.includes(word));
  });

  if (payload.locations.length === 0) {
    return "Error: 有効な出力対象拠点がありませんでした。（すべて管理用シートとして除外されました）";
  }
  // =========================================================

  payload.totalCount = payload.locations.length;
  payload.completedCount = 0;
  payload.globalStartTime = Date.now(); 
  payload.currentMonthIndex = 0; 
  
  props.setProperty('BOSHUKUN_BATCH_QUEUE', JSON.stringify(payload));
  deleteTriggers();

  if (typeof createProgressMonitor === 'function') createProgressMonitor(ss, payload.totalCount);
  if (typeof updateProgressMonitor === 'function') {
    updateProgressMonitor(ss, 0, payload.totalCount, payload.totalCount * 2, "準備中...（全拠点のデータを一括ダウンロード中...しばらくお待ちください）");
  }

  // =========================================================
  // ★【究極のタイムアウト対策】最初の1回だけ全データを一括取得し、シートにキャッシュする
  // =========================================================
  try {
    const targetYear = parseInt(payload.year, 10);
    const baseLocs = [...new Set(payload.locations.map(loc => loc.replace(/（.*?）/, '')))];
    
    // 外部から一括取得
    let extractedData = typeof fetchAndOrganizeData === 'function' ? fetchAndOrganizeData(targetYear, payload.term, baseLocs) : {};
    let rawCache = typeof getMasterRawData === 'function' ? getMasterRawData(targetYear) : {};
    let specialtyMap = typeof buildDoctorSpecialtyMap === 'function' ? buildDoctorSpecialtyMap(targetYear) : {};

    // 隠しシートに保存
    _saveBatchCache(ss, {
      extractedData: extractedData,
      rawCache: rawCache,
      specialtyMap: specialtyMap
    });
  } catch(e) {
    if (typeof updateProgressMonitor === 'function') {
      updateProgressMonitor(ss, 0, payload.totalCount, 0, `🚨 データの一括取得に失敗しました。\n時間をおいて再度お試しください。\nエラー: ${e.message}`);
    }
    props.deleteProperty('BOSHUKUN_BATCH_QUEUE');
    return "Error: " + e.message;
  }
  
  // ★修正：Googleの書き込み完了（ロック解除）を待つため 10秒(10000ms) に延長
  ScriptApp.newTrigger('processBatchQueue').timeBased().after(10000).create();
  return "Success";
}

function processBatchQueue() {
  const BATCH_START_TIME = Date.now();
  
  // =========================================================
  // ★ 大修正：INTERNALエラー（V8メモリパンク）を完全に防ぐため、
  // 連続稼働の限界を 180000(3分) → 90000(1.5分) に超短縮！
  // =========================================================
  const SAFE_TIME_LIMIT = 90000; 

  deleteTriggers(); 
  const props = PropertiesService.getScriptProperties();
  const queueStr = props.getProperty('BOSHUKUN_BATCH_QUEUE');
  if (!queueStr) return;
  
  let queue = JSON.parse(queueStr);
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  if (!queue.locations || queue.locations.length === 0) {
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
  // リトライカウントの初期化
  if (queue.retryCount === undefined) queue.retryCount = 0;

  // =========================================================
  // ★ マスタから開院日を取得し、時間軸フィルターをかける
  // =========================================================
  const openDatesMap = getClinicOpeningDates(ss);
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

  if (sheetExists) {
    let currentYearMonthVal = (currentYearNum * 100) + currentMonthNum;
    targetMonths = targetMonths.filter(monthStr => {
      let parts = monthStr.split('/');
      let targetYearMonthVal = (parseInt(parts[0], 10) * 100) + parseInt(parts[1], 10);
      return targetYearMonthVal >= currentYearMonthVal;
    });
  }

  let cached = _loadBatchCache(ss);
  if (!cached) {
    console.error(`[致命的エラー] キャッシュの読み込みに失敗。白紙化を防ぐためスキップします。`);
    queue.locations.shift();
    queue.currentMonthIndex = 0;
    props.setProperty('BOSHUKUN_BATCH_QUEUE', JSON.stringify(queue));
    ScriptApp.newTrigger('processBatchQueue').timeBased().after(10000).create();
    return;
  }

  let shiftData = (cached.extractedData && cached.extractedData.shifts) ? cached.extractedData.shifts : {};
  let rawCache = cached.rawCache || {};
  let specialtyMap = cached.specialtyMap || {};

  let originalDataForMonth = shiftData[baseLocName] || {};
  let isSplitTarget = (baseLocName === "亀有" || baseLocName === "北葛西");
  let targetCat = subLocName.includes("内科") ? "内科" : (subLocName.includes("小児科") ? "小児科" : "");

  let isRenderedAny = false;
  let isTimeoutOccurred = false; 

  if (queue.currentMonthIndex === 0 && typeof updateProgressMonitor === 'function') {
    let msg = targetMonths.length > 0 ?
      `[ ${subLocName} ] を展開中... (${targetMonths.length}ヶ月分)` : `[ ${subLocName} ] は出力不要のためスキップします`;
    let remainMin = Math.ceil((queue.locations.length * Math.max(1, targetMonths.length)) * 0.15);
    updateProgressMonitor(ss, queue.completedCount, queue.totalCount, remainMin, msg);
  }

  while (queue.currentMonthIndex < targetMonths.length) {
    if (Date.now() - BATCH_START_TIME > SAFE_TIME_LIMIT) {
      props.setProperty('BOSHUKUN_BATCH_QUEUE', JSON.stringify(queue));
      ScriptApp.newTrigger('processBatchQueue').timeBased().after(10000).create();
      return; 
    }

    let yearMonthStr = targetMonths[queue.currentMonthIndex];

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
              if (targetCat === "内科") {
                if (!isSubjectMatch) return false; 
              } else {
                if (isOtherSubjectMatch) return false; 
              }
              return true;
            }
          });
        }
        if (shifts.length > 0) filteredDataForMonth[dStr] = shifts;
      }

      const isRendered = renderShiftBlock(ss, subLocName, finalSheetName, yearMonthStr, filteredDataForMonth, targetTerm);
      if (isRendered) {
        isRenderedAny = true;
        queue.retryCount = 0; // 成功したらリセット
      }
      
    } catch(e) {
      console.error(`[エラー] ${subLocName} (${yearMonthStr}): ${e.message}`);
      if (e.message.includes("タイムアウト") || e.message.includes("timeout") || e.message.includes("Timeout")) {
        isTimeoutOccurred = true;
        break; 
      }
    }

    queue.currentMonthIndex++;
  }

  // =========================================================
  // ★ ゴミ掃除 ＆ リトライ処理（二次災害の完全ブロック）
  // =========================================================
  if (isTimeoutOccurred) {
    queue.retryCount = (queue.retryCount || 0) + 1;
    console.warn(`🚨 タイムアウト検知: ${subLocName} の処理を一時中断します。(リトライ: ${queue.retryCount}回目)`);
    
    // 二次災害（クラッシュ）を防ぐバリア
    try {
      if (!isRenderedAny) {
        let garbageSheet = ss.getSheetByName(finalSheetName);
        if (garbageSheet) ss.deleteSheet(garbageSheet);
      }
    } catch (e2) {
      console.warn("通信遮断状態のためゴミ削除をスキップしました（クラッシュ回避）。");
    }

    // 2回までは同じ月をリトライする。3回目で諦めてスキップする。
    if (queue.retryCount > 2) {
      console.error(`❌ リトライ上限到達: ${subLocName} を諦めてスキップします。`);
      queue.locations.shift();
      queue.currentMonthIndex = 0;
      queue.retryCount = 0;
    }
    
    props.setProperty('BOSHUKUN_BATCH_QUEUE', JSON.stringify(queue));
    // バックグラウンドの通信詰まりを解消するため長めに深呼吸
    ScriptApp.newTrigger('processBatchQueue').timeBased().after(15000).create();
    return;
  }

  // =========================================================
  // ★ 1拠点完了時
  // =========================================================
  if (queue.currentMonthIndex >= targetMonths.length) {
    let finalSheet = ss.getSheetByName(finalSheetName);
    if (finalSheet && isRenderedAny) {
      
      if (finalSheet.getMaxColumns() > 16) finalSheet.deleteColumns(17, finalSheet.getMaxColumns() - 16);
      
      let lastDataRow = finalSheet.getLastRow();
      let maxSheetRows = finalSheet.getMaxRows();
      let targetMaxRows = lastDataRow + 3;
      if (maxSheetRows > targetMaxRows) {
        finalSheet.deleteRows(targetMaxRows + 1, maxSheetRows - targetMaxRows);
      }
      
      safeExecute(() => syncSheetIndependent(finalSheet), 3, "書式・プルダウン同期");
      
      if (typeof computeStableHash === 'function') {
         let newHash = computeStableHash(targetYear, subLocName, rawCache);
         props.setProperty('HASH_' + targetYear + '_' + subLocName, newHash);
      }
    }

    queue.locations.shift(); 
    queue.currentMonthIndex = 0;
    queue.retryCount = 0;
    queue.completedCount++;
  }

  if (queue.locations.length > 0) {
    props.setProperty('BOSHUKUN_BATCH_QUEUE', JSON.stringify(queue));
    ScriptApp.newTrigger('processBatchQueue').timeBased().after(10000).create();
  } else {
    if (typeof updateProgressMonitor === 'function') {
      updateProgressMonitor(ss, queue.totalCount, queue.totalCount, 0, "🎉 すべての出力が正常に完了しました！");
    }
    queue.status = "COMPLETED";
    props.setProperty('BOSHUKUN_BATCH_QUEUE', JSON.stringify(queue));

    let cacheSheet = ss.getSheetByName("⚙️通信キャッシュ");
    if (cacheSheet) ss.deleteSheet(cacheSheet);
  }
}

// =========================================================
// ★ ローカルキャッシュ用 ヘルパー関数（データ容量制限の突破）
// =========================================================
function _saveBatchCache(ss, dataObj) {
  let sheet = ss.getSheetByName("⚙️通信キャッシュ");
  if (!sheet) {
    sheet = ss.insertSheet("⚙️通信キャッシュ");
    sheet.hideSheet();
  } else {
    sheet.clear();
  }
  let jsonStr = JSON.stringify(dataObj);
  let chunks = [];
  for (let i = 0; i < jsonStr.length; i += 45000) {
    chunks.push([jsonStr.substring(i, i + 45000)]);
  }
  if (chunks.length > 0) {
    sheet.getRange(1, 1, chunks.length, 1).setValues(chunks);
  }
}

function _loadBatchCache(ss) {
  let sheet = ss.getSheetByName("⚙️通信キャッシュ");
  if (!sheet) return null;
  let maxRow = sheet.getLastRow();
  if (maxRow === 0) return null;
  let data = sheet.getRange(1, 1, maxRow, 1).getValues();
  let jsonStr = data.map(r => r[0]).join('');
  if (!jsonStr) return null;
  try {
    return JSON.parse(jsonStr);
  } catch(e) {
    return null;
  }
}

// =========================================================
// ★ 開院日マップの取得
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
          map[locName] = null;
        }
      }
    }
  }
  return map;
}

// =========================================================
// ★ 専門科目マスタの取得
// =========================================================
function buildDoctorSpecialtyMap(year) {
  const masterUrl = 'https://docs.google.com/spreadsheets/d/1aEjphEv_63SeWQmwiOy9sx7IrMfawU01sHbKd_Ki4iA/edit';
  let map = {};
  try {
    const masterSs = safeOpenByUrl(masterUrl);
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