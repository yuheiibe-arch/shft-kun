/**
 * ==========================================
 * 04C_Render_Sync.gs
 * 自動同期・条件付き書式の反映・onEditイベント
 * ★ プルダウン一括設定（2次元配列によるユーザー様オリジナル爆速版 完全復活）
 * ★ ダッシュボードは8列制限（L列保護）、カレンダーは12列適用（色抜け修正）
 * ★ 【常勤 ＞ 非常勤 ＞ 先行】の優先順位で色を統一
 * ==========================================
 */

function onEdit(e) {
  if (!e) return;
  const range = e.range;
  const sheet = range.getSheet();
  const sheetName = sheet.getName();

  try {
    if (typeof CONFIG !== 'undefined') {
      if (sheetName === CONFIG.SETTING_SHEET_NAME || sheetName === CONFIG.TEMPLATE_SHEET_NAME) return;
    }
  } catch(e) {}
  if (sheetName === "🗂️シフト目次") return;

  const row = range.getRow();
  const col = range.getColumn();
  if (col >= 4 && col <= 11) {
    const labels = sheet.getRange(row, 1, 1, 3).getValues()[0];
    const rowLabelA = String(labels[0]).trim();
    
    if (rowLabelA === "常勤医師" || rowLabelA === "非常勤医師" || rowLabelA === "先行応募" || rowLabelA === "先行応募医師") {
      syncSheetIndependent(sheet);
    }
  }
}

function applyMasterListToAll(sheet, joukinList, teikiList, senkouList) {
  const data = sheet.getDataRange().getValues();
  const maxRows = data.length;
  if (maxRows === 0) return;
  
  const finalJ = [...new Set(joukinList)];
  const finalT = [...new Set(teikiList)];
  const finalS = [...new Set(senkouList || [])];
  
  const dropdownList = ["募集", "休館日", "未開院", ...finalJ, ...finalT, ...finalS].filter(Boolean);
  const newRule = SpreadsheetApp.newDataValidation().requireValueInList(dropdownList, true).build();

  const joukinColor = "#fce5cd"; 
  const teikiColor  = "#d9d2e9"; 
  const senkouColor = "#d9ead3";
  const emptyColor  = "#ffffff"; 

  let alignRanges = [];
  
  // ★ ダッシュボード上部（名前一覧）は8列に制限してL列を守る
  const MAX_DASHBOARD_COLS = 8;
  // ★ カレンダー（シフト枠）は色抜けを防ぐため12列（D〜O列）まで適用する
  const CALENDAR_COLS = 12;
  
  const targetRange = sheet.getRange(1, 4, maxRows, CALENDAR_COLS); 
  let currentRules = targetRange.getDataValidations();
  let rulesModified = false;
  
  for (let r = 0; r < maxRows; r++) {
    let rowNum = r + 1;
    let labelA = String(data[r][0]).trim(); 
    let labelC = String(data[r][2]).trim(); 
    
    // ダッシュボードへの書き込み（8列のみ）
    if (labelA === "常勤医師") {
      let vals = new Array(MAX_DASHBOARD_COLS).fill("");
      let bgs = new Array(MAX_DASHBOARD_COLS).fill(emptyColor);
      for (let i = 0; i < finalJ.length && i < MAX_DASHBOARD_COLS; i++) { vals[i] = finalJ[i]; bgs[i] = joukinColor; }
      sheet.getRange(rowNum, 4, 1, MAX_DASHBOARD_COLS).setValues([vals]).setBackgrounds([bgs]).setHorizontalAlignment("left");
    } 
    else if (labelA === "非常勤医師") {
      let vals = new Array(MAX_DASHBOARD_COLS).fill("");
      let bgs = new Array(MAX_DASHBOARD_COLS).fill(emptyColor);
      for (let i = 0; i < finalT.length && i < MAX_DASHBOARD_COLS; i++) { vals[i] = finalT[i]; bgs[i] = teikiColor; }
      sheet.getRange(rowNum, 4, 1, MAX_DASHBOARD_COLS).setValues([vals]).setBackgrounds([bgs]).setHorizontalAlignment("left");
    }
    else if (labelA === "先行応募" || labelA === "先行応募医師") {
      let vals = new Array(MAX_DASHBOARD_COLS).fill("");
      let bgs = new Array(MAX_DASHBOARD_COLS).fill(emptyColor);
      for (let i = 0; i < finalS.length && i < MAX_DASHBOARD_COLS; i++) { vals[i] = finalS[i]; bgs[i] = senkouColor; }
      sheet.getRange(rowNum, 4, 1, MAX_DASHBOARD_COLS).setValues([vals]).setBackgrounds([bgs]).setHorizontalAlignment("left");
    }
    
    // カレンダーへのプルダウン設定（12列全て）
    if (labelC === "1診目" || labelC === "2診目") {
      alignRanges.push(`D${rowNum}:O${rowNum}`);
      for (let c = 0; c < CALENDAR_COLS; c++) {
        currentRules[r][c] = newRule;
      }
      rulesModified = true;
    }
  }

  if (alignRanges.length > 0) {
    sheet.getRangeList(alignRanges).setHorizontalAlignment("left");
  }
  
  if (rulesModified) {
    targetRange.setDataValidations(currentRules);
  }

  updateSheetWideCF(sheet, finalJ, finalT, finalS, maxRows);
}

function syncSheetIndependent(sheet) {
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  
  let joukinList = [];
  let teikiList = [];
  let senkouList = [];

  for (let r = 0; r < data.length; r++) {
    let rowLabel = String(data[r][0]).trim();
    if (rowLabel === "常勤医師") {
      for (let c = 3; c <= 10; c++) { 
        let name = String(data[r][c]).trim();
        if (name && name !== "undefined" && name !== "休") joukinList.push(name);
      }
    } else if (rowLabel === "非常勤医師") {
      for (let c = 3; c <= 10; c++) {
        let name = String(data[r][c]).trim();
        if (name && name !== "undefined" && name !== "休") teikiList.push(name);
      }
    } else if (rowLabel === "先行応募" || rowLabel === "先行応募医師") {
      for (let c = 3; c <= 10; c++) {
        let name = String(data[r][c]).trim();
        if (name && name !== "undefined" && name !== "休") senkouList.push(name);
      }
    }
  }

  applyMasterListToAll(sheet, joukinList, teikiList, senkouList);
}

function updateSheetWideCF(sheet, joukinList, teikiList, senkouList, maxRows) {
  let rules = [];
  if (maxRows < 5) return;
  
  // ★ 条件付き書式をカレンダー部分（5行目以降）の「12列分（D〜O列）」に適用（色抜け修正）
  const targetRange = sheet.getRange(5, 4, maxRows - 4, 12); 

  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("募集").setBackground("#ffff00").setFontColor("#000000").setRanges([targetRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("休").setBackground("#cccccc").setFontColor("#cccccc").setRanges([targetRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("休館日").setBackground("#cccccc").setFontColor("#666666").setRanges([targetRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("未開院").setBackground("#e0e0e0").setFontColor("#999999").setRanges([targetRange]).build());
  
  // ★ 色の優先順位：常勤（肌色）＞ 非常勤（薄紫）＞ 先行（薄緑）
  joukinList.forEach(doc => {
    if(doc) rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(doc).setBackground("#fce5cd").setFontColor("#000000").setRanges([targetRange]).build());
  });
  
  teikiList.forEach(doc => {
    if(doc) rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(doc).setBackground("#d9d2e9").setFontColor("#000000").setRanges([targetRange]).build());
  });
  
  senkouList.forEach(doc => {
    // 常勤や非常勤に既に含まれている場合は、薄緑ルールを追加しない（全て緑になるのを回避）
    if(doc && !joukinList.includes(doc) && !teikiList.includes(doc)) {
      rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(doc).setBackground("#d9ead3").setFontColor("#000000").setRanges([targetRange]).build());
    }
  });
  
  sheet.setConditionalFormatRules(rules);
}