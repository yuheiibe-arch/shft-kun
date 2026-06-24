/**
 * ==========================================
 * 04C_Render_Sync.gs
 * 自動同期・条件付き書式の反映・onEditイベント
 * ★プルダウン一括設定（API通信1回）究極対応版
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
  if (col >= 4 && col <= 10) {
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
  
  // ====================================================================
  // ★ 諸悪の根源（個別ループ通信）を撤廃し、2次元配列で一気に処理する
  // ====================================================================
  const targetRange = sheet.getRange(1, 4, maxRows, 12); // D列〜O列の全行を一気に取得
  let currentRules = targetRange.getDataValidations();
  let rulesModified = false;

  for (let r = 0; r < maxRows; r++) {
    let rowNum = r + 1;
    let labelA = String(data[r][0]).trim(); 
    let labelC = String(data[r][2]).trim(); 
    
    if (labelA === "常勤医師") {
      let vals = new Array(7).fill("");
      let bgs = new Array(7).fill(emptyColor);
      for (let i = 0; i < finalJ.length && i < 7; i++) { vals[i] = finalJ[i]; bgs[i] = joukinColor; }
      sheet.getRange(rowNum, 4, 1, 7).setValues([vals]).setBackgrounds([bgs]).setHorizontalAlignment("left");
    } 
    else if (labelA === "非常勤医師") {
      let vals = new Array(7).fill("");
      let bgs = new Array(7).fill(emptyColor);
      for (let i = 0; i < finalT.length && i < 7; i++) { vals[i] = finalT[i]; bgs[i] = teikiColor; }
      sheet.getRange(rowNum, 4, 1, 7).setValues([vals]).setBackgrounds([bgs]).setHorizontalAlignment("left");
    }
    
    if (labelC === "1診目" || labelC === "2診目") {
      alignRanges.push(`D${rowNum}:O${rowNum}`);
      // 裏側（メモリ上）で配列にルールをセットする（通信は発生しない）
      for (let c = 0; c < 12; c++) {
        currentRules[r][c] = newRule;
      }
      rulesModified = true;
    }
  }

  // 左詰めの一括適用
  if (alignRanges.length > 0) {
    sheet.getRangeList(alignRanges).setHorizontalAlignment("left");
  }
  
  // ★ 最後に1回だけ、書き換えたルール配列をガバッと被せる（通信1回で終了！）
  if (rulesModified) {
    targetRange.setDataValidations(currentRules);
  }

  updateSheetWideCF(sheet, finalJ, finalT, finalS);
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
      for (let c = 3; c <= 9; c++) {
        let name = String(data[r][c]).trim();
        if (name && name !== "undefined" && name !== "休") joukinList.push(name);
      }
    } else if (rowLabel === "非常勤医師") {
      for (let c = 3; c <= 9; c++) {
        let name = String(data[r][c]).trim();
        if (name && name !== "undefined" && name !== "休") teikiList.push(name);
      }
    } else if (rowLabel === "先行応募" || rowLabel === "先行応募医師") {
      for (let c = 3; c <= 9; c++) {
        let name = String(data[r][c]).trim();
        if (name && name !== "undefined" && name !== "休") senkouList.push(name);
      }
    }
  }

  applyMasterListToAll(sheet, joukinList, teikiList, senkouList);
}

function updateSheetWideCF(sheet, joukinList, teikiList, senkouList) {
  let rules = [];
  const targetRange = sheet.getRange("D:O"); 

  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("募集").setBackground("#ffff00").setFontColor("#000000").setRanges([targetRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("休").setBackground("#cccccc").setFontColor("#cccccc").setRanges([targetRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("休館日").setBackground("#cccccc").setFontColor("#666666").setRanges([targetRange]).build());
  rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo("未開院").setBackground("#e0e0e0").setFontColor("#999999").setRanges([targetRange]).build());
  
  joukinList.forEach(doc => {
    if(doc) rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(doc).setBackground("#fce5cd").setFontColor("#000000").setRanges([targetRange]).build());
  });
  
  teikiList.forEach(doc => {
    if(doc) rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(doc).setBackground("#d9d2e9").setFontColor("#000000").setRanges([targetRange]).build());
  });
  
  if(senkouList) {
    senkouList.forEach(doc => {
      if(doc) rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(doc).setBackground("#d9ead3").setFontColor("#000000").setRanges([targetRange]).build());
    });
  }
  
  sheet.setConditionalFormatRules(rules);
}