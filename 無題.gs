// ==========================================
// 1. まず東京・埼玉を実行し、モニターが完了したら次へ
// ==========================================
function fixTokyoSaitama() { 
  safeFixAndBatch("1rScroDlMNiRxThbaxGEuvhyCH2b9RoM6BadNSCAWsvI", "東京・埼玉"); 
}

// ==========================================
// 2. 関東を実行し、モニターが完了したら次へ
// ==========================================
function fixKanto() { 
  safeFixAndBatch("19Q-xVsMX0thz_rvmdo_GERJvhjFLQgw2pUkWaTYxQhE", "関東"); 
}

// ==========================================
// 3. 最後に関西を実行
// ==========================================
function fixKansai() { 
  safeFixAndBatch("1fIFvTck_g9-Hp8MSpY7hIWwE5L2buJjBHIY4GBEf_MY", "関西"); 
}


// --- 以下、共通の掃除＆バッチ起動エンジン（変更不要） ---
function safeFixAndBatch(fileId, areaName) {
  const targetYear = "2026";
  let allLocs = [];

  console.log(`=== 🧹 【${areaName}】のフォーマット掃除と再構築を開始します ===`);

  try {
    let ss = SpreadsheetApp.openById(fileId);
    ss.getSheets().forEach(sheet => {
      let sName = sheet.getName();
      
      // 2026から始まるシートのみ対象
      if (sName.startsWith(targetYear)) {
        let locName = sName.replace(targetYear, "");
        allLocs.push(locName);

        // 1. Q列以降のはみ出したゴミを完全に消去
        const maxCols = sheet.getMaxColumns();
        if (maxCols >= 17) {
          sheet.getRange(1, 17, sheet.getMaxRows(), maxCols - 16).clearFormat().clearContent();
        }

        // 2. 右側表のリセット ＆ ヘッダーの1行上への移動
        const data = sheet.getDataRange().getValues();
        for (let r = 0; r < data.length; r++) {
          let label = String(data[r][0]).trim();
          if (label === "先行応募" || label === "先行応募医師") {
            let startRow = r + 1; 
            
            // L〜P列のデータを消去し白背景化
            sheet.getRange(startRow, 12, 16, 5)
                 .clearContent()
                 .setBackground("#ffffff")
                 .setBorder(false, false, false, false, false, false);
            
            // ヘッダーを1行上にセット（8pt）
            let headers = [["医師名", "稼働時間", "契約内容", "時給", "月間コスト"]];
            sheet.getRange(startRow, 12, 1, 5)
                 .setValues(headers)
                 .setBackground("#e4efff")
                 .setBorder(true, true, true, true, true, true, null, SpreadsheetApp.BorderStyle.SOLID)
                 .setFontSize(8);
          }
        }
      }
    });
  } catch(e) {
    console.log(`❌ ${areaName} の処理中にエラー: ` + e.message);
    return;
  }

  if (allLocs.length > 0) {
    console.log(`✅ 掃除完了。【${areaName}】計 ${allLocs.length} 拠点のバッチをスタートします！`);
    
    // お客様のバッチシステムへキューを投げる
    startBackgroundBatch({
      year: targetYear,
      term: "通年", 
      locations: allLocs
    });
  } else {
    console.log(`⚠️ 対象となる ${targetYear} のシートが見つかりませんでした。`);
  }
}