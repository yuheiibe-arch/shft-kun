function debugWhySakugyouhyou_Level2() {
  console.log("=== 🔍 [誤爆原因追跡・完全版] スプレッドシート全体の徹底捜索 ===");
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const targetStr = "下期作業表";

  console.log("【フェーズ1：全シートの全セルを物理検索】");
  let foundInCells = false;
  
  sheets.forEach(sheet => {
    const sName = sheet.getName();
    // データがない空シートはスキップ
    if (sheet.getLastRow() === 0) return; 
    
    const data = sheet.getDataRange().getDisplayValues();
    for (let r = 0; r < data.length; r++) {
      for (let c = 0; c < data[r].length; c++) {
        if (data[r][c].includes(targetStr)) {
           console.log(` 🚨 発見: シート「${sName}」の ${sheet.getRange(r+1, c+1).getA1Notation()} セル -> "${data[r][c]}"`);
           foundInCells = true;
        }
      }
    }
  });
  
  if (!foundInCells) {
    console.log(" ✅ どのシートのセルにも文字列は見つかりませんでした。");
  }

  console.log("\n【フェーズ2：シート名からの自動取得ロジックの可能性】");
  let foundInSheetName = false;
  sheets.forEach(sheet => {
    if (sheet.getName().includes(targetStr)) {
      console.log(` 💡 発見: シート名「${sheet.getName()}」が存在します。`);
      foundInSheetName = true;
    }
  });

  console.log("\n=== 調査完了 ===");
}