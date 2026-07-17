/**
 * テスト①：スマート・トリミングの確証テスト
 * 空白行が「3行」だけ残して綺麗に削除されるか、目に見える形で証明します。
 */
function testSmartTrimming() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = "🔧トリミングテスト";
  
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) ss.deleteSheet(sheet);
  
  // 1. わざと1000行の巨大なシートを作る
  sheet = ss.insertSheet(sheetName);
  if (sheet.getMaxRows() < 1000) sheet.insertRowsAfter(sheet.getMaxRows(), 1000 - sheet.getMaxRows());
  
  // 2. 50行目までダミーデータを入れる
  sheet.getRange(1, 1, 50, 1).setValue("データ");
  
  console.log(`[トリミング前] データ最終行: ${sheet.getLastRow()} / シート全体の行数: ${sheet.getMaxRows()}`);
  
  // 3. バッチエンジンに組み込んだものと【全く同じ】トリミング処理を実行
  let lastDataRow = sheet.getLastRow();
  let maxSheetRows = sheet.getMaxRows();
  let targetMaxRows = lastDataRow + 3; // 50行 + 余白3行 = 53行になるはず
  
  if (maxSheetRows > targetMaxRows) {
    sheet.deleteRows(targetMaxRows + 1, maxSheetRows - targetMaxRows);
  }
  
  console.log(`[トリミング後] データ最終行: ${sheet.getLastRow()} / シート全体の行数: ${sheet.getMaxRows()}`);
  
  if (sheet.getMaxRows() === 53) {
    console.log("✅ 【大成功】余白3行だけを残して、不要な空白行が完全に切り落とされました！");
  } else {
    console.log("❌ トリミングに失敗しました。");
  }
}


/**
 * テスト②：本番トリガーバッチの直接テスト
 * UIを介さず、バックグラウンドの「トリガー連鎖（10秒待機）」が正常に回るかテストします。
 */
function testRealTriggerBatch() {
  console.log("=== 🚀 本番トリガーバッチの直接テストを開始 ===");
  
  // UIから送られてくるのと同じデータを擬似的に作成
  const payload = {
    year: "2026",
    term: "通年",
    locations: ["長吉長原"] // ここをテストしたい拠点に変えてもOKです
  };
  
  // 本番のバッチエンジンを直接起動
  const result = startBackgroundBatch(payload);
  
  console.log(`起動結果: ${result}`);
  if (result === "Success") {
    console.log("✅ バッチエンジンが正常に起動しました。");
    console.log("💡 【確認手順】");
    console.log("1. スプレッドシートに「🗂️進行中モニター」が表示されているか確認してください。");
    console.log("2. GASエディタの左側メニュー「実行数 (時計マーク)」を開いてください。");
    console.log("3. 『processBatchQueue』がタイムアウトせずに、10秒間隔でバトンタッチしながら完走するか見届けてください。");
  } else {
    console.log("❌ バッチの起動に失敗しました。");
  }
}