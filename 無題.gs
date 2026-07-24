function debugAndFetchResults() {
  const EXCEPTION_SS_ID = "10yPdoOOgqSSGKwoiPAXi83YM9vb_Em8r6Ex-bLfg28M";
  const ss = SpreadsheetApp.openById(EXCEPTION_SS_ID);
  const originalSheet = ss.getSheetByName("お休み情報");
  if (!originalSheet) return;

  // 1. テストシートリセット
  let testSheet = ss.getSheetByName("お休み情報_TEST");
  if (testSheet) ss.deleteSheet(testSheet);
  testSheet = originalSheet.copyTo(ss);
  testSheet.setName("お休み情報_TEST");

  const cleanStr = (str) => String(str).replace(/[\s 【】\u200B\n\r]/g, "").trim();
  const lastCol = testSheet.getLastColumn();
  const headers = testSheet.getRange(1, 1, 1, lastCol).getValues()[0].map(cleanStr);
  
  const cDate = headers.findIndex(h => /日付|勤務日|対象日|年月日|日時/.test(h));
  const cName = headers.findIndex(h => /氏名|名前|医師名|ドクター/.test(h));
  const cStart = headers.findIndex(h => /開始|出勤|入/.test(h));
  const cEnd = headers.findIndex(h => /終了|退勤|退/.test(h));
  const cLoc = headers.findIndex(h => /拠点|クリニック|勤務先|店舗|勤務地|場所/.test(h));
  const cType = headers.findIndex(h => /種別|タイプ|理由|事由|区分|休み/.test(h));
  const cMedId = headers.findIndex(h => /医籍番号|ID|id/i.test(h));

  const startRow = testSheet.getLastRow() + 1;

  const today = new Date();
  const futureDate = new Date(today); futureDate.setDate(today.getDate() + 10);
  const pastDate = new Date(today); pastDate.setDate(today.getDate() - 10);
  const fmt = d => Utilities.formatDate(d, "JST", "yyyy/MM/dd");

  // 見つけやすいように拠点名に目印をつける
  const mockData = [
    { date: fmt(futureDate), name: "山本敬一", start: "9:00", end: "18:00", loc: "【テスト検証】未来", type: "有給" },
    { date: fmt(pastDate), name: "山本敬一", start: "9:00", end: "18:00", loc: "【テスト検証】過去", type: "欠勤" }
  ];
  const numRows = mockData.length;

  const maxRows = testSheet.getMaxRows();
  if (startRow + numRows - 1 > maxRows) {
    testSheet.insertRowsAfter(maxRows, (startRow + numRows - 1) - maxRows + 2);
  }

  // 2. データ書き込み
  const newRowData = [];
  for (let i = 0; i < numRows; i++) {
    let rowArray = new Array(lastCol).fill(null);
    if(cDate > -1) rowArray[cDate] = mockData[i].date;
    if(cName > -1) rowArray[cName] = mockData[i].name;
    if(cStart > -1) rowArray[cStart] = mockData[i].start;
    if(cEnd > -1) rowArray[cEnd] = mockData[i].end;
    if(cLoc > -1) rowArray[cLoc] = mockData[i].loc;
    if(cType > -1) rowArray[cType] = mockData[i].type;
    newRowData.push(rowArray);
  }

  const targetRange = testSheet.getRange(startRow, 1, numRows, lastCol);
  targetRange.setValues(newRowData);

  // 3. 2行目からの完全なフォーマット＆入力規則コピー
  const formatSource = testSheet.getRange(2, 1, 1, lastCol);
  formatSource.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  formatSource.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);

  // 4. 降順ソートと非表示
  const currentLastRow = testSheet.getLastRow();
  testSheet.showRows(2, currentLastRow - 1); 
  
  if (cDate > -1) {
    testSheet.getRange(2, 1, currentLastRow - 1, lastCol).sort({ column: cDate + 1, ascending: false });
    const dateValues = testSheet.getRange(2, cDate + 1, currentLastRow - 1, 1).getValues();
    const todayDate = new Date(); todayDate.setHours(0, 0, 0, 0);

    for (let i = 0; i < dateValues.length; i++) {
      const dVal = dateValues[i][0];
      let rowDate = null;
      if (dVal instanceof Date) rowDate = new Date(dVal.getTime());
      else if (dVal) {
        let strVal = String(dVal).replace(/[年月]/g, '/').replace(/日/g, '');
        let d = new Date(strVal);
        if (!isNaN(d.getTime())) rowDate = d;
      }
      if (rowDate) {
        rowDate.setHours(0, 0, 0, 0);
        if (rowDate.getTime() < todayDate.getTime()) {
          testSheet.hideRows(i + 2);
        }
      }
    }
  }

  // ==========================================
  // 5. 【自動結果取得】スクリプト自身でシートを検査
  // ==========================================
  SpreadsheetApp.flush(); // スプレッドシートへの変更を確定させる
  
  console.log("=== 📊 スクリプトによる自己診断結果 ===");
  const finalData = testSheet.getRange(1, 1, testSheet.getLastRow(), lastCol).getValues();
  const finalDisplayValues = testSheet.getRange(1, 1, testSheet.getLastRow(), lastCol).getDisplayValues();
  const finalValidations = testSheet.getRange(1, 1, testSheet.getLastRow(), lastCol).getDataValidations();
  
  for (let i = 1; i < finalData.length; i++) {
    const locVal = String(finalData[i][cLoc]);
    if (locVal.includes("【テスト検証】")) {
      const rowNum = i + 1;
      const isHidden = testSheet.isRowHiddenByUser(rowNum);
      const hasDropdown = finalValidations[i][cEnd] != null;
      const medIdVal = finalDisplayValues[i][cMedId];
      
      console.log(`\n■ 発見: ${locVal} の行 (シート ${rowNum}行目)`);
      console.log(`・【プルダウン】終了時間に適用されたか？ -> ${hasDropdown ? "✅ はい" : "❌ いいえ"}`);
      console.log(`・【数式】医籍番号の表示結果 -> [ ${medIdVal} ] ${medIdVal.includes("#N/A") ? "❌ エラー" : "✅ 正常取得"}`);
      console.log(`・【非表示】行の表示状態 -> ${isHidden ? "✅ 非表示(隠れている)" : "❌ 表示中(見えている)"}`);
    }
  }
  console.log("\nこちらのログを貼り付けていただけますでしょうか。");
}