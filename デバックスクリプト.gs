function debugNakamuraAndOshima() {
  const targetYear = "2026";
  const masterUrl = 'https://docs.google.com/spreadsheets/d/1aEjphEv_63SeWQmwiOy9sx7IrMfawU01sHbKd_Ki4iA/edit';
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  console.log("=== 🏥 ドクター契約マスタの調査 ===");
  try {
    const masterSs = SpreadsheetApp.openByUrl(masterUrl);
    ["常勤", "定期非常勤"].forEach(type => {
      const sheet = masterSs.getSheetByName(`${type}${targetYear}年度`);
      if (!sheet) return;
      const data = sheet.getDataRange().getValues();
      const headers = data[0].map(h => String(h).replace(/[\s　]+/g, ""));
      const nameCol = headers.findIndex(h => h.includes("氏名") || h.includes("医師名"));
      const bikouCol = headers.findIndex(h => h.includes("勤務備考"));

      for (let i = 1; i < data.length; i++) {
        const rawName = String(data[i][nameCol]);
        if ((rawName.includes("中村") && rawName.includes("千穂")) || (rawName.includes("大島") && rawName.includes("華倫"))) {
          console.log(`[マスタ: ${type}] 行${i+1} | 名前: '${rawName}' | 備考: \n${String(data[i][bikouCol]).split('\n').join(' / ')}`);
        }
      }
    });
  } catch(e) { console.log("マスタ読み込みエラー: " + e.message); }

  console.log("\n=== 📅 例外シート（先行応募・振替勤務）の調査 ===");
  ["先行応募", "振替勤務"].forEach(sName => {
    const sheet = ss.getSheetByName(sName);
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const rowStr = data[i].join(" | ");
      if ((rowStr.includes("中村") && rowStr.includes("千穂")) || (rowStr.includes("大島") && rowStr.includes("華倫"))) {
        console.log(`[例外: ${sName}] 行${i+1}: ${rowStr}`);
      }
    }
  });
}