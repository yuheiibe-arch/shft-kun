function auditMissingMonths() {
  console.log("=== 🔍 [現状監査] 拠点別・欠落月 特定デバックスクリプト ===");
  const targetYear = 2026;

  // 1. マスタから開院日を取得
  let openDatesMap = {};
  try {
    const masterSs = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/14RbsDcv0nXfEwweki8-9cK3lQUg1XUuhozLNF9u2qAs/edit");
    const sheet = masterSs.getSheetByName("拠点名");
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      let name = String(data[i][0]).trim();
      let dVal = data[i][7]; // H列(開院日)
      if (name) {
        if (dVal instanceof Date) openDatesMap[name] = dVal;
        else if (dVal) openDatesMap[name] = new Date(dVal);
        else openDatesMap[name] = null;
      }
    }
  } catch(e) {
    console.log("❌ マスタ取得エラー: " + e.message);
    return;
  }

  // 2. 現在のスプレッドシート内の「2026」から始まるシートを監査
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  let errorCount = 0;

  sheets.forEach(sheet => {
    let sheetName = sheet.getName();
    // 「2026」から始まるシートのみ対象
    if (!sheetName.startsWith(String(targetYear))) return;

    // シート名から拠点名を抽出 (例: "2026北葛西（内科）" -> "北葛西")
    let rawLoc = sheetName.replace(String(targetYear), '').replace(/（.*?）|\(.*?\)|【.*?】/g, "").trim();

    // マスタと名前をマッチング
    let clinicOpenDate = null;
    let matchedMasterName = Object.keys(openDatesMap).find(k => rawLoc.includes(k) || k.includes(rawLoc));
    if (matchedMasterName) {
      clinicOpenDate = openDatesMap[matchedMasterName];
    }

    // 期待される月リストを作成（通年：4月〜翌3月）
    let expectedMonths = [];
    for (let m = 4; m <= 12; m++) expectedMonths.push(`${targetYear}/${('0' + m).slice(-2)}`);
    for (let m = 1; m <= 3; m++) expectedMonths.push(`${targetYear + 1}/${('0' + m).slice(-2)}`);

    // 開院日フィルタ適用（開院前の月は「期待される月」から除外）
    if (clinicOpenDate && !isNaN(clinicOpenDate.getTime())) {
      let openYM = (clinicOpenDate.getFullYear() * 100) + (clinicOpenDate.getMonth() + 1);
      expectedMonths = expectedMonths.filter(ym => {
        let parts = ym.split('/');
        let targetYM = (parseInt(parts[0], 10) * 100) + parseInt(parts[1], 10);
        return targetYM >= openYM;
      });
    }

    // シート内の全データを走査して「実際に描画されている月」を抽出
    let foundMonths = new Set();
    let values = sheet.getDataRange().getValues();
    let displayValues = sheet.getDataRange().getDisplayValues();

    for (let r = 0; r < values.length; r++) {
      for (let c = 0; c < values[r].length; c++) {
        let v = values[r][c];
        let dv = displayValues[r][c];
        
        // Dateオブジェクトとして存在する場合
        if (v instanceof Date) {
          let y = v.getFullYear();
          let m = v.getMonth() + 1;
          if ((y === targetYear && m >= 4) || (y === targetYear + 1 && m <= 3)) {
            foundMonths.add(`${y}/${('0'+m).slice(-2)}`);
          }
        } 
        // 文字列として存在する場合 (2026/04, 2026年4月 などに柔軟に対応)
        else if (typeof dv === 'string') {
          let match = dv.match(/(2026|2027)[\/\-年]\s*(1[0-2]|0?[1-9])/);
          if (match) {
            let y = parseInt(match[1], 10);
            let m = parseInt(match[2], 10);
            if ((y === targetYear && m >= 4 && m <= 12) || (y === targetYear + 1 && m >= 1 && m <= 3)) {
              foundMonths.add(`${y}/${('0'+m).slice(-2)}`);
            }
          }
        }
      }
    }

    // 差分チェック（期待される月リストのうち、シート内に存在しなかった月）
    let missingMonths = expectedMonths.filter(m => !foundMonths.has(m));

    if (missingMonths.length > 0) {
      console.log(`🚨 拠点 [${sheetName}]`);
      console.log(`   -> 開院日: ${clinicOpenDate ? Utilities.formatDate(clinicOpenDate, "JST", "yyyy/MM/dd") : "設定なし（通年対象）"}`);
      console.log(`   -> ❌ 欠落している月: ${missingMonths.join(", ")}`);
      errorCount++;
    } else {
      console.log(`✅ 拠点 [${sheetName}]: 欠落なし`);
    }
  });

  if (errorCount === 0) {
    console.log("\n✨ すべての「2026」シートで、開院日に基づく必要な月がすべて揃っています。");
  } else {
    console.log(`\n⚠️ 合計 ${errorCount} 拠点で「本来あるべき月の欠落」が確認されました。`);
  }
  console.log("=== 調査完了 ===");
}