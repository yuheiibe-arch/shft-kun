function debugMusashikosugiTuesday() {
  const targetLoc = "武蔵小杉";
  const targetDow = "火";
  const targetYear = "2026";
  
  console.log(`=== 🔍 募集くんデバッグ: ${targetLoc} ${targetDow}曜 09:00-13:00 の未検知調査 ===`);

  // 1. 拠点名の解決
  let map = typeof getLocationDictionary === "function" ? getLocationDictionary() : {};
  let cleanLoc = typeof normalizeLocationName === "function" ? normalizeLocationName(targetLoc, map) : targetLoc;
  console.log(`✅ 拠点正規化: ${targetLoc} -> ${cleanLoc}`);

  // 2. マスタから対象拠点の契約情報と退職日を取得
  const masterUrl = 'https://docs.google.com/spreadsheets/d/1aEjphEv_63SeWQmwiOy9sx7IrMfawU01sHbKd_Ki4iA/edit';
  let masterSs;
  try {
    masterSs = SpreadsheetApp.openByUrl(masterUrl);
  } catch(e) {
    console.log("❌ マスタが開けません: " + e.message);
    return;
  }

  let foundContract = false;
  ["常勤", "定期非常勤"].forEach(type => {
    const sheet = masterSs.getSheetByName(`${type}${targetYear}年度`);
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    const nameIdx = data[0].indexOf("医師名");
    const bikouIdx = data[0].findIndex(h => String(h).includes("備考"));
    const retireIdx = data[0].findIndex(h => String(h).includes("退職"));

    for (let r = 1; r < data.length; r++) {
      let docName = String(data[r][nameIdx]).replace(/先生$/, "").trim();
      let bikou = String(data[r][bikouIdx] || "");
      if (!docName || !bikou) continue;
      
      if (bikou.includes(cleanLoc) || bikou.includes(targetLoc)) {
        if (bikou.includes(targetDow) && bikou.includes("09")) {
          foundContract = true;
          let retireVal = retireIdx !== -1 ? data[r][retireIdx] : "";
          let retireDateStr = "未設定";
          if (retireVal instanceof Date) {
            retireDateStr = Utilities.formatDate(retireVal, "JST", "yyyy/MM/dd");
          } else if (retireVal && String(retireVal).trim() !== "") {
            retireDateStr = `[文字列] ${String(retireVal)}`;
          }
          
          console.log(`👨‍⚕️ 医師: ${docName} (${type})`);
          console.log(`   退職日: ${retireDateStr}`);
          console.log(`   契約備考: ${bikou.replace(/\n/g, " / ")}`);
        }
      }
    }
  });
  
  if (!foundContract) {
    console.log(`⚠️ マスタ上で ${targetLoc} の ${targetDow}曜 09:00〜 に該当する医師の契約が見つかりませんでした。`);
  }

  // 3. 時給取得のテスト（ここが空欄だと募集が出力されません）
  console.log(`\n=== 💰 時給取得のテスト ===`);
  try {
    if (typeof initializeWageData === "function") initializeWageData();
    // 2026年8月の火曜日(11日)をサンプルに時給を叩く
    let wageStr = typeof _getWageWrapper === "function" ? _getWageWrapper("2026/08/11", cleanLoc, "小児科", targetDow, 9, 13) : "関数未定義";
    
    console.log(`   取得時給: ${wageStr}`);
    if (!wageStr || wageStr === "") {
      console.log(`   ❌ 原因特定！時給が取得できていません。時給マスターに【${cleanLoc}】のデータがないか、エラーが発生しています。`);
    } else {
      console.log(`   ✅ 時給は正常に取得されています。`);
    }
  } catch(e) {
    console.log(`   ❌ 時給計算エラー: ${e.message}`);
  }
}