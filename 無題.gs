function runTruePerfectSimulation() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 過去に落ちた実績のある設定
  const targetYear = 2026;
  const targetTerm = "通年";
  const subLocName = "北葛西内科"; 
  const baseLocName = subLocName.replace(/（.*?）/, '');
  const finalSheetName = "🔧真・完全シミュレーション";
  
  // 1年分（12ヶ月）の月配列を作成
  const targetMonths = [];
  for (let m = 4; m <= 12; m++) targetMonths.push(`${targetYear}/${('0' + m).slice(-2)}`);
  for (let m = 1; m <= 3; m++) targetMonths.push(`${targetYear + 1}/${('0' + m).slice(-2)}`);

  console.log(`=== 🕵️ 真・完全シミュレーション開始: ${subLocName} ===`);

  try {
    // 1. ゴミシートのお掃除
    let sheet = ss.getSheetByName(finalSheetName);
    if (sheet) ss.deleteSheet(sheet);

    // 2. 本番と同じデータロード
    console.time("Step 1: マスターデータの一括ロード");
    let extractedData = typeof fetchAndOrganizeData === 'function' ? fetchAndOrganizeData(targetYear, targetTerm, [baseLocName]) : {shifts:{}};
    let shiftData = extractedData.shifts || {};
    let originalDataForMonth = shiftData[baseLocName] || {};
    let specialtyMap = typeof buildDoctorSpecialtyMap === 'function' ? buildDoctorSpecialtyMap(targetYear) : {};
    console.timeEnd("Step 1: マスターデータの一括ロード");

    // 3. 12ヶ月連続の本番描画ループ
    for (let i = 0; i < targetMonths.length; i++) {
      let yearMonthStr = targetMonths[i];
      console.log(`\n▶️ [${i+1}/12ヶ月目] ${yearMonthStr} の本番描画処理を開始...`);
      
      // 科目フィルタリング（本番と完全同一）
      let filteredDataForMonth = {};
      let isSplitTarget = (baseLocName === "亀有" || baseLocName === "北葛西");
      let targetCat = subLocName.includes("内科") ? "内科" : (subLocName.includes("小児科") ? "小児科" : "");
      
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

      console.time(`  [${yearMonthStr}] renderShiftBlock（本番関数）の実行`);
      // ★ ここで 04A_Render_Main.gs にある「本物の関数」を直接呼び出します
      let isRendered = renderShiftBlock(ss, subLocName, finalSheetName, yearMonthStr, filteredDataForMonth);
      console.timeEnd(`  [${yearMonthStr}] renderShiftBlock（本番関数）の実行`);
      
      if (!isRendered) {
        console.log(`  ℹ️ ${yearMonthStr} は開院前などの理由でスキップされました。`);
      }
    }

    // 4. 本番と同じ最終の書式同期
    console.time("Step 最終: 書式・プルダウンの同期 (syncSheetIndependent)");
    let finalSheet = ss.getSheetByName(finalSheetName);
    if (finalSheet) {
      if (typeof syncSheetIndependent === 'function') {
        syncSheetIndependent(finalSheet);
      }
    }
    console.timeEnd("Step 最終: 書式・プルダウンの同期 (syncSheetIndependent)");

    console.log("\n✅ === 🏁 真・シミュレーション完了（エラーなし） ===");
    
    // ※今回は結果を確認していただくため、自動削除(deleteSheet)は行いません。
    console.log("💡 スプレッドシート上に「🔧真・完全シミュレーション」というシートが完成しています。中身をご確認ください。");

  } catch (e) {
    console.error(`\n❌ [致命的エラー] 処理中にタイムアウトまたはクラッシュしました: ${e.message}`);
  }
}