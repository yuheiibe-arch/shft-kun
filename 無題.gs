/**
 * ==========================================
 * 全シート透視用 デバッグスクリプト（10月6日 阿波座専用）
 * ==========================================
 */
function debugXRayAwaza() {
  const targetLoc = "阿波座";
  const targetYear = 2026;
  const targetTerm = "下期"; // または通年

  console.log(`🔍 [X-Ray デバッグ] ${targetLoc}のデータを全シートから透視します...`);

  // 1. 勤怠データ（大元）の抽出
  let extracted = fetchAndOrganizeData(targetYear, targetTerm, [targetLoc]);
  let mainShifts = extracted.shifts[targetLoc] || {};

  // 2. 外部シート（先行・振替・お休み）の抽出
  let overrides = getShiftOverrides(SpreadsheetApp.getActiveSpreadsheet(), targetLoc, targetLoc, "2026/10");

  // 10月6日のデータをピンポイントで確認
  const checkDate = "2026/10/06";

  console.log(`\n📅 【 ${checkDate} の抽出結果 】`);

  console.log("■ 1. 勤怠マスタ（常勤・定期非常勤）:");
  let dayMain = mainShifts[checkDate] || [];
  if (dayMain.length === 0) console.log("  なし");
  dayMain.forEach(s => console.log(`  - ${s.doctorName} (${s.type}): ${s.startTime}-${s.endTime} (元テキスト: ${s.rawShift})`));

  console.log("\n■ 2. 先行応募シート:");
  let dayAdv = overrides.advance[checkDate] || [];
  if (dayAdv.length === 0) console.log("  なし");
  dayAdv.forEach(s => console.log(`  - ${s.doc}: ${s.startH}:00-${s.endH}:00`));

  console.log("\n■ 3. 振替勤務シート:");
  let daySub = overrides.substitute[checkDate] || [];
  if (daySub.length === 0) console.log("  なし");
  daySub.forEach(s => console.log(`  - ${s.doc}: ${s.startH}:00-${s.endH}:00`));

  console.log("\n■ 4. お休み情報シート (※ここにあると削除されます):");
  let dayAbs = overrides.absence[checkDate] || [];
  if (dayAbs.length === 0) console.log("  なし");
  dayAbs.forEach(s => console.log(`  - ${s.doc}: ${s.startH}:00-${s.endH}:00`));

  console.log("\n💡 【判定方法】");
  console.log("上記の1〜3の中に、「白濱」先生が合計2つ以上出力されていませんか？");
  console.log("もし2つ出ていれば、それが重複の犯人です。名前の漢字や、開始時間の微妙なズレ（17:00と17:30など）を確認してください。");
}