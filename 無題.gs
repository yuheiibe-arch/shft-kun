function debugProveTheBug() {
  console.log("=== 🔍 getShiftForDay の絶望的なバグ証明 ===");
  
  const line = "【稲毛海岸】第2・第4日曜日：9:00～18:00";
  const targetDayStr = "日";
  const weekNum = 4; // 9/27 (第4週)

  console.log(`📝 テスト文字列: ${line}`);

  // ❌ 今のポンコツロジック
  let isMatchOld = false;
  const matchOld = line.match(/第([\d・･]+)/);
  let weeksOld = [];
  if (matchOld) {
    weeksOld = matchOld[1].split(/[・･]/).map(Number);
    if (weeksOld.includes(weekNum)) isMatchOld = true;
  }

  console.log("\n❌ 【現在のシステムの認識】");
  console.log(`抽出された週: [${weeksOld}]`);
  console.log(`9/27(第4週)にシフトは入るか？: ${isMatchOld ? "✅入る" : "❌休になる"}`);

  // ⭕ 最強パーサーと同じロジック
  let isMatchNew = false;
  let beforeDow = line.split(targetDayStr + "曜")[0];
  let digits = beforeDow.match(/\d/g);
  let weeksNew = [];
  if (digits) {
    weeksNew = digits.map(Number);
    if (weeksNew.includes(weekNum)) isMatchNew = true;
  }
  
  console.log("\n✨ 【新ロジックの認識】");
  console.log(`抽出された週: [${weeksNew}]`);
  console.log(`9/27(第4週)にシフトは入るか？: ${isMatchNew ? "✅入る" : "❌休になる"}`);
  console.log("================================================");
}