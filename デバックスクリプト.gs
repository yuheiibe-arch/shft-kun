function debugShiftLogicWithCorrectStructure() {
  const targetYear = "2026";
  const masterUrl = 'https://docs.google.com/spreadsheets/d/1aEjphEv_63SeWQmwiOy9sx7IrMfawU01sHbKd_Ki4iA/edit';
  let specialtyMap = {};
  
  try {
    const masterSs = SpreadsheetApp.openByUrl(masterUrl);
    ["常勤", "定期非常勤"].forEach(type => {
      const sheet = masterSs.getSheetByName(`${type}${targetYear}年度`);
      if (sheet) {
        const data = sheet.getDataRange().getValues();
        const nameIdx = data[0].indexOf("医師名");
        const subjIdx = data[0].findIndex(h => String(h).includes("専門") || String(h).includes("科目"));
        if (nameIdx !== -1 && subjIdx !== -1) {
          for (let r = 1; r < data.length; r++) {
            let docName = String(data[r][nameIdx]).replace(/先生$/, "").trim();
            let spec = String(data[r][subjIdx]).trim();
            if (docName) {
              if (spec.includes("内科")) specialtyMap[docName] = "内科";
              else if (spec.includes("小児科")) specialtyMap[docName] = "小児科";
              else specialtyMap[docName] = "その他";
            }
          }
        }
      }
    });
  } catch(e) {}

  // わざわざシフトで別科目を指定した場合のテストデータを含める
  const testShifts = [
    { doctorName: "森博子", rawShift: "契約：【北葛西内科】09-18", desc: "小児科医が内科シフトに入った場合" },
    { doctorName: "松本公宏", rawShift: "契約：【北葛西内科】15-20", desc: "内科医が内科シフトに入った場合" },
    { doctorName: "鈴木謙", rawShift: "契約：【北葛西小児科】09-13", desc: "内科医が小児科シフトに入った場合" },
    { doctorName: "北原英晃", rawShift: "契約：【北葛西】09-20:15", desc: "内科医が科目指定なしの場合" },
    { doctorName: "中村多一郎", rawShift: "契約：【北葛西】09-18", desc: "小児科医が科目指定なしの場合" }
  ];

  console.log("=========================================");
  console.log("🏥 拠点: 北葛西 の正しい構造判定テスト");
  console.log("▼ 出力対象シート: 北葛西（内科）の場合");

  const targetCat = "内科"; 
  const otherCat = "小児科";

  testShifts.forEach(s => {
    let docSpec = specialtyMap[s.doctorName] || "不明";
    let text = s.rawShift || "";
    
    // ===============================================
    // ★他のスクリプトの構造を応用した正しいロジック
    // ===============================================
    let isBoxMatch = text.includes(targetCat);
    let isOtherBoxMatch = text.includes(otherCat);
    
    let isSubjectMatch = docSpec.includes(targetCat);
    let isOtherSubjectMatch = docSpec.includes(otherCat);

    let isOutput = false;
    let reason = "";

    // 1. まずシフト文字列（箱）の指定を最優先
    if (isOtherBoxMatch && !isBoxMatch) {
      isOutput = false;
      reason = `シフトに「${otherCat}」と明記されているため除外`;
    } 
    else if (isBoxMatch) {
      isOutput = true;
      reason = `シフトに「${targetCat}」と明記されているため出力`;
    } 
    // 2. シフトの指定がない場合のみ、マスタの専門で判定
    else {
      if (!isSubjectMatch && isOtherSubjectMatch) {
        isOutput = false;
        reason = `シフト指定がなく、マスタ専門が「${otherCat}」のため除外`;
      } else {
        isOutput = true;
        reason = `シフト指定がなく、マスタ専門が「${targetCat}」のため出力`;
      }
    }

    const mark = isOutput ? "⭕️ 出力" : "❌ 除外";
    console.log(`[${s.doctorName}] (${s.desc})\n   マスタ:${docSpec} | シフト:${s.rawShift}\n   -> 判定: ${mark} (${reason})\n`);
  });
}