/**
 * ==========================================
 * 99_Debug_Duplicates.gs
 * シートを作らず「実行ログ」に出力 ＆ 2診目の重複を即座に一括全削除する版
 * ==========================================
 */

function debug_LogAndCleanAllDuplicates() {
  const masterSs = SpreadsheetApp.getActiveSpreadsheet();

  // 先ほど作成した「📋重複チェック結果」シートがあれば自動でお掃除（削除）
  const oldReportSheet = masterSs.getSheetByName("📋重複チェック結果");
  if (oldReportSheet) {
    try {
      masterSs.deleteSheet(oldReportSheet);
      console.log("🗑️ 不要な「📋重複チェック結果」シートを削除しました。");
    } catch (e) {}
  }

  const TARGET_FILES = [
    { name: "東京・埼玉", id: "1rScroDlMNiRxThbaxGEuvhyCH2b9RoM6BadNSCAWsvI" },
    { name: "関東",       id: "19Q-xVsMX0thz_rvmdo_GERJvhjFLQgw2pUkWaTYxQhE" },
    { name: "関西",       id: "1fIFvTck_g9-Hp8MSpY7hIWwE5L2buJjBHIY4GBEf_MY" },
    { name: "本体シート", id: masterSs.getId() }
  ];

  const IGNORE_WORDS = ["🗂️", "初期設定", "テンプレート", "目次", "モニター", "作業表", "募集", "シフト作成", "キャンセル", "お休み", "先行", "振替", "休館日", "重複チェック結果"];
  const IGNORE_CELL_VALUES = new Set(["", "募集", "休", "休館日", "未開院", "undefined", "null"]);
  const TIME_LABELS = ["9:00", "10:00", "11:00", "12:00", "13:00", "14:00", "15:00", "16:00", "17:00", "18:00", "19:00", "20:00"];

  const scannedSheetKeys = new Set();
  let totalCleanedRows = 0;
  let totalCleanedCells = 0;
  let logLines = [];

  console.log("🔍 全ファイルの重複スキャン＆一括削除を開始します...");

  TARGET_FILES.forEach(fileInfo => {
    let extSs;
    try {
      extSs = SpreadsheetApp.openById(fileInfo.id);
    } catch (e) {
      console.warn(`[スキップ] ファイルが開けません (${fileInfo.name}): ${e.message}`);
      return;
    }

    const sheets = extSs.getSheets();
    sheets.forEach(sheet => {
      const sheetName = sheet.getName();
      const uniqueKey = `${fileInfo.id}_${sheetName}`;
      if (scannedSheetKeys.has(uniqueKey)) return;
      scannedSheetKeys.add(uniqueKey);

      if (IGNORE_WORDS.some(w => sheetName.includes(w))) return;

      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      if (lastRow < 5 || lastCol < 15) return;

      // D列〜O列を含む A1:O(最終行) を一括取得
      const range = sheet.getRange(1, 1, lastRow, 15);
      const data = range.getValues();
      let sheetModified = false;

      for (let r = 0; r < data.length - 1; r++) {
        const labelC1 = String(data[r][2]).trim();     // C列: 1診目
        const labelC2 = String(data[r + 1][2]).trim(); // C列: 2診目

        if (labelC1 === "1診目" && labelC2 === "2診目") {
          const dowStr = String(data[r][0]).trim();
          const rawDate = data[r + 1][0];
          const dateStr = (rawDate instanceof Date)
            ? Utilities.formatDate(rawDate, "JST", "M/d")
            : String(rawDate).trim();

          let dupHours = [];
          let dupDocName = "";
          let clearedCellsInRow = [];

          // D列(index 3) 〜 O列(index 14) をチェック
          for (let c = 3; c <= 14; c++) {
            if (c === 7 || c === 8) continue; // 13-15時の休診はスキップ

            const doc1 = String(data[r][c]).trim();
            const doc2 = String(data[r + 1][c]).trim();

            if (!IGNORE_CELL_VALUES.has(doc1) && doc1 === doc2) {
              dupDocName = doc1;
              dupHours.push(TIME_LABELS[c - 3]);
              
              // 2診目側のセルを直接空欄クリア（プルダウン等の書式を壊さないよう1セルずつclearContentではなくsetValue("")）
              const colLetter = String.fromCharCode(65 + c); // D〜O
              clearedCellsInRow.push(`${colLetter}${r + 2}`);
              totalCleanedCells++;
              sheetModified = true;
            }
          }

          if (clearedCellsInRow.length > 0) {
            // RangeListを使って該当行の重複セルを一括で空欄にする（超高速＆書式維持）
            sheet.getRangeList(clearedCellsInRow).setValue("");
            totalCleanedRows++;

            const logEntry = `[削除完了] ${fileInfo.name} | ${sheetName} | ${dateStr}(${dowStr}) | 2診目(${r + 2}行目) | 医師: ${dupDocName} | 時間: ${dupHours.join(", ")}`;
            logLines.push(logEntry);
            console.log(logEntry);
          }
        }
      }

      if (sheetModified) {
        SpreadsheetApp.flush();
      }
    });
  });

  console.log("==========================================");
  if (totalCleanedRows === 0) {
    console.log("✨ 重複データは0件でした（すべて綺麗です）。");
  } else {
    console.log(`✅ 全削除完了！ 合計 ${totalCleanedRows} 日分（${totalCleanedCells} マス）の2診目ダブりを削除しました。`);
    console.log("▼ 削除した一覧（コピペ共有用） ▼\n" + logLines.join("\n"));
  }
  console.log("==========================================");
}