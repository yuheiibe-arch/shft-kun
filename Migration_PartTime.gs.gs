// ==========================================
// 【Migration_PartTime】定期非常勤 次年度移行ロジック（テンプレート活用・行数最適化版）
// ==========================================

function createPartTimeMigrationSheet(targetYear) {
  // 入力: "2026年度" -> 現在: 2025, 次: 2026
  const currentYearNum = parseInt(targetYear.replace("年度", "")) - 1;
  const nextYearNum = parseInt(targetYear.replace("年度", ""));
  
  const srcSheetName = `定期非常勤${currentYearNum}年度`;
  const destSheetName = `(調整中)定期非常勤${nextYearNum}年度`;
  const templateSheetName = "テンプレート"; // 固定のテンプレートシート名

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const srcSheet = ss.getSheetByName(srcSheetName);
  
  if (!srcSheet) throw new Error(`現在のシート「${srcSheetName}」が見つかりません。`);

  // --- 1. テンプレート情報の取得（移行先スプシから） ---
  const destSS = SpreadsheetApp.openById(MIGRATION_DEST_ID);
  
  const templateSheet = destSS.getSheetByName(templateSheetName);
  if (!templateSheet) {
    throw new Error(`移行先スプレッドシートに「${templateSheetName}」シートが見つかりません。`);
  }

  // M列（13列目）の「書き方ルール」テキストを取得 (テンプレートの2行目から取得)
  // データ行作成時に各行へセットするため
  let templateRuleText = "";
  if (templateSheet.getLastRow() >= 2) {
    templateRuleText = templateSheet.getRange(2, 13).getValue();
  }

  // --- 2. データ作成 ---
  const data = srcSheet.getDataRange().getValues();
  const headers = data[0];
  const hMap = {};
  headers.forEach((h, i) => hMap[h] = i);

  const outputRows = [];
  let rowCount = 1;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    const retireDate = row[hMap["退職日"]];
    const nextChange = row[hMap["次年度用\n前年度からの変更"]];
    
    // 除外判定
    if (retireDate instanceof Date) continue;
    if (String(nextChange).includes("退職") || String(nextChange).includes("満了")) continue;

    // 27列構成の配列を作成
    const newRow = new Array(27).fill("");

    newRow[0] = rowCount;                  // [1] 番号
    newRow[1] = row[hMap["医籍番号"]];       // [2] 医籍番号
    newRow[2] = row[hMap["医師名"]];        // [3] 医師名
    newRow[3] = row[hMap["入職日"]];        // [4] 入職日
    newRow[4] = row[hMap["専門"]];          // [5] 専門
    newRow[5] = "";                        // [6] 退職日
    newRow[6] = row[hMap["主務"]] || "";    // [7] 主務
    newRow[7] = row[hMap["祝日"]];          // [8] 祝日
    newRow[8] = row[hMap["年末年始"]];       // [9] 年末年始
    // [10] 保留 (後述)
    // [11] 対応不要 (空欄)

    const currentShiftText = row[hMap["勤務備考"]];
    const isChange = (String(nextChange) === "あり");

    newRow[11] = currentShiftText; // [12] シフト記入例（実績）
    newRow[12] = templateRuleText; // [13] 書き方ルール（テンプレートから引用）

    if (!isChange) {
      // 変更なし -> 保留OFF、上期固定で日付更新
      newRow[9] = false; 
      newRow[13] = generateKamikiShiftText(currentShiftText, nextYearNum); // [14] 提案シフト
    } else {
      // 変更あり -> 保留ON、提案シフト空欄
      newRow[9] = true;  
      newRow[13] = "";   
    }
    
    newRow[16] = "時給表どおり"; // [17] 契約時給
    
    outputRows.push(newRow);
    rowCount++;
  }

  // --- 3. シート作成と設定 ---
  
  // 既存の同名シートがあれば削除
  const existingSheet = destSS.getSheetByName(destSheetName);
  if (existingSheet) destSS.deleteSheet(existingSheet);

  // テンプレートをコピーして新しいシートを作成
  const newSheet = templateSheet.copyTo(destSS);
  newSheet.setName(destSheetName);

  // --- 4. ヘッダーの動的書き換え ---
  // テンプレートのヘッダーを取得
  const headerRange = newSheet.getRange(1, 1, 1, 27);
  const headerValues = headerRange.getValues()[0];

  // 列番号は0始まりのインデックス: 12列目->idx11, 14列目->idx13
  headerValues[11] = `${currentYearNum}年度シフト（通期）記入例`;
  headerValues[13] = `${nextYearNum}年度提案シフト`;

  // 書き戻し
  headerRange.setValues([headerValues]);

  // --- 5. データの書き込み ---
  
  // テンプレートに入っていた既存データ（2行目以降）をクリア
  // ※書式は残したいので clearContent() を使用
  const currentMaxRows = newSheet.getMaxRows();
  if (currentMaxRows > 1) {
    newSheet.getRange(2, 1, currentMaxRows - 1, 27).clearContent();
    // チェックボックスの状態などはクリアされないことがあるため、明示的にUncheckが必要ならここで行う
    // (今回は値の上書きで対応)
  }

  if (outputRows.length > 0) {
    newSheet.getRange(2, 1, outputRows.length, 27).setValues(outputRows);
  }

  // --- 6. 行数の最適化 (削除または追加) ---
  const requiredRows = outputRows.length + 1; // ヘッダー(1) + データ数

  if (currentMaxRows > requiredRows) {
    // 行が余っている場合 -> 削除
    newSheet.deleteRows(requiredRows + 1, currentMaxRows - requiredRows);
  } else if (currentMaxRows < requiredRows) {
    // 行が足りない場合 -> 追加
    newSheet.insertRowsAfter(currentMaxRows, requiredRows - currentMaxRows);
  }

  return outputRows.length;
}

/**
 * 上期固定の日付変換（4/1～9/30）
 */
function generateKamikiShiftText(text, nextYear) {
  if (!text) return "";
  const targetTerm = `${nextYear}/04/01～${nextYear}/09/30`;
  
  // 日付行があれば置換
  if (/\d{4}[\/\-]\d{1,2}[\/\-]\d{1,2}.*?[\n\r]/.test(text)) {
    return text.replace(/\d{4}[\/\-]\d{1,2}[\/\-]\d{1,2}.*?[\n\r]/, targetTerm + "\n");
  }
  // なければ先頭に追加
  return targetTerm + "\n" + text;
}