// ==========================================
// 【第3の機構】手動入力された例外（お休み・振替）をシフト表＆勤怠表に一括送信
// （✅ 同一スクリプト内・直接起動版）
// ==========================================

function syncManualExceptions() {
  const ui = SpreadsheetApp.getUi();
  
  // 勤怠表（大元）のスプレッドシートID
  const ATTENDANCE_SS_ID = "1aEjphEv_63SeWQmwiOy9sx7IrMfawU01sHbKd_Ki4iA"; 

  const activeSs = SpreadsheetApp.getActiveSpreadsheet(); // これがシフト表本体のスプシ
  let exceptionSs = activeSs;
  // もし例外シートが別ファイルなら開く（今回は同じファイルのようなのでここはそのまま通過します）
  if (typeof EXCEPTION_SS_ID !== "undefined" && EXCEPTION_SS_ID !== "") {
    try { exceptionSs = SpreadsheetApp.openById(EXCEPTION_SS_ID); } catch (e) {}
  }
  
  let attendanceSs;
  try {
    attendanceSs = SpreadsheetApp.openById(ATTENDANCE_SS_ID);
  } catch (e) {
    ui.alert("エラー", "勤怠表スプレッドシートが開けません。権限やIDを確認してください。", ui.ButtonSet.OK);
    return;
  }
  const attendanceSheets = attendanceSs.getSheets().filter(s => /^(常勤|定期非常勤)勤怠\d{4}/.test(s.getName()));
  
  const sheetNames = ["お休み情報", "振替勤務"];
  let pendingUpdates = [];

  const getTimeStr = (val) => {
    if (!val) return "";
    if (val instanceof Date) return Utilities.formatDate(val, "JST", "HH:mm");
    return String(val).trim();
  };

  sheetNames.forEach(sName => {
    const sheet = exceptionSs.getSheetByName(sName);
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return;

    const headers = data[0].map(h => String(h).replace(/[\s　]+/g, ""));
    const cDate = headers.findIndex(h => /日付|勤務日/.test(h));
    const cName = headers.findIndex(h => /氏名|名前|医師名/.test(h));
    const cStart = headers.findIndex(h => /開始/.test(h));
    const cEnd = headers.findIndex(h => /終了/.test(h));
    const cLoc = headers.findIndex(h => /拠点|クリニック|店舗/.test(h));
    const cType = headers.findIndex(h => /種別|タイプ/.test(h));

    if (cDate === -1 || cLoc === -1 || cName === -1) return;

    const checkColIdx = (sName === "お休み情報") ? 8 : 9; // I列(8) か J列(9)

    for (let i = 1; i < data.length; i++) {
      const dateCell = data[i][cDate];
      const loc = String(data[i][cLoc]).trim();
      const docName = String(data[i][cName]).trim() || "名前なし";
      
      const statusVal = data[i][checkColIdx];
      const isAlreadyDone = (statusVal === true || String(statusVal).toLowerCase() === "true" || String(statusVal) === "反映済" || String(statusVal) === "済");

      if (dateCell && !isAlreadyDone) {
        const dateObj = new Date(dateCell);
        const year = (dateObj.getMonth() >= 3) ? dateObj.getFullYear() : dateObj.getFullYear() - 1;
        const dateStr = Utilities.formatDate(dateObj, "JST", "yyyy年M月d日");
        const targetDateStr = Utilities.formatDate(dateObj, "JST", "yyyy/MM/dd");
        
        const startStr = cStart > -1 ? getTimeStr(data[i][cStart]) : "";
        const endStr = cEnd > -1 ? getTimeStr(data[i][cEnd]) : "";
        let timeDisplay = (startStr || endStr) ? `${startStr}-${endStr}` : "時間指定なし";
        let absenceType = cType > -1 ? String(data[i][cType]).trim() : "欠勤";
        
        pendingUpdates.push({
          sheetName: sName,
          docName: docName,
          docNameClean: docName.replace(/[\s　]+/g, ""),
          locDisplay: loc.includes("【") ? loc : `【${loc}】`,
          cleanLoc: loc.replace(/[【】]/g, '').split('_')[0],
          dateDisplay: dateStr,
          targetDateStr: targetDateStr,
          timeDisplay: timeDisplay,
          startStr: startStr,
          endStr: endStr,
          absenceType: absenceType || "欠勤",
          rowNum: i + 1,
          year: String(year),
          sheetObj: sheet,
          statusCol: checkColIdx + 1 
        });
      }
    }
  });

  if (pendingUpdates.length === 0) {
    activeSs.toast("新しくシフト表・勤怠表へ追加するデータはありませんでした。", "確認", 3);
    return;
  }

  let confirmMsg = `以下の ${pendingUpdates.length} 件をシフト表 ＆ 勤怠表に反映します。\nよろしいですか？\n\n`;
  let displayCount = 0;

  sheetNames.forEach(sName => {
    let itemsForSheet = pendingUpdates.filter(p => p.sheetName === sName);
    if (itemsForSheet.length > 0) {
      confirmMsg += `■ ${sName}\n`;
      let docs = [...new Set(itemsForSheet.map(i => i.docName))];
      docs.forEach(doc => {
        let docItems = itemsForSheet.filter(i => i.docName === doc);
        docItems.forEach(item => {
          if (displayCount < 15) { 
            confirmMsg += `医師: ${item.docName} | ${item.locDisplay} | ${item.dateDisplay}\n`;
          }
          displayCount++;
        });
      });
      confirmMsg += "\n";
    }
  });

  if (displayCount > 15) confirmMsg += `※他 ${displayCount - 15} 件の追加があります。\n\n`;

  const response = ui.alert('一括反映の確認', confirmMsg, ui.ButtonSet.OK_CANCEL);
  if (response !== ui.Button.OK) {
    activeSs.toast("反映処理をキャンセルしました。", "キャンセル", 3);
    return;
  }

  activeSs.toast("勤怠表の書き換えとシフト表の更新を行っています...", "処理中", 10);

  let affectedLocationsByYear = {};

  pendingUpdates.forEach(u => {
    // --- 【A】勤怠表セルの書き換え ---
    for (let s of attendanceSheets) {
      if (!s.getName().includes(u.year)) continue;

      const attData = s.getDataRange().getValues();
      if (attData.length < 3) continue;

      const namesRow0 = attData[0].map(h => String(h).replace(/[\s　]+/g, ""));
      const namesRow1 = attData[1].map(h => String(h).replace(/[\s　]+/g, ""));
      let docColIdx = namesRow0.indexOf(u.docNameClean);
      if (docColIdx === -1) docColIdx = namesRow1.indexOf(u.docNameClean);

      if (docColIdx === -1) continue;

      let foundRow = -1;
      for (let r = 2; r < attData.length; r++) {
        let attDateCell = attData[r][0];
        let attDateStr = "";
        if (attDateCell instanceof Date) {
          attDateStr = Utilities.formatDate(attDateCell, "JST", "yyyy/MM/dd");
        } else if (attDateCell) {
          let dMatch = String(attDateCell).match(/(\d{4})[\/\-年](\d{1,2})[\/\-月](\d{1,2})/);
          if (dMatch) attDateStr = `${dMatch[1]}/${dMatch[2].padStart(2, '0')}/${dMatch[3].padStart(2, '0')}`;
        }

        if (attDateStr === u.targetDateStr) {
          foundRow = r;
          break;
        }
      }

      if (foundRow !== -1) {
        const cell = s.getRange(foundRow + 1, docColIdx + 1);
        let currentVal = String(cell.getValue() || "");

        if (u.sheetName === "お休み情報") {
          let cleanLines = currentVal.split('\n').filter(line => !line.startsWith('→') && !line.includes('※振替') && !line.includes('半日有給') && !line.includes('有給') && !line.includes('欠勤') && !line.includes('移動依頼'));
          let baseVal = cleanLines.join('\n');
          cell.setValue(baseVal + (baseVal ? "\n" : "") + `→${u.absenceType}`);
          cell.setBackground("#e8f0fe"); 
        } else if (u.sheetName === "振替勤務") {
          cell.setValue(`【${u.cleanLoc}】${u.startStr}-${u.endStr}`);
          cell.setBackground(null); 
        }
        break; 
      }
    }

    // --- 【B】例外シート側に「反映済」の書き込み ＆ チェック ---
    const checkBoxCell = u.sheetObj.getRange(u.rowNum, u.statusCol);
    checkBoxCell.insertCheckboxes();
    checkBoxCell.setValue(true);

    // --- 【C】バッチ送信用に拠点をストック ---
    if (!affectedLocationsByYear[u.year]) affectedLocationsByYear[u.year] = new Set();
    affectedLocationsByYear[u.year].add(u.cleanLoc);
  });

  // =========================================
  // ★ココが変更点：APIをやめて直接 Batch Engine を叩き起こす！
  // =========================================
  for (const year in affectedLocationsByYear) {
    if (affectedLocationsByYear[year].size > 0) {
      const locationsArr = Array.from(affectedLocationsByYear[year]);
      console.log(`🚀 [直接起動] ${year}年度のバッチ処理を開始 => 対象拠点: ${locationsArr.join(",")}`);
      
      // 同じGAS内にある 00_Batch_Engine の関数に、必要な情報を渡して直接実行！
      startBackgroundBatch({
        year: year,
        term: "通年", // 通年指定にして対象拠点のシフトを再描画させる
        locations: locationsArr
      });
    }
  }

  activeSs.toast(`${pendingUpdates.length}件のデータを勤怠表に反映し、シフト表の更新バッチを起動しました！モニターを確認してください。`, "✅ 処理開始", 8);
}
// ==========================================
// 【第3の機構：自動監視用】例外データの自動反映（UI・確認なし）
// ※この関数を時間主導型トリガー（1時間おき等）に設定してください
// ==========================================

function autoSyncManualExceptions() {
  const ATTENDANCE_SS_ID = "1aEjphEv_63SeWQmwiOy9sx7IrMfawU01sHbKd_Ki4iA"; 

  const activeSs = SpreadsheetApp.getActiveSpreadsheet();
  let exceptionSs = activeSs;
  if (typeof EXCEPTION_SS_ID !== "undefined" && EXCEPTION_SS_ID !== "") {
    try { exceptionSs = SpreadsheetApp.openById(EXCEPTION_SS_ID); } catch (e) {}
  }
  
  let attendanceSs;
  try {
    attendanceSs = SpreadsheetApp.openById(ATTENDANCE_SS_ID);
  } catch (e) {
    console.error("❌ 勤怠表スプレッドシートが開けません。");
    return;
  }
  const attendanceSheets = attendanceSs.getSheets().filter(s => /^(常勤|定期非常勤)勤怠\d{4}/.test(s.getName()));
  
  const sheetNames = ["お休み情報", "振替勤務"];
  let pendingUpdates = [];

  const getTimeStr = (val) => {
    if (!val) return "";
    if (val instanceof Date) return Utilities.formatDate(val, "JST", "HH:mm");
    return String(val).trim();
  };

  sheetNames.forEach(sName => {
    const sheet = exceptionSs.getSheetByName(sName);
    if (!sheet) return;
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return;

    const headers = data[0].map(h => String(h).replace(/[\s　]+/g, ""));
    const cDate = headers.findIndex(h => /日付|勤務日/.test(h));
    const cName = headers.findIndex(h => /氏名|名前|医師名/.test(h));
    const cStart = headers.findIndex(h => /開始/.test(h));
    const cEnd = headers.findIndex(h => /終了/.test(h));
    const cLoc = headers.findIndex(h => /拠点|クリニック|店舗/.test(h));
    const cType = headers.findIndex(h => /種別|タイプ/.test(h));

    if (cDate === -1 || cLoc === -1 || cName === -1) return;

    const checkColIdx = (sName === "お休み情報") ? 8 : 9; // I列(8) か J列(9)

    for (let i = 1; i < data.length; i++) {
      const dateCell = data[i][cDate];
      const loc = String(data[i][cLoc]).trim();
      const docName = String(data[i][cName]).trim() || "名前なし";
      
      const statusVal = data[i][checkColIdx];
      const isAlreadyDone = (statusVal === true || String(statusVal).toLowerCase() === "true" || String(statusVal) === "反映済" || String(statusVal) === "済");

      if (dateCell && !isAlreadyDone) {
        const dateObj = new Date(dateCell);
        const year = (dateObj.getMonth() >= 3) ? dateObj.getFullYear() : dateObj.getFullYear() - 1;
        const targetDateStr = Utilities.formatDate(dateObj, "JST", "yyyy/MM/dd");
        
        const startStr = cStart > -1 ? getTimeStr(data[i][cStart]) : "";
        const endStr = cEnd > -1 ? getTimeStr(data[i][cEnd]) : "";
        let absenceType = cType > -1 ? String(data[i][cType]).trim() : "欠勤";
        
        pendingUpdates.push({
          sheetName: sName,
          docNameClean: docName.replace(/[\s　]+/g, ""),
          cleanLoc: loc.replace(/[【】]/g, '').split('_')[0],
          targetDateStr: targetDateStr,
          startStr: startStr,
          endStr: endStr,
          absenceType: absenceType || "欠勤",
          rowNum: i + 1,
          year: String(year),
          sheetObj: sheet,
          statusCol: checkColIdx + 1 
        });
      }
    }
  });

  // 対象がなければ無言でスキップして終了
  if (pendingUpdates.length === 0) {
    console.log("ℹ️ 追加・反映する例外データはありませんでした（スキップ）。");
    return;
  }

  console.log(`🚀 ${pendingUpdates.length}件の未反映データを検知。自動反映を開始します。`);

  let affectedLocationsByYear = {};

  pendingUpdates.forEach(u => {
    // --- 【A】勤怠表セルの書き換え ---
    for (let s of attendanceSheets) {
      if (!s.getName().includes(u.year)) continue;

      const attData = s.getDataRange().getValues();
      if (attData.length < 3) continue;

      const namesRow0 = attData[0].map(h => String(h).replace(/[\s　]+/g, ""));
      const namesRow1 = attData[1].map(h => String(h).replace(/[\s　]+/g, ""));
      let docColIdx = namesRow0.indexOf(u.docNameClean);
      if (docColIdx === -1) docColIdx = namesRow1.indexOf(u.docNameClean);

      if (docColIdx === -1) continue;

      let foundRow = -1;
      for (let r = 2; r < attData.length; r++) {
        let attDateCell = attData[r][0];
        let attDateStr = "";
        if (attDateCell instanceof Date) {
          attDateStr = Utilities.formatDate(attDateCell, "JST", "yyyy/MM/dd");
        } else if (attDateCell) {
          let dMatch = String(attDateCell).match(/(\d{4})[\/\-年](\d{1,2})[\/\-月](\d{1,2})/);
          if (dMatch) attDateStr = `${dMatch[1]}/${dMatch[2].padStart(2, '0')}/${dMatch[3].padStart(2, '0')}`;
        }

        if (attDateStr === u.targetDateStr) {
          foundRow = r;
          break;
        }
      }

      if (foundRow !== -1) {
        const cell = s.getRange(foundRow + 1, docColIdx + 1);
        let currentVal = String(cell.getValue() || "");

        if (u.sheetName === "お休み情報") {
          let cleanLines = currentVal.split('\n').filter(line => !line.startsWith('→') && !line.includes('※振替') && !line.includes('半日有給') && !line.includes('有給') && !line.includes('欠勤') && !line.includes('移動依頼'));
          let baseVal = cleanLines.join('\n');
          cell.setValue(baseVal + (baseVal ? "\n" : "") + `→${u.absenceType}`);
          cell.setBackground("#e8f0fe"); 
        } else if (u.sheetName === "振替勤務") {
          cell.setValue(`【${u.cleanLoc}】${u.startStr}-${u.endStr}`);
          cell.setBackground(null); 
        }
        break; 
      }
    }

    // --- 【B】例外シート側に「反映済」のチェック ---
    const checkBoxCell = u.sheetObj.getRange(u.rowNum, u.statusCol);
    checkBoxCell.insertCheckboxes();
    checkBoxCell.setValue(true);

    // --- 【C】バッチ送信用に拠点をストック ---
    if (!affectedLocationsByYear[u.year]) affectedLocationsByYear[u.year] = new Set();
    affectedLocationsByYear[u.year].add(u.cleanLoc);
  });

  // --- 【D】バッチ処理の起動 ---
  for (const year in affectedLocationsByYear) {
    if (affectedLocationsByYear[year].size > 0) {
      const locationsArr = Array.from(affectedLocationsByYear[year]);
      console.log(`🚀 [自動起動] ${year}年度バッチ開始 => 拠点: ${locationsArr.join(",")}`);
      
      startBackgroundBatch({
        year: year,
        term: "通年", 
        locations: locationsArr
      });
    }
  }
}