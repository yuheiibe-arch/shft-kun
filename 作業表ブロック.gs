function updateShiftStatus() {
  // =========================================================
  // 設定項目
  // =========================================================
  
  // 外部シートの設定
  const EXTERNAL_SS_URL = 'https://docs.google.com/spreadsheets/d/13PlAElj8SODdBNME28TIDpifxG0oxS3wAnKIsFTa53k/edit';
  
  // (調整中)常勤2026年度 シート
  const EXT_JOUKIN_SHEET_NAME = '(調整中)常勤2026年度';
  const EXT_JOUKIN_NAME_COL = 4;   // D列: 医師名
  const EXT_JOUKIN_STATUS_COL = 56; // BD列: 医師からの承諾

  // (調整中)定期非常勤2026年度 シート
  const EXT_HIJOUKIN_SHEET_NAME = '(調整中)定期非常勤2026年度';
  const EXT_HIJOUKIN_NAME_COL = 3;   // C列: 医師名
  const EXT_HIJOUKIN_STATUS_COL = 25; // Y列: 対応状況

  // 現在のシート（2026作業表）の設定
  const WORK_SHEET_NAME = '2026作業表';
  const WORK_NAME_COL = 4;     // D列: 医師名
  const WORK_CHECK_COL = 5;    // E列: チェックボックス
  const WORK_PENDING_COL = 2;  // B列: 保留チェックボックス
  const WORK_BASE_COL = 1;     // ★A列: 拠点名（3から1に修正しました！）
  const WORK_START_ROW = 3;    // 3行目からデータ開始

  // 確定シフト作成シートの設定
  const SHIFT_SHEET_NAME = '確定シフト作成';
  const SHIFT_NAME_COL = 4;    // D列: 医師名（〇〇先生）
  const SHIFT_BASE_COL = 2;    // B列: 拠点名
  const SHIFT_START_ROW = 2;   // 2行目から開始

  // 定期募集シートの設定
  const REGULAR_SHEET_NAME = '定期募集';
  const REGULAR_BASE_COL = 2;  // B列: 募集枠の文字列（拠点名／時間...）
  const REGULAR_START_ROW = 2; // データ開始行

  // 色の設定
  const COLOR_LIGHT_RED = '#f4cccc';
  const COLOR_LIGHT_BLUE = '#cfe2f3'; 
  const COLOR_WHITE = null; // 背景色リセット

  // =========================================================
  // 名前のゆらぎを吸収する関数（空白と「先生」を消去して比較）
  // =========================================================
  function normalizeName(name) {
    if (!name) return '';
    return String(name).replace(/[\s ]+/g, '').replace(/先生$/, '');
  }

  // =========================================================
  // 1. 外部シートから「全対象医師」と「完了」のリストを取得
  // =========================================================
  // ★ ここを safeOpenByUrl に変更
  const extSs = safeOpenByUrl(EXTERNAL_SS_URL);
  
  const targetDoctors = new Set();    // 常勤・定期非常勤の全医師リスト
  const completedDoctors = new Set(); // そのうち「完了」している医師リスト

  // 常勤シートの処理
  const sheetJoukin = extSs.getSheetByName(EXT_JOUKIN_SHEET_NAME);
  if (sheetJoukin) {
    const dataJoukin = sheetJoukin.getDataRange().getValues();
    for (let i = 1; i < dataJoukin.length; i++) { 
      let doctorName = normalizeName(dataJoukin[i][EXT_JOUKIN_NAME_COL - 1]);
      let status = String(dataJoukin[i][EXT_JOUKIN_STATUS_COL - 1]).trim();
      
      if (doctorName !== '') {
        targetDoctors.add(doctorName); 
        if (status === '完了') {
          completedDoctors.add(doctorName); 
        }
      }
    }
  }

  // 定期非常勤シートの処理
  const sheetHijoukin = extSs.getSheetByName(EXT_HIJOUKIN_SHEET_NAME);
  if (sheetHijoukin) {
    const dataHijoukin = sheetHijoukin.getDataRange().getValues();
    for (let i = 1; i < dataHijoukin.length; i++) {
      let doctorName = normalizeName(dataHijoukin[i][EXT_HIJOUKIN_NAME_COL - 1]);
      let status = String(dataHijoukin[i][EXT_HIJOUKIN_STATUS_COL - 1]).trim();
      
      if (doctorName !== '') {
        targetDoctors.add(doctorName); 
        if (status === '完了') {
          completedDoctors.add(doctorName); 
        }
      }
    }
  }

  // =========================================================
  // 2. 「2026作業表」シートの処理
  // =========================================================
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetWork = ss.getSheetByName(WORK_SHEET_NAME);
  const pendingBases = new Set(); // 保留になっている拠点を記録するリスト
  
  if (sheetWork) {
    const lastRow = sheetWork.getLastRow();
    if (lastRow >= WORK_START_ROW) {
      const workDataRange = sheetWork.getRange(WORK_START_ROW, 1, lastRow - WORK_START_ROW + 1, Math.max(WORK_NAME_COL, WORK_CHECK_COL, WORK_PENDING_COL, WORK_BASE_COL));
      const workValues = workDataRange.getValues();
      const checkValues = []; 

      for (let i = 0; i < workValues.length; i++) {
        let doctorName = normalizeName(workValues[i][WORK_NAME_COL - 1]);
        let isPending = workValues[i][WORK_PENDING_COL - 1]; 
        let baseName = String(workValues[i][WORK_BASE_COL - 1]).trim(); 

        let isCompleted = completedDoctors.has(doctorName) ? true : false;
        checkValues.push([isCompleted]);

        if (isPending === true && baseName !== '') {
          pendingBases.add(baseName); // 保留の拠点をリストに追加
        }
      }

      sheetWork.getRange(WORK_START_ROW, WORK_CHECK_COL, checkValues.length, 1).setValues(checkValues);

      const colorValues = []; 
      for (let i = 0; i < workValues.length; i++) {
        let baseName = String(workValues[i][WORK_BASE_COL - 1]).trim();
        if (baseName !== '' && pendingBases.has(baseName)) {
          colorValues.push([COLOR_LIGHT_RED]); 
        } else {
          colorValues.push([COLOR_WHITE]);     
        }
      }
      sheetWork.getRange(WORK_START_ROW, WORK_BASE_COL, colorValues.length, 1).setBackgrounds(colorValues);
    }
  }

  // =========================================================
  // 3. 「確定シフト作成」シートの処理
  // =========================================================
  const sheetShift = ss.getSheetByName(SHIFT_SHEET_NAME);
  if (sheetShift) {
    const lastRowShift = sheetShift.getLastRow();
    if (lastRowShift >= SHIFT_START_ROW) {
      // 拠点列(B)と医師名列(D)を含む範囲を一括で取得
      const maxCol = Math.max(SHIFT_NAME_COL, SHIFT_BASE_COL);
      const shiftData = sheetShift.getRange(SHIFT_START_ROW, 1, lastRowShift - SHIFT_START_ROW + 1, maxCol).getValues();
      
      const shiftNameColors = [];
      const shiftBaseColors = [];

      for (let i = 0; i < shiftData.length; i++) {
        let baseName = String(shiftData[i][SHIFT_BASE_COL - 1]).trim();
        let doctorName = normalizeName(shiftData[i][SHIFT_NAME_COL - 1]);

        // ① 拠点名が保留になっているかチェック
        let isBasePending = (baseName !== '' && pendingBases.has(baseName));

        // 拠点名の色の設定
        if (isBasePending) {
          shiftBaseColors.push([COLOR_LIGHT_RED]); 
        } else {
          shiftBaseColors.push([COLOR_WHITE]); 
        }

        // ② 医師名の色の設定
        if (doctorName === '') {
          shiftNameColors.push([COLOR_WHITE]); 
        } else if (isBasePending) {
          shiftNameColors.push([COLOR_LIGHT_RED]); // ★拠点が赤色なら、医師名も赤色にする
        } else if (completedDoctors.has(doctorName)) {
          shiftNameColors.push([COLOR_LIGHT_BLUE]); 
        } else if (targetDoctors.has(doctorName)) {
          shiftNameColors.push([COLOR_LIGHT_RED]); 
        } else {
          shiftNameColors.push([COLOR_LIGHT_BLUE]); 
        }
      }
      
      // 背景色をそれぞれ更新
      sheetShift.getRange(SHIFT_START_ROW, SHIFT_BASE_COL, shiftBaseColors.length, 1).setBackgrounds(shiftBaseColors);
      sheetShift.getRange(SHIFT_START_ROW, SHIFT_NAME_COL, shiftNameColors.length, 1).setBackgrounds(shiftNameColors);
    }
  }

  // =========================================================
  // 4. 「定期募集」シートの処理
  // =========================================================
  const sheetRegular = ss.getSheetByName(REGULAR_SHEET_NAME);
  if (sheetRegular) {
    const lastRowReg = sheetRegular.getLastRow();
    if (lastRowReg >= REGULAR_START_ROW) {
      // 文字列の処理 (B列)
      const regBaseRange = sheetRegular.getRange(REGULAR_START_ROW, REGULAR_BASE_COL, lastRowReg - REGULAR_START_ROW + 1, 1);
      const regBaseValues = regBaseRange.getValues();
      const regBaseColors = [];

      for (let i = 0; i < regBaseValues.length; i++) {
        let rawText = String(regBaseValues[i][0]).trim();
        
        // スラッシュ（全角・半角）で分割し、最初の部分（拠点名）だけを取得
        let baseName = rawText.split(/[／/]/)[0].trim();

        if (baseName !== '' && pendingBases.has(baseName)) {
          regBaseColors.push([COLOR_LIGHT_RED]); // 保留リストにあれば薄い赤
        } else {
          regBaseColors.push([COLOR_WHITE]); // それ以外は白（リセット）
        }
      }
      regBaseRange.setBackgrounds(regBaseColors);
    }
  }
}