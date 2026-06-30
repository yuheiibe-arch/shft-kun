/**
 * ====================================================================
 * 03_WageEngine.gs
 * 【汎用モジュール】外部スクリプトから呼び出すための時給計算コアエンジン
 * ★ 拠点ID（番号）未設定時は処理をスキップする安全対応版
 * ====================================================================
 */

const NewWageEngine = (function() {
  const CONFIG = {
    WAGE_MASTER_URL: "https://docs.google.com/spreadsheets/d/14RbsDcv0nXfEwweki8-9cK3lQUg1XUuhozLNF9u2qAs/edit",
    SPECIAL_WAGE_SHEET: "特別時給医師"
  };

  let _isInitialized = false;
  let _locDicts = null;
  let _baseWageDB = null;
  let _specialWageDB = null;

  /**
   * マスタデータの初期読み込み（初回のみ実行）
   */
  function _init() {
    if (_isInitialized) return;
    try {
      const wageMasterSs = safeOpenByUrl(CONFIG.WAGE_MASTER_URL);
      _locDicts = _buildLocationDictionary(wageMasterSs);
      _baseWageDB = _buildBaseWageDB(wageMasterSs, _locDicts.nameMap);
      _specialWageDB = _buildSpecialWageDB(wageMasterSs);

      _isInitialized = true;
    } catch (e) {
      throw new Error(`【システムエラー】時給マスタの初期化に失敗しました。URLやシート名を確認してください。詳細: ${e.message}`);
    }
  }

  /**
   * 【メインAPI】各コマの時給を算出する
   */
  function calculate(medId, rawLoc, docSpecialty, workDate, rawStartTime, rawEndTime, isHol = false, shiftFullTimeStr = "", contractText = "") {
    _init(); 

    // 1. 日付から曜日区分を判定
    if (!(workDate instanceof Date) || isNaN(workDate.getTime())) {
      throw new Error("【システムエラー】有効なDateオブジェクトが渡されていません。");
    }

    const day = workDate.getDay();
    let dayType = "平日";
    if (day === 6) dayType = "土曜";
    if (day === 0 || isHol) dayType = "日祝";

    // 2. 拠点の正規化と存在チェック
    const locName = _normalizeLocForChecker(rawLoc, _locDicts.nameMap);
    
    // 拠点名が見つからない場合はエラー
    if (!locName) {
      throw new Error(`【重大エラー: 拠点不明】入力された拠点「${rawLoc}」は拠点名マスタに登録されていません。`);
    }

    const clinicId = _locDicts.idMap[locName];
    // ★ ご指示の通り：ID（番号）が設定されていない拠点は「未登録」とみなし、エラーを出さずに処理をスキップする
    if (!clinicId || clinicId === "") {
      return []; 
    }

    // 診療科の自動判定（亀有・北葛西の特例ルール）
    const isNaika = String(rawLoc).includes("内科") || String(docSpecialty).includes("内科");
    const deptReq = (locName === "亀有" || locName === "北葛西") ? (isNaika ? "内科" : "小児科") : "小児科";

    // 専門科目の一致チェック
    if (deptReq !== docSpecialty) {
      throw new Error(`【重大エラー: 専門外勤務】拠点「${locName}」の要請科目(${deptReq})に対し、医師の専門(${docSpecialty})が不一致です。`);
    }

    // 3. 通し時間をコマに分割（13:00-15:00の休診時間は自動除外）
    const slots = _splitIntoSlots(rawStartTime, rawEndTime);
    
    if (slots.length === 0) {
      throw new Error(`【重大エラー: シフト時間異常】指定時間（${rawStartTime}-${rawEndTime}）が規定コマに存在しないか、絶対的休診時間(13-15時)に該当します。`);
    }

    // シフト全体の時間文字列を生成（マスタ検索用）
    if (!shiftFullTimeStr) {
      const normalizeTime = (t) => {
        if (!t) return "";
        let m = String(t).match(/(\d{1,2})[:：](\d{2})/);
        return m ? `${('0' + m[1]).slice(-2)}:${m[2]}` : String(t);
      };
      shiftFullTimeStr = `${normalizeTime(rawStartTime)}-${normalizeTime(rawEndTime)}`;
    }

    // 4. 各コマの時給を算出
    const results = [];
    slots.forEach(slot => {
      const wageResult = _calculateSingleSlotWage({
        medId, clinicId, locName, deptReq, docSpecialty, 
        workDate, dayType, 
        slotStartTime: slot.start, slotEndTime: slot.end, slotName: slot.name,
        shiftFullTimeStr: shiftFullTimeStr,
        contractText: contractText,
        baseWageDB: _baseWageDB, specialWageDB: _specialWageDB
      });

      results.push({
        slotName: slot.name,
        startTime: slot.start,
        endTime: slot.end,
        wage: wageResult.wage,
        baseWage: wageResult.baseWage,
        wageType: wageResult.type
      });
    });

    return results;
  }

  // ====================================================================
  // 以下、内部ヘルパー関数
  // ====================================================================

  function _buildLocationDictionary(ss) {
    const sheet = ss.getSheetByName("拠点名");
    if (!sheet) throw new Error("拠点名マスタシートが見つかりません。");
    const data = sheet.getDataRange().getValues();
    let nameMap = {}, idMap = {};
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      let canonical = row[0] ? String(row[0]).trim() : "";
      let cId = row[9] ? String(row[9]).trim() : ""; 
      if (!canonical) continue;
      
      nameMap[canonical] = canonical;
      if (cId) idMap[canonical] = cId;
      
      for (let j = 1; j <= 4; j++) {
        let variant = row[j] ? String(row[j]).replace(/[\s ]+/g, "") : "";
        if (variant) nameMap[variant] = canonical;
      }
    }
    return { nameMap, idMap };
  }

  function _normalizeLocForChecker(rawText, dict) {
    if (!rawText) return "";
    let cleanName = String(rawText).replace(/[【】\(（]?(内科|小児科)[\)）]?/g, "").replace(/\/.*/, "").replace(/[\s ]+/g, "");
    return dict[cleanName] || cleanName;
  }

  function _buildBaseWageDB(ss, nameMap) {
    const db = { '2025': {}, '2026上期': {}, '2026下期': {} };
    const targetSheets = ['2025時給', '2026上期時給', '2026下期時給'];
    
    targetSheets.forEach(sheetName => {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet) return; 
      
      let yearKey = sheetName.replace('時給', '');
      const values = sheet.getDataRange().getValues();
      
      for (let i = 1; i < values.length; i++) {
        const rawLoc = String(values[i][1]).trim();
        const dept = String(values[i][2]).trim();
        if (!rawLoc || !dept) continue;
        
        const canonicalLoc = _normalizeLocForChecker(rawLoc, nameMap);
        db[yearKey][`${canonicalLoc}_${dept}`] = {
          rates: {
            '平日': { '午前': values[i][3], '午後': values[i][4], '夜間': values[i][5] },
            '土曜': { '午前': values[i][6], '午後': values[i][7], '夜間': values[i][8] },
            '日祝': { '午前': values[i][9], '午後': values[i][10], '夜間': values[i][11] }
          }
        };
      }
    });
    return db;
  }

  function _buildSpecialWageDB(ss) {
    const db = {};
    const sheet = ss.getSheetByName(CONFIG.SPECIAL_WAGE_SHEET);
    if (!sheet) return db; 
    
    const values = sheet.getDataRange().getValues();

    const addWage = (key, wage) => {
      if (!db[key]) db[key] = [];
      db[key].push(wage);
    };

    for (let i = 1; i < values.length; i++) {
      const medId = String(values[i][0]).trim();
      const locName = String(values[i][3]).trim(); 
      const cId = String(values[i][4]).trim();
      let dayType = String(values[i][5]).trim();
      
      let rawTime = String(values[i][6]).replace(/[～〜]/g, "-").replace(/[\s ]+/g, "").trim();
      let timeParts = rawTime.split('-');
      if (timeParts.length === 2) {
         const pad = (t) => {
            let m = t.match(/(\d{1,2})[:：](\d{2})/);
            return m ? `${('0' + m[1]).slice(-2)}:${m[2]}` : t;
         };
         rawTime = `${pad(timeParts[0])}-${pad(timeParts[1])}`;
      }
      let timeStr = rawTime;

      const wage = parseInt(values[i][7], 10);
      
      if (isNaN(wage)) continue; 
      if (dayType === "全日") dayType = "ALL";
      
      if (cId) addWage(`${medId}_${cId}_${dayType}_${timeStr}`, wage);
      if (locName) addWage(`${medId}_NAME:${locName}_${dayType}_${timeStr}`, wage);
    }
    return db;
  }

  function _splitIntoSlots(rawStart, rawEnd) {
    const toMinutes = t => { 
      const [h, m] = String(t).split(':').map(Number); 
      return (h || 0) * 60 + (m || 0); 
    };
    
    const startMin = toMinutes(rawStart);
    let endMin = toMinutes(rawEnd);
    if (endMin < startMin && endMin !== 0) endMin += 24 * 60; 
    
    const DEFINED_SLOTS = [
      { name: "午前", start: "09:00", end: "13:00", sMin: 9*60, eMin: 13*60 },
      { name: "午後", start: "15:00", end: "18:00", sMin: 15*60, eMin: 18*60 },
      { name: "夜間", start: "18:00", end: "21:00", sMin: 18*60, eMin: 21*60 }
    ];
    
    return DEFINED_SLOTS.filter(s => startMin < s.eMin && endMin > s.sMin);
  }

  function _calculateSingleSlotWage(p) {
    let checkEndTime = p.slotEndTime;
    if (p.locName.includes("北葛西") && p.slotName === "夜間") checkEndTime = "20:00";

    const timeStr = `${p.slotStartTime}-${checkEndTime}`;
    const fullTimeStr = p.shiftFullTimeStr;

    let searchDayType = p.dayType === "平日" ? "平日" : "土日祝";
    
    if (p.dayType === "土曜" && p.contractText) {
      if (p.contractText.includes("祝日") && p.contractText.includes("平日") && !p.contractText.includes("土曜")) {
        searchDayType = "平日";
      }
    }

    const searchKeys = [
      `${p.medId}_${p.clinicId}_${searchDayType}_${timeStr}`,
      `${p.medId}_NAME:${p.locName}_${searchDayType}_${timeStr}`,
      `${p.medId}_${p.clinicId}_ALL_${timeStr}`,
      `${p.medId}_NAME:${p.locName}_ALL_${timeStr}`,
      `${p.medId}_${p.clinicId}_${searchDayType}_${fullTimeStr}`,
      `${p.medId}_NAME:${p.locName}_${searchDayType}_${fullTimeStr}`,
      `${p.medId}_${p.clinicId}_ALL_${fullTimeStr}`,
      `${p.medId}_NAME:${p.locName}_ALL_${fullTimeStr}`,
      `${p.medId}_${p.clinicId}_ALL_全時間`
    ];

    for (let key of searchKeys) {
      if (p.specialWageDB[key]) {
        const wages = p.specialWageDB[key];
        let finalWage = wages[0];
        
        if (wages.length > 1) {
          if (searchDayType === "平日" && p.dayType === "土曜") finalWage = Math.min(...wages);
          else if (searchDayType === "土日祝" && p.dayType === "土曜") finalWage = Math.min(...wages);
          else if (searchDayType === "土日祝" && p.dayType === "日祝") finalWage = Math.max(...wages);
        }
        return { wage: finalWage, baseWage: finalWage, type: "特別時給" };
      }
    }

    const year = p.workDate.getFullYear();
    const month = p.workDate.getMonth() + 1;
    let termKey = '2026下期'; 
    
    if ((year === 2025 && month >= 4) || (year === 2026 && month <= 3)) {
      termKey = '2025';
    } else if (year === 2026 && month >= 4 && month <= 9) {
      termKey = '2026上期';
    }

    const baseDB = p.baseWageDB[termKey];
    if (!baseDB) throw new Error(`【重大エラー: 時給表欠落】「${termKey}」の時給マスターシートが存在しません。`);

    const wageRecord = baseDB[`${p.locName}_${p.deptReq}`];
    if (!wageRecord) {
      throw new Error(`【重大エラー: マスタ未定義】「${termKey}」の時給表に「${p.locName}」の「${p.deptReq}」が存在しません。`);
    }

    let wageCurrent = parseInt(wageRecord.rates[p.dayType][p.slotName], 10);
    if (isNaN(wageCurrent)) {
      throw new Error(`【重大エラー: 時給空欄】「${p.locName}」の「${p.dayType}・${p.slotName}」の時給が空欄です。`);
    }

    if (p.contractText && (p.contractText.includes("2025年度") || p.contractText.includes("2025年"))) {
      let wage2025 = 0;
      const baseDB_2025 = p.baseWageDB['2025'];
      
      let targetLocFor2025 = p.contractText.includes("流山") ? "流山おおたかの森" : p.locName;
      if (baseDB_2025 && baseDB_2025[`${targetLocFor2025}_${p.deptReq}`]) {
        wage2025 = parseInt(baseDB_2025[`${targetLocFor2025}_${p.deptReq}`].rates[p.dayType][p.slotName], 10) || 0;
      }
      
      const finalWage = Math.max(wageCurrent, wage2025);
      return { wage: finalWage, baseWage: wageCurrent, type: "基本時給" };
    }

    return { wage: wageCurrent, baseWage: wageCurrent, type: "基本時給" };
  }

  return {
    calculate: calculate
  };
})();