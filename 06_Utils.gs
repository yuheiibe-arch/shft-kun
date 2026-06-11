/**
 * ==========================================
 * 06_Utils.gs
 * 共通定数、日付操作、辞書引きなどのユーティリティ
 * ==========================================
 */

// --- システム全体で使う定数 ---
const CONFIG = {
  // 勤怠データの大元シート
  EXTERNAL_URL: "https://docs.google.com/spreadsheets/d/1aEjphEv_63SeWQmwiOy9sx7IrMfawU01sHbKd_Ki4iA/edit",
  // 表記ブレを吸収する辞書マスタシート (開院日もここから取得)
  MAP_URL: "https://docs.google.com/spreadsheets/d/14RbsDcv0nXfEwweki8-9cK3lQUg1XUuhozLNF9u2qAs/edit", 
  // 時給マスタシート
  WAGE_URL: "https://docs.google.com/spreadsheets/d/1eqejNaKWSuHVnRwxaGT-RHgOnlsHkcXYQ5J32B8T_XM/edit",
  
  SETTING_SHEET_NAME: "初期設定",
  MAP_SHEET_NAME: "拠点名",
  TEMPLATE_SHEET_NAME: "テンプレート"
};

/**
 * 外部シートのURLからスプレッドシートオブジェクトを取得
 */
function getExternalSpreadsheet() {
  try {
    return SpreadsheetApp.openByUrl(CONFIG.EXTERNAL_URL);
  } catch (e) {
    throw new Error("外部シートの読み込みに失敗しました。URLとアクセス権限を確認してください。");
  }
}

/**
 * 日付オブジェクトを yyyy/MM/dd 形式に変換
 */
function formatDateString(date) {
  if (!date || isNaN(date.getTime())) return "";
  return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy/MM/dd");
}

/**
 * 指定した年月（YYYY/MM）の初日と末日を取得
 * 戻り値例: { start: "2025/06/01", end: "2025/06/30", daysInMonth: 30 }
 */
function getMonthEdges(yearMonthStr) {
  const [year, month] = yearMonthStr.split('/').map(Number);
  const startDate = new Date(year, month - 1, 1);
  const endDate = new Date(year, month, 0); // 翌月の0日目 ＝ 今月の末日
  
  return {
    start: formatDateString(startDate),
    end: formatDateString(endDate),
    daysInMonth: endDate.getDate()
  };
}

/**
 * 対象年月から時給マスタの適切なシート名（上期/下期）を動的判定して返す
 * 下期が存在しない場合は上期にフォールバックする
 */
function getValidWageSheetName(yearMonthStr) {
  const [yearStr, monthStr] = yearMonthStr.split('/');
  const month = parseInt(monthStr, 10);
  let year = parseInt(yearStr, 10);
  
  // 対象月が1〜3月の場合は「前年の年度」として扱う（例：2026/01 -> 2025年度）
  const fiscalYear = (month >= 4 && month <= 12) ? year : year - 1;
  
  // 上期（4〜9月）か下期（10〜3月）かを判定
  const isSecondHalf = (month >= 10 || month <= 3);
  const targetSuffix = isSecondHalf ? "（下期）" : "（上期）";
  let expectedSheetName = fiscalYear + "年度" + targetSuffix;
  
  const wageSs = SpreadsheetApp.openByUrl(CONFIG.WAGE_URL);
  let sheet = wageSs.getSheetByName(expectedSheetName);
  
  // 下期の対象月だが、まだ下期のシートが作られていない場合は「上期」にフォールバック
  if (!sheet && isSecondHalf) {
    expectedSheetName = fiscalYear + "年度（上期）";
    sheet = wageSs.getSheetByName(expectedSheetName);
  }
  
  // それでも見つからない場合はエラー
  if (!sheet) {
    throw new Error(`時給マスタに「${expectedSheetName}」シートが見つかりません。`);
  }
  
  return expectedSheetName;
}

/**
 * 外部の正規表現シートから表記ブレ辞書を動的に取得する
 */
function getLocationDictionary() {
  const ss = SpreadsheetApp.openByUrl(CONFIG.MAP_URL);
  const sheet = ss.getSheetByName(CONFIG.MAP_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  let map = {};
  
  data.forEach(row => {
    let canonical = row[0] ? row[0].toString().trim() : "";
    if (!canonical) return;
    
    // B列〜E列の表記ブレを正規名に紐付ける
    for (let j = 1; j <= 4; j++) {
      let variant = row[j] ? row[j].toString() : "";
      variant = variant.replace(/[\s　]+/g, ""); // 全角半角スペースを除去
      if (variant) map[variant] = canonical;
    }
  });
  
  return map;
}

/**
 * ★新設：外部の正規表現シートから「開院日」を取得する
 * 空欄の場合は「未定(null)」として扱う
 */
function getLocationOpenDates() {
  const ss = SpreadsheetApp.openByUrl(CONFIG.MAP_URL);
  const sheet = ss.getSheetByName(CONFIG.MAP_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  let openDates = {};
  
  if (data.length < 2) return openDates;
  
  const headers = data[0];
  const dateIdx = headers.indexOf("開院日");
  
  if (dateIdx === -1) {
    Logger.log("[警告] 拠点名シートに「開院日」というヘッダーが見つかりません。");
    // ヘッダーがない場合は一旦すべてを未定(null)として返す
    for (let i = 1; i < data.length; i++) {
      const locName = data[i][0] ? data[i][0].toString().trim() : "";
      if (locName) openDates[locName] = null;
    }
    return openDates;
  }
  
  for (let i = 1; i < data.length; i++) {
    const locName = data[i][0] ? data[i][0].toString().trim() : "";
    if (!locName) continue;
    
    let dateVal = data[i][dateIdx];
    
    // 空欄の場合は「未定」
    if (dateVal === "" || dateVal == null) {
      openDates[locName] = null;
    } else if (dateVal instanceof Date) {
      openDates[locName] = dateVal;
    } else {
      let parsed = new Date(dateVal);
      openDates[locName] = isNaN(parsed.getTime()) ? null : parsed;
    }
  }
  
  return openDates;
}

/**
 * テキストから正式な拠点名を抽出・変換する
 */
function normalizeLocationName(rawText, dict) {
  const map = dict || getLocationDictionary();
  let cleanName = rawText.replace(/[\(（]?(内科|小児科)[\)）]?/g, "").replace(/\/.*/, "");
  cleanName = cleanName.replace(/[\s　]+/g, "");
  return map[cleanName] || cleanName;
}

/**
 * ====================================================================
 * 【基幹エンジン】時給データ取得・検索モジュール
 * ====================================================================
 */

const WAGE_CONFIG = {
  MASTER_URL: 'https://docs.google.com/spreadsheets/d/1eqejNaKWSuHVnRwxaGT-RHgOnlsHkcXYQ5J32B8T_XM/edit',
  SHEET_2025: '2025年度（年間）',
  SHEET_2026: '2026年度（上期）'
};

let _wageMasterMap2025 = null;
let _wageData2026 = null;
let _locationMasterMap = null;

function initializeWageData() {
  if (_wageMasterMap2025 && _wageData2026 && _locationMasterMap) return;

  const masterSs = SpreadsheetApp.openByUrl(WAGE_CONFIG.MASTER_URL);
  
  _locationMasterMap = getLocationDictionary(); 

  // 2025年度マップの構築
  const sheet2025 = masterSs.getSheetByName(WAGE_CONFIG.SHEET_2025);
  _wageMasterMap2025 = {};
  if (sheet2025) {
    const values = sheet2025.getRange('B2:O' + Math.max(2, sheet2025.getLastRow())).getValues();
    values.forEach(row => {
      const clinic = (row[0] || '').toString().trim();
      const dept = (row[1] || '').toString().replace(/\s+/g, '');
      if (clinic) {
        _wageMasterMap2025[`${clinic}||${dept}`] = {
          wd_am: row[3], wd_pm: row[4], wd_nt: row[5],
          hol_am: row[11], hol_pm: row[12], hol_nt: row[13]
        };
      }
    });
  }

  // 2026年度データの取得
  const sheet2026 = masterSs.getSheetByName(WAGE_CONFIG.SHEET_2026);
  _wageData2026 = sheet2026 ? sheet2026.getDataRange().getValues() : [];
}

function _normalizeForSearch(rawName) {
  if (!rawName) return "";
  let name = rawName.toString();
  if (name.normalize) name = name.normalize('NFKC');
  return name
    .replace(/[【】\[\]]/g, "") 
    .replace(/[\(（]?(内科|小児科|皮膚科|整形外科)[\)）]?/g, "") 
    .replace(/(病院|クリニック|診療所|モール)$/g, "") 
    .replace(/\s+/g, "") 
    .replace(/ヶ/g, "ケ") 
    .trim();
}

function _findRate2025(formalName, deptName) {
  const exact = _wageMasterMap2025[`${formalName}||${deptName}`] || 
                _wageMasterMap2025[`${formalName}||共通`] || 
                _wageMasterMap2025[`${formalName}||`];
  if (exact) return exact;

  for (const key in _wageMasterMap2025) {
    const [c, d] = key.split("||");
    if (c.includes(formalName) || formalName.includes(c)) {
      if (deptName === "小児科" && (d.includes("小児") || c.includes("小児"))) return _wageMasterMap2025[key];
      if (deptName === "内科" && (!d.includes("小児") && !c.includes("小児"))) return _wageMasterMap2025[key];
    }
  }
  return null;
}

function _findRate2026(formalName, deptName) {
  let found = _wageData2026.find(row => {
    const c = _normalizeForSearch(row[1]);
    const d = (row[2] || "").toString();
    if (!c) return false;
    if (!c.includes(formalName) && !formalName.includes(c)) return false;

    if (deptName === "小児科") return c.includes("小児") || d.includes("小児");
    if (deptName === "内科") return (c.includes("内科") || d.includes("内科")) || (!c.includes("小児") && !d.includes("小児"));
    return false;
  });

  if (!found) {
    found = _wageData2026.find(row => {
      const c = _normalizeForSearch(row[1]);
      return c && c.includes(formalName);
    });
  }

  if (found) {
    return {
      wd_am: found[3], wd_pm: found[4], wd_nt: found[5],
      hol_am: found[6], hol_pm: found[7], hol_nt: found[8]
    };
  }
  return null;
}

function getClinicWages(rawClinicName) {
  if (!_wageMasterMap2025) initializeWageData();

  const cleanLoc = _normalizeForSearch(rawClinicName);
  const formalName = _locationMasterMap[cleanLoc] || cleanLoc;

  const targetDepts = (formalName.includes("北葛西") || formalName.includes("亀有")) 
                      ? ["内科", "小児科"] 
                      : ["小児科"];

  const results = [];

  targetDepts.forEach(dept => {
    const rate2025 = _findRate2025(formalName, dept);
    const rate2026 = _findRate2026(formalName, dept);

    results.push({
      clinicName: formalName,
      department: dept,
      rates: {
        y2025: rate2025 || null,
        y2026: rate2026 || null
      }
    });
  });

  return results;
}