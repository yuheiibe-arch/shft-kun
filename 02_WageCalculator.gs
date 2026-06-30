/**
 * ====================================================================
 * 02_WageCalculator.gs (新アダプター版 - 全面差し替え)
 * 新型時給エンジン(NewWageEngine)と募集くんシステムを繋ぐブリッジ
 * ★曜日ズレ補正（定期募集の代表日バグ防止）対応版
 * ====================================================================
 */

// グローバルな医師情報キャッシュ（08A_DoctorMasterでセットされる想定）
var GLOBAL_DOC_MASTER = GLOBAL_DOC_MASTER || {};

/**
 * 医師名から「医籍番号」を取得するヘルパー
 */
function _resolveMedId(docName) {
  if (!docName || docName === "募集" || docName === "休") return "BOSHU";
  const cleanName = String(docName).replace(/先生$/, "").replace(/[\s ]+/g, "").trim();
  
  if (GLOBAL_DOC_MASTER[cleanName] && GLOBAL_DOC_MASTER[cleanName].medId) {
    return GLOBAL_DOC_MASTER[cleanName].medId;
  }
  // マスタに医籍番号がない場合は、フェイルセーフとして名前をそのままエンジンに渡す
  return cleanName; 
}

/**
 * 新エンジンを安全に呼び出すコア・ラッパー（究極ロジック連動版）
 */
function _callEngineSafe(dateStr, locName, category, sH, eH, isHol, docName = null) {
  try {
    const workDate = new Date(dateStr);
    const medId = _resolveMedId(docName);
    const rawStart = `${('0' + sH).slice(-2)}:00`;
    const rawEnd = `${('0' + eH).slice(-2)}:00`;
    
    // ★ 究極ロジック対応：フルタイム文字列を生成
    const fullTimeStr = `${rawStart}-${rawEnd}`;

    // ★ 究極ロジック対応：GLOBAL_DOC_MASTERから「内訳テキスト(specialDetail)」を引っ張ってくる
    let contractText = "";
    if (docName && docName !== "募集" && docName !== "休") {
      const cleanName = String(docName).replace(/先生$/, "").replace(/[\s ]+/g, "").trim();
      if (GLOBAL_DOC_MASTER[cleanName] && GLOBAL_DOC_MASTER[cleanName].specialDetail) {
        contractText = GLOBAL_DOC_MASTER[cleanName].specialDetail;
      }
    }
    
    // ★ 修正：第8引数(fullTimeStr)と第9引数(contractText)をエンジンに渡す
    return NewWageEngine.calculate(medId, locName, category, workDate, rawStart, rawEnd, isHol, fullTimeStr, contractText);
  } catch (e) {
    console.log(`[時給計算スキップ] ${dateStr} ${locName} ${docName || '募集'}: ${e.message}`);
    return []; // エラー時は空配列を返す
  }
}

// ====================================================================
// 以下、既存の08系システムが呼び出してくる関数に対する受け口
// （※フォーマットの互換性を完全に維持します）
// ====================================================================

function _getWageWrapper(dStr, loc, cat, dow, startH, endH) {
  const isHol = (typeof _debug_isTrueHoliday === "function") ? _debug_isTrueHoliday(dStr) : false;
  
  // ★ バグ修正：代表日（10/1など）が渡された場合、指定曜日(dow)に合わせて日付を未来へシフトする
  let targetDateStr = dStr;
  if (dow && !isHol) {
    const dowMap = {"日":0, "月":1, "火":2, "水":3, "木":4, "金":5, "土":6};
    let matchDow = String(dow).replace(/曜[日]?/, ""); // 「土曜日」などで来ても「土」にする
    if (dowMap[matchDow] !== undefined) {
       let workDate = new Date(dStr);
       let currentDow = workDate.getDay();
       let diff = dowMap[matchDow] - currentDow;
       if (diff < 0) diff += 7; // 期を跨いで過去に戻らないよう未来へシフト
       workDate.setDate(workDate.getDate() + diff);
       targetDateStr = Utilities.formatDate(workDate, "JST", "yyyy/MM/dd");
    }
  }

  const slots = _callEngineSafe(targetDateStr, loc, cat, startH, endH, isHol, null);
  
  if (slots.length === 0) return "";
  // 既存システムとの互換性：1コマだけなら生の数字、複数コマなら時間と合体させたテキストを返す
  if (slots.length === 1) {
    return slots[0].wage ? slots[0].wage.toString() : "";
  } else {
    return slots.map(s => `${s.startTime}-${s.endTime}：￥${Number(s.wage).toLocaleString()}`).join('\n');
  }
}

function _getHolidayWageWrapper(dStr, loc, cat, startH, endH) {
  // 祝日強制フラグ(true)を立てて計算
  const slots = _callEngineSafe(dStr, loc, cat, startH, endH, true, null);
  if (slots.length === 0) return "";
  
  if (slots.length === 1) {
    return slots[0].wage ? slots[0].wage.toString() : "";
  } else {
    return slots.map(s => `${s.startTime}-${s.endTime}：￥${Number(s.wage).toLocaleString()}`).join('\n');
  }
}

function _getFinalContractWage(dStr, loc, cat, c) {
  const isHol = (typeof _debug_isTrueHoliday === "function") ? _debug_isTrueHoliday(dStr) : false;
  const slots = _callEngineSafe(dStr, loc, cat, c.sH, c.eH, isHol, c.docName);
  
  if (slots.length === 0) return "";
  // 既存システムとの互換性：確定シフト側はスラッシュ(/)区切りのフォーマットを返す
  return slots.map(s => `${s.startTime}-${s.endTime}/¥${Number(s.wage).toLocaleString()}`).join('\n');
}

function _getDailyCost(dStr, loc, cat, dow, startH, endH) {
  const isHol = (typeof _debug_isTrueHoliday === "function") ? _debug_isTrueHoliday(dStr) : false;

  // ★ バグ修正：コスト計算側も同様に曜日シフトを適用する
  let targetDateStr = dStr;
  if (dow && !isHol) {
    const dowMap = {"日":0, "月":1, "火":2, "水":3, "木":4, "金":5, "土":6};
    let matchDow = String(dow).replace(/曜[日]?/, "");
    if (dowMap[matchDow] !== undefined) {
       let workDate = new Date(dStr);
       let currentDow = workDate.getDay();
       let diff = dowMap[matchDow] - currentDow;
       if (diff < 0) diff += 7; 
       workDate.setDate(workDate.getDate() + diff);
       targetDateStr = Utilities.formatDate(workDate, "JST", "yyyy/MM/dd");
    }
  }

  const slots = _callEngineSafe(targetDateStr, loc, cat, startH, endH, isHol, null);
  
  if (slots.length === 0) return 0;
  let totalCost = 0;
  slots.forEach(s => {
    const hStart = parseInt(s.startTime.split(':')[0], 10);
    const hEnd = parseInt(s.endTime.split(':')[0], 10);
    totalCost += (hEnd - hStart) * s.wage;
  });
  return totalCost;
}

// ※07に存在していた古い関数のダミー（念のためシステムクラッシュを防ぐ安全装置）
function _debug_getCalculatedWage() { return 0; }
function _debug_getMultiZoneWageString(dateStr, dow, sH, eH, locName, category, docName = null) {
  return _getWageWrapper(dateStr, locName, category, dow, sH, eH);
}