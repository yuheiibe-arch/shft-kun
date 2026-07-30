/**
 * ====================================================================
 * 10_Boshukun_Writer.gs
 * 出来上がった各種リストをスプレッドシートの該当シートへ書き出す処理
 * ====================================================================
 */

function _writeBoshukunSheets(ctx) {
  const headersDef = {
    regular: ["エリア", "シフトタイトル", "期間", "開始時間", "終了時間", "繰り返し曜日", "時給", "祝日時給", "該当日", "募集時間", "コスト", "先行・振替", "対応済"],
    single: ["エリア", "拠点名", "理由", "期間", "開始時間", "終了時間", "繰り返し曜日", "該当日", "時給", "募集時間", "コスト", "対応済", "注意", "担当者"],
    // ★ 修正箇所：末尾に「作業メモ」を追加
    confirm: ["エリア", "拠点名", "種別", "医師名", "シフトタイトル", "設定期間", "設定時間", "設定曜日", "時給", "契約内容", "注意箇所", "対応済", "担当者", "作業メモ"],
    cancel: ["医師名", "該当日", "理由", "対象拠点", "対象勤務時間", "募集シフト作成指示", "対応済", "対応者", "GASチェック"]
  };

  const _k = (val) => {
    if (val === null || val === undefined) return "";
    if (val instanceof Date) {
      if (val.getFullYear() <= 1970) return Utilities.formatDate(val, "JST", "HH:mm");
      return Utilities.formatDate(val, "JST", "yyyy/MM/dd");
    }
    let str = String(val).trim();

    str = str.replace(/[\s ]+/g, "");
    str = str.replace(/先生$/g, "");
    str = str.replace(/[【】\(（]?(内科|小児科)[\)）]?/g, "");
    str = str.replace(/[－ー〜～]/g, "-");
    str = str.replace(/\b(\d{1}):(\d{2})\b/g, "0$1:$2");
    str = str.replace(/(\d{4})年(\d{1,2})月(\d{1,2})日/g, (match, y, m, d) => {
      return `${y}/${m.padStart(2, '0')}/${d.padStart(2, '0')}`;
    });

    return str;
  };

  // ★ 修正箇所：シート上の値が「院外勤務」であっても、処理対象が「MQC」なら一致とみなす判定ロジック
  const checkTargetMatch = (sheetValue, targetLocs) => {
    const s = String(sheetValue);
    return targetLocs.some(loc => s.includes(loc) || (loc === "MQC" && s.includes("院外勤務")));
  };

  writeWithProtectionObj("定期募集", ctx.masterRegularList, headersDef.regular, "対応済", 
    row => checkTargetMatch(row["シフトタイトル"], ctx.targetDisplayLocs), 
    row => `${_k(row["シフトタイトル"])}|${_k(row["期間"])}|${_k(row["開始時間"])}|${_k(row["終了時間"])}|${_k(row["繰り返し曜日"])}`
  );
  
  writeWithProtectionObj("単独募集", ctx.singleList, headersDef.single, "対応済", 
    row => checkTargetMatch(row["拠点名"], ctx.targetDisplayLocs), 
    row => `${_k(row["拠点名"])}|${_k(row["理由"])}|${_k(row["期間"])}|${_k(row["開始時間"])}|${_k(row["終了時間"])}|${_k(row["該当日"])}`
  );
  
  writeWithProtectionObj("確定シフト作成", ctx.confirmList, headersDef.confirm, "対応済", 
    row => checkTargetMatch(row["拠点名"], ctx.targetDisplayLocs), 
    row => `${_k(row["拠点名"])}|${_k(row["種別"])}|${_k(row["医師名"])}|${_k(row["設定期間"])}|${_k(row["設定時間"])}|${_k(row["設定曜日"])}`
  );
  
  writeWithProtectionObj("欠勤・シフトキャンセル作成", ctx.cancelList, headersDef.cancel, "対応済", 
    row => checkTargetMatch(row["対象拠点"], ctx.targetDisplayLocs), 
    row => `${_k(row["医師名"])}|${_k(row["該当日"])}|${_k(row["理由"])}|${_k(row["対象拠点"])}|${_k(row["対象勤務時間"])}`, ["対応済", "GASチェック"]
  );
}