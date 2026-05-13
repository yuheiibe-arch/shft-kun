/**
 * ====================================================================
 * 10_Boshukun_Writer.gs
 * 出来上がった各種リストをスプレッドシートの該当シートへ書き出す処理
 * ====================================================================
 */

function _writeBoshukunSheets(ctx) {
  const headersDef = {
    // ★「担当者」を削除
    regular: ["エリア", "シフトタイトル", "期間", "開始時間", "終了時間", "繰り返し曜日", "時給", "祝日時給", "該当日", "募集時間", "コスト", "先行・振替", "対応済"],
    single: ["エリア", "拠点名", "理由", "期間", "開始時間", "終了時間", "繰り返し曜日", "該当日", "時給", "募集時間", "コスト", "対応済", "注意", "担当者"],
    confirm: ["エリア", "拠点名", "種別", "医師名", "シフトタイトル", "設定期間", "設定時間", "設定曜日", "時給", "契約内容", "注意箇所", "対応済", "担当者"],
    cancel: ["医師名", "該当日", "理由", "対象拠点", "対象勤務時間", "募集シフト作成指示", "対応済", "対応者", "GASチェック"]
  };

  const _k = (val) => {
    if (val === null || val === undefined) return "";
    if (val instanceof Date) {
      if (val.getFullYear() <= 1970) return Utilities.formatDate(val, "JST", "HH:mm");
      return Utilities.formatDate(val, "JST", "yyyy/MM/dd");
    }
    let str = String(val).trim();

    // ① 全角・半角スペースをすべて削除（"吉澤 和希" → "吉澤和希"）
    str = str.replace(/[\s　]+/g, "");

    // ② 名前の末尾の「先生」を削除（"阿部優作先生" → "阿部優作"）
    str = str.replace(/先生$/g, "");

    // ③ 拠点名のカッコや(内科)(小児科)表記を削除して統一（"亀有（内科）" → "亀有"）
    str = str.replace(/[【】\(（]?(内科|小児科)[\)）]?/g, "");

    // ④ 時間の揺れ（〜、～、ー、-）をすべて半角ハイフンに統一
    str = str.replace(/[－ー〜～]/g, "-");

    // ⑤ 時刻のゼロ埋め（"9:00" → "09:00"）
    str = str.replace(/\b(\d{1}):(\d{2})\b/g, "0$1:$2");

    // ⑥ 年月日をスラッシュ形式に統一（"2026年4月1日" → "2026/04/01"）
    str = str.replace(/(\d{4})年(\d{1,2})月(\d{1,2})日/g, (match, y, m, d) => {
      return `${y}/${m.padStart(2, '0')}/${d.padStart(2, '0')}`;
    });

    return str;
  };

  writeWithProtectionObj("定期募集", ctx.masterRegularList, headersDef.regular, "対応済", 
    row => ctx.targetDisplayLocs.some(loc => String(row["シフトタイトル"]).includes(loc)), 
    row => `${_k(row["シフトタイトル"])}|${_k(row["期間"])}|${_k(row["開始時間"])}|${_k(row["終了時間"])}|${_k(row["繰り返し曜日"])}`
  );
  
  writeWithProtectionObj("単独募集", ctx.singleList, headersDef.single, "対応済", 
    row => ctx.targetDisplayLocs.some(loc => String(row["拠点名"]).includes(loc)), 
    row => `${_k(row["拠点名"])}|${_k(row["理由"])}|${_k(row["期間"])}|${_k(row["開始時間"])}|${_k(row["終了時間"])}|${_k(row["該当日"])}`
  );
  
  writeWithProtectionObj("確定シフト作成", ctx.confirmList, headersDef.confirm, "対応済", 
    row => ctx.targetDisplayLocs.some(loc => String(row["拠点名"]).includes(loc)), 
    row => `${_k(row["拠点名"])}|${_k(row["種別"])}|${_k(row["医師名"])}|${_k(row["設定期間"])}|${_k(row["設定時間"])}|${_k(row["設定曜日"])}`
  );
  
  writeWithProtectionObj("欠勤・シフトキャンセル作成", ctx.cancelList, headersDef.cancel, "対応済", 
    row => ctx.targetDisplayLocs.some(loc => String(row["対象拠点"]).includes(loc)), 
    row => `${_k(row["医師名"])}|${_k(row["該当日"])}|${_k(row["理由"])}|${_k(row["対象拠点"])}|${_k(row["対象勤務時間"])}`, ["対応済", "GASチェック"]
  );
}