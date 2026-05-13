/**
 * ==========================================
 * 03_Logic_Tetris.gs
 * シフトのパズル配置（1診目・2診目のテトリスロジック）
 * ==========================================
 */

/**
 * 抽出したシフトデータを1日ごと・時間帯ごとのグリッドに変換する
 * @param {Array} dailyShifts その日の全ドクターのシフト配列
 * @param {string} locationName 拠点名（北葛西の20時終了などの判定用）
 * @return {Object} { line1: [], line2: [], docTypes: {} }
 */
function calculateTetrisAllocation(dailyShifts, locationName) {
  // D列(9:00)〜O列(20:00)の12マス分を初期化（nullで埋める）
  let line1 = new Array(12).fill(null);
  let line2 = new Array(12).fill(null);
  let docTypes = {}; // 医師名から「常勤/定期非常勤」を引けるようにする辞書

  // 休診時間のインデックス (H列=4, I列=5) は最初から固定で「休」扱いにしておく
  line1[4] = line1[5] = line2[4] = line2[5] = "休";

  // 【北葛西】の場合は営業が20時までなので、O列(インデックス11)は「休」にする
  if (locationName.includes("北葛西")) {
    line1[11] = line2[11] = "休";
  }

  // もしその日に誰もシフトが入っていなければ、後続の空き枠埋め処理へ直行
  if (dailyShifts && dailyShifts.length > 0) {
    dailyShifts.forEach(shift => {
      docTypes[shift.doctorName] = shift.type; // 属性を記録
      
      let startHour = 9;
      let endHour = 18;

      // ★修正：開始時間の安全な取得（Dateオブジェクト、文字列、数値すべてに対応）
      if (shift.startTime instanceof Date) {
        startHour = shift.startTime.getHours();
      } else if (shift.startTime) {
        let sStr = String(shift.startTime);
        if (sStr.includes(':')) {
          startHour = parseInt(sStr.split(':')[0], 10);
        } else {
          startHour = parseInt(sStr, 10);
        }
      } else if (shift.startHour !== undefined) {
        startHour = shift.startHour;
      }

      // ★修正：終了時間の安全な取得
      if (shift.endTime instanceof Date) {
        endHour = shift.endTime.getHours();
      } else if (shift.endTime) {
        let eStr = String(shift.endTime);
        if (eStr.includes(':')) {
          endHour = parseInt(eStr.split(':')[0], 10);
        } else {
          endHour = parseInt(eStr, 10);
        }
      } else if (shift.hours !== undefined) {
        endHour = startHour + shift.hours;
      }
      
      // 万が一NaNになった場合のフェイルセーフ
      if (isNaN(startHour)) startHour = 9;
      if (isNaN(endHour)) endHour = startHour;

      // 時間帯をインデックス(0〜11)に変換してマス目を埋める
      for (let h = startHour; h < endHour; h++) {
        let idx = h - 9; // 9時ならインデックス0
        if (idx < 0 || idx > 11 || idx === 4 || idx === 5) continue; // 範囲外や休診時間はスキップ
        
        // テトリスロジック：1診目が空いていれば1診目へ、埋まっていれば2診目へ
        if (line1[idx] === null || line1[idx] === "募集") {
          line1[idx] = shift.doctorName;
        } else if (line2[idx] === null || line2[idx] === "募集") {
          line2[idx] = shift.doctorName;
        }
      }
    });
  }

  // 【空き枠の募集化】
  // line1 (1診目) の null の部分を「募集」に書き換える
  for (let i = 0; i < 12; i++) {
    if (line1[i] === null) {
      line1[i] = "募集";
    }
  }

  return { line1: line1, line2: line2, docTypes: docTypes };
}