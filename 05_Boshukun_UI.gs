/**
 * ====================================================================
 * 05_Boshukun_UI.gs
 * UI用の初期データ取得 ＆ 拠点マスターからの抽出
 * ====================================================================
 */

/**
 * UI起動時に呼ばれる：年度とエリアの一覧を返す
 */
function getBoshukunInitData() {
  const years = getAvailableYears(); 
  const latestYear = years.length > 0 ? years[0] : new Date().getFullYear().toString();
  const defaultTerm = "上期";
  
  const areas = ["東京", "神奈川", "千葉", "埼玉", "関西", "その他"];
  
  // 初期表示（最新年度・上期）の時点で開院している拠点だけを抽出
  const allLocs = _getAvailableLocationsForUI(latestYear, defaultTerm);
  
  return {
    years: years,
    latestYear: latestYear,
    areas: areas,
    allLocs: allLocs
  };
}

/**
 * UI上で年度や期間が変更された時に呼ばれる：プルダウン用の拠点リストを再取得
 */
function updateLocationListForUI(yearStr, term) {
  return _getAvailableLocationsForUI(yearStr, term);
}

/**
 * (内部関数) 指定された年度・期に開院している拠点リストを取得
 */
function _getAvailableLocationsForUI(yearStr, term) {
  const allLocs = [];
  const y = parseInt(yearStr, 10);
  const termEndDate = term === "上期" ? new Date(y, 8, 30) : new Date(y + 1, 2, 31);
  
  try {
    const masterUrl = 'https://docs.google.com/spreadsheets/d/14RbsDcv0nXfEwweki8-9cK3lQUg1XUuhozLNF9u2qAs/edit';
    const ss = SpreadsheetApp.openByUrl(masterUrl);
    const sheet = ss.getSheetByName('拠点名');
    if (sheet) {
      const data = sheet.getDataRange().getValues();
      for (let i = 1; i < data.length; i++) {
        const name = String(data[i][0]).trim();
        const openDateVal = data[i][7]; // H列: 開院日
        
        // MQC と 空欄 を除外
        if (!name || name.toUpperCase() === "MQC") continue;
        
        // 開院日の判定（期末日までに開院していなければ除外）
        if (openDateVal) {
           let openDate = openDateVal instanceof Date ? openDateVal : new Date(openDateVal);
           if (!isNaN(openDate.getTime()) && openDate > termEndDate) {
             continue; 
           }
        }
        allLocs.push(name);
      }
    }
  } catch(e) {
    Logger.log("マスター読み込みエラー: " + e.message);
  }
  return allLocs;
}

/**
 * メイン処理から呼ばれる：選択されたエリアに該当する拠点を抽出
 */
function getTargetLocationsFromMaster(yearStr, term, targetAreas) {
  const masterUrl = 'https://docs.google.com/spreadsheets/d/14RbsDcv0nXfEwweki8-9cK3lQUg1XUuhozLNF9u2qAs/edit';
  let ss;
  try { ss = SpreadsheetApp.openByUrl(masterUrl); } 
  catch(e) { throw new Error("拠点マスターシートへのアクセス権限がありません。"); }
  
  const sheet = ss.getSheetByName('拠点名');
  if (!sheet) throw new Error("マスターに「拠点名」シートが見つかりません。");

  const data = sheet.getDataRange().getValues();
  const locNames = [];
  const y = parseInt(yearStr, 10);
  const termEndDate = term === "上期" ? new Date(y, 8, 30) : new Date(y + 1, 2, 31);

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const name = String(row[0]).trim();            // A列: 正規記載
    const group = String(row[5]).trim();           // F列: 拠点グループ
    const areaCol = String(row[6]).trim();         // G列: エリア
    const openDateVal = row[7];                    // H列: 開院日

    // MQC と 空欄 を除外
    if (!name || name.toUpperCase() === "MQC") continue;
    
    // ★ 拠点グループ（F列）とエリア（G列）からUI用のエリアカテゴリに変換
    let mappedArea = "その他";
    if (group === "関東第一" || group === "関東第二") {
      mappedArea = "東京";
    } else if (group.includes("神奈川")) {
      mappedArea = "神奈川";
    } else if (group.includes("千葉")) {
      mappedArea = "千葉";
    } else if (group.includes("埼玉")) {
      mappedArea = "埼玉";
    } else if (areaCol.includes("関西") || group.includes("大阪") || group.includes("関西")) {
      mappedArea = "関西"; 
    }
    
    // UIで選択されたエリアに含まれていなければスキップ
    if (!targetAreas.includes(mappedArea)) continue;
    
    // 開院日の判定（期末日までに開院していなければ除外）
    if (openDateVal) {
      let openDate = openDateVal instanceof Date ? openDateVal : new Date(openDateVal);
      if (!isNaN(openDate.getTime()) && openDate > termEndDate) {
        continue;
      }
    }
    
    locNames.push(name);
  }
  return locNames;
}