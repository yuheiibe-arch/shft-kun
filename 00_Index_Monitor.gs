/**
 * ==========================================
 * 00_Index_Monitor.gs
 * 進行中モニターと目次シートの生成
 * ★ 3ファイルダイレクトリンク対応 ＆ 超軽量化パッチ
 * ★ 単独実行時エラー回避（セーフティネット）搭載版
 * ★ バッチ完了時自動連動（RUN_INDEX_UPDATE）統合版
 * ==========================================
 */

function createProgressMonitor(ss, totalCount) {
  // シートを削除・追加するとスプレッドシートが数分間フリーズするため、
  // 「すでにあるシートの中身だけを消して使い回す」
  let sheet = ss.getSheetByName("🗂️進行中モニター");
  
  if (!sheet) {
    sheet = ss.insertSheet("🗂️進行中モニター", 0);
  } else {
    sheet.clear(); // 中身だけ一瞬で消す（フリーズしない）
    ss.setActiveSheet(sheet);
    ss.moveActiveSheet(1); // 一番左に移動させる
  }
  
  sheet.getRange("A1:B1").mergeAcross().setValue(`⏳ 処理の準備中... (0 / ${totalCount} 拠点)\nしばらくお待ちください...`)
       .setBackground("#202124").setFontColor("#ffffff").setFontWeight("bold").setFontSize(18)
       .setVerticalAlignment("middle").setHorizontalAlignment("center");
  
  // チェックボックス設置（A2セル）
  sheet.getRange("A2").insertCheckboxes().setValue(false).setHorizontalAlignment("center");
  sheet.getRange("B2").setValue("← 🛑 処理を緊急停止する場合は、ここにチェックを入れてください")
       .setFontColor("#ea4335").setFontWeight("bold").setFontSize(14).setVerticalAlignment("middle");

  sheet.setRowHeight(1, 200);
  sheet.setRowHeight(2, 40);
  sheet.setColumnWidth(1, 60);  
  sheet.setColumnWidth(2, 740); 
  
  // 削除前に「現在シートにある最大行列」をチェックしてから安全に削る
  let maxCols = sheet.getMaxColumns();
  let maxRows = sheet.getMaxRows();
  if (maxCols > 2) sheet.deleteColumns(3, maxCols - 2);
  if (maxRows > 2) sheet.deleteRows(3, maxRows - 2);
}

function updateProgressMonitor(ss, completed, total, remainMin, currentLoc) {
  try {
    let sheet = ss.getSheetByName("🗂️進行中モニター");
    if (sheet) {
      let msg = `⏳ 処理中: ${completed} / ${total} 拠点完了\n(残り目安: 約 ${remainMin} 分)\n\n現在 [ ${currentLoc} ] を展開中...`;
      sheet.getRange("A1:B1").setValue(msg);
    }
  } catch(e) {}
}


// ==========================================
// ▼▼ 目次自動生成・リンク更新機能 ▼▼
// ==========================================

function RUN_INDEX_UPDATE() {
  console.log("🚀 目次更新処理をスタートします...");
  
  // シフトくん本体のID（絶対に失敗しないように直接指定）
  const MASTER_ID = '10yPdoOOgqSSGKwoiPAXi83YM9vb_Em8r6Ex-bLfg28M';
  const ss = SpreadsheetApp.openById(MASTER_ID);
  
  generateAreaIndexSheets(ss);
  
  console.log("✅ 目次の更新が完了しました！スプレッドシートをご確認ください。");
}

let cachedAreaMap = null;

function getAreaMapping() {
  if (cachedAreaMap) return cachedAreaMap;

  let map = {};
  try {
    const ss = SpreadsheetApp.openByUrl("https://docs.google.com/spreadsheets/d/14RbsDcv0nXfEwweki8-9cK3lQUg1XUuhozLNF9u2qAs/edit");
    const data = ss.getSheetByName("拠点名").getDataRange().getValues();
    const headers = data[0];
    const grpIdx = headers.indexOf("拠点グループ");
    const areaIdx = headers.indexOf("エリア");
    
    for (let i = 1; i < data.length; i++) {
      let loc = String(data[i][0]).trim();
      if (!loc) continue;
      let grp = grpIdx !== -1 ? String(data[i][grpIdx]).trim() : "";
      let area = areaIdx !== -1 ? String(data[i][areaIdx]).trim() : "";
      
      let finalArea = "その他";
      if (grp.includes("関東第一") || grp.includes("関東第二")) finalArea = "東京";
      else if (grp.includes("神奈川") || area.includes("神奈川")) finalArea = "神奈川";
      else if (grp.includes("千葉") || area.includes("千葉")) finalArea = "千葉";
      else if (grp.includes("埼玉") || area.includes("埼玉")) finalArea = "埼玉";
      else if (grp.includes("茨城") || area.includes("茨城")) finalArea = "茨城";
      else if (area.includes("関西") || grp.includes("関西")) finalArea = "関西";
      
      map[loc] = finalArea;
    }
  } catch (e) {}
  
  cachedAreaMap = map;
  return map;
}

function generateAreaIndexSheets(ss) {
  if (!ss) ss = SpreadsheetApp.openById('10yPdoOOgqSSGKwoiPAXi83YM9vb_Em8r6Ex-bLfg28M');

  try {
    let monitor = ss.getSheetByName("🗂️進行中モニター");
    if (monitor) ss.deleteSheet(monitor);
  } catch(e) {}

  const TARGET_IDS = [
    "1rScroDlMNiRxThbaxGEuvhyCH2b9RoM6BadNSCAWsvI", // 東京・埼玉
    "19Q-xVsMX0thz_rvmdo_GERJvhjFLQgw2pUkWaTYxQhE", // 関東
    "1fIFvTck_g9-Hp8MSpY7hIWwE5L2buJjBHIY4GBEf_MY"  // 関西
  ];

  const areaMap = getAreaMapping();
  let shiftSheets = [];
  
  TARGET_IDS.forEach(id => {
    try {
      let extSs = SpreadsheetApp.openById(id);
      extSs.getSheets().forEach(s => {
        let name = s.getName();
        if (name.includes("🗂️") || name === "初期設定" || name === "テンプレート" || name === "目次") return;
        
        let locMatch = name.match(/^\d{4}(.*)$/);
        let locName = locMatch ? locMatch[1] : name;
        let cleanLoc = locName.replace(/（.*?）/g, '').trim();
        
        let area = areaMap[cleanLoc] || "その他";
        let order = 99;
        let mainGroup = "その他";
        
        if (area === "東京") { order = 1; mainGroup = "関東"; }
        else if (area === "神奈川") { order = 2; mainGroup = "関東"; }
        else if (area === "千葉") { order = 3; mainGroup = "関東"; }
        else if (area === "埼玉") { order = 4; mainGroup = "関東"; }
        else if (area === "茨城") { order = 5; mainGroup = "関東"; } 
        else if (area === "関西") { order = 6; mainGroup = "関西"; }
        
        shiftSheets.push({
          sheetName: name,
          url: `https://docs.google.com/spreadsheets/d/${id}/edit#gid=${s.getSheetId()}`,
          area: area,
          mainGroup: mainGroup,
          order: order
        });
      });
    } catch(e) {
      console.warn("外部ファイルの読み込みに失敗しました: " + id);
    }
  });

  shiftSheets.sort((a, b) => {
    if (a.order !== b.order) return a.order - b.order;
    return a.sheetName.localeCompare(b.sheetName, 'ja');
  });

  let groups = { "関東": [], "関西": [], "その他": [] };
  shiftSheets.forEach(item => {
    if (groups[item.mainGroup]) groups[item.mainGroup].push(item);
  });

  createIndexSheetForGroup(ss, "🗂️その他目次", groups["その他"]);
  createIndexSheetForGroup(ss, "🗂️関西目次", groups["関西"]);
  createIndexSheetForGroup(ss, "🗂️関東目次", groups["関東"]);
}

function createIndexSheetForGroup(ss, sheetName, items) {
  if (items.length === 0) {
    let old = ss.getSheetByName(sheetName);
    if (old) ss.deleteSheet(old);
    return;
  }
  
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName, ss.getSheets().length);
  } else {
    sheet.clear();
  }
  
  let data = [[`✨ ${sheetName.replace("🗂️", "")}`, "", ""], ["エリア", "拠点・シート名", "リンク"]];
  items.forEach(item => data.push([item.area, item.sheetName, `=HYPERLINK("${item.url}", "開く")`]));
  
  sheet.getRange(1, 1, data.length, 3).setValues(data);
  sheet.getRange(1, 1, 1, 3).setBackground("#34a853").setFontColor("white").setFontWeight("bold").setFontSize(14).mergeAcross();
  sheet.getRange(2, 1, 1, 3).setBackground("#4a86e8").setFontColor("white").setFontWeight("bold");
  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 250);
  sheet.setColumnWidth(3, 80);
  
  SpreadsheetApp.flush(); 
}