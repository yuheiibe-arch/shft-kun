/**
 * ==========================================
 * 01_UI_Main.gs (または 00_UI_Menu.gs)
 * スプレッドシートのメニューとUIダイアログ管理
 * ==========================================
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('シフト管理')
    .addItem('シフト表を作成', 'showShiftCreationDialog')
    .addItem('⚡ 差分のみ同期（スマート更新）', 'promptSmartSync')
    .addSeparator()
    .addItem('募集くんを起動', 'showBoshukunDialog')
    .addToUi();
}

function showBoshukunDialog() {
  const html = HtmlService.createTemplateFromFile('Dialog_Boshukun');
  const ui = html.evaluate().setWidth(400).setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(ui, '募集くん 出力ウィザード');
}

function showShiftCreationDialog() {
  const html = HtmlService.createTemplateFromFile('Dialog_ShiftCreate');
  const ui = html.evaluate().setWidth(450).setHeight(550);
  SpreadsheetApp.getUi().showModalDialog(ui, 'シフト表作成ウィザード');
}

function promptSmartSync() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let maxYear = 0;
  
  // シート名から最新年度を自動取得
  ss.getSheets().forEach(s => {
    const match = s.getName().match(/^(\d{4})/);
    if (match) {
      const y = parseInt(match[1], 10);
      if (y > maxYear) maxYear = y;
    }
  });

  // 対象シートがなければ無言で終了（アラートを出さない）
  if (maxYear === 0) return; 

  // ★爆速化：画面上でそのまま処理せず、doGetと同じように「裏側（バックグラウンド）」に任せる！
  let props = PropertiesService.getScriptProperties();
  props.setProperty('PENDING_SYNC_YEAR', String(maxYear));

  // 重複実行を防ぐため古いトリガーを掃除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'processSmartSyncBackground') {
      ScriptApp.deleteTrigger(t);
    }
  });

  // 1秒後に裏側で差分チェックを開始させる
  ScriptApp.newTrigger('processSmartSyncBackground').timeBased().after(1000).create();
  
  // 画面にはメッセージだけ出して、即座に終了する（これでタイムアウト回避）
  ss.toast("⚡ 差分チェックを裏側で開始しました！\n数秒後にモニター画面が自動で立ち上がります。", "スマート更新", 8);
}

// ★追加：裏側でバトンを受け取って動く関数
function processSmartSyncBackground() {
  // 自分のトリガーを掃除
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(t => {
    if (t.getHandlerFunction() === 'processSmartSyncBackground') {
      ScriptApp.deleteTrigger(t);
    }
  });

  let props = PropertiesService.getScriptProperties();
  let targetYear = props.getProperty('PENDING_SYNC_YEAR');
  if (!targetYear) return;

  // ここで本命の差分チェックを実行（制限時間6分をフルに使える）
  checkAndStartSmartSync(targetYear, "通年");
}

function getAvailableYears() {
  const extSs = typeof getExternalSpreadsheet === 'function' ? getExternalSpreadsheet() : SpreadsheetApp.getActiveSpreadsheet();
  const sheets = extSs.getSheets().map(s => s.getName());
  let joukinYears = [], teikiYears = [];
  sheets.forEach(name => {
    const match = name.match(/(\d{4})年度/);
    if (match) {
      if (name.includes('常勤') && !name.includes('定期非常勤')) joukinYears.push(match[1]);
      if (name.includes('定期非常勤')) teikiYears.push(match[1]);
    }
  });
  const pairedYears = joukinYears.filter(year => teikiYears.includes(year));
  return pairedYears.sort((a, b) => b - a);
}

function getLocationListForUI(params) {
  const targetYear = parseInt(params.year, 10);
  const targetTerm = params.term;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  let termEndDate = targetTerm === "上期" ? new Date(targetYear, 8, 30) : new Date(targetYear + 1, 2, 31);
  const openDatesMap = typeof getLocationOpenDates === 'function' ? getLocationOpenDates() : {};
  let locationList = [];
  
  for (const locName in openDatesMap) {
    let openDate = openDatesMap[locName];
    if (!openDate || openDate > termEndDate) continue;
    
    if (locName === "亀有" || locName === "北葛西") {
      [`${locName}（内科）`, `${locName}（小児科）`].forEach(sub => {
        const isCreated = (ss.getSheetByName(`${targetYear}${sub}`) !== null);
        locationList.push({ name: sub, isCreated: isCreated });
      });
    } else {
      const isCreated = (ss.getSheetByName(`${targetYear}${locName}`) !== null);
      locationList.push({ name: locName, isCreated: isCreated });
    }
  }
  return locationList;
}