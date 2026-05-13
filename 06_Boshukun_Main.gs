/**
 * ====================================================================
 * 06_Boshukun_Main.gs
 * メインコントローラー（処理の司令塔）
 * ※実際の処理は 07_DataLoader と 08_Generator に分割されています
 * ====================================================================
 */

function executeBoshukunMain(year, term, mode, areas, singleLoc) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast("マスタデータを解析してリストを生成しています...", "募集くん", 5);

  // 全ファイルで使い回す「データ格納用の箱（Context）」を作成
  const ctx = {
    ss: ss,
    year: year,
    term: term,
    mode: mode,
    areas: areas,
    singleLoc: singleLoc,
    dowNames: GLOBAL_DOW_NAMES,
    
    // データローダーで格納される変数群
    targetDisplayLocs: [], locNames: [],
    _areaMap: {}, _openDateMap: {},
    startDate: null, endDate: null, startDStr: "",
    p1Start: null, p1End: null, p1TitleStr: "",
    p2Start: null, p2End: null, p2TitleStr: "",
    calendarCache: [], absences: [], advances: [], substitutes: [], kyukans: [],
    contractsByLoc: {},
    
    // ジェネレーターで格納される出力用リスト
    masterRegularList: [], singleList: [], confirmList: [], cancelList: [],
    tempSingles: [], pushedSingles: new Set()
  };

  try {
    // STEP 1: データの読み込みと初期化 (07_Boshukun_DataLoader.gs)
    _loadBoshukunData(ctx);

    // STEP 2: シフトと募集枠の生成 (08_Boshukun_Generator.gs)
    _generateAllShifts(ctx);

    // STEP 3: スプレッドシートへの書き出し (08_Boshukun_Generator.gs)
    _writeBoshukunSheets(ctx);

    ss.toast("リストの出力が完了しました！", "募集くん", 5);
    return "Success";

  } catch (error) {
    ss.toast(error.message, "募集くんエラー", 10);
    throw error; // エラー時は処理を中断してスロー
  }
}