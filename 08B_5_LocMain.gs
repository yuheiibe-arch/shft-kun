/**
 * ====================================================================
 * 08B_5_LocMain.gs
 * 拠点処理の司令塔（各分割ファイルを呼び出す）
 * ====================================================================
 */

function _processLocation(ctx, locName) {
  const cleanLocName = locName.split('_')[0].replace(/[（\(\)）]/g, "").replace(/内科|小児科/g, "").trim();
  const category = locName.includes("内科") ? "内科" : "小児科";
  let displayLoc = cleanLocName;
  if (displayLoc === "亀有" || displayLoc === "北葛西") displayLoc = `${displayLoc}（${category}）`;
  
  const area = _getAreaHelper(displayLoc, ctx._areaMap);
  const openDate = _getOpenDateHelperMaster(cleanLocName, ctx._openDateMap);
  if (openDate && openDate > ctx.endDate) return;
  
  let actualStartDate = (openDate && openDate > ctx.startDate) ? openDate : ctx.startDate;
  const locCalendar = ctx.calendarCache.filter(c => c.getTime >= actualStartDate.getTime() && c.getTime <= ctx.endDate.getTime());

  let activeContracts = ctx.contractsByLoc[locName] || [];

  // 🚀 爆速化・2診判定のための辞書生成（既存のルールは一切変えません）
  let locDicts = _buildDailyDictsForLoc(ctx, cleanLocName, category, actualStartDate);
  let dailyBusyMap = _buildDailyBusyMap(activeContracts, locCalendar, locDicts);

  // コンテキストをひとまとめにする
  let locCtx = {
    locName, cleanLocName, category, displayLoc, area, actualStartDate, locCalendar, activeContracts, locDicts, dailyBusyMap
  };

  // 分割した各処理を呼び出し（ルールの改変は一切ありません）
  _processLocBoshu(ctx, locCtx);
  _processLocConfirm(ctx, locCtx);
  _processLocCancel(ctx, locCtx);
}