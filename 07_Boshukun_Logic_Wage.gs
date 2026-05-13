/**
 * ====================================================================
 * 07_Boshukun_Logic_Wage.gs (縮小・残存ユーティリティ版)
 * ※時給計算ロジック自体は新エンジン(02/03)へ完全移行したため削除済。
 * ※他ファイルから呼ばれるテトリス関連等の便利関数のみを、新エンジン対応にして残しています。
 * ====================================================================
 */

function _debug_calcBlockWage(block, isHoliday, dateStr, cleanLocName, category) {
  // ★古い計算機ではなく、02_WageCalculator のアダプターを経由して新エンジンに計算させる
  // （時間単価×時間のコストを算出）
  return _getDailyCost(dateStr, cleanLocName, category, block.dow, block.startHour, block.startHour + block.hours);
}

function _debug_extractBoshuBlocksFromTetris(dateStr, dow, locName, tetrisLine1) {
  const blocks = [];
  let currentBlock = null;
  const getTimeZone = (idx) => (idx < 4) ? "AM" : (idx < 6) ? "REST" : (idx < 9) ? "PM" : "NT";

  const pushBlock = (b) => {
    const startH = b.startIdx + 9;
    const endH = b.endIdx + 10;
    blocks.push({
      loc: locName, date: dateStr, dow: dow,
      start: `${('0'+startH).slice(-2)}:00`, end: `${('0'+endH).slice(-2)}:00`,
      startHour: startH, hours: endH - startH
    });
  };

  for (let i = 0; i <= 12; i++) {
    const cell = (i < 12) ? tetrisLine1[i] : null;
    if (cell === "募集") {
      const tz = getTimeZone(i);
      if (!currentBlock) {
        currentBlock = { startIdx: i, endIdx: i, tz: tz };
      } else if (currentBlock.tz !== tz) {
        pushBlock(currentBlock);
        currentBlock = { startIdx: i, endIdx: i, tz: tz };
      } else {
        currentBlock.endIdx = i;
      }
    } else if (currentBlock) {
      pushBlock(currentBlock);
      currentBlock = null;
    }
  }
  return blocks;
}