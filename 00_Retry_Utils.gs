/**
 * ==========================================
 * 00_Retry_Utils.gs (アップデート版)
 * Google APIの気まぐれな通信タイムアウトを回避するリトライ機構
 * ==========================================
 */

function safeOpenByUrl(url) {
  return safeExecute(() => SpreadsheetApp.openByUrl(url), 3, `URLの取得(${url})`);
}

function safeExecute(action, maxRetry = 3, actionName = "スプレッドシート操作") {
  for (let i = 0; i < maxRetry; i++) {
    try {
      return action();
    } catch (e) {
      if (i === maxRetry - 1) {
        console.error(`[致命的エラー] ${actionName}に${maxRetry}回失敗しました: ${e.message}`);
        throw e;
      }
      console.warn(`[通信エラー] ${actionName}に失敗。${i + 1}回目のリトライを行います...`);
      // 2秒, 4秒, 6秒と待機時間を指数関数的に延ばしてリトライ
      Utilities.sleep(2000 + (i * 2000)); 
    }
  }
}