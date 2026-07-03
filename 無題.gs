function debugExtractWeeklyBlocks() {
  console.log("=========================================");
  console.log("🧪 定期募集の 15-21 分割テスト（本命関数）");
  
  try {
    // 15:00〜21:00が空いている状態をシミュレーション
    let arr = new Array(13).fill(null);
    for(let i=6; i<=11; i++) arr[i] = new Set([1,2,3,4,5]);
    
    let result = _extractWeeklyBlocks(arr);
    console.log("【抽出結果】:");
    result.forEach(b => console.log(` ${b.sH}:00 - ${b.eH}:00`));
    
    if (result.length === 2 && result[0].sH === 15 && result[1].sH === 18) {
      console.log("👉 関数は最新です！シートの残骸が原因です。");
    } else {
      console.log("👉 ❌ 危険：GASが古い関数を読み込んでいます（分割されていません）。");
    }
  } catch(e) {
    console.error("エラー: " + e.message);
  }
  console.log("=========================================");
}