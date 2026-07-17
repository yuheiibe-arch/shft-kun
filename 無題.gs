function test_UI_to_Batch_Connection() {
  console.log("=== 🔬 [導通テスト] UI指示受け取り ＆ バッチ起動テスト ===");
  
  // 1. UIから送られてくるデータを擬似的に作成（代官山・通年）
  const mockPayload = {
    year: "2026",
    term: "通年",
    locations: ["代官山"]
  };
  console.log(`📡 [テスト1/4] UIからの擬似指示（Payload）を作成しました: ${JSON.stringify(mockPayload)}`);
  
  // 2. startBackgroundBatch の内部ロジックを段階的にテスト
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  console.log("📡 [テスト2/4] スプレッドシートの取得完了");
  
  // 誤爆防衛ガードのテスト
  let validLocations = mockPayload.locations.filter(loc => {
    let invalidWords = [
      "その他", "キャッシュ", "作業表", "お休み", "確定シフト", 
      "欠勤", "キャンセル", "手順書", "原本", "振替勤務", 
      "先行応募", "単独募集", "定期募集", "目次", "設定", 
      "テンプレート", "ダッシュボード"
    ];
    return !invalidWords.some(word => loc.includes(word));
  });
  console.log(`🛡️ [テスト3/4] ガード通過後のターゲット拠点: ${JSON.stringify(validLocations)}`);
  
  if (validLocations.length === 0) {
    console.log("❌ テスト終了: ターゲット拠点が0になりました。");
    return;
  }
  
  // 3. キャッシュ作成処理の限界テスト（ここがフリーズの疑いがある場所）
  console.log("⏳ [テスト4/4] 外部データの一括ダウンロードとシートキャッシュの作成を開始します...");
  console.log("   （※ここで処理が止まる、または数分かかる場合は、この処理がタイムアウトの元凶です）");
  
  const startTime = Date.now();
  
  try {
    const targetYear = parseInt(mockPayload.year, 10);
    const baseLocs = [...new Set(validLocations.map(loc => loc.replace(/（.*?）/, '')))];
    
    // 外部から一括取得（※依存関数が存在するかチェック）
    let extractedData = {};
    if (typeof fetchAndOrganizeData === 'function') {
       console.log("   -> fetchAndOrganizeData を実行中...");
       extractedData = fetchAndOrganizeData(targetYear, mockPayload.term, baseLocs);
    } else {
       console.log("   ⚠️ fetchAndOrganizeData 関数が見つかりません。");
    }
    
    let rawCache = {};
    if (typeof getMasterRawData === 'function') {
       console.log("   -> getMasterRawData を実行中...");
       rawCache = getMasterRawData(targetYear);
    } else {
       console.log("   ⚠️ getMasterRawData 関数が見つかりません。");
    }
    
    let specialtyMap = {};
    if (typeof buildDoctorSpecialtyMap === 'function') {
       console.log("   -> buildDoctorSpecialtyMap を実行中...");
       specialtyMap = buildDoctorSpecialtyMap(targetYear);
    } else {
       console.log("   ⚠️ buildDoctorSpecialtyMap 関数が見つかりません。");
    }
    
    console.log("   -> データの取得完了。隠しシート（⚙️通信キャッシュ）への保存を試みます...");
    
    // 隠しシートに保存
    if (typeof _saveBatchCache === 'function') {
      _saveBatchCache(ss, {
        extractedData: extractedData,
        rawCache: rawCache,
        specialtyMap: specialtyMap
      });
      const endTime = Date.now();
      console.log(`✅ キャッシュの作成と保存に成功しました！ (所要時間: ${(endTime - startTime) / 1000} 秒)`);
      console.log("💡 結論: UIからの受け取りからキャッシュ保存まで、正常に通過できます。トリガー起動に進めます。");
    } else {
      console.log("   ⚠️ _saveBatchCache 関数が見つかりません。");
    }
    
  } catch(e) {
    const errorTime = Date.now();
    console.log(`❌ 【エラー発生】キャッシュ作成中にフリーズ（またはクラッシュ）しました。`);
    console.log(`   発生までの時間: ${(errorTime - startTime) / 1000} 秒`);
    console.log(`   エラー内容: ${e.message}`);
    console.log("💡 結論: 予想通り、この一括ダウンロード＆シート保存処理が重すぎてシステムを止めています。");
  }
  
  console.log("=== 導通テスト完了 ===");
}