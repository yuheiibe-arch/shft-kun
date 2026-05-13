/**
 * ====================================================================
 * 08_Boshukun_Generator.gs
 * シフト作成の進行管理（オーケストレーター）★拠点別マージ対応版
 * ====================================================================
 */

function _generateAllShifts(ctx) {
  // 1. 医師の雇用形態（常勤・定期非常勤）を取得して ctx に保存
  _fetchDoctorTypes(ctx);

  // 2. 拠点ごとのシフト処理ループ（★要塞化：エラーが起きても止まらない）
  ctx.locNames.forEach(locName => {
    try {
      _processLocation(ctx, locName);
    } catch (e) {
      console.error(`[警告] ${locName} のデータ生成中にエラーが発生しました（スキップして継続します）: ${e.message}`);
    }
  });

  // 3. 単独（スポット・追加募集）シフトのグループ化処理
  if (typeof _processGroupedSingles === "function") {
    try {
      _processGroupedSingles(ctx);
    } catch(e) {
      console.error(`[警告] 単独シフトのグループ化中にエラーが発生しました: ${e.message}`);
    }
  }

  // 4. 確定シフト作成リストの同一医師・拠点・契約マージ処理
  // ★拠点名をキーに含めることで、笠井先生のような「同日別拠点」の反映漏れを防止します
  ctx.confirmList = _mergeConfirmList(ctx.confirmList);

  // 5. 確定シフト作成リストを「拠点名」＞「医師名」＞「日付」の順にソートする
  ctx.confirmList.sort((a, b) => {
    if (a["拠点名"] !== b["拠点名"]) return String(a["拠点名"]).localeCompare(String(b["拠点名"]), 'ja');
    if (a["医師名"] !== b["医師名"]) return String(a["医師名"]).localeCompare(String(b["医師名"]), 'ja');
    return a._sortDate - b._sortDate;
  });

  // ソート用の隠しプロパティをお掃除
  ctx.confirmList.forEach(item => {
    delete item._sortDate;
  });

  // 6. 定期募集リストの同時間帯・曜日マージ処理
  if (typeof _mergeMasterRegularList === "function") {
    ctx.masterRegularList = _mergeMasterRegularList(ctx.masterRegularList);
  }

  // 7. 欠勤・シフトキャンセルリストのメモリ上での重複排除のみを行う
  ctx.cancelList = _deduplicateCancelList(ctx.cancelList);
}

/**
 * 確定シフトリストのマージ処理
 * ★修正：拠点が異なる場合は合体させず、独立した行として維持する（笠井先生対応）
 */
function _mergeConfirmList(confirmList) {
  const mergedMap = new Map();
  const dowOrder = ["月", "火", "水", "木", "金", "土", "日"];

  // 1. 「医師・種別・期間・拠点」をセットでグループ化する
  confirmList.forEach(item => {
    // ★重要：キーに拠点を追加することで、同日の別拠点シフトが上書き消滅するのを防ぎます
    const key = `${item["医師名"]}_${item["種別"]}_${item["設定期間"]}_${item["拠点名"]}`;

    if (!mergedMap.has(key)) {
      mergedMap.set(key, { ...item, _items: [item] });
    } else {
      mergedMap.get(key)._items.push(item);
    }
  });

  const result = [];
  
  // 2. グループごとにデータを結合して整形
  mergedMap.forEach(group => {
    const convertMQC = (name) => name === "MQC" ? "院外勤務（小児科）" : name;

    let uniqueAreas = [...new Set(group._items.map(it => it["エリア"]))];
    uniqueAreas = uniqueAreas.filter(area => area && area !== "その他");
    if (uniqueAreas.length === 0) uniqueAreas = ["その他"]; 
    
    const uniqueLocs = [...new Set(group._items.map(it => convertMQC(it["拠点名"])))];
    
    group["エリア"] = uniqueAreas.join(" / ");
    group["拠点名"] = uniqueLocs.join(" / ");

    let rawContract = group["契約内容"] || "";
    if (!rawContract) {
      const foundItem = group._items.find(it => it["契約内容"]);
      if (foundItem) rawContract = foundItem["契約内容"];
    }
    
    if (rawContract) {
      group["契約内容"] = String(rawContract).replace(/【MQC】[ 　]*(.*)/g, "【院外勤務（小児科）】$1\n※MQC業務");
    }

    group._items.forEach(it => {
      if (!it._dow) {
        const match = String(it["設定曜日"]).match(/[月火水木金土日]/);
        it._dow = match ? match[0] : "";
      }
    });

    group._items.sort((a, b) => dowOrder.indexOf(a._dow) - dowOrder.indexOf(b._dow));

    let mergedTimes = [];
    let mergedDows = [];
    let mergedWages = [];
    let mergedCautions = [];

    const isMultiLoc = uniqueLocs.length > 1;

    group._items.forEach(it => {
      const dispLoc = convertMQC(it["拠点名"]);
      const locStr = isMultiLoc ? `${dispLoc}・` : "";
      const pfx = `[${locStr}${it._dow}曜] `;
      
      mergedTimes.push(`${pfx}${it["設定時間"]}`);
      
      let originalDowStr = String(it["設定曜日"]);
      if (originalDowStr.length === 1) originalDowStr += "曜日";
      mergedDows.push(`${pfx}${originalDowStr}`);
      
      if (it["時給"]) {
        mergedWages.push(`[${locStr}${it._dow}曜(平日等)]\n${it["時給"]}`);
      }

      if (it["注意箇所"]) {
        mergedCautions.push(it["注意箇所"]);
      }
    });

    group["設定時間"] = mergedTimes.join("\n");
    group["設定曜日"] = mergedDows.join("\n");
    group["時給"] = mergedWages.join("\n\n");

    if (mergedCautions.length > 0) {
      let combinedCaution = mergedCautions.join("\n\n");
      group["注意箇所"] = combinedCaution.replace(/\n{3,}/g, '\n\n').trim();
    }

    delete group._items;
    result.push(group);
  });

  return result;
}

/**
 * メモリ上でのみ欠勤・キャンセルリストの重複を排除する
 */
function _deduplicateCancelList(cancelList) {
  const newCancelMap = new Map();
  
  cancelList.forEach(item => {
    const targetDate = item["該当日"] || item["该当日"] || "";
    const docNameClean = String(item["医師名"]).replace(/[\s　]+/g, "").replace(/先生$/, "");
    
    let dateStr = targetDate;
    if (targetDate instanceof Date) {
      dateStr = Utilities.formatDate(targetDate, "JST", "yyyy年MM月dd日");
    } else {
      dateStr = String(targetDate).trim();
    }
    
    const reason = String(item["理由"]).trim();
    const loc = String(item["対象拠点"]).trim();
    const timeStr = String(item["対象勤務時間"]).replace(/[-〜~]/g, "〜").trim();
    
    const key = `${docNameClean}_${dateStr}_${reason}_${loc}_${timeStr}`;
    
    item["医師名"] = `${docNameClean}先生`;
    item["該当日"] = dateStr;
    item["対象勤務時間"] = timeStr;
    
    if (!newCancelMap.has(key)) {
      newCancelMap.set(key, item);
    }
  });

  return Array.from(newCancelMap.values());
}