function debugImazatoDynamicFinal() {
  const targetId = "1fIFvTck_g9-Hp8MSpY7hIWwE5L2buJjBHIY4GBEf_MY";
  const ss = SpreadsheetApp.openById(targetId);
  const sheet = ss.getSheetByName("2026今里 のコピー");

  if (!sheet) {
    console.log("❌ シートが見つかりません。");
    return;
  }

  console.log("=== 🔍 ハードコーディングを完全排除した動的解析テストを開始します ===");

  const data = sheet.getDataRange().getValues();
  let count = 0;

  for (let r = 0; r < data.length; r++) {
    let label = String(data[r][0]).trim();
    if (label === "先行応募" || label === "先行応募医師") {
      let startRow = r; 
      
      // -------------------------------------------------------------
      // ① 動的読み取り：固定値を使わず、シートのダッシュボードから所属を記憶する
      // -------------------------------------------------------------
      let masterS = []; // 先行応募
      let masterJ = []; // 常勤
      let masterT = []; // 非常勤

      for (let c = 3; c <= 10; c++) {
        let valS = String(data[startRow][c]).trim();
        if (valS && !["募集", "休"].includes(valS)) masterS.push(valS);
        
        let valJ = String(data[startRow + 1][c]).trim();
        if (valJ && !["募集", "休"].includes(valJ)) masterJ.push(valJ);
        
        let valT = String(data[startRow + 2][c]).trim();
        if (valT && !["募集", "休"].includes(valT)) masterT.push(valT);
      }

      // -------------------------------------------------------------
      // ② カレンダー解析：実際にシフトに入っている医師と時間を抽出
      // -------------------------------------------------------------
      let docShifts = {}; 
      let currentDow = "";
      let currentWeek = "";
      let isWeekend = false;

      for (let calR = startRow + 17; calR < data.length; calR++) {
        if (calR > startRow + 20 && String(data[calR][0]).trim().includes("先行応募")) break;

        let colC = String(data[calR][2]).trim();
        if (colC === "1診目" || colC === "2診目") {
          
          if (colC === "1診目") {
            let rawDow = data[calR][0];
            let rawWeek = data[calR][1];

            if (rawDow instanceof Date) {
              const dows = ["日","月","火","水","木","金","土"];
              currentDow = dows[rawDow.getDay()];
            } else {
              currentDow = String(rawDow).trim();
            }
            if (rawWeek) currentWeek = String(rawWeek).replace(/[^0-9]/g, '');

            let nextDayCell = String(data[calR+1] ? data[calR+1][0] : "");
            isWeekend = (currentDow.includes("土") || currentDow.includes("日") || currentDow.includes("祝") || nextDayCell.includes("祝"));
            currentDow = currentDow.replace(/[0-9\/]/g, '').replace(/\(祝\)/, '').replace(/祝/g, '').trim();
            if(!currentDow) currentDow = "祝";
          }

          for (let c = 3; c <= 14; c++) {
            let docName = String(data[calR][c]).trim();
            if (docName && !["募集", "休", "休館日", "未開院"].includes(docName)) { 
              if (!docShifts[docName]) docShifts[docName] = [];
              docShifts[docName].push({ week: currentWeek, dow: currentDow, hour: c + 6, isWeekend: isWeekend });
            }
          }
        }
      }

      // -------------------------------------------------------------
      // ③ フィルタリング：記憶した所属をもとに、稼働者だけを再分類
      // -------------------------------------------------------------
      let activeDocs = Object.keys(docShifts);
      
      // 動的マスターを使って分類（これで宮田先生も先行の先生も正確に振り分けられます）
      let finalS = activeDocs.filter(d => masterS.includes(d));
      let finalJ = activeDocs.filter(d => masterJ.includes(d));
      // 先行にも常勤にもいない人を非常勤とする
      let finalT = activeDocs.filter(d => !masterS.includes(d) && !masterJ.includes(d)); 

      // -------------------------------------------------------------
      // ④ ダッシュボード再描画（空欄の色抜け対策済み）
      // -------------------------------------------------------------
      const setDash = (rowOffset, docs, bgColor) => {
        let vals = new Array(8).fill("");
        let bgs = new Array(8).fill("#ffffff"); // 空欄は白
        
        for(let i=0; i<8; i++) {
          if (i < docs.length) {
            vals[i] = docs[i];
            bgs[i] = bgColor; // 名前があるセルだけ色付け
          }
        }
        sheet.getRange(startRow + 1 + rowOffset, 4, 1, 8).setValues([vals]).setBackgrounds([bgs]);
      };
      
      setDash(0, finalS, "#d9ead3"); // 先行
      setDash(1, finalJ, "#fce5cd"); // 常勤
      setDash(2, finalT, "#d9d2e9"); // 非常勤

      // -------------------------------------------------------------
      // ⑤ コスト計算（非常勤＝finalT のみ）
      // -------------------------------------------------------------
      let costOutput = [];
      finalT.forEach(doc => {
        let shifts = docShifts[doc];
        if (!shifts || shifts.length === 0) return;

        let totalHours = shifts.length;
        let totalCost = 0;
        let ratesUsed = new Set();
        let dayMap = {}; 

        shifts.forEach(s => {
          let key = `${s.week}_${s.dow}`;
          if(!dayMap[key]) dayMap[key] = { week: s.week, dow: s.dow, hours: [], isWeekend: s.isWeekend };
          dayMap[key].hours.push(s.hour);
        });

        let scheduleMap = {};
        Object.values(dayMap).forEach(day => {
          let sortedHours = day.hours.sort((a,b)=>a-b);
          let blocks = [];
          let startH = sortedHours[0];
          let prevH = startH;

          for(let i=1; i<=sortedHours.length; i++) {
            if(i === sortedHours.length || sortedHours[i] !== prevH + 1) {
              blocks.push(`${String(startH).padStart(2,'0')}:00～${String(prevH+1).padStart(2,'0')}:00`);
              if(i < sortedHours.length) { startH = sortedHours[i]; prevH = startH; }
            } else { prevH = sortedHours[i]; }
          }

          let schedKey = `${day.dow}曜日：${blocks.join(" / ")}`;
          if(!scheduleMap[schedKey]) scheduleMap[schedKey] = new Set();
          scheduleMap[schedKey].add(day.week);

          // シミュレーション計算
          sortedHours.forEach(h => {
             let rate = 14000; 
             if(day.isWeekend) rate = 16000;
             else if(h >= 18) rate = 16000; 
             totalCost += rate;
             ratesUsed.add(rate.toLocaleString());
          });
        });

        let contractLines = [];
        for (let [sched, weeksSet] of Object.entries(scheduleMap)) {
          let weeksArr = Array.from(weeksSet).sort();
          let weekStr = weeksArr.length >= 4 ? "毎週" : `第${weeksArr.join("・")}`;
          contractLines.push(`【今里】${weekStr}${sched}`);
        }

        costOutput.push([
          `${doc}（定非）`, totalHours, contractLines.join("\n"), Array.from(ratesUsed).join("/"), totalCost.toLocaleString()
        ]);
      });

      // -------------------------------------------------------------
      // ⑥ 右側表の出力
      // -------------------------------------------------------------
      sheet.getRange(startRow + 1, 12, 16, 5).clearContent().setBorder(false,false,false,false,false,false);
      let headers = [["医師名", "稼働時間", "契約内容", "時給", "月間コスト"]];
      sheet.getRange(startRow + 2, 12, 1, 5).setValues(headers).setBackground("#e4efff").setBorder(true, true, true, true, true, true, null, SpreadsheetApp.BorderStyle.SOLID);
           
      if (costOutput.length > 0) {
        sheet.getRange(startRow + 3, 12, costOutput.length, 5).setValues(costOutput).setBorder(true, true, true, true, true, true, null, SpreadsheetApp.BorderStyle.SOLID).setHorizontalAlignment("left");
      }
      count++;
    }
  }
  SpreadsheetApp.flush();
  console.log(`✅ ${count}ヶ月分の完全動的解析・出力が完了しました！`);
}