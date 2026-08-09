/**
 * ==========================================
 * 02_FetchData.gs
 * 外部シートからのデータ取得・解析・構造化
 * ★空白除去＆完全同一人物認識・強化版
 * ★セル内改行（複数シフト）の完全読み取り対応版
 * ★【追加】同セル内に「契約」と「確定」が混在する場合、「契約」行を優先抽出する機能
 * ==========================================
 */

function fetchAndOrganizeData(year, term, targetLocations) {
  const extSs = getExternalSpreadsheet();
  const doctorMaster = fetchDoctorMaster(extSs, year);
  const locationDict = getLocationDictionary();
  const shiftData = fetchKintaiData(extSs, year, term, targetLocations, doctorMaster, locationDict);
  
  return {
    master: doctorMaster,
    shifts: shiftData
  };
}

function fetchDoctorMaster(extSs, year) {
  const master = {};
  const types = ["常勤", "定期非常勤"];
  
  types.forEach(type => {
    const sheetName = `${type}${year}年度`;
    const sheet = extSs.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log(`[警告] 医師マスタシートが見つかりません: ${sheetName}`);
      return;
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return;
    
    const headers = data[0];
    const nameIdx = headers.indexOf("医師名");
    const wageIdx = headers.indexOf("契約時給") !== -1 ? headers.indexOf("契約時給") : headers.indexOf("時給");
    const specialWageIdx = headers.indexOf("特別時給の内訳") !== -1 ? headers.indexOf("特別時給の内訳") : headers.indexOf("特別時給");
    const bikouIdx = headers.indexOf("勤務備考") !== -1 ? headers.indexOf("勤務備考") : headers.indexOf("備考"); 
    
    if (nameIdx === -1) {
      Logger.log(`[警告] ${sheetName} に「医師名」列がありません。`);
      return;
    }

    for (let r = 1; r < data.length; r++) {
      let rawName = String(data[r][nameIdx]);
      if (!rawName) continue;
      
      const docNameClean = rawName.replace(/[\s ]+/g, "").replace(/先生$/, "");
      
      master[docNameClean] = {
        name: docNameClean, 
        originalName: docNameClean, 
        type: type,
        contractType: type, 
        wageInfo: (wageIdx !== -1) ? data[r][wageIdx] : "",
        specialWageDetail: (specialWageIdx !== -1) ? data[r][specialWageIdx] : "",
        bikou: (bikouIdx !== -1) ? data[r][bikouIdx] : "",       
        remarks: (bikouIdx !== -1) ? data[r][bikouIdx] : "",     
        contractInfo: (bikouIdx !== -1) ? data[r][bikouIdx] : "",
        contract: (bikouIdx !== -1) ? data[r][bikouIdx] : ""
      };
    }
  });
  
  return master;
}

function fetchKintaiData(extSs, year, term, targetLocations, doctorMaster, locationDict) {
  let shiftData = {};
  targetLocations.forEach(loc => { shiftData[loc] = {}; });
  
  const types = ["常勤", "定期非常勤"];
  
  types.forEach(type => {
    const sheetName = `${type}勤怠${year}`;
    const sheet = extSs.getSheetByName(sheetName);
    if (!sheet) {
      Logger.log(`[警告] 勤怠シートが見つかりません: ${sheetName}`);
      return;
    }
    
    const data = sheet.getDataRange().getValues();
    if (data.length < 3) return;
    
    let dateIdx = -1;
    let dateRowIdx = -1;
    for (let r = 0; r < 3; r++) {
      for (let c = 0; c < 5; c++) {
        if (String(data[r][c]).replace(/[\s ]+/g, "").includes("日付")) {
          dateIdx = c;
          dateRowIdx = r;
          break;
        }
      }
      if (dateIdx !== -1) break;
    }
    if (dateIdx === -1) throw new Error(`${sheetName} に「日付」列がありません。`);
    
    const excludeHeaders = ["日付", "曜日", "平日判定", "祝日判定", "祝日備考", "年末年始判定"];
    const doctorCols = [];
    
    let nameRowIdx = dateRowIdx > 0 ? dateRowIdx - 1 : 0;
    
    for (let c = dateIdx + 1; c < data[nameRowIdx].length; c++) {
      let docName = String(data[nameRowIdx][c] || "").replace(/[\s ]+/g, "").replace(/先生$/, ""); 
      if (docName && !excludeHeaders.includes(docName) && isNaN(Number(docName))) { 
        doctorCols.push({ index: c, name: docName });
      }
    }
    
    let startRow = dateRowIdx + 1;
    
    for (let r = startRow; r < data.length; r++) {
      const row = data[r];
      let dateVal = row[dateIdx];
      if (!dateVal) continue;
      
      let month, day;
      if (dateVal instanceof Date) {
        month = dateVal.getMonth() + 1;
        day = dateVal.getDate();
      } else {
        const parts = String(dateVal).split('/');
        if (parts.length >= 3) {
          month = parseInt(parts[1], 10);
          day = parseInt(parts[2], 10);
        } else {
          month = parseInt(parts[0], 10);
          day = parseInt(parts[1], 10);
        }
      }
      if (isNaN(month) || isNaN(day)) continue;
      
      const actualYear = (month >= 4 && month <= 12) ? parseInt(year, 10) : parseInt(year, 10) + 1;
      const dateStr = `${actualYear}/${('0' + month).slice(-2)}/${('0' + day).slice(-2)}`;
      
      if (!isDateInTerm(actualYear, month, year, term)) continue;
      
      doctorCols.forEach(doc => {
        const cellVal = row[doc.index];
        if (!cellVal || String(cellVal).trim() === "" || String(cellVal).trim() === "休") return;
        
        const rawCellText = String(cellVal).trim();
        let lines = rawCellText.split(/\r\n|\n|\r/);
        
        // ==========================================
        // ★追加：セル内に「契約」と「確定」などが混在している場合、
        // 「契約」が含まれる行のみを抽出して優先する。
        // （別拠点の兼務で「契約」が2行ある場合は両方残る）
        // ==========================================
        const contractLines = lines.filter(line => line.includes("契約"));
        if (contractLines.length > 0) {
          lines = contractLines;
        }
        
        lines.forEach(shiftStr => {
          shiftStr = shiftStr.trim();
          if (!shiftStr || shiftStr === "休") return;
          
          const timeMatch = shiftStr.match(/([0-9]{1,2})[:：]?([0-9]{2})?\s*[-~～]\s*([0-9]{1,2})[:：]?([0-9]{2})?/);
          if (!timeMatch) return;
          
          const startH = parseInt(timeMatch[1], 10);
          const startM = timeMatch[2] ? parseInt(timeMatch[2], 10) : 0;
          const endH = parseInt(timeMatch[3], 10);
          const endM = timeMatch[4] ? parseInt(timeMatch[4], 10) : 0;

          if (isNaN(startH) || isNaN(endH)) return;

          const startTime = `${('0' + startH).slice(-2)}:${('0' + startM).slice(-2)}`;
          const endTime = `${('0' + endH).slice(-2)}:${('0' + endM).slice(-2)}`;
          
          let cleanLoc = null;
          const bracketMatch = shiftStr.match(/【(.*?)】/);
          if (bracketMatch && bracketMatch[1].trim() !== "") {
            cleanLoc = normalizeLocationName(bracketMatch[1].trim(), locationDict);
          } else {
            for (const loc of targetLocations) {
              if (shiftStr.includes(loc)) {
                cleanLoc = loc;
                break;
              }
            }
            if (!cleanLoc) {
              const strWithoutTime = shiftStr.replace(timeMatch[0], '').trim();
              cleanLoc = normalizeLocationName(strWithoutTime, locationDict);
            }
          }
          
          if (!cleanLoc || !targetLocations.includes(cleanLoc)) return;
          
          if (!shiftData[cleanLoc][dateStr]) {
            shiftData[cleanLoc][dateStr] = [];
          }
          
          let docMasterEntry = doctorMaster[doc.name];
          let finalDocName = docMasterEntry ? docMasterEntry.originalName : doc.name;
          finalDocName = finalDocName.replace(/[\s ]+/g, "").replace(/先生$/, "");
          
          shiftData[cleanLoc][dateStr].push({
            doctorName: finalDocName,
            type: type,
            startTime: startTime,
            endTime: endTime,
            startHour: startH,
            hours: (endH - startH),
            rawShift: shiftStr,
            wageInfo: docMasterEntry ? docMasterEntry.wageInfo : "",
            specialWageDetail: docMasterEntry ? docMasterEntry.specialWageDetail : ""
          });
        });
      });
    }
  });
  
  return shiftData;
}

function isDateInTerm(actualYear, month, baseYearStr, term) {
  if (term === "通年") return true;
  const baseYear = parseInt(baseYearStr, 10);
  if (term === "上期") {
    return (actualYear === baseYear && month >= 4 && month <= 9);
  }
  if (term === "下期") {
    if (actualYear === baseYear && month >= 10 && month <= 12) return true;
    if (actualYear === baseYear + 1 && month >= 1 && month <= 3) return true;
    return false;
  }
  return true;
}