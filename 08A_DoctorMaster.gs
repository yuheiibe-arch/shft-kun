/**
 * ====================================================================
 * 08A_DoctorMaster.gs (バグ修正・完全版)
 * 外部マスタから「常勤」「定期非常勤」の医籍番号・名前・時給情報を確実に取得する処理
 * ====================================================================
 */

var GLOBAL_DOC_MASTER = GLOBAL_DOC_MASTER || {};

function _fetchDoctorTypes(ctx) {
  const getColIdxSmart = (sheet, keyword) => {
    if (!sheet) return -1;
    const maxCol = sheet.getLastColumn();
    if (maxCol === 0) return -1;
    const data = sheet.getRange(1, 1, 5, maxCol).getValues();
    for (let r = 0; r < 5; r++) {
      for (let c = 0; c < maxCol; c++) {
        if (String(data[r][c]).includes(keyword)) return { row: r, col: c };
      }
    }
    return -1;
  };
  
  const normalizeId = (id) => String(id).replace(/[０-９]/g, s => String.fromCharCode(s.charCodeAt(0) - 0xFEE0)).replace(/[^0-9]/g, '');

  ctx.jokinDocNames = new Set(); 
  ctx.hijokinDocNames = new Set(); 
  ctx.docMasterData = {}; 
  if (typeof GLOBAL_DOC_MASTER === 'undefined') GLOBAL_DOC_MASTER = {};
  
  try {
    const extSsUrl = "https://docs.google.com/spreadsheets/d/1aEjphEv_63SeWQmwiOy9sx7IrMfawU01sHbKd_Ki4iA/edit";
    const extSs = SpreadsheetApp.openByUrl(extSsUrl);
    
    const getLatestSheetByPrefix = (ss, prefix) => {
      const sheets = ss.getSheets().filter(s => s.getName().startsWith(prefix));
      if (sheets.length === 0) return null;
      sheets.sort((a, b) => b.getName().localeCompare(a.getName(), 'ja'));
      return sheets[0];
    };
    
    // 【常勤の取得】
    const jokinSheet = getLatestSheetByPrefix(extSs, "常勤");
    let jokinIds = new Set();
    if (jokinSheet) {
      let idPos = getColIdxSmart(jokinSheet, "医籍番号");
      let namePos = getColIdxSmart(jokinSheet, "氏名");
      if (namePos === -1) namePos = getColIdxSmart(jokinSheet, "医師名"); // バグ修正: 確実なフォールバック
      let spPos = getColIdxSmart(jokinSheet, "特別時給の内訳");
      let conPos = getColIdxSmart(jokinSheet, "契約時給");

      if (idPos !== -1) {
        let data = jokinSheet.getDataRange().getValues();
        for (let i = idPos.row + 1; i < data.length; i++) {
          let id = normalizeId(data[i][idPos.col]);
          if (id) jokinIds.add(id);
          if (namePos !== -1) {
            let n = String(data[i][namePos.col]).replace(/先生$/, "").replace(/\s+/g, "").trim();
            if (n) {
              ctx.jokinDocNames.add(n);
              ctx.docMasterData[n] = {
                medId: id,
                specialWageDetail: spPos !== -1 ? String(data[i][spPos.col]) : "",
                contractType: conPos !== -1 ? String(data[i][conPos.col]) : ""
              };
              GLOBAL_DOC_MASTER[n] = { medId: id };
            }
          }
        }
      }
    }

    // 【定期非常勤の取得】
    const hijokinSheet = getLatestSheetByPrefix(extSs, "定期非常勤");
    let hijokinIds = new Set();
    if (hijokinSheet) {
      let idPos = getColIdxSmart(hijokinSheet, "医籍番号");
      let namePos = getColIdxSmart(hijokinSheet, "氏名");
      if (namePos === -1) namePos = getColIdxSmart(hijokinSheet, "医師名"); // バグ修正
      let spPos = getColIdxSmart(hijokinSheet, "特別時給の内訳");
      let conPos = getColIdxSmart(hijokinSheet, "契約時給");

      if (idPos !== -1) {
        let data = hijokinSheet.getDataRange().getValues();
        for (let i = idPos.row + 1; i < data.length; i++) {
          let id = normalizeId(data[i][idPos.col]);
          if (id) hijokinIds.add(id);
          if (namePos !== -1) {
            let n = String(data[i][namePos.col]).replace(/先生$/, "").replace(/\s+/g, "").trim();
            if (n) {
              ctx.hijokinDocNames.add(n);
              ctx.docMasterData[n] = {
                medId: id,
                specialWageDetail: spPos !== -1 ? String(data[i][spPos.col]) : "",
                contractType: conPos !== -1 ? String(data[i][conPos.col]) : ""
              };
              GLOBAL_DOC_MASTER[n] = { medId: id };
            }
          }
        }
      }
    }
    
    const activeSs = SpreadsheetApp.getActiveSpreadsheet();
    ["先行応募", "先行応募表", "振替", "振替勤務", "単発・スポット"].forEach(sName => {
      let sheet = activeSs.getSheetByName(sName);
      if (sheet) {
        let idPos = getColIdxSmart(sheet, "医籍番号");
        if (idPos === -1) idPos = { row: 2, col: 2 }; 
        let namePos = getColIdxSmart(sheet, "氏名");
        if (namePos === -1) namePos = getColIdxSmart(sheet, "医師名");
        if (namePos === -1) namePos = { row: 2, col: 5 }; 
        
        let data = sheet.getDataRange().getValues();
        let startRow = idPos.row + 1;
        for (let i = startRow; i < data.length; i++) {
          if (!data[i]) continue;
          let id = normalizeId(data[i][idPos.col]);
          let name = String(data[i][namePos.col]).replace(/先生$/, "").replace(/\s+/g, "").trim();
          if (id && name) {
            if (jokinIds.has(id)) ctx.jokinDocNames.add(name);
            if (hijokinIds.has(id)) ctx.hijokinDocNames.add(name);
            if (!GLOBAL_DOC_MASTER[name]) GLOBAL_DOC_MASTER[name] = { medId: id };
          }
        }
      }
    });
  } catch(e) {
    console.log("常勤判定マスタの読み込みエラー: " + e.message);
  }

  if (!ctx.docMasterData) ctx.docMasterData = {};
  
  if (ctx.contractsByLoc) {
    for (let loc in ctx.contractsByLoc) {
      ctx.contractsByLoc[loc].forEach(c => {
        let n = String(c.docName).replace(/先生$/, "").replace(/\s+/g, "").trim();
        if (n) {
           if (c.type === "常勤") ctx.jokinDocNames.add(n);
           if (c.type === "定期非常勤") ctx.hijokinDocNames.add(n);
           if (!ctx.docMasterData[n]) {
             ctx.docMasterData[n] = {
               medId: c.medId || (GLOBAL_DOC_MASTER[n] ? GLOBAL_DOC_MASTER[n].medId : ""),
               specialWageDetail: c.specialWageDetail || "",
               contractType: c.contractType || ""
             };
           }
        }
      });
    }
  }
}