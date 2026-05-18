// ==========================================
// 【Import Module】外部シートからの転記モジュール
// ==========================================

/**
 * カスタムメニューから呼ばれるダイアログ起動関数
 */
function openExternalImportDialog() {
  ExternalDataImporter.showDialog();
}

/**
 * HTML側から呼ばれる、既存シートの存在チェック関数
 */
function checkExistingSheets(selectedSheets, targetYear) {
  return ExternalDataImporter.checkExisting(selectedSheets, targetYear);
}

/**
 * HTML側から呼ばれる転記実行関数
 */
function runExternalImportBatch(selectedSheets, targetYear) {
  return ExternalDataImporter.processImport(selectedSheets, targetYear);
}

/**
 * 処理本体をカプセル化したオブジェクト
 */
const ExternalDataImporter = {
  
  EXTERNAL_SS_ID: '13PlAElj8SODdBNME28TIDpifxG0oxS3wAnKIsFTa53k',

  /**
   * 1. 外部シートを読み込み、ダイアログを表示する
   */
  showDialog: function() {
    const ui = SpreadsheetApp.getUi();
    try {
      const extSs = SpreadsheetApp.openById(this.EXTERNAL_SS_ID);
      const extSheets = extSs.getSheets();
      
      const targetSheets = extSheets
        .map(s => s.getName())
        .filter(name => name.includes("調整"));

      if (targetSheets.length === 0) {
        ui.alert("お知らせ", "外部ファイルに「調整」を含むシートが見つかりませんでした。", ui.ButtonSet.OK);
        return;
      }

      // 年度ロジック：1〜3月なら「今年」、4〜12月なら「来年」が次年度になる
      const now = new Date();
      const currentMonth = now.getMonth() + 1;
      const currentYear = now.getFullYear();
      const nextYear = (currentMonth <= 3) ? currentYear : currentYear + 1;

      // HTMLテンプレート
      const htmlTemplate = HtmlService.createTemplate(`
        <!DOCTYPE html>
        <html>
          <head>
            <base target="_top">
            <style>
              body { font-family: sans-serif; padding: 10px; color: #333; }
              .container { display: flex; flex-direction: column; gap: 15px; }
              .form-group { display: flex; flex-direction: column; gap: 5px; }
              label { font-weight: bold; font-size: 14px; }
              input[type="number"] { padding: 8px; font-size: 14px; border-radius: 4px; border: 1px solid #ccc; width: 100px; }
              .sheet-list { border: 1px solid #ccc; padding: 10px; border-radius: 4px; max-height: 120px; overflow-y: auto; background: #f9f9f9; }
              .sheet-item { display: flex; align-items: center; gap: 8px; margin-bottom: 5px; font-size: 13px; }
              .buttons { margin-top: 10px; display: flex; justify-content: flex-end; gap: 10px; }
              button { padding: 8px 16px; border: none; border-radius: 4px; cursor: pointer; font-size: 14px; }
              .btn-cancel { background-color: #f3f3f3; color: #333; }
              .btn-submit { background-color: #4285f4; color: white; font-weight: bold; }
              .btn-submit:hover { background-color: #357ae8; }
              .btn-submit:disabled { background-color: #a0c1f9; cursor: not-allowed; }
            </style>
          </head>
          <body>
            <div class="container">
              <div class="form-group">
                <label for="year">① 作成する年度</label>
                <div><input type="number" id="year" value="<?= nextYear ?>"> 年度</div>
              </div>

              <div class="form-group">
                <label>② 転記するシートを選択</label>
                <div class="sheet-list">
                  <? for (let i = 0; i < sheets.length; i++) { ?>
                    <label class="sheet-item">
                      <input type="checkbox" name="sheetCheck" value="<?= sheets[i] ?>" checked>
                      <?= sheets[i] ?>
                    </label>
                  <? } ?>
                </div>
              </div>

              <div class="buttons">
                <button class="btn-cancel" onclick="google.script.host.close()">キャンセル</button>
                <button class="btn-submit" onclick="runScript()">転 記</button>
              </div>
            </div>

            <script>
              function runScript() {
                const year = document.getElementById('year').value;
                const checkboxes = document.querySelectorAll('input[name="sheetCheck"]:checked');
                const selectedSheets = Array.from(checkboxes).map(cb => cb.value);

                if (selectedSheets.length === 0) {
                  alert("シートを1つ以上選択してください。");
                  return;
                }
                if (!year) {
                  alert("年度を入力してください。");
                  return;
                }

                const btn = document.querySelector('.btn-submit');
                btn.disabled = true;
                btn.innerText = '確認中...';

                google.script.run
                  .withSuccessHandler(function(existingSheets) {
                    if (existingSheets.length > 0) {
                      const msg = "以下のシートは既に存在します。上書き（既存のシートを削除して再作成）しますか？\\n\\n" + existingSheets.join("\\n");
                      if (!confirm(msg)) {
                        btn.disabled = false;
                        btn.innerText = '転 記';
                        return;
                      }
                    }

                    btn.innerText = '処理中...';
                    google.script.run
                      .withSuccessHandler(function(res) {
                        google.script.host.close();
                      })
                      .withFailureHandler(function(err) {
                        alert('エラーが発生しました: ' + err);
                        btn.disabled = false;
                        btn.innerText = '転 記';
                      })
                      .runExternalImportBatch(selectedSheets, year);
                  })
                  .withFailureHandler(function(err) {
                    alert('エラーが発生しました: ' + err);
                    btn.disabled = false;
                    btn.innerText = '転 記';
                  })
                  .checkExistingSheets(selectedSheets, year);
              }
            </script>
          </body>
        </html>
      `);

      htmlTemplate.nextYear = nextYear;
      htmlTemplate.sheets = targetSheets;
      
      const html = htmlTemplate.evaluate().setWidth(400).setHeight(350);
      ui.showModalDialog(html, '外部シートからの転記');

    } catch (e) {
      ui.alert("エラー", "外部ファイルへのアクセスに失敗しました。\n" + e.message, ui.ButtonSet.OK);
    }
  },

  /**
   * 作成予定のシートが既に存在するかをチェックしてリストを返す
   */
  checkExisting: function(selectedSheets, targetYear) {
    const activeSs = SpreadsheetApp.getActiveSpreadsheet();
    const existingNames = [];

    for (const sheetName of selectedSheets) {
      const isPartTime = sheetName.includes("非常勤");
      const isFullTime = !isPartTime && sheetName.includes("常勤");
      
      const targetName = (isFullTime ? "常勤" : isPartTime ? "定期非常勤" : "不明") + targetYear + "年度";
      
      if (activeSs.getSheetByName(targetName)) {
        existingNames.push(targetName);
      }
    }
    return [...new Set(existingNames)];
  },

  /**
   * 2. 実際の転記処理を行う
   */
  processImport: function(selectedSheets, targetYear) {
    const extSs = SpreadsheetApp.openById(this.EXTERNAL_SS_ID);
    const activeSs = SpreadsheetApp.getActiveSpreadsheet();
    let successCount = 0;

    for (const sheetName of selectedSheets) {
      const sourceSheet = extSs.getSheetByName(sheetName);
      if (!sourceSheet) continue;

      const isPartTime = sheetName.includes("非常勤");
      const isFullTime = !isPartTime && sheetName.includes("常勤");
      
      const targetName = (isFullTime ? "常勤" : isPartTime ? "定期非常勤" : "不明") + targetYear + "年度";

      // 上書き対応：同名のシートがすでに存在する場合は削除する
      const oldSheet = activeSs.getSheetByName(targetName);
      if (oldSheet) {
        activeSs.deleteSheet(oldSheet);
      }

      // Activeなスプレッドシートへ丸ごとコピー
      const newSheet = sourceSheet.copyTo(activeSs);

      if (isFullTime) {
        this.formatFullTimeSheet_(newSheet);
      } else if (isPartTime) {
        // 定期非常勤の場合は、選択された年度(targetYear)も渡す
        this.formatPartTimeSheet_(newSheet, targetYear);
      }

      newSheet.setName(targetName);
      successCount++;
    }

    activeSs.toast(`${successCount}件のシートを転記（完了）しました！`, "完了", 5);
    return true;
  },

  /**
   * 3. 常勤用のデータ成形
   */
  formatFullTimeSheet_: function(sheet) {
    const maxCols = sheet.getMaxColumns();
    if (maxCols > 25) {
      sheet.deleteColumns(26, maxCols - 25);
    }

    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    
    let retireIdx = 5; 
    for (let c = 0; c < headers.length; c++) {
      if (String(headers[c]).includes("退職")) {
        retireIdx = c;
        break;
      }
    }

    for (let i = data.length - 1; i > 0; i--) {
      const retirementDate = String(data[i][retireIdx]).trim();
      if (retirementDate !== "") {
        sheet.deleteRow(i + 1);
      }
    }

    const newLastRow = sheet.getLastRow();
    const newMaxCols = sheet.getMaxColumns();
    if (newLastRow > 1) {
      sheet.getRange(2, 14, newLastRow - 1, 2).clearContent();
      sheet.getRange(2, 1, newLastRow - 1, newMaxCols).setBackground(null);
    }
  },

  /**
   * 4. 定期非常勤用のデータ成形
   */
  formatPartTimeSheet_: function(sheet, targetYear) {
    const data = sheet.getDataRange().getValues();
    const headers = data[0]; 
    
    // 各項目の列インデックスを動的に検索
    let retireIdx = 5;     // デフォルトF列
    let notNeededIdx = 10; // デフォルトK列
    let hireIdx = -1;      // 入職日（デフォルトなし）

    for (let c = 0; c < headers.length; c++) {
      const h = String(headers[c]);
      if (h.includes("退職")) retireIdx = c;
      if (h.includes("対応不要")) notNeededIdx = c;
      if (h.includes("入職")) hireIdx = c;
    }
    
    // 不要な行の除外（下から上へ）
    for (let i = data.length - 1; i > 0; i--) {
      const retirementDate = String(data[i][retireIdx]).trim();
      const isNotNeeded = data[i][notNeededIdx];
      const isSkip = (isNotNeeded === true || String(isNotNeeded).toUpperCase() === "TRUE");

      if (retirementDate !== "" || isSkip) {
        sheet.deleteRow(i + 1);
      }
    }

    const newLastRow = sheet.getLastRow();

    // 【追加ルール1】入職日を選択した年度の 4/1 に一括変更
    if (hireIdx !== -1 && newLastRow > 1) {
      const newHireDate = `${targetYear}/04/01`;
      sheet.getRange(2, hireIdx + 1, newLastRow - 1, 1).setValue(newHireDate);
    }

    // 【追加ルール2,3】列の整理 (Q, R, S列を残して左に詰める)
    // 現在: A~I(1~9), J~M(10~13), N(14), O~P(15~16), Q(17), R(18), S(19)...
    
    // 1. J, K, L, M列 (10~13) を削除
    sheet.deleteColumns(10, 4);
    // これにより、旧N列が10列目(J)、旧O, P列が11, 12列目になる

    // 2. 旧O, P列にあたる11列目と12列目を削除
    sheet.deleteColumns(11, 2);
    // これにより、旧Q, R, S列がそれぞれ 11(K), 12(L), 13(M)列目として並ぶ

    // 3. J列（旧N列）のヘッダーを「勤務備考」に変更
    sheet.getRange(1, 10).setValue("勤務備考");

    // 4. 残したS列（現在の13列目/M列）より右側をすべて削除
    const maxCols = sheet.getMaxColumns();
    if (maxCols > 13) {
      sheet.deleteColumns(14, maxCols - 13);
    }
    
    // 背景色をクリア（白に戻す）
    const finalMaxCols = sheet.getMaxColumns();
    if (newLastRow > 1) {
      sheet.getRange(2, 1, newLastRow - 1, finalMaxCols).setBackground(null);
    }
  }
};