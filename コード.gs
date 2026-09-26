function doGet(e) {
  const prop = PropertiesService.getScriptProperties().getProperties();
  const Spread_ID = prop.Spread_ID;
  let rowData = {};  

  if (!e || !e.parameter || Object.keys(e.parameter).length === 0) {
    rowData.value = "undefined";
    return ContentService.createTextOutput(JSON.stringify(rowData))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const sheetName = e.parameter.p1;
  
  if (!sheetName) {
    rowData.value = "missing_sheet_name";
    return ContentService.createTextOutput(JSON.stringify(rowData))
      .setMimeType(ContentService.MimeType.JSON);
  }

  try {
    const spreadsheet = SpreadsheetApp.openById(Spread_ID);
    const sheet = spreadsheet.getSheetByName(String(sheetName));

    if (!sheet) {
      rowData.value = "sheet_not_found";
      return ContentService.createTextOutput(JSON.stringify(rowData))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // 1. GAS側で現在日時を取得（フォーマット: YYYY-MM-DD HH:mm）
    const now = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy-MM-dd HH:mm");

    // 2. e.parameter に含まれる p2, p3, p4 ... の最大インデックスを探す
    let maxIndex = 1;
    Object.keys(e.parameter).forEach(key => {
      const match = key.match(/^p(\d+)$/);
      if (match) {
        const num = parseInt(match[1], 10);
        if (num > maxIndex) {
          maxIndex = num;
        }
      }
    });

    // 3. p2 から maxIndex まで、空文字も含めて順番に配列化
    const pValues = [];
    for (let i = 2; i <= maxIndex; i++) {
      const val = e.parameter[`p${i}`];
      // パラメータ自体が存在しない場合は空文字、存在する場合はその値を保持
      pValues.push(val !== undefined ? val : "");
    }

    // 4. スプレッドシートへの書き込み用配列を作成
    // 1列目: 日時 (now)
    // 2列目: シート名 (p1)
    // 3列目以降: 受信データ (p2, p3, ..., p13, ...)
    const array = [now, sheetName, ...pValues];

    // シートの最下行へ追記
    sheet.appendRow(array);

    rowData.value = "ok";
    return ContentService.createTextOutput(JSON.stringify(rowData))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    rowData.value = "error: " + err.message;
    return ContentService.createTextOutput(JSON.stringify(rowData))
      .setMimeType(ContentService.MimeType.JSON);
  }
}