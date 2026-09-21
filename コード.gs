function doGet(e) {
  // JSONオブジェクト格納用の入れ物
  let prop = PropertiesService.getScriptProperties().getProperties();
  const Spread_ID = prop.Spread_ID;
  var rowData = {};  

  if (e.parameter == undefined) {

      // パラメータ不良の場合はundefinedで返す
      var getvalue = "undefined";

      // エラーはJSONで返すつもりなので
      rowData.value = getvalue;
      var result = JSON.stringify(rowData);
      return ContentService.createTextOutput(result);

  } else {

      // 書込先スプレッドシートのIDを入力
      var id = Spread_ID;

      // スプレッドシート名指定
      var sheet = SpreadsheetApp.openById(id).getSheetByName(String(e.parameter.p7));

      // GAS側で現在日時を取得（フォーマット指定: YYYY-MM-DD HH:mm）
      // スプレッドシートの日時型オブジェクトとして保持したい場合は var now = new Date(); でも構いません
      var now = Utilities.formatDate(new Date(), "Asia/Tokyo", "yyyy-MM-dd HH:mm");

      // 日時（now）を先頭にし、p2〜p6の測定データを配列にセット
      // p1（日時データ）はクライアント側から渡さなくてもGASで自動生成されます
      var array = [ 
        now, 
        e.parameter.p2, 
        e.parameter.p3, 
        e.parameter.p4, 
        e.parameter.p5, 
        e.parameter.p6 
      ];

      // シートに配列を書き込み
      sheet.appendRow(array);

      // 書き込み終わったらOKを返す
      var getvalue = "ok";

      // エラーはJSONで返すつもりなので
      rowData.value = getvalue;
      var result = JSON.stringify(rowData);
      return ContentService.createTextOutput(result);

  }
}
