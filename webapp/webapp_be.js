// --- webapp専用の関数 ---
// このファイルはwebappでのみ使用される関数を含みます

// Webアプリとして開いたときに呼ばれる
function doGet() {
  // 常にマスタシートを対象とする
  return HtmlService.createHtmlOutputFromFile('webapp')
      .setTitle('仕訳入力')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// スプレッドシート情報取得
function getSpreadsheetInfo() {
  const ss = getMasterSpreadsheet();
  return {
    name: ss.getName(),
    url: ss.getUrl()
  };
}
