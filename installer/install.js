// トリガーで呼ばれる関数
function openSidebarOnOpen(e) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  SpreadsheetApp.getUi().createMenu('仕訳入力')
    .addItem('サイドバーを開く', 'showSidebar')
    .addToUi();
  showSidebar(); // スプレッドシートを開いたときにもサイドバー表示
}

// サイドバー表示
function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('sidebar')
      .setTitle('仕訳入力');
  SpreadsheetApp.getUi().showSidebar(html);
}

// インストール型トリガーを作成する関数（1回だけ実行）
function createOnOpenTriggerOnce() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty("TRIGGER_CREATED")) return; // すでに作成済みならスキップ

  ScriptApp.newTrigger('openSidebarOnOpen')
    .forSpreadsheet(ss)
    .onOpen()
    .create();

  props.setProperty("TRIGGER_CREATED", "true"); // 作成済みフラグを設定
}