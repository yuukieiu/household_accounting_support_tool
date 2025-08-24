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

  // 現在のスプレッドシートのIDをMASTER_SHEET_IDとして設定
  const currentSheetId = ss.getId();
  props.setProperty("MASTER_SHEET_ID", currentSheetId);

  // 設定完了を通知
  const ui = SpreadsheetApp.getUi();
  ui.alert(
    '設定完了',
    `MASTER_SHEET_IDを現在のスプレッドシートID（${currentSheetId}）に設定しました。`,
    ui.ButtonSet.OK
  );

  // トリガーを作成
  ScriptApp.newTrigger('openSidebarOnOpen')
    .forSpreadsheet(ss)
    .onOpen()
    .create();

  props.setProperty("TRIGGER_CREATED", "true"); // 作成済みフラグを設定

  ui.alert('完了', 'トリガーが正常に作成されました。', ui.ButtonSet.OK);
}