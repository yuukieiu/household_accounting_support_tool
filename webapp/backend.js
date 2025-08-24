// --- マスタシートを取得する共通関数 ---
function getMasterSpreadsheet() {
  const props = PropertiesService.getScriptProperties();
  const sheetId = props.getProperty("MASTER_SHEET_ID");
  if (!sheetId) throw new Error("MASTER_SHEET_ID が設定されていません");
  return SpreadsheetApp.openById(sheetId);
}

// 勘定科目取得
function getAccounts() {
  const sheet = getMasterSpreadsheet().getSheetByName("勘定科目");
  const data = sheet.getDataRange().getValues().slice(1); // ヘッダー除く
  return data
    .filter(row => row[0]) // 空行除外
    .map(row => ({
      category: row[0],               // 分類
      code: row[1],                   // 勘定コード
      name: row[2],                   // 勘定科目名
      display: row[1] + " " + row[2]  // プルダウン表示
    }));
}
// 仕訳日記帳に追加
function addJournalEntry(entry) {
  const sheet = getMasterSpreadsheet().getSheetByName("仕訳日記帳");
  const lastRow = sheet.getLastRow() + 1;

  sheet.getRange(lastRow, 3).setValue(entry.date);        // 取引日
  sheet.getRange(lastRow, 4).setFormula(`=query('勘定科目'!$B$2:$C$1004,"select B where C = '${entry.debitName}'")`);
  sheet.getRange(lastRow, 5).setValue(entry.debitName);   // 借方科目
  sheet.getRange(lastRow, 6).setValue(entry.debitAmount); // 借方金額
  sheet.getRange(lastRow, 7).setFormula(`=query('勘定科目'!$B$2:$C$1004,"select B where C = '${entry.creditName}'")`);
  sheet.getRange(lastRow, 8).setValue(entry.creditName);  // 貸方科目
  sheet.getRange(lastRow, 9).setValue(entry.debitAmount); // 貸方金額
  sheet.getRange(lastRow, 10).setValue(entry.description); // 摘要

  // 自動関数系
  sheet.getRange(lastRow, 1).setFormula(`=iferror(if(AND(D${lastRow} < "600", D${lastRow} >= "500"),query('計上月'!$A$3:$D$13,"select A WHERE C <= date '" & TEXT(C${lastRow},"YYYY-MM-DD") & "' and D >= date '" & TEXT(C${lastRow},"YYYY-MM-DD") & "'"),query('計上月'!$E$3:$H$114,"select E WHERE G <= date '" & TEXT(C${lastRow},"YYYY-MM-DD") & "' and H >= date '" & TEXT(C${lastRow},"YYYY-MM-DD") & "'")),"")`);
  sheet.getRange(lastRow, 2).setFormula(`=iferror(if(AND(G${lastRow} < "600", G${lastRow} >= "500"),query('計上月'!$A$3:$D$13,"select A WHERE C <= date '" & TEXT(C${lastRow},"YYYY-MM-DD") & "' and D >= date '" & TEXT(C${lastRow},"YYYY-MM-DD") & "'"),query('計上月'!$E$3:$H$114,"select E WHERE G <= date '" & TEXT(C${lastRow},"YYYY-MM-DD") & "' and H >= date '" & TEXT(C${lastRow},"YYYY-MM-DD") & "'")),"")`);
  sheet.getRange(lastRow, 11).setFormula(`=CONCAT(A${lastRow},E${lastRow})`);
  sheet.getRange(lastRow, 12).setFormula(`=CONCAT(B${lastRow},H${lastRow})`);

  return true; // successHandler 発火用
}
