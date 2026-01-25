// --- マスタシートを取得する共通関数 ---
function getMasterSpreadsheet() {
  const props = PropertiesService.getScriptProperties();
  const sheetId = props.getProperty(PROPERTY_KEYS.MASTER_SHEET_ID);
  if (!sheetId) throw new Error("MASTER_SHEET_ID が設定されていません");
  return SpreadsheetApp.openById(sheetId);
}

// 勘定科目取得
function getAccounts() {
  const sheet = getMasterSpreadsheet().getSheetByName(SHEET_NAMES.ACCOUNTS);
  const data = sheet.getDataRange().getValues().slice(1); // ヘッダー除く
  return data
    .filter(row => row[ACCOUNT_COLUMNS.CATEGORY]) // 空行除外
    .map(row => ({
      category: row[ACCOUNT_COLUMNS.CATEGORY],               // 分類
      code: row[ACCOUNT_COLUMNS.CODE],                       // 勘定コード
      name: row[ACCOUNT_COLUMNS.NAME],                       // 勘定科目名
      display: row[ACCOUNT_COLUMNS.CODE] + " " + row[ACCOUNT_COLUMNS.NAME]  // プルダウン表示
    }));
}
// 仕訳日記帳に追加
function addJournalEntry(entry) {
  const sheet = getMasterSpreadsheet().getSheetByName(SHEET_NAMES.JOURNAL);
  const lastRow = sheet.getLastRow() + 1;

  // 列文字を事前に生成（数式内で使用）
  const colDate = getColumnLetter(JOURNAL_COLUMNS.DATE);
  const colDebitCode = getColumnLetter(JOURNAL_COLUMNS.DEBIT_CODE);
  const colDebitName = getColumnLetter(JOURNAL_COLUMNS.DEBIT_NAME);
  const colDebitAmount = getColumnLetter(JOURNAL_COLUMNS.DEBIT_AMOUNT);
  const colCreditCode = getColumnLetter(JOURNAL_COLUMNS.CREDIT_CODE);
  const colCreditName = getColumnLetter(JOURNAL_COLUMNS.CREDIT_NAME);
  const colDebitPeriod = getColumnLetter(JOURNAL_COLUMNS.DEBIT_PERIOD);
  const colCreditPeriod = getColumnLetter(JOURNAL_COLUMNS.CREDIT_PERIOD);

  // 基本データの設定
  sheet.getRange(lastRow, JOURNAL_COLUMNS.DATE).setValue(entry.date);        // 取引日
  sheet.getRange(lastRow, JOURNAL_COLUMNS.DEBIT_NAME).setValue(entry.debitName);   // 借方科目
  sheet.getRange(lastRow, JOURNAL_COLUMNS.DEBIT_AMOUNT).setValue(entry.debitAmount); // 借方金額
  sheet.getRange(lastRow, JOURNAL_COLUMNS.CREDIT_NAME).setValue(entry.creditName);  // 貸方科目
  sheet.getRange(lastRow, JOURNAL_COLUMNS.DESCRIPTION).setValue(entry.description); // 摘要

  // 自動関数の設定
  // 借方勘定コード取得
  sheet.getRange(lastRow, JOURNAL_COLUMNS.DEBIT_CODE).setFormula(
    `=query(${RANGES.ACCOUNTS_LOOKUP},"select B where C = '" & ${colDebitName}${lastRow} & "'")`
  );
  // 貸方勘定コード取得
  sheet.getRange(lastRow, JOURNAL_COLUMNS.CREDIT_CODE).setFormula(
    `=query(${RANGES.ACCOUNTS_LOOKUP},"select B where C = '" & ${colCreditName}${lastRow} & "'")`
  );
  // 貸方金額（借方金額を参照）
  sheet.getRange(lastRow, JOURNAL_COLUMNS.CREDIT_AMOUNT).setFormula(
    `=${colDebitAmount}${lastRow}`
  );

  // 計上月の自動判定（借方）
  sheet.getRange(lastRow, JOURNAL_COLUMNS.DEBIT_PERIOD).setFormula(
    `=iferror(if(AND(${colDebitCode}${lastRow} < "${ACCOUNT_CODE_RANGES.EXPENSE_MAX}", ${colDebitCode}${lastRow} >= "${ACCOUNT_CODE_RANGES.EXPENSE_MIN}"),query(${RANGES.PERIOD_DEBIT},"select A WHERE C <= date '" & TEXT(${colDate}${lastRow},"YYYY-MM-DD") & "' and D >= date '" & TEXT(${colDate}${lastRow},"YYYY-MM-DD") & "'"),query(${RANGES.PERIOD_CREDIT},"select E WHERE G <= date '" & TEXT(${colDate}${lastRow},"YYYY-MM-DD") & "' and H >= date '" & TEXT(${colDate}${lastRow},"YYYY-MM-DD") & "'")),"")`
  );
  // 計上月の自動判定（貸方）
  sheet.getRange(lastRow, JOURNAL_COLUMNS.CREDIT_PERIOD).setFormula(
    `=iferror(if(AND(${colCreditCode}${lastRow} < "${ACCOUNT_CODE_RANGES.EXPENSE_MAX}", ${colCreditCode}${lastRow} >= "${ACCOUNT_CODE_RANGES.EXPENSE_MIN}"),query(${RANGES.PERIOD_DEBIT},"select A WHERE C <= date '" & TEXT(${colDate}${lastRow},"YYYY-MM-DD") & "' and D >= date '" & TEXT(${colDate}${lastRow},"YYYY-MM-DD") & "'"),query(${RANGES.PERIOD_CREDIT},"select E WHERE G <= date '" & TEXT(${colDate}${lastRow},"YYYY-MM-DD") & "' and H >= date '" & TEXT(${colDate}${lastRow},"YYYY-MM-DD") & "'")),"")`
  );
  // 連結文字列
  sheet.getRange(lastRow, JOURNAL_COLUMNS.DEBIT_CONCAT).setFormula(
    `=CONCAT(${colDebitPeriod}${lastRow},${colDebitName}${lastRow})`
  );
  sheet.getRange(lastRow, JOURNAL_COLUMNS.CREDIT_CONCAT).setFormula(
    `=CONCAT(${colCreditPeriod}${lastRow},${colCreditName}${lastRow})`
  );

  return true; // successHandler 発火用
}

/**
 * 列番号を列文字（A, B, C...）に変換するヘルパー関数
 * @param {number} columnNumber - 列番号（1始まり）
 * @return {string} 列文字（A, B, C...）
 */
function getColumnLetter(columnNumber) {
  let result = '';
  while (columnNumber > 0) {
    columnNumber--;
    result = String.fromCharCode(65 + (columnNumber % 26)) + result;
    columnNumber = Math.floor(columnNumber / 26);
  }
  return result;
}

// フロントエンド用の設定値を取得
function getFrontendConfig() {
  return {
    paymentCategories: PAYMENT_CATEGORIES,
    purposeCategories: PURPOSE_CATEGORIES
  };
}
