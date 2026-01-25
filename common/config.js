// --- 設定定数 ---
// このファイルにはアプリケーション全体で使用する設定値を定義します

/**
 * シート名の定数
 */
const SHEET_NAMES = {
  ACCOUNTS: "勘定科目",
  JOURNAL: "仕訳日記帳",
  PERIOD: "計上月"
};

/**
 * 仕訳日記帳の列番号
 */
const JOURNAL_COLUMNS = {
  DEBIT_PERIOD: 1,        // 借方計上月（自動関数）
  CREDIT_PERIOD: 2,       // 貸方計上月（自動関数）
  DATE: 3,                // 取引日
  DEBIT_CODE: 4,          // 借方勘定コード（自動関数）
  DEBIT_NAME: 5,          // 借方科目名
  DEBIT_AMOUNT: 6,        // 借方金額
  CREDIT_CODE: 7,         // 貸方勘定コード（自動関数）
  CREDIT_NAME: 8,         // 貸方科目名
  CREDIT_AMOUNT: 9,       // 貸方金額（自動関数）
  DESCRIPTION: 10,        // 摘要
  DEBIT_CONCAT: 11,       // 借方連結（自動関数）
  CREDIT_CONCAT: 12       // 貸方連結（自動関数）
};

/**
 * 勘定科目シートの列インデックス（0始まり）
 */
const ACCOUNT_COLUMNS = {
  CATEGORY: 0,            // 分類
  CODE: 1,                // 勘定コード
  NAME: 2                 // 勘定科目名
};

/**
 * 勘定科目のカテゴリ名
 */
const ACCOUNT_CATEGORIES = {
  ASSET: "資産",
  LIABILITY: "負債",
  EQUITY: "純資産",
  EXPENSE: "費用",
  REVENUE: "収益"
};

/**
 * 取引手段として使用可能なカテゴリ
 */
const PAYMENT_CATEGORIES = [
  ACCOUNT_CATEGORIES.ASSET,
  ACCOUNT_CATEGORIES.LIABILITY,
  ACCOUNT_CATEGORIES.EQUITY
];

/**
 * 相手科目として使用可能なカテゴリ
 */
const PURPOSE_CATEGORIES = [
  ACCOUNT_CATEGORIES.ASSET,
  ACCOUNT_CATEGORIES.EXPENSE,
  ACCOUNT_CATEGORIES.REVENUE,
  ACCOUNT_CATEGORIES.EQUITY,
  ACCOUNT_CATEGORIES.LIABILITY
];

/**
 * スプレッドシートの範囲設定
 */
const RANGES = {
  // 勘定科目シートの参照範囲
  ACCOUNTS_LOOKUP: "'" + SHEET_NAMES.ACCOUNTS + "'!$B$2:$C$1004",
  // 計上月シートの範囲（借方用）
  PERIOD_DEBIT: "'" + SHEET_NAMES.PERIOD + "'!$A$3:$D$13",
  // 計上月シートの範囲（貸方用）
  PERIOD_CREDIT: "'" + SHEET_NAMES.PERIOD + "'!$E$3:$H$114"
};

/**
 * 勘定コードの範囲（費用科目の判定用）
 */
const ACCOUNT_CODE_RANGES = {
  EXPENSE_MIN: 500,
  EXPENSE_MAX: 600
};

/**
 * PropertiesServiceのキー名
 */
const PROPERTY_KEYS = {
  MASTER_SHEET_ID: "MASTER_SHEET_ID",
  TRIGGER_CREATED: "TRIGGER_CREATED"
};
