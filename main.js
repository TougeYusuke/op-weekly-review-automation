/**
 * 週次レビュー自動化ツール
 * Operation Phoenixの週次レビュー（金曜日の作業）を自動化するツール
 * 
 * 機能:
 * 1. 週次データの自動集計（記事数、収益、作業時間）
 * 2. 週次レビュー用レポートの自動生成
 * 3. 来週の計画テンプレートの自動生成
 */

// ===== 設定 =====
const SHEET_NAMES = {
  DAILY_INPUT: '日次データ入力',      // 日次データ入力シート
  WEEKLY_SUMMARY: '週次集計',         // 週次集計シート
  REPORT: 'レポート',                 // レポートシート
  NEXT_WEEK_PLAN: '来週の計画'        // 来週の計画シート
};

const TARGET_REVENUE = 80000; // 目標収益（8万円）

// ===== メイン処理 =====

/**
 * 週次レビューを実行（メイン関数）
 * 金曜日に実行する関数
 */
function runWeeklyReview() {
  try {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. 週次データを集計
    const weeklyData = aggregateWeeklyData(spreadsheet);
    
    // 2. 週次集計シートに出力
    outputWeeklySummary(spreadsheet, weeklyData);
    
    // 3. レポートを生成
    generateReport(spreadsheet, weeklyData);
    
    // 4. 来週の計画テンプレートを生成
    generateNextWeekPlan(spreadsheet, weeklyData);
    
    Logger.log('週次レビューが完了しました');
    SpreadsheetApp.getUi().alert('週次レビューが完了しました！');
  } catch (error) {
    Logger.log('エラー: ' + error);
    SpreadsheetApp.getUi().alert('エラーが発生しました: ' + error);
  }
}

// ===== 週次データ集計 =====

/**
 * 週次データを集計
 * @param {Spreadsheet} spreadsheet - スプレッドシートオブジェクト
 * @return {Object} 週次データ
 */
function aggregateWeeklyData(spreadsheet) {
  const dailySheet = spreadsheet.getSheetByName(SHEET_NAMES.DAILY_INPUT);
  if (!dailySheet) {
    throw new Error('日次データ入力シートが見つかりません');
  }
  
  const data = dailySheet.getDataRange().getValues();
  const headers = data[0];
  
  // ヘッダーのインデックスを取得
  const dateIndex = headers.indexOf('日付');
  const noteArticleIndex = headers.indexOf('Note記事');
  const wpArticleIndex = headers.indexOf('WordPress記事');
  const gasTemplateIndex = headers.indexOf('GASテンプレート');
  const noteRevenueIndex = headers.indexOf('Note有料記事収益');
  const wpRevenueIndex = headers.indexOf('WordPressアフィリエイト収益');
  const gasRevenueIndex = headers.indexOf('GAS販売収益');
  const workTimeIndex = headers.indexOf('作業時間');
  
  // 今週の日付範囲を取得（月曜日〜日曜日）
  const today = new Date();
  const dayOfWeek = today.getDay(); // 0=日曜日, 1=月曜日, ..., 6=土曜日
  const monday = new Date(today);
  monday.setDate(today.getDate() - dayOfWeek + (dayOfWeek === 0 ? -6 : 1)); // 今週の月曜日
  const sunday = new Date(monday);
  sunday.setDate(monday.getDate() + 6); // 今週の日曜日
  
  // 今週のデータを集計
  let noteArticles = 0;
  let wpArticles = 0;
  let gasTemplates = 0;
  let noteRevenue = 0;
  let wpRevenue = 0;
  let gasRevenue = 0;
  let workTime = 0;
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowDate = new Date(row[dateIndex]);
    
    // 今週のデータかチェック
    if (rowDate >= monday && rowDate <= sunday) {
      noteArticles += row[noteArticleIndex] || 0;
      wpArticles += row[wpArticleIndex] || 0;
      gasTemplates += row[gasTemplateIndex] || 0;
      noteRevenue += row[noteRevenueIndex] || 0;
      wpRevenue += row[wpRevenueIndex] || 0;
      gasRevenue += row[gasRevenueIndex] || 0;
      workTime += row[workTimeIndex] || 0;
    }
  }
  
  const totalRevenue = noteRevenue + wpRevenue + gasRevenue;
  
  // 累計収益を取得（前週までの累計 + 今週の収益）
  const cumulativeRevenue = getCumulativeRevenue(spreadsheet) + totalRevenue;
  const remainingRevenue = TARGET_REVENUE - cumulativeRevenue;
  
  return {
    weekStart: monday,
    weekEnd: sunday,
    noteArticles: noteArticles,
    wpArticles: wpArticles,
    gasTemplates: gasTemplates,
    noteRevenue: noteRevenue,
    wpRevenue: wpRevenue,
    gasRevenue: gasRevenue,
    totalRevenue: totalRevenue,
    cumulativeRevenue: cumulativeRevenue,
    remainingRevenue: remainingRevenue,
    workTime: workTime
  };
}

/**
 * 累計収益を取得
 * @param {Spreadsheet} spreadsheet - スプレッドシートオブジェクト
 * @return {number} 累計収益
 */
function getCumulativeRevenue(spreadsheet) {
  const summarySheet = spreadsheet.getSheetByName(SHEET_NAMES.WEEKLY_SUMMARY);
  if (!summarySheet) {
    return 0;
  }
  
  const data = summarySheet.getDataRange().getValues();
  if (data.length <= 1) {
    return 0;
  }
  
  // 最後の行の累計収益を取得
  const headers = data[0];
  const cumulativeIndex = headers.indexOf('累計収益');
  if (cumulativeIndex === -1) {
    return 0;
  }
  
  const lastRow = data[data.length - 1];
  return lastRow[cumulativeIndex] || 0;
}

// ===== 週次集計シートへの出力 =====

/**
 * 週次集計シートに出力
 * @param {Spreadsheet} spreadsheet - スプレッドシートオブジェクト
 * @param {Object} weeklyData - 週次データ
 */
function outputWeeklySummary(spreadsheet, weeklyData) {
  let summarySheet = spreadsheet.getSheetByName(SHEET_NAMES.WEEKLY_SUMMARY);
  
  // シートが存在しない場合は作成
  if (!summarySheet) {
    summarySheet = spreadsheet.insertSheet(SHEET_NAMES.WEEKLY_SUMMARY);
    // ヘッダーを設定
    summarySheet.getRange(1, 1, 1, 12).setValues([[
      '週開始日', '週終了日', 'Note記事数', 'WordPress記事数', 'GASテンプレート数',
      'Note有料記事収益', 'WordPressアフィリエイト収益', 'GAS販売収益',
      '今週の合計収益', '累計収益', '目標までの残り', '作業時間'
    ]]);
    // ヘッダーの書式設定
    summarySheet.getRange(1, 1, 1, 12).setFontWeight('bold');
    summarySheet.getRange(1, 1, 1, 12).setBackground('#4285f4');
    summarySheet.getRange(1, 1, 1, 12).setFontColor('#ffffff');
  }
  
  // データを追加
  const row = [
    Utilities.formatDate(weeklyData.weekStart, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    Utilities.formatDate(weeklyData.weekEnd, Session.getScriptTimeZone(), 'yyyy-MM-dd'),
    weeklyData.noteArticles,
    weeklyData.wpArticles,
    weeklyData.gasTemplates,
    weeklyData.noteRevenue,
    weeklyData.wpRevenue,
    weeklyData.gasRevenue,
    weeklyData.totalRevenue,
    weeklyData.cumulativeRevenue,
    weeklyData.remainingRevenue,
    weeklyData.workTime
  ];
  
  summarySheet.appendRow(row);
  
  // 列幅を自動調整
  summarySheet.autoResizeColumns(1, 12);
}

// ===== レポート生成 =====

/**
 * レポートを生成
 * @param {Spreadsheet} spreadsheet - スプレッドシートオブジェクト
 * @param {Object} weeklyData - 週次データ
 */
function generateReport(spreadsheet, weeklyData) {
  let reportSheet = spreadsheet.getSheetByName(SHEET_NAMES.REPORT);
  
  // シートが存在しない場合は作成
  if (!reportSheet) {
    reportSheet = spreadsheet.insertSheet(SHEET_NAMES.REPORT);
  } else {
    // 既存のシートをクリア
    reportSheet.clear();
  }
  
  // レポートを生成
  const report = [
    ['週次レビューレポート'],
    [''],
    ['週期間: ' + Utilities.formatDate(weeklyData.weekStart, Session.getScriptTimeZone(), 'yyyy-MM-dd') + ' 〜 ' + Utilities.formatDate(weeklyData.weekEnd, Session.getScriptTimeZone(), 'yyyy-MM-dd')],
    [''],
    ['## 今週の成果サマリー'],
    [''],
    ['記事数:'],
    ['  - Note記事: ' + weeklyData.noteArticles + '本'],
    ['  - WordPress記事: ' + weeklyData.wpArticles + '本'],
    ['  - GASテンプレート: ' + weeklyData.gasTemplates + '個'],
    [''],
    ['収益:'],
    ['  - Note有料記事: ' + weeklyData.noteRevenue.toLocaleString() + '円'],
    ['  - WordPressアフィリエイト: ' + weeklyData.wpRevenue.toLocaleString() + '円'],
    ['  - GAS販売: ' + weeklyData.gasRevenue.toLocaleString() + '円'],
    ['  - 今週の合計: ' + weeklyData.totalRevenue.toLocaleString() + '円'],
    [''],
    ['累計収益: ' + weeklyData.cumulativeRevenue.toLocaleString() + '円'],
    ['目標までの残り: ' + weeklyData.remainingRevenue.toLocaleString() + '円'],
    ['目標達成率: ' + calculateAchievementRate(weeklyData.cumulativeRevenue, TARGET_REVENUE) + '%'],
    [''],
    ['作業時間: ' + weeklyData.workTime + '時間'],
    [''],
    ['## 改善点'],
    [''],
    ['- [ ] 作業時間の無駄を確認'],
    ['- [ ] テンプレート化できる部分を確認'],
    ['- [ ] 自動化できる部分を確認']
  ];
  
  reportSheet.getRange(1, 1, report.length, 1).setValues(report.map(row => [row[0]]));
  
  // タイトルの書式設定
  reportSheet.getRange(1, 1).setFontSize(16).setFontWeight('bold');
  
  // 列幅を自動調整
  reportSheet.autoResizeColumn(1);
}

// ===== 目標達成率の計算 =====

/**
 * 目標達成率を計算
 * @param {number} actual - 実績値
 * @param {number} target - 目標値
 * @return {number} 達成率（0-100、目標が0の場合は0を返す）
 */
function calculateAchievementRate(actual, target) {
  if (target === 0) {
    return 0;
  }
  const rate = (actual / target) * 100;
  return Math.round(rate * 100) / 100; // 小数点第2位まで
}

/**
 * 週次目標に対する達成率を計算
 * @param {Object} weeklyData - 週次データ
 * @param {Object} weeklyTargets - 週次目標（記事数、収益など）
 * @return {Object} 達成率データ
 */
function calculateWeeklyAchievementRates(weeklyData, weeklyTargets) {
  return {
    noteArticlesRate: calculateAchievementRate(weeklyData.noteArticles, weeklyTargets.noteArticles || 0),
    wpArticlesRate: calculateAchievementRate(weeklyData.wpArticles, weeklyTargets.wpArticles || 0),
    revenueRate: calculateAchievementRate(weeklyData.totalRevenue, weeklyTargets.revenue || 0),
    workTimeRate: calculateAchievementRate(weeklyData.workTime, weeklyTargets.workTime || 0),
  };
}

// ===== 来週の計画テンプレート生成 =====

/**
 * 来週の計画テンプレートを生成
 * @param {Spreadsheet} spreadsheet - スプレッドシートオブジェクト
 * @param {Object} weeklyData - 週次データ
 */
function generateNextWeekPlan(spreadsheet, weeklyData) {
  let planSheet = spreadsheet.getSheetByName(SHEET_NAMES.NEXT_WEEK_PLAN);
  
  // シートが存在しない場合は作成
  if (!planSheet) {
    planSheet = spreadsheet.insertSheet(SHEET_NAMES.NEXT_WEEK_PLAN);
  } else {
    // 既存のシートをクリア
    planSheet.clear();
  }
  
  // 来週の日付範囲を計算
  const nextMonday = new Date(weeklyData.weekEnd);
  nextMonday.setDate(weeklyData.weekEnd.getDate() + 1);
  const nextSunday = new Date(nextMonday);
  nextSunday.setDate(nextMonday.getDate() + 6);
  
  // 来週の目標を算出（今週の実績をベースに）
  const targetNoteArticles = Math.max(weeklyData.noteArticles, 2);
  const targetWpArticles = Math.max(weeklyData.wpArticles, 2);
  const targetRevenue = Math.max(weeklyData.totalRevenue, 15000);
  
  // 計画テンプレートを生成
  const plan = [
    ['来週の計画'],
    [''],
    ['週期間: ' + Utilities.formatDate(nextMonday, Session.getScriptTimeZone(), 'yyyy-MM-dd') + ' 〜 ' + Utilities.formatDate(nextSunday, Session.getScriptTimeZone(), 'yyyy-MM-dd')],
    [''],
    ['## 来週のネタ'],
    [''],
    ['- [ ] Note記事ネタ: （ネタ帳から選ぶ）'],
    ['- [ ] WordPress記事ネタ: （ネタ帳から選ぶ）'],
    ['- [ ] GASテンプレートネタ: （ネタ帳から選ぶ）'],
    [''],
    ['## 来週のスケジュール'],
    [''],
    ['- [ ] 月曜日: Note記事執筆'],
    ['- [ ] 火曜日: WordPress記事執筆'],
    ['- [ ] 水曜日: GASテンプレート開発'],
    ['- [ ] 木曜日: Note記事 or WordPress記事'],
    ['- [ ] 金曜日: 週次レビュー'],
    ['- [ ] 土曜日: まとめ作業（画像編集、公開準備）'],
    ['- [ ] 日曜日: 次週の準備（ネタ帳整理、企画）'],
    [''],
    ['## 来週の目標'],
    [''],
    ['記事数:'],
    ['  - Note記事: ' + targetNoteArticles + '本'],
    ['  - WordPress記事: ' + targetWpArticles + '本'],
    [''],
    ['収益目標: ' + targetRevenue.toLocaleString() + '円']
  ];
  
  planSheet.getRange(1, 1, plan.length, 1).setValues(plan.map(row => [row[0]]));
  
  // タイトルの書式設定
  planSheet.getRange(1, 1).setFontSize(16).setFontWeight('bold');
  
  // 列幅を自動調整
  planSheet.autoResizeColumn(1);
}

// ===== テスト用のエクスポート =====

// テスト用のエクスポート（GASでは無視される）
if (typeof module !== 'undefined' && module.exports) {
  module.exports = {
    aggregateWeeklyData,
    getCumulativeRevenue,
    outputWeeklySummary,
    generateReport,
    generateNextWeekPlan,
    calculateAchievementRate,
    calculateWeeklyAchievementRates
  };
}

