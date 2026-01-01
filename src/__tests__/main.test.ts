/**
 * 週次レビュー自動化ツールのテストファイル
 */

import { describe, it, expect, beforeEach, vi } from 'vitest';
import { SpreadsheetApp, Utilities, Session, createMockSheet } from '../__mocks__/gas-mock';

// グローバル変数としてGAS APIを設定
declare global {
  var SpreadsheetApp: typeof SpreadsheetApp;
  var Utilities: typeof Utilities;
  var Session: typeof Session;
  var Logger: { log: ReturnType<typeof vi.fn> };
}

let gasFunctions: any;

beforeEach(() => {
  // グローバル変数にGAS APIを設定
  global.SpreadsheetApp = SpreadsheetApp as any;
  global.Utilities = Utilities as any;
  global.Session = Session as any;
  global.Logger = { log: vi.fn() };
  
  // モックをリセット
  vi.clearAllMocks();
  
  // main.jsを読み込む（初回のみ）
  if (!gasFunctions) {
    gasFunctions = require('../../main.js');
  }
});

describe('aggregateWeeklyData', () => {
  it('週次データを正しく集計する', () => {
    // モックシートを作成
    const mockSheet = createMockSheet();
    const mockSpreadsheet = {
      getSheetByName: vi.fn(() => mockSheet),
    };
    
    // SpreadsheetAppのモックを設定
    SpreadsheetApp.getActiveSpreadsheet = vi.fn(() => mockSpreadsheet as any);
    
    // テストデータを設定
    const mockData = [
      ['日付', '曜日', 'Note記事', 'WordPress記事', 'GASテンプレート', 'Note有料記事収益', 'WordPressアフィリエイト収益', 'GAS販売収益', '作業時間', 'メモ'],
      ['2026-01-01', '月', 1, 0, 0, 0, 0, 0, 3, ''],
      ['2026-01-02', '火', 0, 1, 0, 0, 0, 0, 3, ''],
      ['2026-01-03', '水', 0, 0, 1, 0, 0, 0, 3, ''],
    ];
    
    (mockSheet.getDataRange as any).mockReturnValue({
      getValues: () => mockData,
    });
    
    // テスト実行
    const result = gasFunctions.aggregateWeeklyData(mockSpreadsheet);
    
    // 検証
    expect(result).toBeDefined();
    expect(result.noteArticles).toBe(1);
    expect(result.wpArticles).toBe(1);
    expect(result.gasTemplates).toBe(1);
    expect(result.workTime).toBe(9);
  });
  
  it('シートが存在しない場合はエラーを投げる', () => {
    const mockSpreadsheet = {
      getSheetByName: vi.fn(() => null),
    };
    
    SpreadsheetApp.getActiveSpreadsheet = vi.fn(() => mockSpreadsheet as any);
    
    expect(() => {
      gasFunctions.aggregateWeeklyData(mockSpreadsheet);
    }).toThrow('日次データ入力シートが見つかりません');
  });
});

describe('getCumulativeRevenue', () => {
  it('累計収益を正しく取得する', () => {
    const mockSheet = createMockSheet();
    const mockSpreadsheet = {
      getSheetByName: vi.fn(() => mockSheet),
    };
    
    SpreadsheetApp.getActiveSpreadsheet = vi.fn(() => mockSpreadsheet as any);
    
    // テストデータを設定（累計収益が50000円）
    const mockData = [
      ['週開始日', '週終了日', 'Note記事数', 'WordPress記事数', 'GASテンプレート数', 'Note有料記事収益', 'WordPressアフィリエイト収益', 'GAS販売収益', '今週の合計収益', '累計収益', '目標までの残り', '作業時間'],
      ['2025-12-25', '2025-12-31', 2, 2, 0, 6000, 9000, 0, 15000, 50000, 30000, 15],
    ];
    
    (mockSheet.getDataRange as any).mockReturnValue({
      getValues: () => mockData,
    });
    
    const result = gasFunctions.getCumulativeRevenue(mockSpreadsheet);
    
    expect(result).toBe(50000);
  });
  
  it('シートが存在しない場合は0を返す', () => {
    const mockSpreadsheet = {
      getSheetByName: vi.fn(() => null),
    };
    
    SpreadsheetApp.getActiveSpreadsheet = vi.fn(() => mockSpreadsheet as any);
    
    const result = gasFunctions.getCumulativeRevenue(mockSpreadsheet);
    
    expect(result).toBe(0);
  });
});

describe('outputWeeklySummary', () => {
  it('週次集計シートに正しく出力する', () => {
    const mockSheet = createMockSheet();
    const mockSpreadsheet = {
      getSheetByName: vi.fn(() => mockSheet),
      insertSheet: vi.fn(() => mockSheet),
    };
    
    SpreadsheetApp.getActiveSpreadsheet = vi.fn(() => mockSpreadsheet as any);
    
    const weeklyData = {
      weekStart: new Date('2026-01-01'),
      weekEnd: new Date('2026-01-07'),
      noteArticles: 2,
      wpArticles: 2,
      gasTemplates: 1,
      noteRevenue: 6000,
      wpRevenue: 9000,
      gasRevenue: 0,
      totalRevenue: 15000,
      cumulativeRevenue: 50000,
      remainingRevenue: 30000,
      workTime: 15,
    };
    
    // シートが存在しない場合をシミュレート
    (mockSpreadsheet.getSheetByName as any).mockReturnValue(null);
    
    gasFunctions.outputWeeklySummary(mockSpreadsheet, weeklyData);
    
    // シートが作成されたことを確認
    expect(mockSpreadsheet.insertSheet).toHaveBeenCalledWith('週次集計');
    // appendRowが呼ばれたことを確認
    expect(mockSheet.appendRow).toHaveBeenCalled();
  });
});

describe('generateReport', () => {
  it('レポートを正しく生成する', () => {
    const mockSheet = createMockSheet();
    const mockSpreadsheet = {
      getSheetByName: vi.fn(() => null),
      insertSheet: vi.fn(() => mockSheet),
    };
    
    SpreadsheetApp.getActiveSpreadsheet = vi.fn(() => mockSpreadsheet as any);
    
    const weeklyData = {
      weekStart: new Date('2026-01-01'),
      weekEnd: new Date('2026-01-07'),
      noteArticles: 2,
      wpArticles: 2,
      gasTemplates: 1,
      noteRevenue: 6000,
      wpRevenue: 9000,
      gasRevenue: 0,
      totalRevenue: 15000,
      cumulativeRevenue: 50000,
      remainingRevenue: 30000,
      workTime: 15,
    };
    
    gasFunctions.generateReport(mockSpreadsheet, weeklyData);
    
    // シートが作成されたことを確認
    expect(mockSpreadsheet.insertSheet).toHaveBeenCalledWith('レポート');
    // getRangeが呼ばれたことを確認（setValuesとsetFontSize/setFontWeightのため）
    expect(mockSheet.getRange).toHaveBeenCalled();
    // getRangeが複数回呼ばれることを確認（setValues用とsetFontSize/setFontWeight用）
    expect(mockSheet.getRange).toHaveBeenCalledTimes(2);
  });

  it('前週比を含むレポートを正しく生成する', () => {
    const mockSheet = createMockSheet();
    const mockSpreadsheet = {
      getSheetByName: vi.fn(() => null),
      insertSheet: vi.fn(() => mockSheet),
    };
    
    SpreadsheetApp.getActiveSpreadsheet = vi.fn(() => mockSpreadsheet as any);
    
    const weeklyData = {
      weekStart: new Date('2026-01-08'),
      weekEnd: new Date('2026-01-14'),
      noteArticles: 3,
      wpArticles: 3,
      gasTemplates: 1,
      noteRevenue: 8000,
      wpRevenue: 12000,
      gasRevenue: 0,
      totalRevenue: 20000,
      cumulativeRevenue: 70000,
      remainingRevenue: 10000,
      workTime: 18,
    };

    const comparisonData = {
      hasPreviousData: true,
      noteArticlesChange: 50,
      wpArticlesChange: 50,
      gasTemplatesChange: 100,
      revenueChange: 33.33,
      workTimeChange: 20,
    };
    
    gasFunctions.generateReport(mockSpreadsheet, weeklyData, comparisonData);
    
    // シートが作成されたことを確認
    expect(mockSpreadsheet.insertSheet).toHaveBeenCalledWith('レポート');
    // getRangeが呼ばれたことを確認
    expect(mockSheet.getRange).toHaveBeenCalled();
  });
});

describe('generateNextWeekPlan', () => {
  it('来週の計画テンプレートを正しく生成する', () => {
    const mockSheet = createMockSheet();
    const mockSpreadsheet = {
      getSheetByName: vi.fn(() => null),
      insertSheet: vi.fn(() => mockSheet),
    };
    
    SpreadsheetApp.getActiveSpreadsheet = vi.fn(() => mockSpreadsheet as any);
    
    const weeklyData = {
      weekStart: new Date('2026-01-01'),
      weekEnd: new Date('2026-01-07'),
      noteArticles: 2,
      wpArticles: 2,
      gasTemplates: 1,
      noteRevenue: 6000,
      wpRevenue: 9000,
      gasRevenue: 0,
      totalRevenue: 15000,
      cumulativeRevenue: 50000,
      remainingRevenue: 30000,
      workTime: 15,
    };
    
    gasFunctions.generateNextWeekPlan(mockSpreadsheet, weeklyData);
    
    // シートが作成されたことを確認
    expect(mockSpreadsheet.insertSheet).toHaveBeenCalledWith('来週の計画');
    // getRangeが呼ばれたことを確認（setValuesとsetFontSize/setFontWeightのため）
    expect(mockSheet.getRange).toHaveBeenCalled();
    // getRangeが複数回呼ばれることを確認（setValues用とsetFontSize/setFontWeight用）
    expect(mockSheet.getRange).toHaveBeenCalledTimes(2);
  });
});

describe('calculateAchievementRate', () => {
  it('目標を達成した場合、100%以上を返す', () => {
    expect(gasFunctions.calculateAchievementRate(100, 80)).toBe(125);
    expect(gasFunctions.calculateAchievementRate(160, 100)).toBe(160);
  });

  it('目標未達成の場合、100%未満を返す', () => {
    expect(gasFunctions.calculateAchievementRate(50, 100)).toBe(50);
    expect(gasFunctions.calculateAchievementRate(75, 100)).toBe(75);
  });

  it('目標をちょうど達成した場合、100%を返す', () => {
    expect(gasFunctions.calculateAchievementRate(100, 100)).toBe(100);
    expect(gasFunctions.calculateAchievementRate(80, 80)).toBe(100);
  });

  it('目標が0の場合、0を返す', () => {
    expect(gasFunctions.calculateAchievementRate(100, 0)).toBe(0);
    expect(gasFunctions.calculateAchievementRate(0, 0)).toBe(0);
  });

  it('実績が0の場合、0%を返す', () => {
    expect(gasFunctions.calculateAchievementRate(0, 100)).toBe(0);
    expect(gasFunctions.calculateAchievementRate(0, 50)).toBe(0);
  });

  it('小数点第2位まで正確に計算する', () => {
    expect(gasFunctions.calculateAchievementRate(33, 100)).toBe(33);
    expect(gasFunctions.calculateAchievementRate(66, 100)).toBe(66);
    expect(gasFunctions.calculateAchievementRate(33.33, 100)).toBe(33.33);
  });

  it('負の値も正しく処理する', () => {
    expect(gasFunctions.calculateAchievementRate(-50, 100)).toBe(-50);
    expect(gasFunctions.calculateAchievementRate(50, -100)).toBe(-50);
  });
});

describe('calculateWeeklyAchievementRates', () => {
  it('週次目標に対する達成率を正しく計算する', () => {
    const weeklyData = {
      noteArticles: 2,
      wpArticles: 2,
      totalRevenue: 15000,
      workTime: 15,
    };

    const weeklyTargets = {
      noteArticles: 2,
      wpArticles: 2,
      revenue: 20000,
      workTime: 20,
    };

    const result = gasFunctions.calculateWeeklyAchievementRates(weeklyData, weeklyTargets);

    expect(result.noteArticlesRate).toBe(100); // 2/2 = 100%
    expect(result.wpArticlesRate).toBe(100); // 2/2 = 100%
    expect(result.revenueRate).toBe(75); // 15000/20000 = 75%
    expect(result.workTimeRate).toBe(75); // 15/20 = 75%
  });

  it('目標が設定されていない場合、0%を返す', () => {
    const weeklyData = {
      noteArticles: 2,
      wpArticles: 2,
      totalRevenue: 15000,
      workTime: 15,
    };

    const weeklyTargets = {
      noteArticles: 0,
      wpArticles: 0,
      revenue: 0,
      workTime: 0,
    };

    const result = gasFunctions.calculateWeeklyAchievementRates(weeklyData, weeklyTargets);

    expect(result.noteArticlesRate).toBe(0);
    expect(result.wpArticlesRate).toBe(0);
    expect(result.revenueRate).toBe(0);
    expect(result.workTimeRate).toBe(0);
  });

  it('一部の目標が設定されていない場合、該当項目のみ0%を返す', () => {
    const weeklyData = {
      noteArticles: 2,
      wpArticles: 2,
      totalRevenue: 15000,
      workTime: 15,
    };

    const weeklyTargets = {
      noteArticles: 2,
      wpArticles: 0, // 目標未設定
      revenue: 20000,
      workTime: 0, // 目標未設定
    };

    const result = gasFunctions.calculateWeeklyAchievementRates(weeklyData, weeklyTargets);

    expect(result.noteArticlesRate).toBe(100);
    expect(result.wpArticlesRate).toBe(0); // 目標未設定
    expect(result.revenueRate).toBe(75);
    expect(result.workTimeRate).toBe(0); // 目標未設定
  });
});

describe('calculateChangeRate', () => {
  it('前週より増加した場合、正の値を返す', () => {
    expect(gasFunctions.calculateChangeRate(110, 100)).toBe(10); // +10%
    expect(gasFunctions.calculateChangeRate(150, 100)).toBe(50); // +50%
  });

  it('前週より減少した場合、負の値を返す', () => {
    expect(gasFunctions.calculateChangeRate(90, 100)).toBe(-10); // -10%
    expect(gasFunctions.calculateChangeRate(50, 100)).toBe(-50); // -50%
  });

  it('前週と同じ場合、0%を返す', () => {
    expect(gasFunctions.calculateChangeRate(100, 100)).toBe(0);
    expect(gasFunctions.calculateChangeRate(50, 50)).toBe(0);
  });

  it('前週が0の場合、今週が0より大きければ100%を返す', () => {
    expect(gasFunctions.calculateChangeRate(100, 0)).toBe(100);
    expect(gasFunctions.calculateChangeRate(50, 0)).toBe(100);
  });

  it('前週が0で今週も0の場合、0%を返す', () => {
    expect(gasFunctions.calculateChangeRate(0, 0)).toBe(0);
  });

  it('小数点第2位まで正確に計算する', () => {
    expect(gasFunctions.calculateChangeRate(33.33, 100)).toBe(-66.67);
    expect(gasFunctions.calculateChangeRate(133.33, 100)).toBe(33.33);
  });
});

describe('formatChangeRate', () => {
  it('正の値の場合、+記号を付ける', () => {
    expect(gasFunctions.formatChangeRate(10)).toBe('+10%');
    expect(gasFunctions.formatChangeRate(50)).toBe('+50%');
    expect(gasFunctions.formatChangeRate(0.5)).toBe('+0.5%');
  });

  it('負の値の場合、-記号を付ける', () => {
    expect(gasFunctions.formatChangeRate(-10)).toBe('-10%');
    expect(gasFunctions.formatChangeRate(-50)).toBe('-50%');
    expect(gasFunctions.formatChangeRate(-0.5)).toBe('-0.5%');
  });

  it('0の場合、0%を返す', () => {
    expect(gasFunctions.formatChangeRate(0)).toBe('0%');
  });
});

describe('getPreviousWeekData', () => {
  it('前週のデータを正しく取得する', () => {
    const mockSheet = createMockSheet();
    const mockSpreadsheet = {
      getSheetByName: vi.fn(() => mockSheet),
    };
    
    SpreadsheetApp.getActiveSpreadsheet = vi.fn(() => mockSpreadsheet as any);
    
    // テストデータを設定（2週間分のデータ）
    const currentWeekStart = new Date('2026-01-08'); // 今週の開始日（月曜日）
    const previousWeekStart = new Date('2026-01-01'); // 前週の開始日
    
    const mockData = [
      ['週開始日', '週終了日', 'Note記事数', 'WordPress記事数', 'GASテンプレート数', 'Note有料記事収益', 'WordPressアフィリエイト収益', 'GAS販売収益', '今週の合計収益', '累計収益', '目標までの残り', '作業時間'],
      [previousWeekStart, '2026-01-07', 2, 2, 0, 6000, 9000, 0, 15000, 50000, 30000, 15],
      [currentWeekStart, '2026-01-14', 3, 3, 1, 8000, 12000, 0, 20000, 70000, 10000, 18],
    ];
    
    (mockSheet.getDataRange as any).mockReturnValue({
      getValues: () => mockData,
    });
    
    const result = gasFunctions.getPreviousWeekData(mockSpreadsheet, currentWeekStart);
    
    expect(result).toBeDefined();
    expect(result.noteArticles).toBe(2);
    expect(result.wpArticles).toBe(2);
    expect(result.totalRevenue).toBe(15000);
    expect(result.workTime).toBe(15);
  });

  it('前週のデータが存在しない場合、nullを返す', () => {
    const mockSheet = createMockSheet();
    const mockSpreadsheet = {
      getSheetByName: vi.fn(() => mockSheet),
    };
    
    SpreadsheetApp.getActiveSpreadsheet = vi.fn(() => mockSpreadsheet as any);
    
    // テストデータを設定（1週間分のデータのみ）
    const currentWeekStart = new Date('2026-01-08');
    const mockData = [
      ['週開始日', '週終了日', 'Note記事数', 'WordPress記事数', 'GASテンプレート数', 'Note有料記事収益', 'WordPressアフィリエイト収益', 'GAS販売収益', '今週の合計収益', '累計収益', '目標までの残り', '作業時間'],
      [currentWeekStart, '2026-01-14', 3, 3, 1, 8000, 12000, 0, 20000, 70000, 10000, 18],
    ];
    
    (mockSheet.getDataRange as any).mockReturnValue({
      getValues: () => mockData,
    });
    
    const result = gasFunctions.getPreviousWeekData(mockSpreadsheet, currentWeekStart);
    
    expect(result).toBeNull();
  });

  it('シートが存在しない場合、nullを返す', () => {
    const mockSpreadsheet = {
      getSheetByName: vi.fn(() => null),
    };
    
    SpreadsheetApp.getActiveSpreadsheet = vi.fn(() => mockSpreadsheet as any);
    
    const result = gasFunctions.getPreviousWeekData(mockSpreadsheet, new Date('2026-01-08'));
    
    expect(result).toBeNull();
  });
});

describe('compareWithPreviousWeek', () => {
  it('前週のデータがある場合、増減率を正しく計算する', () => {
    const currentWeekData = {
      noteArticles: 3,
      wpArticles: 3,
      gasTemplates: 1,
      totalRevenue: 20000,
      workTime: 18,
    };

    const previousWeekData = {
      noteArticles: 2,
      wpArticles: 2,
      gasTemplates: 0,
      totalRevenue: 15000,
      workTime: 15,
    };

    const result = gasFunctions.compareWithPreviousWeek(currentWeekData, previousWeekData);

    expect(result.hasPreviousData).toBe(true);
    expect(result.noteArticlesChange).toBe(50); // (3-2)/2*100 = 50%
    expect(result.wpArticlesChange).toBe(50); // (3-2)/2*100 = 50%
    expect(result.gasTemplatesChange).toBe(100); // (1-0)/0*100 = 100% (前週が0の場合)
    expect(result.revenueChange).toBeCloseTo(33.33, 2); // (20000-15000)/15000*100 ≈ 33.33%
    expect(result.workTimeChange).toBe(20); // (18-15)/15*100 = 20%
  });

  it('前週のデータがない場合、すべて0を返す', () => {
    const currentWeekData = {
      noteArticles: 3,
      wpArticles: 3,
      gasTemplates: 1,
      totalRevenue: 20000,
      workTime: 18,
    };

    const result = gasFunctions.compareWithPreviousWeek(currentWeekData, null);

    expect(result.hasPreviousData).toBe(false);
    expect(result.noteArticlesChange).toBe(0);
    expect(result.wpArticlesChange).toBe(0);
    expect(result.gasTemplatesChange).toBe(0);
    expect(result.revenueChange).toBe(0);
    expect(result.workTimeChange).toBe(0);
  });

  it('前週より減少した場合、負の値を返す', () => {
    const currentWeekData = {
      noteArticles: 1,
      wpArticles: 1,
      gasTemplates: 0,
      totalRevenue: 10000,
      workTime: 10,
    };

    const previousWeekData = {
      noteArticles: 2,
      wpArticles: 2,
      gasTemplates: 1,
      totalRevenue: 15000,
      workTime: 15,
    };

    const result = gasFunctions.compareWithPreviousWeek(currentWeekData, previousWeekData);

    expect(result.hasPreviousData).toBe(true);
    expect(result.noteArticlesChange).toBe(-50); // (1-2)/2*100 = -50%
    expect(result.wpArticlesChange).toBe(-50); // (1-2)/2*100 = -50%
    expect(result.revenueChange).toBeCloseTo(-33.33, 2); // (10000-15000)/15000*100 ≈ -33.33%
    expect(result.workTimeChange).toBeCloseTo(-33.33, 2); // (10-15)/15*100 ≈ -33.33%
  });
});

describe('analyzeRevenueTrend', () => {
  it('増加傾向を正しく判定する', () => {
    const revenueArray = [
      { revenue: 10000 },
      { revenue: 12000 },
      { revenue: 14000 },
      { revenue: 16000 },
    ];
    
    const result = gasFunctions.analyzeRevenueTrend(revenueArray);
    expect(result).toBe('increasing');
  });

  it('減少傾向を正しく判定する', () => {
    const revenueArray = [
      { revenue: 16000 },
      { revenue: 14000 },
      { revenue: 12000 },
      { revenue: 10000 },
    ];
    
    const result = gasFunctions.analyzeRevenueTrend(revenueArray);
    expect(result).toBe('decreasing');
  });

  it('横ばいを正しく判定する', () => {
    const revenueArray = [
      { revenue: 10000 },
      { revenue: 10500 },
      { revenue: 9500 },
      { revenue: 10000 },
    ];
    
    const result = gasFunctions.analyzeRevenueTrend(revenueArray);
    expect(result).toBe('stable');
  });

  it('データが少ない場合、stableを返す', () => {
    expect(gasFunctions.analyzeRevenueTrend([{ revenue: 10000 }])).toBe('stable');
    expect(gasFunctions.analyzeRevenueTrend([])).toBe('stable');
  });
});

describe('analyzeWorkTimeTrend', () => {
  it('増加傾向を正しく判定する', () => {
    const workTimeArray = [
      { workTime: 10 },
      { workTime: 12 },
      { workTime: 14 },
      { workTime: 16 },
    ];
    
    const result = gasFunctions.analyzeWorkTimeTrend(workTimeArray);
    expect(result).toBe('increasing');
  });

  it('減少傾向を正しく判定する', () => {
    const workTimeArray = [
      { workTime: 16 },
      { workTime: 14 },
      { workTime: 12 },
      { workTime: 10 },
    ];
    
    const result = gasFunctions.analyzeWorkTimeTrend(workTimeArray);
    expect(result).toBe('decreasing');
  });

  it('横ばいを正しく判定する', () => {
    const workTimeArray = [
      { workTime: 15 },
      { workTime: 16 },
      { workTime: 14 },
      { workTime: 15 },
    ];
    
    const result = gasFunctions.analyzeWorkTimeTrend(workTimeArray);
    expect(result).toBe('stable');
  });
});

describe('formatTrend', () => {
  it('トレンドを日本語で正しく表示する', () => {
    expect(gasFunctions.formatTrend('increasing')).toBe('📈 増加傾向');
    expect(gasFunctions.formatTrend('decreasing')).toBe('📉 減少傾向');
    expect(gasFunctions.formatTrend('stable')).toBe('➡️ 横ばい');
    expect(gasFunctions.formatTrend('unknown')).toBe('➡️ 横ばい'); // 未知の値の場合
  });
});

describe('getPastWeeksData', () => {
  it('過去N週間のデータを正しく取得する', () => {
    const mockSheet = createMockSheet();
    const mockSpreadsheet = {
      getSheetByName: vi.fn(() => mockSheet),
    };
    
    SpreadsheetApp.getActiveSpreadsheet = vi.fn(() => mockSpreadsheet as any);
    
    // テストデータを設定（4週間分のデータ）
    const mockData = [
      ['週開始日', '週終了日', 'Note記事数', 'WordPress記事数', 'GASテンプレート数', 'Note有料記事収益', 'WordPressアフィリエイト収益', 'GAS販売収益', '今週の合計収益', '累計収益', '目標までの残り', '作業時間'],
      ['2026-01-01', '2026-01-07', 2, 2, 0, 6000, 9000, 0, 15000, 50000, 30000, 15],
      ['2026-01-08', '2026-01-14', 3, 3, 1, 8000, 12000, 0, 20000, 70000, 10000, 18],
      ['2026-01-15', '2026-01-21', 2, 2, 0, 7000, 11000, 0, 18000, 88000, -8000, 16],
      ['2026-01-22', '2026-01-28', 3, 3, 1, 9000, 13000, 0, 22000, 110000, -30000, 20],
    ];
    
    (mockSheet.getDataRange as any).mockReturnValue({
      getValues: () => mockData,
    });
    
    const result = gasFunctions.getPastWeeksData(mockSpreadsheet, 4);
    
    expect(result).toHaveLength(4);
    expect(result[0].revenue).toBe(15000);
    expect(result[1].revenue).toBe(20000);
    expect(result[2].revenue).toBe(18000);
    expect(result[3].revenue).toBe(22000);
  });

  it('シートが存在しない場合、空配列を返す', () => {
    const mockSpreadsheet = {
      getSheetByName: vi.fn(() => null),
    };
    
    SpreadsheetApp.getActiveSpreadsheet = vi.fn(() => mockSpreadsheet as any);
    
    const result = gasFunctions.getPastWeeksData(mockSpreadsheet, 4);
    
    expect(result).toEqual([]);
  });

  it('データが少ない場合、存在するデータのみを返す', () => {
    const mockSheet = createMockSheet();
    const mockSpreadsheet = {
      getSheetByName: vi.fn(() => mockSheet),
    };
    
    SpreadsheetApp.getActiveSpreadsheet = vi.fn(() => mockSpreadsheet as any);
    
    // テストデータを設定（2週間分のデータのみ）
    const mockData = [
      ['週開始日', '週終了日', 'Note記事数', 'WordPress記事数', 'GASテンプレート数', 'Note有料記事収益', 'WordPressアフィリエイト収益', 'GAS販売収益', '今週の合計収益', '累計収益', '目標までの残り', '作業時間'],
      ['2026-01-01', '2026-01-07', 2, 2, 0, 6000, 9000, 0, 15000, 50000, 30000, 15],
      ['2026-01-08', '2026-01-14', 3, 3, 1, 8000, 12000, 0, 20000, 70000, 10000, 18],
    ];
    
    (mockSheet.getDataRange as any).mockReturnValue({
      getValues: () => mockData,
    });
    
    const result = gasFunctions.getPastWeeksData(mockSpreadsheet, 4);
    
    expect(result).toHaveLength(2); // 2週間分のみ
  });
});

describe('analyzeTrends', () => {
  it('トレンド分析を正しく実行する', () => {
    const mockSheet = createMockSheet();
    const mockSpreadsheet = {
      getSheetByName: vi.fn(() => mockSheet),
    };
    
    SpreadsheetApp.getActiveSpreadsheet = vi.fn(() => mockSpreadsheet as any);
    
    // テストデータを設定（増加傾向のデータ）
    const mockData = [
      ['週開始日', '週終了日', 'Note記事数', 'WordPress記事数', 'GASテンプレート数', 'Note有料記事収益', 'WordPressアフィリエイト収益', 'GAS販売収益', '今週の合計収益', '累計収益', '目標までの残り', '作業時間'],
      ['2026-01-01', '2026-01-07', 2, 2, 0, 6000, 9000, 0, 10000, 50000, 30000, 10],
      ['2026-01-08', '2026-01-14', 3, 3, 1, 8000, 12000, 0, 12000, 62000, 18000, 12],
      ['2026-01-15', '2026-01-21', 2, 2, 0, 7000, 11000, 0, 14000, 76000, 4000, 14],
      ['2026-01-22', '2026-01-28', 3, 3, 1, 9000, 13000, 0, 16000, 92000, -12000, 16],
    ];
    
    (mockSheet.getDataRange as any).mockReturnValue({
      getValues: () => mockData,
    });
    
    const result = gasFunctions.analyzeTrends(mockSpreadsheet, 4);
    
    expect(result.hasEnoughData).toBe(true);
    expect(result.weeksAnalyzed).toBe(4);
    expect(result.revenueTrend).toBe('increasing');
    expect(result.workTimeTrend).toBe('increasing');
  });

  it('データが少ない場合、stableを返す', () => {
    const mockSheet = createMockSheet();
    const mockSpreadsheet = {
      getSheetByName: vi.fn(() => mockSheet),
    };
    
    SpreadsheetApp.getActiveSpreadsheet = vi.fn(() => mockSpreadsheet as any);
    
    // テストデータを設定（1週間分のデータのみ）
    const mockData = [
      ['週開始日', '週終了日', 'Note記事数', 'WordPress記事数', 'GASテンプレート数', 'Note有料記事収益', 'WordPressアフィリエイト収益', 'GAS販売収益', '今週の合計収益', '累計収益', '目標までの残り', '作業時間'],
      ['2026-01-01', '2026-01-07', 2, 2, 0, 6000, 9000, 0, 15000, 50000, 30000, 15],
    ];
    
    (mockSheet.getDataRange as any).mockReturnValue({
      getValues: () => mockData,
    });
    
    const result = gasFunctions.analyzeTrends(mockSpreadsheet, 4);
    
    expect(result.hasEnoughData).toBe(false);
    expect(result.revenueTrend).toBe('stable');
    expect(result.workTimeTrend).toBe('stable');
  });
});

