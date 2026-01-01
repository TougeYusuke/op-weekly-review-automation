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

