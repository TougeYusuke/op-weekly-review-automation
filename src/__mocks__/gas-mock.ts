/**
 * Google Apps Scriptのモック実装
 * テスト時にGAS APIをモック化するためのファイル
 * 週次レビュー自動化ツール用に拡張
 */

import { vi } from 'vitest';

// SpreadsheetAppのモック
export const SpreadsheetApp = {
  openById: vi.fn(),
  getActiveSpreadsheet: vi.fn(() => ({
    getSheetByName: vi.fn(),
    insertSheet: vi.fn(),
  })),
  getSpreadsheetTimeZone: vi.fn(() => 'Asia/Tokyo'),
};

// Sheetのモック
export const createMockSheet = () => {
  // getRange()が返すRangeオブジェクトを作成（メソッドチェーン対応）
  const createMockRange = () => {
    const range = {
      setValues: vi.fn().mockReturnThis(),
      setFontSize: vi.fn().mockReturnThis(),
      setFontWeight: vi.fn().mockReturnThis(),
      setBackground: vi.fn().mockReturnThis(),
      setFontColor: vi.fn().mockReturnThis(),
    };
    return range;
  };

  return {
    getDataRange: vi.fn(() => ({
      getValues: vi.fn(() => []),
    })),
    getRange: vi.fn(() => createMockRange()),
    appendRow: vi.fn(),
    clear: vi.fn(),
    autoResizeColumns: vi.fn(),
    autoResizeColumn: vi.fn(),
  };
};

// Loggerのモック
export const Logger = {
  log: vi.fn(),
};

// Utilitiesのモック
export const Utilities = {
  formatDate: vi.fn((date: Date, timezone: string, format: string) => {
    const d = new Date(date);
    const year = d.getFullYear();
    const month = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
  }),
  newBlob: vi.fn(),
  base64Encode: vi.fn(),
};

// ContentServiceのモック
export const ContentService = {
  createTextOutput: vi.fn((content: string) => ({
    setMimeType: vi.fn(),
    setDownloadAsFile: vi.fn(),
  })),
  MimeType: {
    JSON: 'application/json',
    CSV: 'text/csv',
  },
};

// DriveAppのモック
export const DriveApp = {
  getRootFolder: vi.fn(),
  getFolderById: vi.fn(),
  getFilesByName: vi.fn(),
};

// Sessionのモック
export const Session = {
  getScriptTimeZone: vi.fn(() => 'Asia/Tokyo'),
};

// MimeTypeのモック
export const MimeType = {
  CSV: 'text/csv',
};

// SpreadsheetAppのUiのモック
export const Ui = {
  alert: vi.fn(),
};

