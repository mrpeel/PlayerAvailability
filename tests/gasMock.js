/**
 * Jest setup: mock Google Apps Script globals so adapter/integration tests
 * (SheetLogic.gs / Setup.gs) can run locally. Loaded via jest setupFiles.
 */
global.SpreadsheetApp = {
  getActiveSpreadsheet: jest.fn(),
  getUi: jest.fn(() => ({
    alert: jest.fn(),
    ButtonSet: { OK: 0 }
  })),
  newDataValidation: jest.fn().mockReturnValue({
    requireValueInList: jest.fn().mockReturnThis(),
    requireValueInRange: jest.fn().mockReturnThis(),
    build: jest.fn()
  }),
  BorderStyle: { SOLID: 'SOLID' }
};

global.ContentService = {
  createTextOutput: jest.fn().mockReturnValue({
    setMimeType: jest.fn().mockReturnThis()
  }),
  MimeType: { JSON: 'JSON' }
};

global.DriveApp = {
  getFolderById: jest.fn()
};
