/**
 * Jest setup: mock Google Apps Script globals so adapter/integration tests
 * (SheetLogic.gs / Setup.gs) can run locally. Loaded via jest setupFiles.
 */
var scriptPropertiesStore = {};
global.PropertiesService = {
  getScriptProperties: jest.fn(() => ({
    getProperty: jest.fn(k => scriptPropertiesStore[k] || null),
    setProperty: jest.fn((k, v) => { scriptPropertiesStore[k] = String(v); }),
    getProperties: jest.fn(() => scriptPropertiesStore)
  }))
};

global.SlidesApp = {
  openById: jest.fn(),
  PageElementType: {
    GROUP: 'GROUP',
    SHAPE: 'SHAPE',
    IMAGE: 'IMAGE',
    TABLE: 'TABLE'
  },
  ShapeType: {
    ELLIPSE: 'ELLIPSE',
    RECTANGLE: 'RECTANGLE',
    TEXT_BOX: 'TEXT_BOX'
  },
  ContentAlignment: {
    TOP: 'TOP',
    MIDDLE: 'MIDDLE',
    BOTTOM: 'BOTTOM',
    JUSTIFIED: 'JUSTIFIED'
  },
  ParagraphAlignment: {
    START: 'START',
    CENTER: 'CENTER',
    END: 'END',
    JUSTIFIED: 'JUSTIFIED'
  }
};

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
  getFolderById: jest.fn(),
  getFileById: jest.fn()
};

global.HtmlService = {
  createHtmlOutput: jest.fn(() => ({
    setWidth: jest.fn().mockReturnThis(),
    setHeight: jest.fn().mockReturnThis()
  }))
};

global.Utilities = {
  base64Decode: jest.fn(str => Buffer.from(str, 'base64')),
  base64Encode: jest.fn(buf => Buffer.from(buf).toString('base64')),
  newBlob: jest.fn((bytes, contentType, name) => ({
    getBytes: () => bytes,
    getContentType: () => contentType,
    getName: () => name
  }))
};
