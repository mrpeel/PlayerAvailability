const fs = require('fs');
const path = require('path');
const { formatRoundDate, extractSlotNumberForTeam } = require('../src/logic.js');

describe('Google Slide Sync - ROUND_DATE property', () => {
  describe('formatRoundDate format and edge cases', () => {
    test('formats round start date "2026-10-03" to "3 Oct 26"', () => {
      expect(formatRoundDate('2026-10-03')).toBe('3 Oct 26');
    });

    test('formats round start date without leading zeros for day', () => {
      expect(formatRoundDate('2026-04-01')).toBe('1 Apr 26');
      expect(formatRoundDate('2026-12-09')).toBe('9 Dec 26');
    });

    test('formats multi-day fixture date string using start date', () => {
      expect(formatRoundDate('2026-10-03, 2026-10-10')).toBe('3 Oct 26');
    });

    test('formats DD/MM/YYYY date strings', () => {
      expect(formatRoundDate('03/10/2026')).toBe('3 Oct 26');
      expect(formatRoundDate('3/10/2026')).toBe('3 Oct 26');
    });

    test('formats JS Date objects correctly', () => {
      const d = new Date(2026, 9, 3); // Oct 3, 2026
      expect(formatRoundDate(d)).toBe('3 Oct 26');
    });

    test('handles empty and invalid inputs gracefully', () => {
      expect(formatRoundDate('')).toBe('');
      expect(formatRoundDate(null)).toBe('');
      expect(formatRoundDate(undefined)).toBe('');
      expect(formatRoundDate('invalid')).toBe('');
    });
  });

  describe('extractSlotNumberForTeam interaction', () => {
    test('does not match ROUND_DATE as a player slot number', () => {
      expect(extractSlotNumberForTeam('ROUND_DATE', '1ST')).toBeNull();
      expect(extractSlotNumberForTeam('1ST_ROUND_DATE', '1ST')).toBeNull();
      expect(extractSlotNumberForTeam('T20_1ST_ROUND_DATE', 'T20_1ST')).toBeNull();
    });
  });

  describe('Slide element matching for ROUND_DATE', () => {
    // Load SheetLogic.gs into Jest global context
    beforeAll(() => {
      const vm = require('vm');
      const logic = require('../src/logic.js');
      Object.assign(global, logic);
      const sheetLogicCode = fs.readFileSync(path.join(__dirname, '../src/SheetLogic.gs'), 'utf8');
      vm.runInNewContext(sheetLogicCode, global);
    });

    test('resolveRoundStartDate resolves from activeRound date string', () => {
      const formatted = resolveRoundStartDate('2026-10-03');
      expect(formatted).toBe('3 Oct 26');
    });

    test('resolveRoundStartDate resolves from active round sheet B1 cell', () => {
      const mockRoundSheet = {
        getRange: jest.fn(a1 => {
          if (a1 === 'B1') return { getValue: () => '2026-10-03' };
          return { getValue: () => '' };
        })
      };
      const mockSs = {
        getSheetByName: jest.fn(name => {
          if (name === 'Round 1') return mockRoundSheet;
          return null;
        })
      };

      const formatted = resolveRoundStartDate('Round 1', mockSs);
      expect(formatted).toBe('3 Oct 26');
    });

    test('resolveRoundStartDate resolves from Fixtures sheet', () => {
      const mockFixSheet = {
        getLastRow: () => 3,
        getDataRange: () => ({
          getValues: () => [
            ['Game Date', '1st Round', '1st Opponent', '1st Venue', '1st Format'],
            ['2026-10-03', 'Round 1', 'Mitcham', 'Kalang Park', 'Two Day'],
            ['2026-10-17', 'Round 2', 'Box Hill', 'City Oval', 'One Day']
          ]
        })
      };
      const mockSs = {
        getSheetByName: jest.fn(name => {
          if (name === 'Fixtures') return mockFixSheet;
          return null;
        })
      };

      const formatted = resolveRoundStartDate('Round 1', mockSs);
      expect(formatted).toBe('3 Oct 26');
    });

    test('syncPresentationStagingToSlides updates elements with description ROUND_DATE', () => {
      let roundDateShapeText = '';
      const mockRoundDateShape = {
        getPageElementType: () => 'SHAPE',
        getTitle: () => '',
        getDescription: () => 'ROUND_DATE',
        asShape: () => ({
          getText: () => ({
            setText: val => { roundDateShapeText = val; },
            asString: () => roundDateShapeText
          })
        })
      };

      let roundOpponentShapeText = '';
      const mockOpponentShape = {
        getPageElementType: () => 'SHAPE',
        getTitle: () => '1ST_ROUND_OPPONENT',
        getDescription: () => '',
        asShape: () => ({
          getText: () => ({
            setText: val => { roundOpponentShapeText = val; },
            asString: () => roundOpponentShapeText
          })
        })
      };

      let coverRoundDateText = '';
      const mockCoverRoundDateShape = {
        getPageElementType: () => 'SHAPE',
        getTitle: () => '',
        getDescription: () => 'ROUND_DATE',
        asShape: () => ({
          getText: () => ({
            setText: val => { coverRoundDateText = val; },
            asString: () => coverRoundDateText
          })
        })
      };

      const mockCoverSlide = {
        getObjectId: () => 'cover_slide',
        getPageElements: () => [mockCoverRoundDateShape]
      };

      let groupedDateText = '';
      const mockGroupChildShape = {
        getPageElementType: () => 'SHAPE',
        getTitle: () => '',
        getDescription: () => 'ROUND_DATE',
        asShape: () => ({
          getText: () => ({
            setText: val => { groupedDateText = val; },
            asString: () => groupedDateText
          })
        })
      };

      const mockGroup = {
        getPageElementType: () => 'GROUP',
        getTitle: () => 'HEADER_GROUP',
        getDescription: () => '',
        asGroup: () => ({
          getChildren: () => [mockGroupChildShape]
        })
      };

      let tableDateText = '';
      const mockTable = {
        getPageElementType: () => 'TABLE',
        getTitle: () => 'MATCH_INFO_TABLE',
        getDescription: () => '',
        asTable: () => ({
          getNumRows: () => 1,
          getNumColumns: () => 2,
          getCell: (r, c) => ({
            getText: () => ({
              asString: () => (c === 0 ? 'Round Date' : tableDateText),
              setText: val => { tableDateText = val; }
            })
          })
        })
      };

      const mockSlide = {
        getObjectId: () => 'slide_1',
        getPageElements: () => [mockRoundDateShape, mockOpponentShape, mockGroup, mockTable]
      };

      const mockPres = {
        getSlides: () => [mockCoverSlide, mockSlide]
      };

      global.SlidesApp.openById = jest.fn(() => mockPres);
      global.PropertiesService.getScriptProperties = jest.fn(() => ({
        getProperty: jest.fn(k => {
          if (k === 'SLIDES_PRESENTATION_ID') return 'mock_pres_id';
          return null;
        }),
        setProperty: jest.fn()
      }));

      const mockStaging = {
        getName: () => 'Presentation_Staging',
        clear: jest.fn(),
        getRange: jest.fn((r, c) => {
          if (r === 'B1' || (r === 1 && c === 2)) return { getValue: () => '2026-10-03', setValue: jest.fn().mockReturnThis(), setNumberFormat: jest.fn().mockReturnThis(), setBackground: jest.fn().mockReturnThis(), setFontColor: jest.fn().mockReturnThis(), setFontWeight: jest.fn().mockReturnThis(), setHorizontalAlignment: jest.fn().mockReturnThis() };
          if (r === 'A1' || (r === 1 && c === 1)) return { getValue: () => 'Round to present', setValue: jest.fn().mockReturnThis(), setFontWeight: jest.fn().mockReturnThis(), setBackground: jest.fn().mockReturnThis(), setFontColor: jest.fn().mockReturnThis(), setHorizontalAlignment: jest.fn().mockReturnThis() };
          // Metadata rows for 1ST (start: 4): row 5 = Round, 6 = Opponent, 7 = Venue, 8 = Format
          if (r === 5 && c === 2) return { getValue: () => 'Round 1' };
          if (r === 6 && c === 2) return { getValue: () => 'Mitcham' };
          if (r === 7 && c === 2) return { getValue: () => 'Kalang Park' };
          if (r === 8 && c === 2) return { getValue: () => 'Two Day' };
          // Default range mock
          return {
            getValue: () => '',
            setValue: jest.fn().mockReturnThis(),
            setFormula: jest.fn().mockReturnThis(),
            setFontWeight: jest.fn().mockReturnThis(),
            setFontStyle: jest.fn().mockReturnThis(),
            setBackground: jest.fn().mockReturnThis(),
            setFontColor: jest.fn().mockReturnThis(),
            setHorizontalAlignment: jest.fn().mockReturnThis(),
            setBorder: jest.fn().mockReturnThis(),
            setValues: jest.fn().mockReturnThis(),
            merge: jest.fn().mockReturnThis()
          };
        }),
        setColumnWidth: jest.fn(),
        hideColumns: jest.fn()
      };

      const mockSs = {
        getSheetByName: jest.fn(name => {
          if (name === 'Presentation_Staging') return mockStaging;
          if (name === '2026-10-03') return { getRange: () => ({ getValue: () => '' }) };
          return null;
        }),
        getSheets: () => [mockStaging]
      };

      global.SpreadsheetApp.getActiveSpreadsheet = jest.fn(() => mockSs);
      global.SpreadsheetApp.flush = jest.fn();

      const result = syncPresentationStagingToSlides('2026-10-03', mockSs);
      expect(result.success).toBe(true);
      // The slide element with description ROUND_DATE must be replaced with "3 Oct 26"
      expect(roundDateShapeText).toBe('3 Oct 26');
      // Grouped child with description ROUND_DATE is replaced
      expect(groupedDateText).toBe('3 Oct 26');
      // Table cell next to "Round Date" is replaced
      expect(tableDateText).toBe('3 Oct 26');
      // Cover slide element with description ROUND_DATE is replaced
      expect(coverRoundDateText).toBe('3 Oct 26');
    });
  });
});
