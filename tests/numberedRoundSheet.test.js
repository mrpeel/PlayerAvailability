const fs = require('fs');
const path = require('path');
const vm = require('vm');
const logic = require('../src/logic.js');

const {
  getRoundFrames,
  getSlotNumbersForFrame,
  getRoundColumnMapping,
  isNumberedRoundSheetLayout
} = logic;

describe('Numbered Round Selection Sheet Logic', () => {
  describe('getSlotNumbersForFrame', () => {
    test('returns 1-12 for 1st Eleven and 2nd Eleven', () => {
      const frames = getRoundFrames(false);
      const firstXI = frames[0];
      const secondXI = frames[1];

      expect(getSlotNumbersForFrame(firstXI)).toEqual([1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]);
      expect(getSlotNumbersForFrame(secondXI)).toEqual([1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]);
    });

    test('returns 1-13 for 3rd, 4th, and 5th Eleven', () => {
      const frames = getRoundFrames(false);
      const thirdXI = frames[2];
      const fourthXI = frames[3];
      const fifthXI = frames[4];

      expect(getSlotNumbersForFrame(thirdXI)).toEqual([1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13]);
      expect(getSlotNumbersForFrame(fourthXI)).toEqual([1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13]);
      expect(getSlotNumbersForFrame(fifthXI)).toEqual([1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12, 13]);
    });

    test('returns 1-12 for T20 1st and 2nd Eleven', () => {
      const t20Frames = getRoundFrames(true);
      expect(getSlotNumbersForFrame(t20Frames[0])).toEqual([1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]);
      expect(getSlotNumbersForFrame(t20Frames[1])).toEqual([1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11, 12]);
    });
  });

  describe('getRoundColumnMapping', () => {
    test('returns correct 1-based column indices for numbered layout', () => {
      const mapping = getRoundColumnMapping(true);
      expect(mapping.numberCol).toBe(1);        // Col A
      expect(mapping.roleCol).toBe(2);          // Col B
      expect(mapping.playerCol).toBe(3);        // Col C
      expect(mapping.spacer1Col).toBe(4);       // Col D
      expect(mapping.dynamicPoolCol).toBe(5);   // Col E
      expect(mapping.spacer2Col).toBe(6);       // Col F
      expect(mapping.availStartCol).toBe(7);    // Col G (ProfileID)
      expect(mapping.availNameCol).toBe(8);     // Col H (Player Name)
      expect(mapping.availNotesCol).toBe(9);    // Col I (Notes)
      expect(mapping.availSwitcherCol).toBe(10);// Col J (Switcher)
      expect(mapping.spacer3Col).toBe(11);      // Col K
      expect(mapping.unavailStartCol).toBe(12); // Col L (ProfileID)
      expect(mapping.unavailNameCol).toBe(13);  // Col M (Player Name)
      expect(mapping.unavailNotesCol).toBe(14); // Col N (Notes)
      expect(mapping.unavailSwitcherCol).toBe(15);// Col O (Switcher)
      expect(mapping.spacer4Col).toBe(16);      // Col P
      expect(mapping.unknownStartCol).toBe(17); // Col Q (ProfileID)
      expect(mapping.unknownNameCol).toBe(18);  // Col R (Player Name)
      expect(mapping.unknownNotesCol).toBe(19); // Col S (Notes)
      expect(mapping.unknownSwitcherCol).toBe(20);// Col T (Switcher)
      expect(mapping.spacer5Col).toBe(21);      // Col U
      expect(mapping.whatsappMsgCol).toBe(22);  // Col V
      expect(mapping.whatsappShameCol).toBe(23);// Col W
      expect(mapping.spacer6Col).toBe(24);      // Col X
      expect(mapping.helperStartCol).toBe(28);  // Col AB
    });

    test('returns legacy column indices when isNumbered is false', () => {
      const legacy = getRoundColumnMapping(false);
      expect(legacy.numberCol).toBeNull();
      expect(legacy.roleCol).toBe(1);           // Col A
      expect(legacy.playerCol).toBe(2);         // Col B
      expect(legacy.spacer1Col).toBe(3);        // Col C
      expect(legacy.dynamicPoolCol).toBe(4);    // Col D
      expect(legacy.availStartCol).toBe(6);     // Col F
      expect(legacy.availNameCol).toBe(7);      // Col G
      expect(legacy.availSwitcherCol).toBe(9);  // Col I
      expect(legacy.unavailStartCol).toBe(11);  // Col K
      expect(legacy.unavailNameCol).toBe(12);   // Col L
      expect(legacy.unavailSwitcherCol).toBe(14);// Col N
      expect(legacy.unknownStartCol).toBe(16);  // Col P
      expect(legacy.unknownNameCol).toBe(17);   // Col Q
      expect(legacy.unknownSwitcherCol).toBe(19);// Col S
      expect(legacy.whatsappMsgCol).toBe(21);   // Col U
      expect(legacy.whatsappShameCol).toBe(22); // Col V
      expect(legacy.helperStartCol).toBe(27);   // Col AA
    });
  });

  describe('isNumberedRoundSheetLayout', () => {
    test('detects numbered layout from 2D sheet data with Col B "Round:"', () => {
      const grid = [
        ['', 'Match Date:', '2026-10-10'],
        ['', '', ''],
        ['', '', ''],
        ['FIRST ELEVEN', '', ''],
        ['', 'Round:', 'Round 2'],
        ['', 'Opponent:', 'Glen Waverley Hawks - 1st XI'],
        ['', 'Venue:', 'Capital Reserve'],
        ['', 'Format:', 'One Day'],
        [1, 'Captain', 'Ross Theunissen']
      ];
      expect(isNumberedRoundSheetLayout(grid)).toBe(true);
    });

    test('detects legacy layout from 2D sheet data with Col A "Round:"', () => {
      const legacyGrid = [
        ['Match Date:', '2026-10-10'],
        ['', ''],
        ['', ''],
        ['FIRST ELEVEN', ''],
        ['Round:', 'Round 2'],
        ['Opponent:', 'Glen Waverley Hawks - 1st XI'],
        ['Venue:', 'Capital Reserve'],
        ['Format:', 'One Day'],
        ['Captain', 'Ross Theunissen']
      ];
      expect(isNumberedRoundSheetLayout(legacyGrid)).toBe(false);
    });

    test('detects numbered layout from key cell object', () => {
      expect(isNumberedRoundSheetLayout({ b5: 'Round:' })).toBe(true);
      expect(isNumberedRoundSheetLayout({ a9: 1 })).toBe(true);
      expect(isNumberedRoundSheetLayout({ b1: 'Match Date:' })).toBe(true);
      expect(isNumberedRoundSheetLayout({ a5: 'Round:' })).toBe(false);
    });
  });

  describe('GAS Adapter Integration (SheetLogic.gs)', () => {
    let gasContext;

    beforeAll(() => {
      gasContext = Object.assign({}, global, logic);
      const sheetLogicCode = fs.readFileSync(path.join(__dirname, '../src/SheetLogic.gs'), 'utf8');
      vm.runInNewContext(sheetLogicCode, gasContext);
    });

    test('isNumberedRoundSheet detects numbered round sheet', () => {
      const mockNumberedSheet = {
        getRange: jest.fn(a1 => {
          if (a1 === 'B5') return { getValue: () => 'Round:' };
          if (a1 === 'A9') return { getValue: () => 1 };
          if (a1 === 'B1') return { getValue: () => 'Match Date:' };
          return { getValue: () => '' };
        })
      };
      expect(gasContext.isNumberedRoundSheet(mockNumberedSheet)).toBe(true);

      const mockLegacySheet = {
        getRange: jest.fn(a1 => {
          if (a1 === 'A5') return { getValue: () => 'Round:' };
          if (a1 === 'B5') return { getValue: () => 'Round 1' };
          if (a1 === 'A9') return { getValue: () => 'Captain' };
          if (a1 === 'B1') return { getValue: () => '2026-10-04' };
          return { getValue: () => '' };
        })
      };
      expect(gasContext.isNumberedRoundSheet(mockLegacySheet)).toBe(false);
    });

    test('resolveRoundStartDate resolves from C1 when Match Date is in C1', () => {
      const mockRoundSheet = {
        getRange: jest.fn(a1 => {
          if (a1 === 'B1') return { getValue: () => 'Match Date:' };
          if (a1 === 'C1') return { getValue: () => '2026-10-10' };
          return { getValue: () => '' };
        })
      };
      const mockSs = {
        getSheetByName: jest.fn(name => {
          if (name === '2026-10-10') return mockRoundSheet;
          return null;
        })
      };

      const resolved = gasContext.resolveRoundStartDate('2026-10-10', mockSs);
      expect(resolved).toBe('10 Oct 26');
    });

    test('deployVerticalRoundSheet creates numbered round layout', () => {
      const cellValues = {};
      const columnWidths = {};
      const merges = [];
      const hiddenCols = [];

      const createRangeMock = (rowOrA1, col, numRows, numCols) => {
        const key = typeof rowOrA1 === 'string'
          ? rowOrA1
          : `R${rowOrA1}C${col}_${numRows || 1}x${numCols || 1}`;
        const rangeMock = {
          setValue: jest.fn(v => { cellValues[key] = v; return rangeMock; }),
          setValues: jest.fn(arr => { cellValues[key] = arr; return rangeMock; }),
          getValue: jest.fn(() => cellValues[key] || ''),
          getValues: jest.fn(() => cellValues[key] || []),
          setFormula: jest.fn(f => { cellValues[key + '_formula'] = f; return rangeMock; }),
          setFontWeight: jest.fn(() => rangeMock),
          setFontStyle: jest.fn(() => rangeMock),
          setFontSize: jest.fn(() => rangeMock),
          setFontFamily: jest.fn(() => rangeMock),
          setBackground: jest.fn(() => rangeMock),
          setFontColor: jest.fn(() => rangeMock),
          setHorizontalAlignment: jest.fn(() => rangeMock),
          setVerticalAlignment: jest.fn(() => rangeMock),
          setNumberFormat: jest.fn(() => rangeMock),
          setWrap: jest.fn(() => rangeMock),
          setBorder: jest.fn(() => rangeMock),
          setDataValidation: jest.fn(() => rangeMock),
          clear: jest.fn(() => rangeMock),
          clearContent: jest.fn(() => rangeMock),
          merge: jest.fn(() => { merges.push(key); return rangeMock; }),
          breakApart: jest.fn(() => rangeMock),
          isPartOfMerge: jest.fn(() => false)
        };
        return rangeMock;
      };

      const mockSheet = {
        getName: () => '2026-10-10',
        getRange: jest.fn(createRangeMock),
        setColumnWidth: jest.fn((c, w) => { columnWidths[c] = w; }),
        hideColumns: jest.fn((c, num) => { hiddenCols.push({ col: c, count: num || 1 }); }),
        getMaxRows: () => 150,
        getMaxColumns: () => 30,
        insertColumnsAfter: jest.fn()
      };

      const mockSs = {
        getSheetByName: jest.fn(name => {
          if (name === 'Fixtures') return null;
          if (name === 'Players') return null;
          return null;
        }),
        insertSheet: jest.fn(() => mockSheet),
        setActiveSheet: jest.fn(),
        getSheets: () => [mockSheet]
      };

      gasContext.deployVerticalRoundSheet(mockSs, '2026-10-10');

      // Verify Column 1 width is 22px (narrow)
      expect(columnWidths[1]).toBe(22);
      expect(columnWidths[2]).toBe(110);
      expect(columnWidths[3]).toBe(190);
      expect(columnWidths[14]).toBe(163); // Col N: Exemption Notes width

      // Verify Col B is "Match Date:", Col C has the target date
      expect(cellValues['B1']).toBe('Match Date:');

      // Verify team banner merges A, B, C (Range R4C1 numCols 3)
      expect(merges).toContain('R4C1_1x3');

      // Verify Col A slot numbers structure for 1st XI (R9C1: 12 rows, 3 cols)
      const firstXIStructure = cellValues['R9C1_12x3'];
      expect(firstXIStructure).toBeDefined();
      expect(firstXIStructure).toHaveLength(12);
      expect(firstXIStructure[0]).toEqual([1, 'Captain', '']);
      expect(firstXIStructure[1]).toEqual([2, 'Vice Captain', '']);
      expect(firstXIStructure[2]).toEqual([3, 'Wicket Keeper', '']);
      expect(firstXIStructure[11]).toEqual([12, 'Player', '']);

      // Verify Col A slot numbers structure for 3rd XI (R45C1: 13 rows, 3 cols)
      const thirdXIStructure = cellValues['R45C1_13x3'];
      expect(thirdXIStructure).toBeDefined();
      expect(thirdXIStructure).toHaveLength(13);
      expect(thirdXIStructure[0]).toEqual([1, 'Captain', '']);
      expect(thirdXIStructure[12]).toEqual([13, 'Player', '']);

      // Verify Dynamic Virtual Pool in Col E
      expect(cellValues['E3']).toBe('AVAILABLE FOR SELECTION');
      expect(cellValues['E5_formula']).toContain('COUNTIF(C:C, H5:H)=0');

      // Verify ProfileID columns hidden at 7, 12, 17
      const hiddenColIndices = hiddenCols.map(h => h.col);
      expect(hiddenColIndices).toContain(7);
      expect(hiddenColIndices).toContain(12);
      expect(hiddenColIndices).toContain(17);
    });
  });
});
