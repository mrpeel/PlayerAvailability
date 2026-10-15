const fs = require('fs');
const path = require('path');
const vm = require('vm');
const logic = require('../src/logic.js');
const {
  getDayRestriction,
  parseRoundNote,
  buildRoundNote,
  resolveResponseFromRoundNote,
  AVAILABILITY_LOG_HEADERS,
  isLegacyAvailabilityLogHeader,
  buildAvailabilityLogRow,
  migrateLegacyAvailabilityLogRows
} = logic;

describe('Partial-day availability (Day 1 Only / Day 2 Only)', () => {
  test('getDayRestriction recognises partial-day responses only', () => {
    expect(getDayRestriction('Day 1 Only')).toBe('Day 1 Only');
    expect(getDayRestriction('day 2 only')).toBe('Day 2 Only');
    expect(getDayRestriction('Both Days')).toBe('');
    expect(getDayRestriction('Available')).toBe('');
    expect(getDayRestriction('Unavailable')).toBe('');
    expect(getDayRestriction(undefined)).toBe('');
  });

  test('buildRoundNote prefixes the restriction onto the player note', () => {
    expect(buildRoundNote('Day 1 Only', '')).toBe('Day 1 Only');
    expect(buildRoundNote('Day 2 Only', 'arriving 1pm')).toBe('Day 2 Only - arriving 1pm');
    expect(buildRoundNote('Both Days', 'arriving 1pm')).toBe('arriving 1pm');
    expect(buildRoundNote('Unavailable', '')).toBe('');
  });

  test('buildRoundNote never doubles or keeps a stale prefix', () => {
    expect(buildRoundNote('Day 1 Only', 'Day 1 Only - late')).toBe('Day 1 Only - late');
    expect(buildRoundNote('Day 2 Only', 'Day 1 Only - late')).toBe('Day 2 Only - late');
    expect(buildRoundNote('Both Days', 'Day 1 Only')).toBe('');
  });

  test('parseRoundNote splits restriction and user note', () => {
    expect(parseRoundNote('Day 2 Only - arriving 1pm')).toEqual({ dayRestriction: 'Day 2 Only', userNote: 'arriving 1pm' });
    expect(parseRoundNote('Day 1 Only')).toEqual({ dayRestriction: 'Day 1 Only', userNote: '' });
    expect(parseRoundNote('Injured: hamstring')).toEqual({ dayRestriction: '', userNote: 'Injured: hamstring' });
    expect(parseRoundNote('')).toEqual({ dayRestriction: '', userNote: '' });
  });

  test('resolveResponseFromRoundNote reports the partial day back to the player', () => {
    expect(resolveResponseFromRoundNote('Available', 'Day 1 Only')).toBe('Day 1 Only');
    expect(resolveResponseFromRoundNote('Available', 'Day 2 Only - late')).toBe('Day 2 Only');
    expect(resolveResponseFromRoundNote('Available', '')).toBe('Available');
    expect(resolveResponseFromRoundNote('Unavailable', 'Day 1 Only')).toBe('Unavailable');
  });
});

describe('Availability_Log helpers', () => {
  test('headers lead with player name and include submitting phone', () => {
    expect(AVAILABILITY_LOG_HEADERS).toEqual([
      'Timestamp', 'Player Name', 'Match Date', 'Response', 'Notes', 'Submitted By Phone', 'ProfileID'
    ]);
  });

  test('isLegacyAvailabilityLogHeader detects the old layout only', () => {
    expect(isLegacyAvailabilityLogHeader(['Timestamp', 'ProfileID', 'MatchDate', 'Response', 'Notes'])).toBe(true);
    expect(isLegacyAvailabilityLogHeader(AVAILABILITY_LOG_HEADERS)).toBe(false);
    expect(isLegacyAvailabilityLogHeader([])).toBe(false);
  });

  test('buildAvailabilityLogRow orders fields per headers', () => {
    const ts = new Date('2026-10-13T08:26:34');
    expect(buildAvailabilityLogRow({
      timestamp: ts, profileId: ' P1 ', playerName: 'Sam Jones', matchDate: '2026-10-17',
      response: 'Day 1 Only', notes: 'late', phone: '+61400000000'
    })).toEqual([ts, 'Sam Jones', '2026-10-17', 'Day 1 Only', 'late', '+61400000000', 'P1']);
  });

  test('migrateLegacyAvailabilityLogRows back-fills names and preserves data', () => {
    const ts = new Date('2026-10-13T08:26:34');
    const rows = [[ts, 'P1', '2026-10-17', 'Day 1 Only', ''], [ts, 'P9', '2026-10-17', 'Both Days', 'x']];
    expect(migrateLegacyAvailabilityLogRows(rows, { P1: 'Sam Jones' })).toEqual([
      [ts, 'Sam Jones', '2026-10-17', 'Day 1 Only', '', '', 'P1'],
      [ts, '', '2026-10-17', 'Both Days', 'x', '', 'P9']
    ]);
  });
});

describe('Availability_Log GAS adapters (SheetLogic.gs)', () => {
  let gas;

  // Minimal in-memory sheet supporting the calls the adapters make.
  function makeSheet(name, initialRows) {
    let rows = initialRows.map(r => r.slice());
    const sheet = {
      getName: () => name,
      getLastRow: () => rows.length,
      getLastColumn: () => rows.reduce((m, r) => Math.max(m, r.length), 0),
      getDataRange: () => ({ getValues: () => rows.map(r => r.slice()) }),
      appendRow: jest.fn(r => { rows.push(r.slice()); }),
      setFrozenRows: jest.fn(),
      setColumnWidth: jest.fn(),
      getRange: (row, col, numRows = 1, numCols = 1) => {
        const range = {
          getValues: () => {
            const out = [];
            for (let i = 0; i < numRows; i++) {
              const src = rows[row - 1 + i] || [];
              const line = [];
              for (let j = 0; j < numCols; j++) line.push(src[col - 1 + j] === undefined ? '' : src[col - 1 + j]);
              out.push(line);
            }
            return out;
          },
          setValues: vals => {
            vals.forEach((v, i) => {
              const r = row - 1 + i;
              while (rows.length <= r) rows.push([]);
              v.forEach((cell, j) => { rows[r][col - 1 + j] = cell; });
            });
            return range;
          },
          clearContent: () => {
            for (let i = 0; i < numRows; i++) {
              const r = rows[row - 1 + i];
              if (r) for (let j = 0; j < numCols; j++) r[col - 1 + j] = '';
            }
            return range;
          },
          setFontWeight: () => range,
          setBackground: () => range,
          setFontColor: () => range
        };
        return range;
      },
      _rows: () => rows
    };
    return sheet;
  }

  beforeAll(() => {
    gas = Object.assign({}, global, logic);
    const code = fs.readFileSync(path.join(__dirname, '../src/SheetLogic.gs'), 'utf8');
    vm.runInNewContext(code, gas);
  });

  test('migrates legacy log rows to the new layout with player names', () => {
    const ts = new Date('2026-10-13T08:26:34');
    const players = makeSheet('Players', [
      ['ProfileID', 'FirstName', 'LastName', 'FullName'],
      ['P1', 'Sam', 'Jones', 'Sam Jones']
    ]);
    const log = makeSheet('Availability_Log', [
      ['Timestamp', 'ProfileID', 'MatchDate', 'Response', 'Notes'],
      [ts, 'P1', '2026-10-17', 'Day 1 Only', '']
    ]);
    const ss = { getSheetByName: n => ({ Players: players, Availability_Log: log }[n] || null) };

    gas.appendAvailabilityLog(ss, {
      profileId: 'P1', matchDate: '2026-10-17', response: 'Day 2 Only', notes: 'late', phone: '+61400000000'
    });

    const rows = log._rows();
    expect(rows[0]).toEqual(AVAILABILITY_LOG_HEADERS);
    expect(rows[1]).toEqual([ts, 'Sam Jones', '2026-10-17', 'Day 1 Only', '', '', 'P1']);
    expect(rows[2].slice(1)).toEqual(['Sam Jones', '2026-10-17', 'Day 2 Only', 'late', '+61400000000', 'P1']);
  });
});
