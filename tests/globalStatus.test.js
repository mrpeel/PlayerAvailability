const {
  GLOBAL_STATUS_OPTIONS,
  normalizeGlobalStatus,
  toDateKey,
  resolveGlobalStatusPlacement,
  isGlobalStatusNote,
  filterPlayersForRoundSnapshot,
  simulatePlayerAvailability
} = require('../src/logic.js');

describe('normalizeGlobalStatus', () => {
  test('maps web-app and legacy variants to canonical values', () => {
    expect(normalizeGlobalStatus('UnavailableUntil')).toBe('Unavailable Until');
    expect(normalizeGlobalStatus('unavailable until')).toBe('Unavailable Until');
    expect(normalizeGlobalStatus('Away')).toBe('Unavailable Until');
    expect(normalizeGlobalStatus('long term away')).toBe('Long-Term Away');
    expect(normalizeGlobalStatus(' injured ')).toBe('Injured');
    expect(normalizeGlobalStatus('')).toBe('Active');
    expect(normalizeGlobalStatus(null)).toBe('Active');
  });

  test('every canonical value is in the validation list and round-trips', () => {
    GLOBAL_STATUS_OPTIONS.forEach(s => expect(normalizeGlobalStatus(s)).toBe(s));
    expect(GLOBAL_STATUS_OPTIONS).toContain('Unavailable Until');
  });
});

describe('toDateKey', () => {
  test('handles ISO strings, Dates and blanks', () => {
    expect(toDateKey('2026-11-01')).toBe('2026-11-01');
    expect(toDateKey('2026-10-10 (T20)')).toBe('2026-10-10');
    expect(toDateKey(new Date(2026, 10, 1))).toBe('2026-11-01');
    expect(toDateKey('')).toBe('');
    expect(toDateKey(null)).toBe('');
    expect(toDateKey('not a date')).toBe('');
  });
});

describe('resolveGlobalStatusPlacement', () => {
  test('Unavailable Until puts player in Unavailable before the return date, with a note', () => {
    const p = resolveGlobalStatusPlacement('UnavailableUntil', new Date(2026, 10, 1), '2026-10-10');
    expect(p.list).toBe('Unavailable');
    expect(p.note).toBe('Unavailable until 2026-11-01');
  });

  test('player is back in Unknown on/after the return date', () => {
    expect(resolveGlobalStatusPlacement('Unavailable Until', '2026-11-01', '2026-11-01').list).toBe('Unknown');
    expect(resolveGlobalStatusPlacement('Injured', '2026-11-01', '2026-11-08').list).toBe('Unknown');
  });

  test('no return date means unavailable indefinitely with pure reason as note', () => {
    expect(resolveGlobalStatusPlacement('Injured', '', '2026-10-10')).toEqual({
      list: 'Unavailable',
      note: 'Injured',
      status: 'Injured'
    });
    expect(resolveGlobalStatusPlacement('Inactive', '2026-01-01', '2026-10-10')).toEqual({
      list: 'Unavailable',
      note: 'Inactive',
      status: 'Inactive'
    });
  });

  test('active players are Unknown', () => {
    expect(resolveGlobalStatusPlacement('Active', '', '2026-10-10')).toEqual({ list: 'Unknown', note: '', status: 'Active' });
  });
});

describe('isGlobalStatusNote', () => {
  test('recognizes standard and legacy global exemption notes', () => {
    expect(isGlobalStatusNote('Inactive')).toBe(true);
    expect(isGlobalStatusNote('Injured')).toBe(true);
    expect(isGlobalStatusNote('Injured until 2026-11-01')).toBe(true);
    expect(isGlobalStatusNote('Unavailable')).toBe(true);
    expect(isGlobalStatusNote('Unavailable until 2026-11-01')).toBe(true);
    expect(isGlobalStatusNote('Long-Term Away')).toBe(true);
    expect(isGlobalStatusNote('Long-Term Away until 2026-11-01')).toBe(true);
    expect(isGlobalStatusNote('Global Exemption: Inactive')).toBe(true);
    expect(isGlobalStatusNote('Global Exemption: Unavailable until 2026-11-01')).toBe(true);

    expect(isGlobalStatusNote('Unavailable for Round')).toBe(false);
    expect(isGlobalStatusNote('Holiday with family')).toBe(false);
    expect(isGlobalStatusNote('')).toBe(false);
    expect(isGlobalStatusNote(null)).toBe(false);
  });
});

describe('filterPlayersForRoundSnapshot with Unavailable Until', () => {
  const data = [
    ['ProfileID', 'FirstName', 'LastName', 'FullName', 'JuniorLevel', 'T20Squad', 'GlobalStatus', 'ExpectedReturnDate'],
    ['A1', 'Neil', 'Kloot', 'Neil Kloot', '', '', 'Active', ''],
    ['A2', 'Aidan', 'Meehan', 'Aidan Meehan', '', '', 'UnavailableUntil', new Date(2026, 10, 1)],
    ['A3', 'Michael', 'Meehan', 'Michael Meehan', '', '', 'Unavailable Until', '2026-11-01'],
    ['A4', 'Old', 'Injury', 'Old Injury', '', '', 'Injured', '2026-09-01']
  ];

  test('no player silently disappears from the round', () => {
    const res = filterPlayersForRoundSnapshot(data, false, '2026-10-10');
    const ids = res.unavailableSnapshot.concat(res.unknownSnapshot).map(r => r[0]).sort();
    expect(ids).toEqual(['A1', 'A2', 'A3', 'A4']);
  });

  test('Unavailable Until players appear in Unavailable with a notation without prefix', () => {
    const res = filterPlayersForRoundSnapshot(data, false, '2026-10-10');
    expect(res.unavailableSnapshot.map(r => r[0])).toEqual(['A2', 'A3']);
    expect(res.unavailableSnapshot[0][2]).toBe('Unavailable until 2026-11-01');
    // Return date passed -> back in Unknown rather than dropped
    expect(res.unknownSnapshot.map(r => r[0])).toEqual(['A1', 'A4']);
  });

  test('rounds after the return date treat them as normal', () => {
    const res = filterPlayersForRoundSnapshot(data, false, '2026-11-07');
    expect(res.unavailableSnapshot).toHaveLength(0);
    expect(res.unknownSnapshot).toHaveLength(4);
  });
});

describe('simulatePlayerAvailability', () => {
  test('treats Unavailable Until as unavailable', () => {
    const [r] = simulatePlayerAvailability([{ profileId: 'X', fullName: 'X', globalStatus: 'UnavailableUntil' }], 'One Day', () => 0);
    expect(r.response).toBe('Unavailable');
  });
});
