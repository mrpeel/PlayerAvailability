const {
  formatRoundOpponent,
  getRoundFrames,
  extractSlotNumberForTeam,
  isT20Fixture,
  filterPlayersForRoundSnapshot
} = require('../src/logic.js');

describe('T20 Round & Selection Support', () => {
  describe('formatRoundOpponent with T20', () => {
    test('formats T20 1st XI correctly', () => {
      expect(formatRoundOpponent('Round 1', 'T20_1ST', 'Box Hill T20 1st XI'))
        .toBe('Round 1: LCC T20 1st XI vs Box Hill T20 1st XI');
      expect(formatRoundOpponent('1', 'T20 1st', 'Box Hill'))
        .toBe('Round 1: LCC T20 1st XI vs Box Hill');
    });

    test('formats T20 2nd XI correctly', () => {
      expect(formatRoundOpponent('Round 2', 'T20_2ND', 'Wyclif T20 2nd XI'))
        .toBe('Round 2: LCC T20 2nd XI vs Wyclif T20 2nd XI');
    });

    test('maintains standard 1st-5th XI formatting', () => {
      expect(formatRoundOpponent('Round 3', '1ST', 'Mitcham - 2nd XI'))
        .toBe('Round 3: LCC 1st XI vs Mitcham - 2nd XI');
      expect(formatRoundOpponent('Round 4', '5TH', 'Vermont South'))
        .toBe('Round 4: LCC 5th XI vs Vermont South');
    });
  });

  describe('getRoundFrames', () => {
    test('returns 2 frames for T20 round with correct names and slot starts', () => {
      const frames = getRoundFrames(true);
      expect(frames).toHaveLength(2);
      expect(frames[0]).toEqual({
        name: 'T20 1st ELEVEN',
        prefix: 'T20 1st',
        altPrefix: 'T20_1ST',
        defaultIdx: 0,
        start: 4,
        slots: 12
      });
      expect(frames[1]).toEqual({
        name: 'T20 SECOND ELEVEN',
        prefix: 'T20 2nd',
        altPrefix: 'T20_2ND',
        defaultIdx: 1,
        start: 22,
        slots: 12
      });
    });

    test('returns 5 frames for standard Saturday round', () => {
      const frames = getRoundFrames(false);
      expect(frames).toHaveLength(5);
      expect(frames[0].name).toBe('FIRST ELEVEN');
      expect(frames[1].name).toBe('SECOND ELEVEN');
      expect(frames[2].name).toBe('THIRD ELEVEN');
      expect(frames[3].name).toBe('FOURTH ELEVEN');
      expect(frames[4].name).toBe('FIFTH ELEVEN');
    });
  });

  describe('isT20Fixture', () => {
    const fixturesData = [
      ['Game Date', '1st Round', '1st Format', '1st Opponent', '1st Venue', 'T20 1st Round', 'T20 1st Format', 'T20 1st Opponent', 'T20 1st Venue'],
      ['10/01/2026', 'Round 10', 'Two Day', 'Glen Waverley Cougars', 'Kalang Park', '', '', '', ''],
      ['20/01/2026', '', '', '', '', 'Round 4', 'T20', 'Box Hill T20 1st XI', 'Kalang Park']
    ];

    test('detects T20 from fixture row object with format or T20 keys', () => {
      expect(isT20Fixture({ '1st format': 'Two Day', '1st opponent': 'Mitcham' })).toBe(false);
      expect(isT20Fixture({ 't20 1st format': 'T20', 't20 1st opponent': 'Box Hill' })).toBe(true);
      expect(isT20Fixture({ '1st format': 'T20', '1st opponent': 'Box Hill' })).toBe(true);
    });

    test('detects T20 from date lookup against 2D fixtures data', () => {
      expect(isT20Fixture('2026-01-10', fixturesData)).toBe(false);
      expect(isT20Fixture('2026-01-20', fixturesData)).toBe(true);
    });

    test('detects T20 from explicit tab name string containing T20', () => {
      expect(isT20Fixture('2026-01-20 (T20)')).toBe(true);
    });
  });

  describe('filterPlayersForRoundSnapshot', () => {
    const masterPlayersData = [
      ['ProfileID', 'FirstName', 'LastName', 'FullName', 'JuniorLevel', 'T20Squad', 'GlobalStatus', 'ExpectedReturnDate'],
      ['P001', 'Neil', 'Kloot', 'Neil Kloot', '', 'Yes', 'Active', ''],
      ['P002', 'Jimi', 'Kloot', 'Jimi Kloot', 'U16', 'yes', 'Active', ''],
      ['P003', 'John', 'Smith', 'John Smith', '', 'No', 'Active', ''],
      ['P004', 'Mark', 'Taylor', 'Mark Taylor', '', '  YES  ', 'Injured', '2026-02-01'],
      ['P005', 'Sam', 'Brown', 'Sam Brown', '', '', 'Injured', '2026-02-01'],
      ['P006', 'Aaron', 'Alaimo', 'Aaron Alaimo', '', '', 'Active', ''],
      ['P007', 'Adam', 'Doungas', 'Adam Doungas', '', null, 'Active', ''],
      ['P008', 'Gus', 'Schwarz', 'Gus Schwarz', '', 'NO', 'Active', '']
    ];

    test('strictly filters only players with explicit "Yes" in T20Squad for T20 rounds', () => {
      const result = filterPlayersForRoundSnapshot(masterPlayersData, true, '2026-01-20');
      // Only P001 ('Yes') and P002 ('yes') are active with explicit Yes.
      // P003 ('No'), P006 (''), P007 (null), P008 ('NO') MUST be excluded!
      expect(result.unknownSnapshot).toHaveLength(2);
      expect(result.unknownSnapshot[0][0]).toBe('P001');
      expect(result.unknownSnapshot[1][0]).toBe('P002');

      // Unavailable: P004 ('  YES  ') is included. P005 ('') is excluded!
      expect(result.unavailableSnapshot).toHaveLength(1);
      expect(result.unavailableSnapshot[0][0]).toBe('P004');
    });

    test('includes all active players when isT20 is false (standard round)', () => {
      const result = filterPlayersForRoundSnapshot(masterPlayersData, false, '2026-01-10');
      // All active players: P001, P002, P003, P006, P007, P008
      expect(result.unknownSnapshot).toHaveLength(6);
      expect(result.unknownSnapshot.map(r => r[0])).toEqual(['P001', 'P002', 'P003', 'P006', 'P007', 'P008']);

      // All injured players: P004, P005
      expect(result.unavailableSnapshot).toHaveLength(2);
      expect(result.unavailableSnapshot.map(r => r[0])).toEqual(['P004', 'P005']);
    });
  });

  describe('extractSlotNumberForTeam', () => {
    test('extracts slot for T20_1ST from various tag formats', () => {
      expect(extractSlotNumberForTeam('T20_1ST_SLOT_1', 'T20_1ST')).toBe(1);
      expect(extractSlotNumberForTeam('T20_SLOT_1', 'T20_1ST')).toBe(1);
      expect(extractSlotNumberForTeam('T20_1_SLOT_2', 'T20_1ST')).toBe(2);
      expect(extractSlotNumberForTeam('T20_FIRST_SLOT_3', 'T20_1ST')).toBe(3);
      expect(extractSlotNumberForTeam('T20 1ST XI SLOT 4', 'T20_1ST')).toBe(4);
      expect(extractSlotNumberForTeam('T20_1ST__PLAYER_3', 'T20_1ST')).toBe(3);
      expect(extractSlotNumberForTeam('T20_1ST_PHOTO_12', 'T20_1ST')).toBe(12);
      expect(extractSlotNumberForTeam('T20 1ST PHOTO 4', 'T20_1ST')).toBe(4);
      expect(extractSlotNumberForTeam('T20_1ST_IMAGE_6', 'T20_1ST')).toBe(6);
      expect(extractSlotNumberForTeam('T20_1ST_AVATAR_7', 'T20_1ST')).toBe(7);
      expect(extractSlotNumberForTeam('T20_1ST-HEADSHOT-8', 'T20_1ST')).toBe(8);
      expect(extractSlotNumberForTeam('SLOT_5', 'T20_1ST')).toBe(5);
      expect(extractSlotNumberForTeam('PLAYER_2', 'T20_1ST')).toBe(2);
      expect(extractSlotNumberForTeam('PHOTO 9', 'T20_1ST')).toBe(9);
    });

    test('extracts slot for T20_2ND from various tag formats', () => {
      expect(extractSlotNumberForTeam('T20_2ND_SLOT_1', 'T20_2ND')).toBe(1);
      expect(extractSlotNumberForTeam('T20_2_SLOT_2', 'T20_2ND')).toBe(2);
      expect(extractSlotNumberForTeam('T20_SECOND_SLOT_3', 'T20_2ND')).toBe(3);
      expect(extractSlotNumberForTeam('T20 2ND XI SLOT 4', 'T20_2ND')).toBe(4);
    });

    test('rejects other team prefixes from cross-matching', () => {
      // T20_2ND or 1ST elements should NOT match when processing T20_1ST
      expect(extractSlotNumberForTeam('T20_2ND_SLOT_1', 'T20_1ST')).toBeNull();
      expect(extractSlotNumberForTeam('1ST_SLOT_1', 'T20_1ST')).toBeNull();
      expect(extractSlotNumberForTeam('2ND_SLOT_1', 'T20_1ST')).toBeNull();

      // T20_1ST should NOT match when processing 1ST (Saturday Firsts)
      expect(extractSlotNumberForTeam('T20_1ST_SLOT_1', '1ST')).toBeNull();
      expect(extractSlotNumberForTeam('1ST_SLOT_1', '1ST')).toBe(1);
    });

    test('returns null for match info or unrelated tags', () => {
      expect(extractSlotNumberForTeam('T20_1ST__ROUND_OPPONENT', 'T20_1ST')).toBeNull();
      expect(extractSlotNumberForTeam('T20_1ST_FORMAT_VENUE', 'T20_1ST')).toBeNull();
      expect(extractSlotNumberForTeam('RANDOM_SHAPE', 'T20_1ST')).toBeNull();
      expect(extractSlotNumberForTeam('', 'T20_1ST')).toBeNull();
      expect(extractSlotNumberForTeam(null, 'T20_1ST')).toBeNull();
    });
  });
});
