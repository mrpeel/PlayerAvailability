const {
  getRoundEndDate,
  resolveSmartRoundSelection
} = require('../src/logic.js');

describe('Smart Round Selection & Re-check Logic', () => {
  describe('getRoundEndDate', () => {
    test('returns same date for one-day / T20 matches', () => {
      expect(getRoundEndDate('2026-11-01', false)).toBe('2026-11-01');
      expect(getRoundEndDate('2026-11-01', 'One Day')).toBe('2026-11-01');
      expect(getRoundEndDate('2026-11-01', 'T20')).toBe('2026-11-01');
    });

    test('adds 7 days for two-day matches', () => {
      expect(getRoundEndDate('2026-10-04', true)).toBe('2026-10-11');
      expect(getRoundEndDate('2026-10-04', 'Two Day')).toBe('2026-10-11');
      expect(getRoundEndDate('2026-10-31', 'Two Day')).toBe('2026-11-07');
    });

    test('handles empty or invalid dates gracefully', () => {
      expect(getRoundEndDate('', true)).toBe('');
      expect(getRoundEndDate(null, false)).toBe('');
    });
  });

  describe('resolveSmartRoundSelection', () => {
    const sampleFixtures = [
      {
        date: '2026-10-04',
        endDate: '2026-10-11',
        format: 'Two Day',
        round: 'Round 1',
        roundTitle: 'Round 1'
      },
      {
        date: '2026-10-18',
        endDate: '2026-10-25',
        format: 'Two Day',
        round: 'Round 2',
        roundTitle: 'Round 2'
      },
      {
        date: '2026-11-01',
        endDate: '2026-11-01',
        format: 'One Day',
        round: 'Round 3',
        roundTitle: 'Round 3'
      }
    ];

    test('Requirement 1: Automatically finds next round and redirects if no date in URL', () => {
      // Today is 2026-10-13 (Round 1 ended 11 Oct; next round is Round 2 on 18 Oct)
      const res = resolveSmartRoundSelection(null, sampleFixtures, '2026-10-13');

      expect(res.redirected).toBe(true);
      expect(res.targetDate).toBe('2026-10-18');
      expect(res.action).toBe('REDIRECT_TO_UPCOMING');
      expect(res.redirectReason).toBe('NO_DATE_SPECIFIED');
      expect(res.hasUpcoming).toBe(true);
      expect(res.noUpcomingRounds).toBe(false);
    });

    test('Requirement 1: Reports no upcoming rounds when all fixtures are in the past and no date in URL', () => {
      // Season ended: today is 2026-12-01
      const res = resolveSmartRoundSelection(null, sampleFixtures, '2026-12-01');

      expect(res.noUpcomingRounds).toBe(true);
      expect(res.hasUpcoming).toBe(false);
      expect(res.action).toBe('NO_UPCOMING_ROUNDS');
    });

    test('Requirement 1: Reports no upcoming rounds when fixtures list is empty', () => {
      const res = resolveSmartRoundSelection('', [], '2026-10-13');

      expect(res.noUpcomingRounds).toBe(true);
      expect(res.targetDate).toBeNull();
      expect(res.hasUpcoming).toBe(false);
    });

    test('Requirement 2: Switches to upcoming round if requested date is in the past AND future round exists', () => {
      // User opens link from Round 1 (2026-10-04, which ended 11 Oct) on 2026-10-13
      const res = resolveSmartRoundSelection('2026-10-04', sampleFixtures, '2026-10-13');

      expect(res.redirected).toBe(true);
      expect(res.targetDate).toBe('2026-10-18');
      expect(res.previousDate).toBe('2026-10-04');
      expect(res.action).toBe('SWITCH_TO_UPCOMING');
      expect(res.redirectReason).toBe('PAST_DATE_SWITCHED');
      expect(res.hasUpcoming).toBe(true);
    });

    test('Requirement 2: Keeps requested past date if NO future round exists (e.g. end of season review)', () => {
      // User opens Round 3 (2026-11-01) in December 2026 when no more rounds exist
      const res = resolveSmartRoundSelection('2026-11-01', sampleFixtures, '2026-12-01');

      expect(res.redirected).toBe(false);
      expect(res.targetDate).toBe('2026-11-01');
      expect(res.action).toBe('DISPLAY_PAST_NO_UPCOMING');
      expect(res.noUpcomingRounds).toBe(true);
    });

    test('Requirement 3: Keeps requested date if requested date is current or in the future', () => {
      // Today is 2026-10-13; user requests upcoming Round 2 (2026-10-18)
      const res = resolveSmartRoundSelection('2026-10-18', sampleFixtures, '2026-10-13');

      expect(res.redirected).toBe(false);
      expect(res.targetDate).toBe('2026-10-18');
      expect(res.action).toBe('DISPLAY_REQUESTED');
      expect(res.hasUpcoming).toBe(true);
    });

    test('Two-day match in progress: does NOT switch between Day 1 and Day 2', () => {
      // Round 1 started 2026-10-04, Day 2 is 2026-10-11.
      // Today is 2026-10-07 (mid-week between Day 1 and Day 2).
      const res = resolveSmartRoundSelection('2026-10-04', sampleFixtures, '2026-10-07');

      // Round 1 is not in the past yet because its endDate is 2026-10-11
      expect(res.redirected).toBe(false);
      expect(res.targetDate).toBe('2026-10-04');
      expect(res.action).toBe('DISPLAY_REQUESTED');
    });

    test('Two-day match in progress: no date in URL resolves to the active in-progress round', () => {
      // Today is 2026-10-07 (mid-week between Day 1 and Day 2 of Round 1)
      const res = resolveSmartRoundSelection(null, sampleFixtures, '2026-10-07');

      expect(res.targetDate).toBe('2026-10-04');
      expect(res.action).toBe('REDIRECT_TO_UPCOMING');
    });

    test('Handles simple string array of dates', () => {
      const stringDates = ['2026-10-04', '2026-10-18', '2026-11-01'];
      const res = resolveSmartRoundSelection('2026-10-04', stringDates, '2026-10-13');

      expect(res.redirected).toBe(true);
      expect(res.targetDate).toBe('2026-10-18');
      expect(res.action).toBe('SWITCH_TO_UPCOMING');
    });

    test('Handles PlayHQ comma-separated multi-day fixture dates in objects', () => {
      const rawFixtures = [
        {
          'Game Date': '04/10/2025, 11/10/2025',
          '1st Round': 'Round 1',
          '1st Format': 'Two Day'
        },
        {
          'Game Date': '18/10/2025, 25/10/2025',
          '1st Round': 'Round 2',
          '1st Format': 'Two Day'
        }
      ];

      // Reference date is 2025-10-13 (Round 1 ended 11/10/2025)
      const res = resolveSmartRoundSelection('2025-10-04', rawFixtures, '2025-10-13');
      expect(res.redirected).toBe(true);
      expect(res.targetDate).toBe('2025-10-18');
    });
  });
});
