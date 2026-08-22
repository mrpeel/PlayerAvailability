const { getFixtureOptionsForRoundInit } = require('../src/logic.js');

describe('Round Initialization Fixture Options', () => {
  const sampleFixturesData = [
    [
      'Game Date',
      '1st Round', '1st Format', '1st Opponent', '1st Venue',
      '2nd Round', '2nd Format', '2nd Opponent', '2nd Venue',
      '3rd Round', '3rd Format', '3rd Opponent', '3rd Venue',
      '4th Round', '4th Format', '4th Opponent', '4th Venue',
      '5th Round', '5th Format', '5th Opponent', '5th Venue',
      'T20 1st Round', 'T20 1st Format', 'T20 1st Opponent', 'T20 1st Venue',
      'T20 2nd Round', 'T20 2nd Format', 'T20 2nd Opponent', 'T20 2nd Venue'
    ],
    [
      '16/11/2025',
      'Round 6', 'One Day', 'Blackburn - 2nd XI', 'Kalang Park',
      '', '', '', '',
      '', '', '', '',
      '', '', '', '',
      '', '', '', '',
      '', '', '', '',
      '', '', '', ''
    ],
    [
      '10/01/2026, 17/01/2026',
      'Round 10', 'Two Day', 'Glen Waverley Cougars - 1st XI', 'Kalang Park',
      'Round 10', 'Two Day', 'Glen Waverley Cougars - 2nd XI', 'Glen Waverley Reserve',
      'Round 8', 'Two Day', 'Blackburn North CC - 3rd XI', 'Mirrabooka Reserve',
      'Round 8', 'Two Day', 'Kerrimuir United - 4th XI', 'Springfield Park',
      'Round 8', 'Two Day', 'Blackburn North CC - 4th XI', 'Eley Park',
      '', '', '', '',
      '', '', '', ''
    ],
    [
      '20/01/2026',
      '', '', '', '',
      '', '', '', '',
      '', '', '', '',
      '', '', '', '',
      '', '', '', '',
      'Round 4', 'T20', 'Box Hill T20 1st XI', 'Kalang Park',
      'Round 4', 'T20', 'Wyclif T20 2nd XI', 'Jingella Reserve'
    ]
  ];

  test('returns empty array when fixturesData is empty or null', () => {
    expect(getFixtureOptionsForRoundInit(null, [])).toEqual([]);
    expect(getFixtureOptionsForRoundInit([], [])).toEqual([]);
    expect(getFixtureOptionsForRoundInit([['Header1']], [])).toEqual([]);
  });

  test('returns empty array when Game Date column is missing', () => {
    const invalidHeaders = [
      ['InvalidCol1', 'InvalidCol2'],
      ['val1', 'val2']
    ];
    expect(getFixtureOptionsForRoundInit(invalidHeaders, [])).toEqual([]);
  });

  test('correctly formats dropdown labels with Day and date (teams list)', () => {
    const options = getFixtureOptionsForRoundInit(sampleFixturesData, []);
    expect(options).toHaveLength(3);

    // 16 Nov 2025: Sunday, single team (1st XI)
    expect(options[0].date).toBe('2025-11-16');
    expect(options[0].displayDate).toBe('Sun 16 Nov 2025');
    expect(options[0].label).toBe('Sun 16 Nov 2025 (1st XI)');
    expect(options[0].teams).toEqual(['1st XI']);

    // 10 Jan 2026: Saturday, 5 teams (1st, 2nd, 3rd, 4th, 5th XI)
    expect(options[1].date).toBe('2026-01-10');
    expect(options[1].displayDate).toBe('Sat 10 Jan 2026');
    expect(options[1].label).toBe('Sat 10 Jan 2026 (1st XI, 2nd XI, 3rd XI, 4th XI, 5th XI)');
    expect(options[1].teams).toEqual(['1st XI', '2nd XI', '3rd XI', '4th XI', '5th XI']);

    // 20 Jan 2026: Tuesday, T20 teams (T20 1st XI, T20 2nd XI)
    expect(options[2].date).toBe('2026-01-20');
    expect(options[2].displayDate).toBe('Tue 20 Jan 2026');
    expect(options[2].label).toBe('Tue 20 Jan 2026 (T20 1st XI, T20 2nd XI)');
    expect(options[2].teams).toEqual(['T20 1st XI', 'T20 2nd XI']);
  });

  test('accurately identifies when a tab already exists in the workbook', () => {
    const existingSheets = ['Players', 'Fixtures', '2025-11-16', 'Config'];
    const options = getFixtureOptionsForRoundInit(sampleFixturesData, existingSheets);

    expect(options[0].tabExists).toBe(true);
    expect(options[0].label).toBe('Sun 16 Nov 2025 (1st XI) [Tab Initialised]');

    expect(options[1].tabExists).toBe(false);
    expect(options[1].label).not.toContain('[Tab Initialised]');
  });

  test('includes team matches breakdown per fixture date', () => {
    const options = getFixtureOptionsForRoundInit(sampleFixturesData, []);
    const t20Date = options[2];
    expect(t20Date.teamMatches).toHaveLength(2);
    expect(t20Date.teamMatches[0]).toEqual({
      team: 'T20 1st XI',
      round: 'Round 4',
      format: 'T20',
      opponent: 'Box Hill T20 1st XI',
      venue: 'Kalang Park'
    });
    expect(t20Date.teamMatches[1]).toEqual({
      team: 'T20 2nd XI',
      round: 'Round 4',
      format: 'T20',
      opponent: 'Wyclif T20 2nd XI',
      venue: 'Jingella Reserve'
    });
  });

  test('handles out-of-order rows by sorting them by normalized date', () => {
    const reversedData = [
      sampleFixturesData[0],
      sampleFixturesData[3],
      sampleFixturesData[1],
      sampleFixturesData[2]
    ];
    const options = getFixtureOptionsForRoundInit(reversedData, []);
    expect(options[0].date).toBe('2025-11-16');
    expect(options[1].date).toBe('2026-01-10');
    expect(options[2].date).toBe('2026-01-20');
  });
});

