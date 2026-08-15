const fs = require('fs');
const path = require('path');
const {
  parseCsvString,
  formatTeamPrefix,
  buildFixtureHeaders,
  readTeamConfigsFromSheet,
  readTemplatesFromSheet,
  parseFixtureCsv,
  mergeFixturesIntoMatrix,
  DEFAULT_TEAM_CONFIGS
} = require('../src/logic.js');

describe('Fixture Logic & Import Engine', () => {
  const homeCsvPath = path.join(__dirname, '../sample_data/Fixtures for Selection - home.csv');
  const awayCsvPath = path.join(__dirname, '../sample_data/Fixtures for Selection - away.csv');

  const homeCsvContent = fs.readFileSync(homeCsvPath, 'utf8');
  const awayCsvContent = fs.readFileSync(awayCsvPath, 'utf8');

  describe('CSV Parser (parseCsvString)', () => {
    test('parses standard comma-separated lines', () => {
      const input = 'a,b,c\n1,2,3';
      expect(parseCsvString(input)).toEqual([
        ['a', 'b', 'c'],
        ['1', '2', '3']
      ]);
    });

    test('parses quoted values containing commas and escaped quotes', () => {
      const input = 'Name,Bio\n"Smith, John","He said ""Hello!"""';
      const rows = parseCsvString(input);
      expect(rows).toHaveLength(2);
      expect(rows[1][0]).toBe('Smith, John');
      expect(rows[1][1]).toBe('He said "Hello!"');
    });

    test('handles CRLF line breaks and empty lines cleanly', () => {
      const input = "Round 1,Kalang Park\r\n\r\nRound 2,Eley Park\r\n";
      const rows = parseCsvString(input);
      expect(rows).toEqual([
        ['Round 1', 'Kalang Park'],
        ['Round 2', 'Eley Park']
      ]);
    });

    test('returns empty array for empty or non-string input', () => {
      expect(parseCsvString('')).toEqual([]);
      expect(parseCsvString(null)).toEqual([]);
      expect(parseCsvString(undefined)).toEqual([]);
    });
  });

  describe('Config & Template Readers', () => {
    test('extracts team configs from 2D sheet array', () => {
      const sheetData = [
        ['TEAM CONFIGURATION', '', ''],
        ['Internal Team Name', 'Competition', 'Play Cricket Team Name'],
        ['1st XI', 'BHRDCA Senior Competition', 'Laburnum - 1st XI'],
        ['2nd XI', 'BHRDCA Senior Competition', 'Laburnum - 2nd XI'],
        ['', '', ''],
        ['WHATSAPP & MESSAGE TEMPLATES', '', ''],
        ['Template Name', 'Template Text', '']
      ];

      const configs = readTeamConfigsFromSheet(sheetData);
      expect(configs).toHaveLength(2);
      expect(configs[0]).toEqual({
        internalName: '1st XI',
        competition: 'BHRDCA Senior Competition',
        playCricketName: 'Laburnum - 1st XI'
      });
      expect(configs[1]).toEqual({
        internalName: '2nd XI',
        competition: 'BHRDCA Senior Competition',
        playCricketName: 'Laburnum - 2nd XI'
      });
    });

    test('falls back to default team configs if Config sheet is empty or header missing', () => {
      expect(readTeamConfigsFromSheet([])).toEqual(DEFAULT_TEAM_CONFIGS);
      expect(readTeamConfigsFromSheet([['Random', 'Data']])).toEqual(DEFAULT_TEAM_CONFIGS);
    });

    test('extracts templates from 2D sheet array', () => {
      const sheetData = [
        ['WHATSAPP & MESSAGE TEMPLATES', ''],
        ['Template Name', 'Template Text'],
        ['Availability Callout', 'Custom callout message {url}'],
        ['Wall of Shame', 'Custom shame message {players}']
      ];

      const templates = readTemplatesFromSheet(sheetData);
      expect(templates['Availability Callout']).toBe('Custom callout message {url}');
      expect(templates['Wall of Shame']).toBe('Custom shame message {players}');
      expect(templates['Selection Announcement']).toBeDefined();
    });
  });

  describe('Header Builder (buildFixtureHeaders)', () => {
    test('builds wide header array with Game Date as PK and per-team Round, Format, Opponent, Venue', () => {
      const headers = buildFixtureHeaders(DEFAULT_TEAM_CONFIGS);
      expect(headers).toEqual([
        'Game Date',
        '1st Round', '1st Format', '1st Opponent', '1st Venue',
        '2nd Round', '2nd Format', '2nd Opponent', '2nd Venue',
        '3rd Round', '3rd Format', '3rd Opponent', '3rd Venue',
        '4th Round', '4th Format', '4th Opponent', '4th Venue',
        '5th Round', '5th Format', '5th Opponent', '5th Venue'
      ]);
    });

    test('normalizes custom team names in prefix', () => {
      const customConfigs = [
        { internalName: '1st XI', competition: 'Comp', playCricketName: 'Team 1' },
        { internalName: 'Sunday XI', competition: 'Comp', playCricketName: 'Team Sun' }
      ];
      const headers = buildFixtureHeaders(customConfigs);
      expect(headers).toEqual([
        'Game Date',
        '1st Round', '1st Format', '1st Opponent', '1st Venue',
        'Sunday Round', 'Sunday Format', 'Sunday Opponent', 'Sunday Venue'
      ]);
    });
  });

  describe('PlayHQ Fixture Parsing (parseFixtureCsv)', () => {
    test('parses Home fixtures CSV accurately with sample data', () => {
      const games = parseFixtureCsv(homeCsvContent, DEFAULT_TEAM_CONFIGS);
      expect(games.length).toBeGreaterThan(0);

      // Verify a specific home game
      const round1Team1 = games.find(g => g.startDate === '2025-10-04' && g.internalTeamName === '1st XI');
      expect(round1Team1).toBeDefined();
      expect(round1Team1.side).toBe('Home');
      expect(round1Team1.opponent).toBe('Mitcham - 2nd XI');
      expect(round1Team1.venue).toBe('Kalang Park');
      expect(round1Team1.format).toBe('Two Day');
      expect(round1Team1.round).toBe('Round 1');
    });

    test('parses Away fixtures CSV accurately with sample data', () => {
      const games = parseFixtureCsv(awayCsvContent, DEFAULT_TEAM_CONFIGS);
      expect(games.length).toBeGreaterThan(0);

      // Verify a specific away game (2nd XI in Round 1 was away at Ted Ajani Reserve)
      const round1Team2 = games.find(g => g.startDate === '2025-10-04' && g.internalTeamName === '2nd XI');
      expect(round1Team2).toBeDefined();
      expect(round1Team2.side).toBe('Away');
      expect(round1Team2.opponent).toBe('Bulleen Templestowe - 2nd XI');
      expect(round1Team2.venue).toBe('Ted Ajani Reserve');
      expect(round1Team2.format).toBe('Two Day');
      expect(round1Team2.round).toBe('Round 1');
    });

    test('handles empty or invalid CSV text gracefully', () => {
      expect(parseFixtureCsv('', DEFAULT_TEAM_CONFIGS)).toEqual([]);
      expect(parseFixtureCsv('Just,One,Header\n', DEFAULT_TEAM_CONFIGS)).toEqual([]);
    });
  });

  describe('Phased Matrix Merging (mergeFixturesIntoMatrix)', () => {
    test('merges Home fixtures first, then Away fixtures second', () => {
      const homeGames = parseFixtureCsv(homeCsvContent, DEFAULT_TEAM_CONFIGS);
      const awayGames = parseFixtureCsv(awayCsvContent, DEFAULT_TEAM_CONFIGS);

      const matrixPhase1 = mergeFixturesIntoMatrix([], homeGames, DEFAULT_TEAM_CONFIGS);
      expect(matrixPhase1.length).toBeGreaterThan(1);

      const matrixPhase2 = mergeFixturesIntoMatrix(matrixPhase1, awayGames, DEFAULT_TEAM_CONFIGS);
      expect(matrixPhase2.length).toBeGreaterThan(1);

      // Verify headers
      const headers = matrixPhase2[0];
      expect(headers[0]).toBe('Game Date');
      expect(headers[1]).toBe('1st Round');
      expect(headers[2]).toBe('1st Format');
      expect(headers[3]).toBe('1st Opponent');
      expect(headers[4]).toBe('1st Venue');

      // Check Round 1 (2025-10-04): all 5 teams should be populated across Home & Away
      const round1Row = matrixPhase2.find(r => r[0] === '2025-10-04');
      expect(round1Row).toBeDefined();

      // 1st XI (Home)
      expect(round1Row[1]).toBe('Round 1');          // 1st Round
      expect(round1Row[2]).toBe('Two Day');          // 1st Format
      expect(round1Row[3]).toBe('Mitcham - 2nd XI'); // 1st Opponent
      expect(round1Row[4]).toBe('Kalang Park');      // 1st Venue

      // 2nd XI (Away)
      expect(round1Row[5]).toBe('Round 1');                      // 2nd Round
      expect(round1Row[6]).toBe('Two Day');                      // 2nd Format
      expect(round1Row[7]).toBe('Bulleen Templestowe - 2nd XI'); // 2nd Opponent
      expect(round1Row[8]).toBe('Ted Ajani Reserve');            // 2nd Venue

      // 3rd XI (Home)
      expect(round1Row[9]).toBe('Round 1');                          // 3rd Round
      expect(round1Row[10]).toBe('Two Day');                         // 3rd Format
      expect(round1Row[11]).toBe('Box Hill North Super Kings - 4th XI'); // 3rd Opponent
      expect(round1Row[12]).toBe('Eley Park');                       // 3rd Venue

      // 4th XI (Away)
      expect(round1Row[13]).toBe('Round 1');              // 4th Round
      expect(round1Row[14]).toBe('Two Day');              // 4th Format
      expect(round1Row[15]).toBe('Nunawading - 5th XI');  // 4th Opponent
      expect(round1Row[16]).toBe('Mahoneys Reserve');     // 4th Venue

      // 5th XI (Away)
      expect(round1Row[17]).toBe('Round 1');                   // 5th Round
      expect(round1Row[18]).toBe('Two Day');                   // 5th Format
      expect(round1Row[19]).toBe('Blackburn North CC - 4th XI'); // 5th Opponent
      expect(round1Row[20]).toBe('Koonung Reserve');           // 5th Venue
    });

    test('correctly captures differing round numbers across teams on the same match date', () => {
      const homeGames = parseFixtureCsv(homeCsvContent, DEFAULT_TEAM_CONFIGS);
      const awayGames = parseFixtureCsv(awayCsvContent, DEFAULT_TEAM_CONFIGS);

      const matrix = mergeFixturesIntoMatrix(
        mergeFixturesIntoMatrix([], homeGames, DEFAULT_TEAM_CONFIGS),
        awayGames,
        DEFAULT_TEAM_CONFIGS
      );

      // On 2025-11-08: 1st and 2nd XI played Round 4, while 3rd, 4th, 5th XI played Round 3
      const splitRoundRow = matrix.find(r => r[0] === '2025-11-08');
      expect(splitRoundRow).toBeDefined();

      expect(splitRoundRow[1]).toBe('Round 4'); // 1st Round
      expect(splitRoundRow[5]).toBe('Round 4'); // 2nd Round
      expect(splitRoundRow[9]).toBe('Round 3'); // 3rd Round
      expect(splitRoundRow[13]).toBe('Round 3'); // 4th Round
      expect(splitRoundRow[17]).toBe('Round 3'); // 5th Round
    });

    test('is order-invariant (Away then Home produces identical result)', () => {
      const homeGames = parseFixtureCsv(homeCsvContent, DEFAULT_TEAM_CONFIGS);
      const awayGames = parseFixtureCsv(awayCsvContent, DEFAULT_TEAM_CONFIGS);

      const mHomeThenAway = mergeFixturesIntoMatrix(
        mergeFixturesIntoMatrix([], homeGames, DEFAULT_TEAM_CONFIGS),
        awayGames,
        DEFAULT_TEAM_CONFIGS
      );

      const mAwayThenHome = mergeFixturesIntoMatrix(
        mergeFixturesIntoMatrix([], awayGames, DEFAULT_TEAM_CONFIGS),
        homeGames,
        DEFAULT_TEAM_CONFIGS
      );

      expect(mHomeThenAway).toEqual(mAwayThenHome);
    });

    test('accurately captures per-team match formats on mixed format dates', () => {
      const mockGames = [
        {
          internalTeamName: '1st XI',
          playCricketName: 'Laburnum - 1st XI',
          competition: 'Comp',
          startDate: '2026-11-01',
          rawDate: '2026-11-01',
          round: 'Round 4',
          opponent: 'Opponent A',
          venue: 'Ground A',
          format: 'One Day',
          side: 'Home'
        },
        {
          internalTeamName: '4th XI',
          playCricketName: 'Laburnum - 4th XI',
          competition: 'Comp',
          startDate: '2026-11-01',
          rawDate: '2026-11-01, 2026-11-08',
          round: 'Round 4',
          opponent: 'Opponent B',
          venue: 'Ground B',
          format: 'Two Day',
          side: 'Away'
        }
      ];

      const matrix = mergeFixturesIntoMatrix([], mockGames, DEFAULT_TEAM_CONFIGS);
      const row = matrix.find(r => r[0] === '2026-11-01');
      expect(row).toBeDefined();

      // 1st XI format
      expect(row[2]).toBe('One Day');
      // 4th XI format
      expect(row[14]).toBe('Two Day');
    });

    test('sorts match dates chronologically', () => {
      const mockGames = [
        { internalTeamName: '1st XI', startDate: '2026-12-01', round: 'R3', opponent: 'B', venue: 'V', format: 'T20', side: 'Home' },
        { internalTeamName: '1st XI', startDate: '2026-10-01', round: 'R1', opponent: 'A', venue: 'V', format: 'One Day', side: 'Home' },
        { internalTeamName: '1st XI', startDate: '2026-11-01', round: 'R2', opponent: 'C', venue: 'V', format: 'Two Day', side: 'Home' }
      ];

      const matrix = mergeFixturesIntoMatrix([], mockGames, DEFAULT_TEAM_CONFIGS);
      const dates = matrix.slice(1).map(r => r[0]);
      expect(dates).toEqual(['2026-10-01', '2026-11-01', '2026-12-01']);
    });

    test('guarantees a SINGLE ROW per round even when Sheet returns JS Date objects or DD/MM/YYYY strings', () => {
      const homeGames = parseFixtureCsv(homeCsvContent, DEFAULT_TEAM_CONFIGS);
      const awayGames = parseFixtureCsv(awayCsvContent, DEFAULT_TEAM_CONFIGS);

      // 1. Phase 1 (Home)
      const phase1Matrix = mergeFixturesIntoMatrix([], homeGames, DEFAULT_TEAM_CONFIGS);

      // 2. Simulate Google Sheets converting dates to JS Date objects and DD/MM/YYYY strings
      const sheetStoredMatrix = phase1Matrix.map((row, idx) => {
        if (idx === 0) return row;
        const copy = [...row];
        if (idx % 2 === 0) {
          // e.g. "04/10/2025"
          const parts = copy[0].split('-');
          copy[0] = `${parts[2]}/${parts[1]}/${parts[0]}`;
        } else {
          // JS Date object
          const parts = copy[0].split('-');
          copy[0] = new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
        }
        return copy;
      });

      // 3. Phase 2 (Away) reads back from simulated sheet
      const phase2Matrix = mergeFixturesIntoMatrix(sheetStoredMatrix, awayGames, DEFAULT_TEAM_CONFIGS);

      // Verify that every single round date exists in EXACTLY ONE row
      const roundRows = phase2Matrix.slice(1);
      const uniqueDates = new Set(roundRows.map(r => r[0]));
      expect(roundRows.length).toBe(uniqueDates.size);

      // Total senior match dates for the season is exactly 14 rounds
      expect(roundRows.length).toBe(14);

      // Check Round 1 combines 1st, 2nd, 3rd, 4th, and 5th XI all on row 1
      const round1 = roundRows.find(r => r[0] === '2025-10-04');
      expect(round1).toBeDefined();
      expect(round1[1]).toBe('Round 1');
      expect(round1[3]).toBe('Mitcham - 2nd XI');                     // 1st XI (Home)
      expect(round1[7]).toBe('Bulleen Templestowe - 2nd XI');        // 2nd XI (Away)
      expect(round1[11]).toBe('Box Hill North Super Kings - 4th XI'); // 3rd XI (Home)
      expect(round1[15]).toBe('Nunawading - 5th XI');                // 4th XI (Away)
      expect(round1[19]).toBe('Blackburn North CC - 4th XI');        // 5th XI (Away)
    });
  });
});

