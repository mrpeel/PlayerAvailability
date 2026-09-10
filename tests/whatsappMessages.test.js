const {
  formatShortPlayerName,
  generateAvailabilityCalloutMessage,
  generateWallOfShameMessage
} = require('../src/logic');

describe('WhatsApp Messages Generator', () => {
  describe('formatShortPlayerName', () => {
    test('formats standard first and last name', () => {
      expect(formatShortPlayerName('Liam Walker')).toBe('Liam W');
      expect(formatShortPlayerName('Daniel Edwards')).toBe('Daniel E');
      expect(formatShortPlayerName('Aaron Adams')).toBe('Aaron A');
    });

    test('strips junior / age group tags', () => {
      expect(formatShortPlayerName('Sam Harris (U16)')).toBe('Sam H');
      expect(formatShortPlayerName('Ned Adams (U18)')).toBe('Ned A');
      expect(formatShortPlayerName('Ian Miller (u16)')).toBe('Ian M');
    });

    test('handles single word names', () => {
      expect(formatShortPlayerName('Cher')).toBe('Cher');
      expect(formatShortPlayerName('Madonna (U16)')).toBe('Madonna');
    });

    test('handles multi-part last names', () => {
      expect(formatShortPlayerName('Jean-Luc Van Damme')).toBe('Jean-Luc V');
    });

    test('handles empty or non-string gracefully', () => {
      expect(formatShortPlayerName('')).toBe('');
      expect(formatShortPlayerName(null)).toBe('');
      expect(formatShortPlayerName(undefined)).toBe('');
    });
  });

  describe('generateAvailabilityCalloutMessage', () => {
    test('generates expected callout message with date', () => {
      const msg = generateAvailabilityCalloutMessage('2025-10-18');
      expect(msg).toContain('🏏 *LABURNUM CC ROUND AVAILABILITY* 🏏');
      expect(msg).toContain('upcoming round (2025-10-18):');
      expect(msg).toContain('https://lcc-availability.web.app/?round=2025-10-18');
    });
  });

  describe('generateWallOfShameMessage', () => {
    test('generates expected wall of shame message with declared count and short names', () => {
      const unknownList = [
        'Liam Walker',
        'Sam Harris (U16)',
        'Harry Edwards (U18)',
        'Daniel Edwards'
      ];
      const msg = generateWallOfShameMessage('2025-10-18', 35, unknownList);

      expect(msg).toContain('✅ Thanks to the 35 players who have confirmed');
      expect(msg).toContain('🏏 Yet to Get off the Mark 🏏');
      expect(msg).toContain("The following players currently won't be troubling the scorers this round: Liam W, Sam H, Harry E, Daniel E");
      expect(msg).toContain('⚡ Declare your availability in 10 seconds:');
      expect(msg).toContain('https://lcc-availability.web.app/?round=2025-10-18');
    });

    test('handles 2D array input from Google Sheets ranges', () => {
      const sheetRange = [
        ['Liam Walker'],
        ['Sam Harris (U16)'],
        [''],
        [null],
        ['Alex Taylor']
      ];
      const msg = generateWallOfShameMessage('2025-10-18', 42, sheetRange);
      expect(msg).toContain('✅ Thanks to the 42 players who have confirmed');
      expect(msg).toContain("The following players currently won't be troubling the scorers this round: Liam W, Sam H, Alex T");
    });

    test('handles zero unknown players gracefully', () => {
      const msg = generateWallOfShameMessage('2025-10-18', 65, []);
      expect(msg).toContain("The following players currently won't be troubling the scorers this round: None! Everyone has responded!");
    });
  });
});
