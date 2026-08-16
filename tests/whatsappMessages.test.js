const {
  formatShortPlayerName,
  generateAvailabilityCalloutMessage,
  generateWallOfShameMessage
} = require('../src/logic');

describe('WhatsApp Messages Generator', () => {
  describe('formatShortPlayerName', () => {
    test('formats standard first and last name', () => {
      expect(formatShortPlayerName('Liam Wootten')).toBe('Liam W');
      expect(formatShortPlayerName('Daniel Elias')).toBe('Daniel E');
      expect(formatShortPlayerName('Aaron Alaimo')).toBe('Aaron A');
    });

    test('strips junior / age group tags', () => {
      expect(formatShortPlayerName('Shahmeer Hassaan (U16)')).toBe('Shahmeer H');
      expect(formatShortPlayerName('Neddy Alaimo (U18)')).toBe('Neddy A');
      expect(formatShortPlayerName('Ishan Mallick (u16)')).toBe('Ishan M');
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
        'Liam Wootten',
        'Shahmeer Hassaan (U16)',
        'Heath Elias (U18)',
        'Daniel Elias'
      ];
      const msg = generateWallOfShameMessage('2025-10-18', 35, unknownList);

      expect(msg).toContain('✅ Thanks to the 35 players who have confirmed');
      expect(msg).toContain('🏏 Yet to Get off the Mark 🏏');
      expect(msg).toContain("The following players currently won't be troubling the scorers this round: Liam W, Shahmeer H, Heath E, Daniel E");
      expect(msg).toContain('⚡ Declare your availability in 10 seconds:');
      expect(msg).toContain('https://lcc-availability.web.app/?round=2025-10-18');
    });

    test('handles 2D array input from Google Sheets ranges', () => {
      const sheetRange = [
        ['Liam Wootten'],
        ['Shahmeer Hassaan (U16)'],
        [''],
        [null],
        ['Neil Kloot']
      ];
      const msg = generateWallOfShameMessage('2025-10-18', 42, sheetRange);
      expect(msg).toContain('✅ Thanks to the 42 players who have confirmed');
      expect(msg).toContain("The following players currently won't be troubling the scorers this round: Liam W, Shahmeer H, Neil K");
    });

    test('handles zero unknown players gracefully', () => {
      const msg = generateWallOfShameMessage('2025-10-18', 65, []);
      expect(msg).toContain("The following players currently won't be troubling the scorers this round: None! Everyone has responded!");
    });
  });
});
