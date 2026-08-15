const { simulatePlayerAvailability } = require('../src/logic.js');

describe('Player Availability Simulation (simulatePlayerAvailability)', () => {
  const mockPlayers = [
    { profileId: 'P001', fullName: 'Neil Kloot', globalStatus: 'Active' },
    { profileId: 'P002', fullName: 'Jimi Kloot', globalStatus: 'Active' },
    { profileId: 'P003', fullName: 'Sam Jones', globalStatus: 'Active' },
    { profileId: 'P004', fullName: 'Injured Player', globalStatus: 'Injured' },
    { profileId: 'P005', fullName: 'Inactive Player', globalStatus: 'Inactive' }
  ];

  describe('1-Day format (90% available, 10% unavailable)', () => {
    test('marks active player available when rng < 0.90', () => {
      const results = simulatePlayerAvailability(mockPlayers.slice(0, 1), 'One Day', () => 0.50);
      expect(results[0]).toEqual({
        profileId: 'P001',
        fullName: 'Neil Kloot',
        response: 'Available',
        notes: ''
      });
    });

    test('marks active player unavailable when rng >= 0.90', () => {
      const results = simulatePlayerAvailability(mockPlayers.slice(0, 1), 'One Day', () => 0.95);
      expect(results[0]).toEqual({
        profileId: 'P001',
        fullName: 'Neil Kloot',
        response: 'Unavailable',
        notes: 'Unavailable for Round'
      });
    });

    test('always respects Injured and Inactive global status', () => {
      const results = simulatePlayerAvailability(mockPlayers.slice(3), 'One Day', () => 0.10);
      expect(results[0].response).toBe('Unavailable');
      expect(results[0].notes).toBe('Injured');
      expect(results[1].response).toBe('Unavailable');
      expect(results[1].notes).toBe('Inactive in Master');
    });
  });

  describe('2-Day format (85% available, 5% Day 1, 5% Day 2, 5% unavailable)', () => {
    test('marks active player fully available when rng < 0.85', () => {
      const results = simulatePlayerAvailability(mockPlayers.slice(0, 1), 'Two Day', () => 0.80);
      expect(results[0]).toEqual({
        profileId: 'P001',
        fullName: 'Neil Kloot',
        response: 'Available',
        notes: ''
      });
    });

    test('marks active player Day 1 Only when 0.85 <= rng < 0.90', () => {
      const results = simulatePlayerAvailability(mockPlayers.slice(0, 1), 'Two Day', () => 0.87);
      expect(results[0]).toEqual({
        profileId: 'P001',
        fullName: 'Neil Kloot',
        response: 'Available',
        notes: 'Day 1 Only'
      });
    });

    test('marks active player Day 2 Only when 0.90 <= rng < 0.95', () => {
      const results = simulatePlayerAvailability(mockPlayers.slice(0, 1), 'Two Day', () => 0.92);
      expect(results[0]).toEqual({
        profileId: 'P001',
        fullName: 'Neil Kloot',
        response: 'Available',
        notes: 'Day 2 Only'
      });
    });

    test('marks active player Unavailable when rng >= 0.95', () => {
      const results = simulatePlayerAvailability(mockPlayers.slice(0, 1), 'Two Day', () => 0.98);
      expect(results[0]).toEqual({
        profileId: 'P001',
        fullName: 'Neil Kloot',
        response: 'Unavailable',
        notes: 'Unavailable for Round'
      });
    });
  });
});
