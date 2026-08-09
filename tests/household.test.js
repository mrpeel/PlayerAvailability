const { findHouseholdPlayers } = require('../src/logic.js');

describe('findHouseholdPlayers()', () => {
  const mockRoster = [
    { ProfileID: 'GUID-1', FullName: 'Neil Kloot', Phone: '0417663518', Phone2: '', Phone3: '', Phone4: '' },
    { ProfileID: 'GUID-2', FullName: 'Jimi Kloot', Phone: '0479157372', Phone2: '0417663518', Phone3: '', Phone4: '' },
    { ProfileID: 'GUID-3', FullName: 'Harvey Kloot', Phone: '0493674547', Phone2: '0417663518', Phone3: '0422433824', Phone4: '' },
    { ProfileID: 'GUID-4', FullName: 'Aaron Alaimo', Phone: '0412146758', Phone2: '', Phone3: '', Phone4: '' }
  ];

  test('matches parent phone across Phone2 and Phone3 for juniors', () => {
    const parentPhone = '0417663518';
    const result = findHouseholdPlayers(mockRoster, parentPhone);

    expect(result).toHaveLength(3); // Neil, Jimi, and Harvey
    expect(result.map(p => p.FullName)).toEqual(['Neil Kloot', 'Jimi Kloot', 'Harvey Kloot']);
  });

  test('matches formatted input phone (spaces) against stored E164', () => {
    const result = findHouseholdPlayers(mockRoster, '0417 663 518');
    expect(result).toHaveLength(3);
  });

  test('returns empty array if no phone matches', () => {
    const result = findHouseholdPlayers(mockRoster, '0400000000');
    expect(result).toEqual([]);
  });

  test('returns empty array for blank input', () => {
    expect(findHouseholdPlayers(mockRoster, '')).toEqual([]);
    expect(findHouseholdPlayers(mockRoster, null)).toEqual([]);
  });
});
