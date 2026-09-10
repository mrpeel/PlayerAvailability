const { findHouseholdPlayers } = require('../src/logic.js');

describe('findHouseholdPlayers()', () => {
  const mockRoster = [
    { ProfileID: 'GUID-1', FullName: 'Alex Taylor', Phone: '0412345678', Phone2: '', Phone3: '', Phone4: '' },
    { ProfileID: 'GUID-2', FullName: 'Jordan Taylor', Phone: '0411111111', Phone2: '0412345678', Phone3: '', Phone4: '' },
    { ProfileID: 'GUID-3', FullName: 'Sam Taylor', Phone: '0422222222', Phone2: '0412345678', Phone3: '0433333333', Phone4: '' },
    { ProfileID: 'GUID-4', FullName: 'Chris Morgan', Phone: '0444444444', Phone2: '', Phone3: '', Phone4: '' }
  ];

  test('matches parent phone across Phone2 and Phone3 for juniors', () => {
    const parentPhone = '0412345678';
    const result = findHouseholdPlayers(mockRoster, parentPhone);

    expect(result).toHaveLength(3); // Alex, Jordan, and Sam
    expect(result.map(p => p.FullName)).toEqual(['Alex Taylor', 'Jordan Taylor', 'Sam Taylor']);
  });

  test('matches formatted input phone (spaces) against stored E164', () => {
    const result = findHouseholdPlayers(mockRoster, '0412 345 678');
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
