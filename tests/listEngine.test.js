const { moveItemBetweenLists } = require('../src/logic.js');

describe('moveItemBetweenLists()', () => {
  const sourceUnknown = [
    ['GUID-100', 'Ross Theunissen', ''],
    ['GUID-200', 'Ollie Jenkins', ''],
    ['GUID-300', 'Taylor Merrigan', '']
  ];
  const destAvailable = [];

  test('moves player by ProfileID and updates notes in destination list', () => {
    const { sourceList, destList, moved } = moveItemBetweenLists(
      sourceUnknown,
      destAvailable,
      'GUID-200',
      'Available Week 1 only'
    );

    expect(moved).toBe(true);
    expect(sourceList).toHaveLength(2);
    expect(sourceList.map(r => r[1])).toEqual(['Ross Theunissen', 'Taylor Merrigan']);
    expect(destList).toHaveLength(1);
    expect(destList[0]).toEqual(['GUID-200', 'Ollie Jenkins', 'Available Week 1 only']);
  });

  test('returns original lists unchanged when player not found', () => {
    const { sourceList, destList, moved } = moveItemBetweenLists(sourceUnknown, destAvailable, 'GUID-999');
    expect(moved).toBe(false);
    expect(sourceList).toHaveLength(3);
    expect(destList).toEqual([]);
  });

  test('compacts source list upward (no holes left)', () => {
    const source = [['A', 'Alpha', ''], ['B', 'Bravo', ''], ['C', 'Charlie', '']];
    const { sourceList } = moveItemBetweenLists(source, [], 'B');
    expect(sourceList).toEqual([['A', 'Alpha', ''], ['C', 'Charlie', '']]);
  });

  test('does not mutate the original arrays (pure)', () => {
    const source = [['A', 'Alpha', ''], ['B', 'Bravo', '']];
    const dest = [['C', 'Charlie', '']];
    moveItemBetweenLists(source, dest, 'A');
    expect(source).toHaveLength(2);
    expect(dest).toHaveLength(1);
  });
});
