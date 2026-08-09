const { formatNameWithJuniorTag, pickFirstName } = require('../src/logic.js');

describe('formatNameWithJuniorTag()', () => {
  test('blank JuniorLevel returns no suffix', () => {
    expect(formatNameWithJuniorTag('Jane Doe', '')).toBe('Jane Doe');
  });

  test('U18 gets (U18) suffix', () => {
    expect(formatNameWithJuniorTag('Jimi Kloot', 'U18')).toBe('Jimi Kloot (U18)');
  });

  test('U16 gets (U16) suffix', () => {
    expect(formatNameWithJuniorTag('Player One', 'U16')).toBe('Player One (U16)');
  });

  test('legacy U16_Y2 maps to (U16)', () => {
    expect(formatNameWithJuniorTag('Legacy Y2', 'U16_Y2')).toBe('Legacy Y2 (U16)');
  });

  test('legacy U16_Y1 maps to (U16)', () => {
    expect(formatNameWithJuniorTag('Legacy Y1', 'U16_Y1')).toBe('Legacy Y1 (U16)');
  });

  test('U14 gets (U14) suffix', () => {
    expect(formatNameWithJuniorTag('Jane Doe', 'U14')).toBe('Jane Doe (U14)');
  });

  test('unknown value returns no suffix', () => {
    expect(formatNameWithJuniorTag('Jane Doe', 'SomeOtherValue')).toBe('Jane Doe');
  });
});

describe('pickFirstName()', () => {
  test('prefers Preferred Name when shorter (genuine nickname)', () => {
    expect(pickFirstName('Frederick', 'Freddy')).toBe('Freddy');
    expect(pickFirstName('Benjamin', 'Ben')).toBe('Ben');
    expect(pickFirstName('Augustus', 'Gus')).toBe('Gus');
  });

  test('ignores Preferred Name when same length or longer (redundant copy)', () => {
    expect(pickFirstName('Isaac', 'Isaac Wicklein')).toBe('Isaac');
    expect(pickFirstName('Matt', 'Matthew')).toBe('Matt');
  });

  test('uses First Name when Preferred Name is empty', () => {
    expect(pickFirstName('James', '')).toBe('James');
    expect(pickFirstName('James', null)).toBe('James');
    expect(pickFirstName('James', undefined)).toBe('James');
  });
});
