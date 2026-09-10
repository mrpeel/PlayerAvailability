const { formatNameWithJuniorTag, pickFirstName, stripJuniorTag } = require('../src/logic.js');

describe('stripJuniorTag()', () => {
  test('strips U16 suffix', () => {
    expect(stripJuniorTag('Junior One (U16)')).toBe('Junior One');
  });

  test('strips U18 suffix', () => {
    expect(stripJuniorTag('Junior Two (U18)')).toBe('Junior Two');
  });

  test('strips U14 suffix', () => {
    expect(stripJuniorTag('Johnny Doe (U14)')).toBe('Johnny Doe');
  });

  test('leaves names without suffix unchanged', () => {
    expect(stripJuniorTag('Player One')).toBe('Player One');
    expect(stripJuniorTag('Player Two')).toBe('Player Two');
  });

  test('handles blank or non-string input safely', () => {
    expect(stripJuniorTag('')).toBe('');
    expect(stripJuniorTag(null)).toBe('');
    expect(stripJuniorTag(undefined)).toBe('');
  });
});

describe('formatNameWithJuniorTag()', () => {
  test('blank JuniorLevel returns no suffix', () => {
    expect(formatNameWithJuniorTag('Jane Doe', '')).toBe('Jane Doe');
  });

  test('U18 gets (U18) suffix', () => {
    expect(formatNameWithJuniorTag('Junior Player', 'U18')).toBe('Junior Player (U18)');
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
    expect(pickFirstName('Isaac', 'Isaac Smith')).toBe('Isaac');
    expect(pickFirstName('Matt', 'Matthew')).toBe('Matt');
  });

  test('uses First Name when Preferred Name is empty', () => {
    expect(pickFirstName('James', '')).toBe('James');
    expect(pickFirstName('James', null)).toBe('James');
    expect(pickFirstName('James', undefined)).toBe('James');
  });
});
