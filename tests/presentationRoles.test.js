const { formatPlayerPresentationName } = require('../src/logic');

describe('formatPlayerPresentationName', () => {
  test('returns empty string for empty inputs', () => {
    expect(formatPlayerPresentationName('', 'Captain')).toBe('');
    expect(formatPlayerPresentationName(null, 'Captain')).toBe('');
  });

  test('formats regular player with no role badges', () => {
    expect(formatPlayerPresentationName('Isaac Wicklein', 'Player')).toBe('Isaac Wicklein');
    expect(formatPlayerPresentationName('Isaac Wicklein', '4. Player')).toBe('Isaac Wicklein');
    expect(formatPlayerPresentationName('Isaac Wicklein', '')).toBe('Isaac Wicklein');
  });

  test('formats Captain correctly', () => {
    expect(formatPlayerPresentationName('Adam Doungas', 'Captain')).toBe('Adam Doungas (C)');
    expect(formatPlayerPresentationName('Adam Doungas', '1. Captain')).toBe('Adam Doungas (C)');
    expect(formatPlayerPresentationName('Adam Doungas', 'Captain (C)')).toBe('Adam Doungas (C)');
  });

  test('formats Vice Captain correctly', () => {
    expect(formatPlayerPresentationName('Neil Kloot', 'VC')).toBe('Neil Kloot (VC)');
    expect(formatPlayerPresentationName('Neil Kloot', 'Vice Captain')).toBe('Neil Kloot (VC)');
    expect(formatPlayerPresentationName('Neil Kloot', '2. VC')).toBe('Neil Kloot (VC)');
  });

  test('formats standalone Wicket Keeper', () => {
    expect(formatPlayerPresentationName('Jimi Kloot', 'WK')).toBe('Jimi Kloot (Wk)');
    expect(formatPlayerPresentationName('Jimi Kloot', 'Wicket Keeper')).toBe('Jimi Kloot (Wk)');
    expect(formatPlayerPresentationName('Jimi Kloot', '3. WK')).toBe('Jimi Kloot (Wk)');
  });

  test('formats Dual Role: Vice Captain & Wicket Keeper (VC & WK)', () => {
    expect(formatPlayerPresentationName('Neil Kloot', 'VC & WK')).toBe('Neil Kloot (VC) (Wk)');
    expect(formatPlayerPresentationName('Neil Kloot', 'VC / WK')).toBe('Neil Kloot (VC) (Wk)');
    expect(formatPlayerPresentationName('Neil Kloot', 'Vice Captain & Wicket Keeper')).toBe('Neil Kloot (VC) (Wk)');
  });

  test('formats Dual Role: Captain & Wicket Keeper (Captain & WK)', () => {
    expect(formatPlayerPresentationName('Adam Doungas', 'Captain & WK')).toBe('Adam Doungas (C) (Wk)');
    expect(formatPlayerPresentationName('Adam Doungas', 'Captain / WK')).toBe('Adam Doungas (C) (Wk)');
  });

  test('strips existing junior tags from player name before applying role', () => {
    expect(formatPlayerPresentationName('Shahmeer Hassaan (U16)', 'WK')).toBe('Shahmeer Hassaan (Wk)');
    expect(formatPlayerPresentationName('Heath Elias (U18)', 'VC & WK')).toBe('Heath Elias (VC) (Wk)');
  });
});
