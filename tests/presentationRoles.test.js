const {
  formatPlayerPresentationName,
  formatRoundOpponent,
  formatFormatVenue
} = require('../src/logic');

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

describe('formatRoundOpponent', () => {
  test('formats 1st XI correctly with numeric round', () => {
    expect(formatRoundOpponent('1', '1ST', 'Mitcham - 2nd XI'))
      .toBe('Round 1: LCC 1st XI vs Mitcham - 2nd XI');
  });

  test('formats 1st XI correctly with string Round prefix', () => {
    expect(formatRoundOpponent('Round 1', '1ST', 'Mitcham - 2nd XI'))
      .toBe('Round 1: LCC 1st XI vs Mitcham - 2nd XI');
  });

  test('formats other teams (2nd, 3rd, 4th, 5th XI)', () => {
    expect(formatRoundOpponent('2', '2ND', 'Blackburn - 3rd XI'))
      .toBe('Round 2: LCC 2nd XI vs Blackburn - 3rd XI');
    expect(formatRoundOpponent('3', '3RD', 'Box Hill'))
      .toBe('Round 3: LCC 3rd XI vs Box Hill');
    expect(formatRoundOpponent('4', '4TH', 'Surrey Hills'))
      .toBe('Round 4: LCC 4th XI vs Surrey Hills');
    expect(formatRoundOpponent('5', '5TH', 'Vermont'))
      .toBe('Round 5: LCC 5th XI vs Vermont');
  });

  test('handles missing opponent or round gracefully', () => {
    expect(formatRoundOpponent('', '1ST', 'Mitcham'))
      .toBe('LCC 1st XI vs Mitcham');
    expect(formatRoundOpponent('Round 1', '1ST', ''))
      .toBe('Round 1: LCC 1st XI');
  });
});

describe('formatFormatVenue', () => {
  test('formats Two Day game at venue', () => {
    expect(formatFormatVenue('Two Day', 'Kalang Park'))
      .toBe('Two Day game at Kalang Park');
  });

  test('formats One Day game at venue', () => {
    expect(formatFormatVenue('One Day', 'Morton Park'))
      .toBe('One Day game at Morton Park');
  });

  test('avoids duplicate word "game" if already present in format', () => {
    expect(formatFormatVenue('One Day Game', 'Kalang Park'))
      .toBe('One Day Game at Kalang Park');
    expect(formatFormatVenue('T20 Match', 'Laburnum Reserve'))
      .toBe('T20 Match at Laburnum Reserve');
  });

  test('handles missing format or venue gracefully', () => {
    expect(formatFormatVenue('', 'Kalang Park')).toBe('at Kalang Park');
    expect(formatFormatVenue('Two Day', '')).toBe('Two Day game');
  });
});

