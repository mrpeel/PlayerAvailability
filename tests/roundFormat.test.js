const { isOneDayFormat, resolveRoundFormat } = require('../src/logic.js');

describe('isOneDayFormat', () => {
  test.each(['One Day', 'one day', '1-Day', 'T20'])('%s is one-day', (f) => {
    expect(isOneDayFormat(f)).toBe(true);
  });

  test.each(['Two Day', '2-Day', '', null, undefined])('%s is not one-day', (f) => {
    expect(isOneDayFormat(f)).toBe(false);
  });
});

describe('resolveRoundFormat', () => {
  test('defaults to Two Day when no formats set', () => {
    expect(resolveRoundFormat({})).toBe('Two Day');
    expect(resolveRoundFormat(null)).toBe('Two Day');
    expect(resolveRoundFormat({ 'Game Date': '2026-10-10', '1st Format': '' })).toBe('Two Day');
  });

  test('all teams one-day resolves to one-day', () => {
    expect(resolveRoundFormat({
      '1st Format': 'One Day',
      '2nd Format': 'One Day',
      '3rd Format': 'One Day'
    })).toBe('One Day');
  });

  test('all teams two-day resolves to two-day', () => {
    expect(resolveRoundFormat({ '1st Format': 'Two Day', '2nd Format': 'Two Day' })).toBe('Two Day');
  });

  test('1sts two-day, others one-day resolves to two-day', () => {
    expect(resolveRoundFormat({
      '1st Format': 'Two Day',
      '2nd Format': 'One Day',
      '3rd Format': 'One Day'
    })).toBe('Two Day');
  });

  test('1sts one-day, a lower team two-day resolves to two-day', () => {
    expect(resolveRoundFormat({
      '1st Format': 'One Day',
      '2nd Format': 'One Day',
      '4th Format': 'Two Day'
    })).toBe('Two Day');
  });

  test('blank 1st Format falls back to other teams', () => {
    expect(resolveRoundFormat({ '1st Format': '', '2nd Format': 'One Day' })).toBe('One Day');
  });

  test('ignores non-format columns', () => {
    expect(resolveRoundFormat({
      '1st Format': 'One Day',
      '1st Opponent': 'Two Day Rovers',
      '1st Venue': 'Two Day Park'
    })).toBe('One Day');
  });
});
