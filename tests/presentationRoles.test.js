const {
  formatPlayerPresentationName,
  getPlayerInitials,
  calculateImageSlotBounds,
  calculateHeadGuideMetrics,
  calculateScaleTransform,
  formatRoundOpponent,
  formatFormatVenue,
  parseColorHex,
  normalizeHexColor
} = require('../src/logic');

describe('formatPlayerPresentationName', () => {
  test('returns empty string for empty inputs', () => {
    expect(formatPlayerPresentationName('', 'Captain')).toBe('');
    expect(formatPlayerPresentationName(null, 'Captain')).toBe('');
  });

  test('formats regular player with no role badges', () => {
    expect(formatPlayerPresentationName('Player One', 'Player')).toBe('Player One');
    expect(formatPlayerPresentationName('Player One', '4. Player')).toBe('Player One');
    expect(formatPlayerPresentationName('Player One', '')).toBe('Player One');
  });

  test('formats Captain correctly', () => {
    expect(formatPlayerPresentationName('Player Captain', 'Captain')).toBe('Player Captain (C)');
    expect(formatPlayerPresentationName('Player Captain', '1. Captain')).toBe('Player Captain (C)');
    expect(formatPlayerPresentationName('Player Captain', 'Captain (C)')).toBe('Player Captain (C)');
  });

  test('formats Vice Captain correctly', () => {
    expect(formatPlayerPresentationName('Alex Taylor', 'VC')).toBe('Alex Taylor (VC)');
    expect(formatPlayerPresentationName('Alex Taylor', 'Vice Captain')).toBe('Alex Taylor (VC)');
    expect(formatPlayerPresentationName('Alex Taylor', '2. VC')).toBe('Alex Taylor (VC)');
  });

  test('formats standalone Wicket Keeper', () => {
    expect(formatPlayerPresentationName('Jordan Taylor', 'WK')).toBe('Jordan Taylor (Wk)');
    expect(formatPlayerPresentationName('Jordan Taylor', 'Wicket Keeper')).toBe('Jordan Taylor (Wk)');
    expect(formatPlayerPresentationName('Jordan Taylor', '3. WK')).toBe('Jordan Taylor (Wk)');
  });

  test('formats Dual Role: Vice Captain & Wicket Keeper (VC & WK)', () => {
    expect(formatPlayerPresentationName('Alex Taylor', 'VC & WK')).toBe('Alex Taylor (VC) (Wk)');
    expect(formatPlayerPresentationName('Alex Taylor', 'VC / WK')).toBe('Alex Taylor (VC) (Wk)');
    expect(formatPlayerPresentationName('Alex Taylor', 'Vice Captain & Wicket Keeper')).toBe('Alex Taylor (VC) (Wk)');
  });

  test('formats Dual Role: Captain & Wicket Keeper (Captain & WK)', () => {
    expect(formatPlayerPresentationName('Player Captain', 'Captain & WK')).toBe('Player Captain (C) (Wk)');
    expect(formatPlayerPresentationName('Player Captain', 'Captain / WK')).toBe('Player Captain (C) (Wk)');
  });

  test('strips existing junior tags from player name before applying role', () => {
    expect(formatPlayerPresentationName('Player Junior 1 (U16)', 'WK')).toBe('Player Junior 1 (Wk)');
    expect(formatPlayerPresentationName('Player Junior 2 (U18)', 'VC & WK')).toBe('Player Junior 2 (VC) (Wk)');
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

describe('parseColorHex', () => {
  test('parses default 8-character hex 666666ff correctly', () => {
    const res = parseColorHex('666666ff');
    expect(res).toEqual({ hex: '#666666', alpha: 1, isTransparent: false });
  });

  test('parses with leading # symbol', () => {
    const res = parseColorHex('#666666ff');
    expect(res).toEqual({ hex: '#666666', alpha: 1, isTransparent: false });
  });

  test('parses 6-character hex correctly defaulting to alpha 1.0', () => {
    const res = parseColorHex('666666');
    expect(res).toEqual({ hex: '#666666', alpha: 1, isTransparent: false });
    const white = parseColorHex('ffffff');
    expect(white).toEqual({ hex: '#ffffff', alpha: 1, isTransparent: false });
    const maroon = parseColorHex('#4d0012');
    expect(maroon).toEqual({ hex: '#4d0012', alpha: 1, isTransparent: false });
  });

  test('parses 8-character hex with custom alpha', () => {
    const res = parseColorHex('ffffff80');
    expect(res.hex).toBe('#ffffff');
    expect(res.alpha).toBeCloseTo(0.5, 2);
    expect(res.isTransparent).toBe(false);
  });

  test('detects transparent keyword or 0 alpha', () => {
    expect(parseColorHex('transparent').isTransparent).toBe(true);
    expect(parseColorHex('none').isTransparent).toBe(true);
    expect(parseColorHex('66666600').isTransparent).toBe(true);
  });

  test('handles empty or invalid strings with default fallback', () => {
    expect(parseColorHex('')).toEqual({ hex: '#666666', alpha: 1, isTransparent: false });
    expect(parseColorHex(null)).toEqual({ hex: '#666666', alpha: 1, isTransparent: false });
    expect(parseColorHex('invalid')).toEqual({ hex: '#666666', alpha: 1, isTransparent: false });
  });

  test('respects custom fallback when provided', () => {
    expect(parseColorHex('', 'transparent')).toEqual({ hex: '#000000', alpha: 0, isTransparent: true });
    expect(parseColorHex(null, 'none')).toEqual({ hex: '#000000', alpha: 0, isTransparent: true });
    expect(parseColorHex('', '#fac218')).toEqual({ hex: '#fac218', alpha: 1, isTransparent: false });
    expect(parseColorHex('invalid', 'transparent')).toEqual({ hex: '#000000', alpha: 0, isTransparent: true });
  });
});

describe('normalizeHexColor', () => {
  test('normalizes 8-char hex correctly', () => {
    expect(normalizeHexColor('666666ff')).toBe('666666ff');
    expect(normalizeHexColor('#666666FF')).toBe('666666ff');
  });

  test('appends ff to 6-char hex', () => {
    expect(normalizeHexColor('666666')).toBe('666666ff');
    expect(normalizeHexColor('#ffffff')).toBe('ffffffff');
    expect(normalizeHexColor('4d0012')).toBe('4d0012ff');
  });

  test('handles transparent keyword', () => {
    expect(normalizeHexColor('transparent')).toBe('transparent');
    expect(normalizeHexColor('none')).toBe('transparent');
  });

  test('handles empty or missing input', () => {
    expect(normalizeHexColor('')).toBe('666666ff');
    expect(normalizeHexColor(null)).toBe('666666ff');
  });

  test('respects custom fallback when provided', () => {
    expect(normalizeHexColor('', 'transparent')).toBe('transparent');
    expect(normalizeHexColor(null, 'none')).toBe('transparent');
    expect(normalizeHexColor('', '#fac218')).toBe('fac218ff');
    expect(normalizeHexColor('invalid', 'transparent')).toBe('transparent');
  });
});

describe('getPlayerInitials', () => {
  test('returns empty string for empty/null inputs', () => {
    expect(getPlayerInitials('')).toBe('');
    expect(getPlayerInitials(null)).toBe('');
    expect(getPlayerInitials(undefined)).toBe('');
    expect(getPlayerInitials('   ')).toBe('');
  });

  test('extracts initials for standard two-word player names', () => {
    expect(getPlayerInitials('Isaac Wicklein')).toBe('IW');
    expect(getPlayerInitials('George Doungas')).toBe('GD');
    expect(getPlayerInitials('Adam Doungas')).toBe('AD');
    expect(getPlayerInitials('Boyd Eggleston')).toBe('BE');
    expect(getPlayerInitials('Palash Desai')).toBe('PD');
    expect(getPlayerInitials('Joel Cheetham')).toBe('JC');
    expect(getPlayerInitials('Shahmeer Hassaan')).toBe('SH');
    expect(getPlayerInitials('Derek Taylor')).toBe('DT');
  });

  test('strips role tags and badges in parentheses', () => {
    expect(getPlayerInitials('Aaron Alaimo (C)')).toBe('AA');
    expect(getPlayerInitials('Neil Kloot (VC) (Wk)')).toBe('NK');
    expect(getPlayerInitials('Alex Taylor (VC & WK)')).toBe('AT');
    expect(getPlayerInitials('Jordan Taylor (Wk)')).toBe('JT');
    expect(getPlayerInitials('Player Junior (U16)')).toBe('PJ');
    expect(getPlayerInitials('(C)')).toBe('');
  });

  test('handles single word names', () => {
    expect(getPlayerInitials('Madonna')).toBe('MA');
    expect(getPlayerInitials('A')).toBe('A');
    expect(getPlayerInitials('Jo')).toBe('JO');
  });

  test('handles multi-word names by taking first of first word and first of last word', () => {
    expect(getPlayerInitials('John Paul Smith')).toBe('JS');
    expect(getPlayerInitials('Alexander van der Bilt')).toBe('AB');
  });

  test('handles punctuation, hyphens, and apostrophes', () => {
    expect(getPlayerInitials('Jean-Luc Picard')).toBe('JP');
    expect(getPlayerInitials("Tim O'Connor")).toBe('TO');
    expect(getPlayerInitials('A. Alaimo')).toBe('AA');
    expect(getPlayerInitials('I. Wicklein')).toBe('IW');
  });
});

describe('calculateImageSlotBounds', () => {
  test('calculates correct dimensions and centering for a standard circle with 2px border', () => {
    // 36x36 shape at (100, 100), 2px border
    const bounds = calculateImageSlotBounds(100, 100, 36, 36, 2, false);
    expect(bounds).toEqual({
      left: 100,
      top: 100,
      width: 36,
      height: 36
    });
    // Check center matches exactly: 100 + 18 = 118
    expect(bounds.left + bounds.width / 2).toBe(118);
    expect(bounds.top + bounds.height / 2).toBe(118);
  });

  test('preserves 1:1 shape dimensions regardless of border weight so image is not shrunk', () => {
    const b1 = calculateImageSlotBounds(100, 100, 36, 36, 1, false);
    const b3 = calculateImageSlotBounds(100, 100, 36, 36, 3, false);
    const b4 = calculateImageSlotBounds(100, 100, 36, 36, 4, false);
    expect(b1).toEqual({ left: 100, top: 100, width: 36, height: 36 });
    expect(b3).toEqual({ left: 100, top: 100, width: 36, height: 36 });
    expect(b4).toEqual({ left: 100, top: 100, width: 36, height: 36 });
  });

  test('returns full shape dimensions when border is transparent', () => {
    const bounds = calculateImageSlotBounds(100, 100, 36, 36, 2, true);
    expect(bounds).toEqual({
      left: 100,
      top: 100,
      width: 36,
      height: 36
    });
  });

  test('handles asymmetric shape dimensions preserving square 1:1 ratio and center alignment', () => {
    const bounds = calculateImageSlotBounds(50, 60, 42, 38, 2, false);
    // size = min(42, 38) = 38
    // targetLeft = 50 + (42 - 38) / 2 = 52
    // targetTop = 60 + (38 - 38) / 2 = 60
    expect(bounds).toEqual({
      left: 52,
      top: 60,
      width: 38,
      height: 38
    });
    // Centers match:
    expect(bounds.left + bounds.width / 2).toBe(50 + 42 / 2);
    expect(bounds.top + bounds.height / 2).toBe(60 + 38 / 2);
  });

  test('handles 0 or missing border weight safely', () => {
    const bounds = calculateImageSlotBounds(100, 100, 36, 36, 0, false);
    expect(bounds).toEqual({
      left: 100,
      top: 100,
      width: 36,
      height: 36
    });
  });
});

describe('calculateHeadGuideMetrics', () => {
  test('calculates correct metrics for 80% default fill', () => {
    const m = calculateHeadGuideMetrics(80);
    expect(m.percent).toBe(80);
    expect(m.chinY).toBe(330);
    expect(m.targetHeight).toBe(307);
    expect(m.topY).toBe(23);
    expect(m.eyeY).toBe(169);
    expect(m.ovalRy).toBe(154);
    expect(m.ovalRx).toBe(113);
  });

  test('calculates correct metrics across all selectable percentage options (60% to 90%)', () => {
    const options = [60, 65, 70, 75, 80, 85, 90];
    options.forEach(pct => {
      const m = calculateHeadGuideMetrics(pct);
      expect(m.percent).toBe(pct);
      expect(m.chinY).toBe(330);
      expect(m.topY).toBe(330 - m.targetHeight);
      expect(m.topY).toBeLessThan(m.chinY);
      expect(m.eyeY).toBeGreaterThan(m.topY);
      expect(m.eyeY).toBeLessThan(m.chinY);
      expect(m.ovalRx).toBeLessThan(m.ovalRy);
    });
  });

  test('handles decimal fraction inputs (e.g. 0.75)', () => {
    const m = calculateHeadGuideMetrics(0.75);
    expect(m.percent).toBe(75);
    expect(m.targetHeight).toBe(288);
    expect(m.topY).toBe(42);
  });

  test('clamps out-of-bounds or invalid inputs safely', () => {
    const mInvalid = calculateHeadGuideMetrics('invalid');
    expect(mInvalid.percent).toBe(80);

    const mLow = calculateHeadGuideMetrics(20);
    expect(mLow.percent).toBe(50);

    const mHigh = calculateHeadGuideMetrics(120);
    expect(mHigh.percent).toBe(95);
  });
});

describe('calculateScaleTransform', () => {
  test('calculates enlargement transform anchored at chin (Y=330)', () => {
    // Upsize 75% -> 80% with factor 1.067
    const t = calculateScaleTransform(1.067, 'chin', 400, 400);
    expect(t.newWidth).toBe(426.8);
    expect(t.newHeight).toBe(426.8);
    expect(t.newX).toBe(-13.4);
    expect(t.newY).toBe(-22.11);

    // Verify center remains at 200
    expect(Math.round(t.newX + 200 * t.multiplier)).toBe(200);

    // Verify chin remains at 330
    expect(Math.round(t.newY + 330 * t.multiplier)).toBe(330);
  });

  test('calculates shrink transform anchored at chin (Y=330)', () => {
    // Shrink 85% -> 80% with factor 0.941
    const t = calculateScaleTransform(0.941, 'chin', 400, 400);
    expect(t.newWidth).toBe(376.4);
    expect(t.newHeight).toBe(376.4);
    expect(t.newX).toBe(11.8);
    expect(t.newY).toBe(19.47);

    // Verify center remains at 200
    expect(Math.round(t.newX + 200 * t.multiplier)).toBe(200);

    // Verify chin remains at 330
    expect(Math.round(t.newY + 330 * t.multiplier)).toBe(330);
  });

  test('calculates transform centered at canvas center (200, 200)', () => {
    const t = calculateScaleTransform(1.2, 'center', 400, 400);
    expect(t.newWidth).toBe(480);
    expect(t.newHeight).toBe(480);
    expect(t.newX).toBe(-40);
    expect(t.newY).toBe(-40);

    // Centers match
    expect(t.newX + t.newWidth / 2).toBe(200);
    expect(t.newY + t.newHeight / 2).toBe(200);
  });

  test('handles invalid multiplier safely', () => {
    const t = calculateScaleTransform('invalid', 'chin', 400, 400);
    expect(t.multiplier).toBe(1.0);
    expect(t.newWidth).toBe(400);
    expect(t.newHeight).toBe(400);
    expect(t.newX).toBe(0);
    expect(t.newY).toBe(0);

    const tZero = calculateScaleTransform(0, 'chin');
    expect(tZero.multiplier).toBe(1.0);

    const tNeg = calculateScaleTransform(-0.5, 'chin');
    expect(tNeg.multiplier).toBe(1.0);
  });

  test('verifies crown coordinate transformation between 75% and 80% head fill', () => {
    // 75% fill crown is at Y = 42
    const m75 = calculateHeadGuideMetrics(75);
    expect(m75.topY).toBe(42);

    // 80% fill crown is at Y = 23
    const m80 = calculateHeadGuideMetrics(80);
    expect(m80.topY).toBe(23);

    // Scaling 75% -> 80% (1.067x) anchored at chin (330)
    const tEnlarge = calculateScaleTransform(1.067, 'chin', 400, 400);
    const scaledCrownY = Math.round(tEnlarge.newY + m75.topY * tEnlarge.multiplier);
    expect(scaledCrownY).toBe(m80.topY); // Exactly Y=23!

    // Scaling 85% -> 80% (0.941x) anchored at chin (330)
    const m85 = calculateHeadGuideMetrics(85);
    const tShrink = calculateScaleTransform(0.941, 'chin', 400, 400);
    const shrunkCrownY = Math.round(tShrink.newY + m85.topY * tShrink.multiplier);
    expect(shrunkCrownY).toBe(m80.topY); // Exactly Y=23!
  });
});



