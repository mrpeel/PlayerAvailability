const {
  formatPlayerPresentationName,
  getPlayerInitials,
  calculateImageSlotBounds,
  calculateHeadGuideMetrics,
  calculateScaleTransform,
  formatRoundOpponent,
  formatFormatVenue,
  formatRoundDate,
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

describe('formatRoundDate', () => {
  test('formats standard ISO YYYY-MM-DD date to "d mmm yy"', () => {
    // Exact user requirement example:
    expect(formatRoundDate('2026-10-03')).toBe('3 Oct 26');
    expect(formatRoundDate('2026-10-31')).toBe('31 Oct 26');
    expect(formatRoundDate('2026-01-05')).toBe('5 Jan 26');
    expect(formatRoundDate('2025-11-16')).toBe('16 Nov 25');
    expect(formatRoundDate('2026-02-09')).toBe('9 Feb 26');
  });

  test('formats multi-date strings taking the first/start match date', () => {
    expect(formatRoundDate('2026-10-03, 2026-10-10')).toBe('3 Oct 26');
    expect(formatRoundDate('2025-11-16, 2025-11-23')).toBe('16 Nov 25');
  });

  test('formats DD/MM/YYYY date strings', () => {
    expect(formatRoundDate('03/10/2026')).toBe('3 Oct 26');
    expect(formatRoundDate('3/10/2026')).toBe('3 Oct 26');
    expect(formatRoundDate('15/02/2026')).toBe('15 Feb 26');
  });

  test('formats JavaScript Date objects', () => {
    // Note: Month 9 in JS Date is October
    const d = new Date(2026, 9, 3);
    expect(formatRoundDate(d)).toBe('3 Oct 26');
  });

  test('handles empty, null, undefined or invalid dates gracefully', () => {
    expect(formatRoundDate('')).toBe('');
    expect(formatRoundDate(null)).toBe('');
    expect(formatRoundDate(undefined)).toBe('');
    expect(formatRoundDate('invalid-date')).toBe('');
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

  test('calculates correct metrics for 90% fill ensuring chin drops and crown stays inside frame', () => {
    const m = calculateHeadGuideMetrics(90);
    expect(m.percent).toBe(90);
    expect(m.targetHeight).toBe(346);
    // Crown at Y=16: comfortably inside the circle (circle top is Y=8, canvas top is Y=0)
    expect(m.topY).toBe(16);
    expect(m.topY).toBeGreaterThan(8);
    // Chin drops to Y=362: accounts for face size without staying stuck at 330
    expect(m.chinY).toBe(362);
    expect(m.chinY).toBeLessThan(392);
    expect(m.chinY - m.topY).toBe(m.targetHeight);
  });

  test('calculates correct metrics across all selectable percentage options (60% to 90%)', () => {
    const options = [60, 65, 70, 75, 80, 85, 90];
    let prevTopY = 999;
    let prevChinY = -1;

    options.forEach(pct => {
      const m = calculateHeadGuideMetrics(pct);
      expect(m.percent).toBe(pct);
      expect(m.chinY - m.topY).toBe(m.targetHeight);
      // Must always remain fully within the circle avatar (Y=8 to Y=392)
      expect(m.topY).toBeGreaterThanOrEqual(8);
      expect(m.chinY).toBeLessThan(392);
      expect(m.eyeY).toBeGreaterThan(m.topY);
      expect(m.eyeY).toBeLessThan(m.chinY);
      expect(m.ovalRx).toBeLessThan(m.ovalRy);

      // As percentage fill increases, crown moves UP (topY decreases) and chin moves DOWN (chinY increases)
      if (prevTopY !== 999) {
        expect(m.topY).toBeLessThan(prevTopY);
        expect(m.chinY).toBeGreaterThan(prevChinY);
      }
      prevTopY = m.topY;
      prevChinY = m.chinY;
    });
  });

  test('handles decimal fraction inputs (e.g. 0.75)', () => {
    const m = calculateHeadGuideMetrics(0.75);
    expect(m.percent).toBe(75);
    expect(m.targetHeight).toBe(288);
    expect(m.topY).toBe(27);
    expect(m.chinY).toBe(315);
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
  test('calculates proportional transform anchored at head (Y=85) by default', () => {
    // Upsize 75% -> 80% with factor 1.067
    const t = calculateScaleTransform(1.067, 'head', 400, 400);
    expect(t.newWidth).toBe(426.8);
    expect(t.newHeight).toBe(426.8);
    expect(t.newX).toBe(-13.4);
    expect(t.newY).toBe(-5.69);

    // Maps 75% crown (Y=27) directly to 80% crown (Y=23)
    expect(Math.round(t.newY + 27 * t.multiplier)).toBe(23);
    // Maps 75% chin (Y=315) directly to 80% chin (Y=330)
    expect(Math.round(t.newY + 315 * t.multiplier)).toBe(330);
  });

  test('calculates enlargement transform anchored at chin (Y=330) when specified', () => {
    const t = calculateScaleTransform(1.067, 'chin', 400, 400);
    expect(t.newWidth).toBe(426.8);
    expect(t.newHeight).toBe(426.8);
    expect(t.newX).toBe(-13.4);
    expect(t.newY).toBe(-22.11);
    expect(Math.round(t.newY + 330 * t.multiplier)).toBe(330);
  });

  test('calculates shrink transform anchored at head (Y=85)', () => {
    // Shrink 85% -> 80% with factor 0.941
    const t = calculateScaleTransform(0.941, 'head', 400, 400);
    expect(t.newWidth).toBe(376.4);
    expect(t.newHeight).toBe(376.4);
    expect(t.newX).toBe(11.8);
    expect(t.newY).toBe(5.02);

    // Maps 85% crown (Y=20) directly to 80% crown (Y=23)
    expect(Math.round(t.newY + 20 * t.multiplier)).toBe(24);
    // Maps 85% chin (Y=346) directly to 80% chin (Y=330)
    expect(Math.round(t.newY + 346 * t.multiplier)).toBe(331);
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
});



