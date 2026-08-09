const { normalizePhone } = require('../src/logic.js');

describe('normalizePhone', () => {
  test('returns empty string for falsy input', () => {
    expect(normalizePhone(null)).toBe('');
    expect(normalizePhone(undefined)).toBe('');
    expect(normalizePhone('')).toBe('');
  });

  test('converts AU mobile 04xx to E164 +614xx', () => {
    expect(normalizePhone('0412345678')).toBe('+61412345678');
  });

  test('strips spaces, dashes and dots', () => {
    expect(normalizePhone('0400 123 456')).toBe('+61400123456');
    expect(normalizePhone('0400-123-456')).toBe('+61400123456');
    expect(normalizePhone('0400.123.456')).toBe('+61400123456');
  });

  test('converts bare 9-digit mobile (no leading zero) to E164', () => {
    expect(normalizePhone('412345678')).toBe('+61412345678');
  });

  test('leaves already-E164 numbers untouched', () => {
    expect(normalizePhone('+61412345678')).toBe('+61412345678');
  });

  test('normalizes a leading 04 with country code prefix handling', () => {
    expect(normalizePhone('61412345678')).toBe('+61412345678');
  });

  test('does not corrupt non-AU formats', () => {
    expect(normalizePhone('+447911123456')).toBe('+447911123456');
  });
});
