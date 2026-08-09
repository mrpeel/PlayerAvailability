/**
 * ============================================================================
 * Pure Business Logic Layer (logic.js)
 * Laburnum CC Availability Tracker
 *
 * RULE: Functions here MUST NOT call SpreadsheetApp, DriveApp, Ui, or any
 * Google Apps Script service. They take plain JS primitives / arrays /
 * objects and return transformed data, so they run unchanged in Node (Jest).
 *
 * This file is part of the clasp project AND is required by the Jest tests:
 *   - Apps Script: functions are hoisted as globals when pushed to Google.
 *   - Node/Jest:   the guarded module.exports at the bottom enables
 *                  `const { normalizePhone } = require('./logic.js')`.
 * ============================================================================
 */

/**
 * Normalizes phone numbers to E164 format (+614...).
 *   "0400 123 456" -> "+61400123456"
 *   "0412345678"   -> "+61412345678"
 *   "412345678"    -> "+61412345678"
 *   "+61412345678" -> "+61412345678"
 */
function normalizePhone(phone) {
  if (!phone) return "";
  let clean = String(phone).replace(/[^\d]/g, '');

  if (clean.startsWith('04') && clean.length === 10) {
    clean = '61' + clean.slice(1);
  } else if (clean.length === 9 && clean.startsWith('4')) {
    clean = '61' + clean;
  }
  return '+' + clean;
}

/**
 * Returns the household for an input phone: every player whose number appears
 * in ANY of the 4 phone columns (Phone, Phone2, Phone3, Phone4).
 *
 * @param {Array<Object>} playersRows Row objects with Phone/Phone2/Phone3/Phone4.
 * @param {string} inputPhone Raw user input (normalized here).
 * @returns {Array<Object>} Matching player rows.
 */
function findHouseholdPlayers(playersRows, inputPhone) {
  const targetPhone = normalizePhone(inputPhone);
  if (!targetPhone) return [];

  return playersRows.filter(player => {
    const phones = [
      normalizePhone(player.Phone),
      normalizePhone(player.Phone2),
      normalizePhone(player.Phone3),
      normalizePhone(player.Phone4)
    ];
    return phones.includes(targetPhone);
  });
}

/**
 * Moves a 3-cell block [ProfileID, Name, Notes] from one list to another,
 * compacting the source list upward. Pure array manipulation.
 *
 * @param {Array<Array>} sourceList
 * @param {Array<Array>} destList
 * @param {string} targetProfileId
 * @param {string|null} newNotes Optional note update applied on move.
 * @returns {{ sourceList: Array, destList: Array, moved: boolean }}
 */
function moveItemBetweenLists(sourceList, destList, targetProfileId, newNotes = null) {
  const sourceIndex = sourceList.findIndex(row => String(row[0]).trim() === targetProfileId);
  if (sourceIndex === -1) return { sourceList, destList, moved: false };

  // Copy item
  const movedItem = [...sourceList[sourceIndex]];
  if (newNotes !== null) {
    movedItem[2] = newNotes; // Update Notes column
  }

  // Remove from source (compacting)
  const newSource = sourceList.filter((_, idx) => idx !== sourceIndex);

  // Append to destination
  const newDest = [...destList, movedItem];

  return { sourceList: newSource, destList: newDest, moved: true };
}

/**
 * Appends a junior-class suffix tag to the player display name.
 *   Adult       → "First Last"
 *   U18         → "First Last (U18)"
 *   U16 / U16_Y2 / U16_Y1  → "First Last (U16)"
 *   U14         → "First Last (U14)"
 */
function formatNameWithJuniorTag(fullName, juniorLevel) {
  if (juniorLevel === "U18") return fullName + " (U18)";
  if (juniorLevel === "U16" || juniorLevel === "U16_Y2" || juniorLevel === "U16_Y1") return fullName + " (U16)";
  if (juniorLevel === "U14") return fullName + " (U14)";
  return fullName;
}

/**
 * Picks the Preferred Name over First Name if it's genuinely a nickname
 * (shorter than the legal first name and not a First+Last copy).
 */
function pickFirstName(firstName, preferredName) {
  if (!preferredName) return firstName;
  if (preferredName.length < firstName.length) return preferredName;
  return firstName;
}

// Node/Jest interop — Apps Script ignores this guard.
if (typeof module !== 'undefined' && module.exports) {
  module.exports = { normalizePhone, findHouseholdPlayers, moveItemBetweenLists, formatNameWithJuniorTag, pickFirstName };
}
