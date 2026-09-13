const {
  cleanContactName,
  isFirstNameMatch,
  parseWhatsAppContactsInput,
  matchContactsToPlayers
} = require('../src/logic.js');

describe('cleanContactName()', () => {
  test('strips emojis, club tags, and brackets', () => {
    expect(cleanContactName('Jack Kloot (LCC) 🏏')).toBe('jack kloot');
    expect(cleanContactName('Terry Hall [Coach]')).toBe('terry hall');
    expect(cleanContactName('Ned Clark (U18)')).toBe('ned clark');
  });

  test('converts Last, First format to First Last', () => {
    expect(cleanContactName('Kloot, Jack')).toBe('jack kloot');
    expect(cleanContactName('Smith, Jordan')).toBe('jordan smith');
  });

  test('strips trailing role and team suffixes', () => {
    expect(cleanContactName('Alex Taylor - 1st XI')).toBe('alex taylor');
    expect(cleanContactName('Sam Wilson - Capt')).toBe('sam wilson');
  });
});

describe('isFirstNameMatch()', () => {
  test('matches identical first names', () => {
    expect(isFirstNameMatch('Jack', 'Jack')).toBe(true);
    expect(isFirstNameMatch('jack', 'JACK')).toBe(true);
  });

  test('matches common nicknames in Australia', () => {
    expect(isFirstNameMatch('Gus', 'Augustus')).toBe(true);
    expect(isFirstNameMatch('Dan', 'Daniel')).toBe(true);
    expect(isFirstNameMatch('Sam', 'Samuel')).toBe(true);
    expect(isFirstNameMatch('Matt', 'Matthew')).toBe(true);
    expect(isFirstNameMatch('Chris', 'Christopher')).toBe(true);
    expect(isFirstNameMatch('Tom', 'Thomas')).toBe(true);
  });

  test('matches name prefixes if >= 4 letters', () => {
    expect(isFirstNameMatch('Alexander', 'Alex')).toBe(true);
  });

  test('rejects unrelated names', () => {
    expect(isFirstNameMatch('Jack', 'Sam')).toBe(false);
    expect(isFirstNameMatch('Terry', 'Alex')).toBe(false);
  });
});

describe('parseWhatsAppContactsInput()', () => {
  test('parses standard CSV with header', () => {
    const csv = `Name,Phone,Source\nJack Kloot,0412 345 678,2nd XI\nJordan Smith,0400000002,1st XI`;
    const result = parseWhatsAppContactsInput(csv, 'Default');
    expect(result).toHaveLength(2);
    expect(result[0]).toEqual({
      rawName: 'Jack Kloot',
      phone: '+61412345678',
      source: '2nd XI'
    });
    expect(result[1]).toEqual({
      rawName: 'Jordan Smith',
      phone: '+61400000002',
      source: '1st XI'
    });
  });

  test('parses freeform lines with colon or hyphen', () => {
    const text = `Jack Kloot: 0412 345 678\nTerry Hall - 0400000011\n0422 999 888 - Gus White`;
    const result = parseWhatsAppContactsInput(text, '3rd XI');
    expect(result).toHaveLength(3);
    expect(result[0].rawName).toBe('Jack Kloot');
    expect(result[0].phone).toBe('+61412345678');
    expect(result[0].source).toBe('3rd XI');
    expect(result[1].rawName).toBe('Terry Hall');
    expect(result[2].rawName).toBe('Gus White');
    expect(result[2].phone).toBe('+61422999888');
  });

  test('parses WhatsApp chat export lines with joined or added numbers', () => {
    const chatExport = `
12/09/2026, 14:24 - Messages are end-to-end encrypted
12/09/2026, 14:25 - +61 412 345 678 joined using this group's invite link
12/09/2026, 14:26 - John Smith added +61 498 765 432
12/09/2026, 14:27 - +61 422 111 222: Hello everyone!
    `;
    const result = parseWhatsAppContactsInput(chatExport, 'Chat Export');
    expect(result.length).toBeGreaterThanOrEqual(3);
    const phones = result.map(c => c.phone);
    expect(phones).toContain('+61412345678');
    expect(phones).toContain('+61498765432');
    expect(phones).toContain('+61422111222');
  });

  test('parses vCard content', () => {
    const vcard = `
BEGIN:VCARD
VERSION:3.0
FN:Jack Kloot
TEL;TYPE=CELL:+61422111222
END:VCARD
    `;
    const result = parseWhatsAppContactsInput(vcard, 'VCard');
    expect(result).toHaveLength(1);
    expect(result[0].rawName).toBe('Jack Kloot');
    expect(result[0].phone).toBe('+61422111222');
  });
});

describe('matchContactsToPlayers()', () => {
  const samplePlayers = [
    {
      profileId: 'GUID-NEIL',
      firstName: 'Neil',
      lastName: 'Kloot',
      fullName: 'Neil Kloot',
      phone: '+61411111111',
      phone2: '',
      phone3: '',
      phone4: ''
    },
    {
      profileId: 'GUID-JACK',
      firstName: 'Jack',
      lastName: 'Kloot',
      fullName: 'Jack Kloot',
      // In PlayHQ, son was registered with father's phone in both Phone and Phone2
      phone: '+61411111111',
      phone2: '+61411111111',
      phone3: '',
      phone4: ''
    },
    {
      profileId: 'GUID-GUS',
      firstName: 'Augustus',
      lastName: 'White',
      fullName: 'Augustus White',
      phone: '+61400000012',
      phone2: '',
      phone3: '',
      phone4: ''
    },
    {
      profileId: 'GUID-ALEX-1',
      firstName: 'Alex',
      lastName: 'Taylor',
      fullName: 'Alex Taylor',
      phone: '+61400000001',
      phone2: '',
      phone3: '',
      phone4: ''
    }
  ];

  test('user scenario: junior player registered with father phone gets own mobile added to Phone3', () => {
    const contacts = [
      {
        rawName: 'Jack Kloot',
        phone: '0422 222 222',
        source: '2nd XI'
      }
    ];

    const result = matchContactsToPlayers(contacts, samplePlayers);
    expect(result.summary.matched).toBe(1);
    expect(result.summary.addedToPlayer).toBe(1);

    const match = result.matchedContacts[0];
    expect(match.matchStatus).toBe('Matched');
    expect(match.matchedProfileId).toBe('GUID-JACK');
    expect(match.matchedPlayerName).toBe('Jack Kloot');
    expect(match.matchMethod).toBe('Exact Full Name');
    expect(match.addedToPlayer).toBe('Yes');

    expect(result.playerMutations).toHaveLength(1);
    expect(['Phone2', 'Phone3']).toContain(result.playerMutations[0].slot);
    expect(result.playerMutations[0].phone).toBe('+61422222222');
    expect(result.playerMutations[0].profileId).toBe('GUID-JACK');
  });

  test('identifies numbers already present on the player record without duplicating', () => {
    const contacts = [
      {
        rawName: 'Alex Taylor',
        phone: '+61400000001',
        source: '1st XI'
      }
    ];

    const result = matchContactsToPlayers(contacts, samplePlayers);
    expect(result.summary.matched).toBe(1);
    expect(result.summary.alreadyPresent).toBe(1);
    expect(result.summary.addedToPlayer).toBe(0);

    const match = result.matchedContacts[0];
    expect(match.matchStatus).toBe('Matched');
    expect(match.addedToPlayer).toBe('Already Present');
    expect(result.playerMutations).toHaveLength(0);
  });

  test('matches nicknames (e.g. Gus White -> Augustus White)', () => {
    const contacts = [
      {
        rawName: 'Gus White',
        phone: '0499 888 777',
        source: '4th XI'
      }
    ];

    const result = matchContactsToPlayers(contacts, samplePlayers);
    expect(result.summary.matched).toBe(1);
    expect(result.summary.addedToPlayer).toBe(1);

    const match = result.matchedContacts[0];
    expect(match.matchedProfileId).toBe('GUID-GUS');
    expect(match.matchedPlayerName).toBe('Augustus White');
    expect(match.matchMethod).toBe('Preferred/Nickname Match');
  });

  test('marks unregistered players as Unmatched and links them when player later registers', () => {
    const contacts = [
      {
        rawName: 'David Warner',
        phone: '0431 111 222',
        source: '1st XI'
      }
    ];

    // Initial run: David Warner is not yet registered in PlayHQ
    const initialRun = matchContactsToPlayers(contacts, samplePlayers);
    expect(initialRun.summary.unmatched).toBe(1);
    expect(initialRun.matchedContacts[0].matchStatus).toBe('Unmatched');
    expect(initialRun.playerMutations).toHaveLength(0);

    // 2 weeks later: David Warner registers in PlayHQ
    const updatedPlayers = [
      ...samplePlayers,
      {
        profileId: 'GUID-WARNER',
        firstName: 'David',
        lastName: 'Warner',
        fullName: 'David Warner',
        phone: '+61400000099',
        phone2: '',
        phone3: '',
        phone4: ''
      }
    ];

    // Re-run matching process
    const secondRun = matchContactsToPlayers(contacts, updatedPlayers);
    expect(secondRun.summary.matched).toBe(1);
    expect(secondRun.summary.addedToPlayer).toBe(1);
    expect(secondRun.matchedContacts[0].matchStatus).toBe('Matched');
    expect(secondRun.matchedContacts[0].matchedProfileId).toBe('GUID-WARNER');
    expect(secondRun.playerMutations[0].phone).toBe('+61431111222');
    expect(secondRun.playerMutations[0].slot).toBe('Phone2');
  });

  test('handles manual match override', () => {
    const contacts = [
      {
        rawName: 'Davey',
        phone: '0431 111 222',
        source: '1st XI',
        matchStatus: 'Manual Match',
        matchedProfileId: 'GUID-ALEX-1'
      }
    ];

    const result = matchContactsToPlayers(contacts, samplePlayers);
    expect(result.summary.matched).toBe(1);
    expect(result.matchedContacts[0].matchedProfileId).toBe('GUID-ALEX-1');
    expect(result.matchedContacts[0].matchMethod).toBe('Manual');
  });
});

describe('filterAndEnrichFromAddressBook()', () => {
  const existingPlayers = [
    { profileId: 'GUID-1', firstName: 'Wayne', lastName: 'Cheetham', fullName: 'Wayne Cheetham', phone: '' },
    { profileId: 'GUID-2', firstName: 'Rhett', lastName: 'Orr', fullName: 'Rhett Orr', phone: '' }
  ];

  const existingWhatsAppContacts = [
    { rawName: 'Ross Digby', phone: '' },
    { rawName: 'Angus', phone: '+61487372922' }
  ];

  test('only retains contacts that match a player or WhatsApp member, strictly discarding personal contacts', () => {
    const rawGoogleContacts = [
      { rawName: 'Wayne Cheetham', phone: '0412 111 222', source: 'Google Contacts' },
      { rawName: 'Rhett Orr', phone: '0412 333 444', source: 'Google Contacts' },
      { rawName: 'Ross Digby', phone: '0412 555 666', source: 'Google Contacts' },
      // Personal contacts that MUST be ignored:
      { rawName: 'Dr. John Smith (Dentist)', phone: '0499 000 111', source: 'Google Contacts' },
      { rawName: 'Dave The Plumber', phone: '0499 222 333', source: 'Google Contacts' },
      { rawName: 'Auntie Mary', phone: '0499 444 555', source: 'Google Contacts' }
    ];

    const { filterAndEnrichFromAddressBook } = require('../src/logic.js');
    const result = filterAndEnrichFromAddressBook(rawGoogleContacts, existingPlayers, existingWhatsAppContacts);

    // Only Wayne, Rhett, and Ross are retained
    expect(result.retainedContacts).toHaveLength(3);
    const retainedNames = result.retainedContacts.map(c => c.rawName);
    expect(retainedNames).toContain('Wayne Cheetham');
    expect(retainedNames).toContain('Rhett Orr');
    expect(retainedNames).toContain('Ross Digby');

    // Dentist, Plumber, Auntie Mary are completely discarded
    expect(retainedNames).not.toContain('Dr. John Smith (Dentist)');
    expect(retainedNames).not.toContain('Dave The Plumber');
    expect(retainedNames).not.toContain('Auntie Mary');
    expect(result.skippedPersonalContactsCount).toBe(3);
  });
});

