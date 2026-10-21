/**
 * CAP base access + graduation guest form (Google Forms + Google Sheets + Apps Script)
 *
 * Builds a Google Form that collects what a military installation's security office needs for base access (photo ID,
 * driver's license, SSN, citizenship), plus the graduation guest list, for every attendee of a CAP activity such as an
 * encampment. Then it checks every answer, emails people exactly what to fix, follows up automatically, and produces a
 * masked roster (last 4 digits only) and a full export for installation security.
 *
 * SET UP: see README.md. In short: make a Google Sheet, paste this into Extensions > Apps Script, fill in the SETTINGS
 * below (event, dates, installation, contacts), replace the installation legal text, then run "setup".
 *
 * PROTECTING SSN / DOB / ID NUMBERS
 *   - Responses are NOT copied into the spreadsheet. Full answers stay inside the Google Form.
 *   - Tabs show only the last 4 of SSN and ID numbers, and age instead of date of birth.
 *   - "Create full SECFOR export" makes a separate, private spreadsheet for installation security. Delete it after sending.
 *   - "Delete ALL form data" wipes the responses after the event.
 *
 * Menus: "Base Access" (roster pull, link emails, reminders, fix requests, follow-ups, export) and
 * "Registration reminders" (unit / parent approval emails from Reg Zone exports; optional).
 */

// ---------------------------------------------------------------- settings
var CONFIG = {
  // ---- the event (shown to attendees and parents)
  EVENT_NAME: '20XX Example Wing Winter Encampment',      // used in emails and on the form
  FORM_TITLE: '20XX Example Wing Winter Encampment: Base Access & Graduation Guests',
  STAFF_START: 'Sat 26 Dec 20XX',           // first day for staff
  STUDENT_START: 'Sun 27 Dec 20XX',         // first day for students
  GRADUATION: 'Sun 3 Jan 20XX',             // graduation day (guests come this day)
  EVENT_END_ISO: '20XX-01-03',              // last day, as YYYY-MM-DD: every ID must still be valid on this date
  DUE_DATE: 'Sun 1 Nov 20XX',               // everyone must submit by this date
  GUEST_DEADLINE: 'Sun 15 Nov 20XX',        // last day to add or change graduation guests (blank = "as soon as you know")
  SLOTTING_NOTE: 'Slotting decisions are announced separately.',   // shown so people know the form is not acceptance
  MAX_GUESTS: 8,                            // guest pages built into the form (more pages = slower setup)
  VISIT_DATES: { Staff: '26 Dec 20XX - 3 Jan 20XX', Encampment: '27 Dec 20XX - 3 Jan 20XX', ATS: '27 Dec 20XX - 3 Jan 20XX',
                 Guest: '3 Jan 20XX' },     // dates on base, per event group (same names as REGZONE_FILES), for the security export
  PURPOSE: { attendee: 'CAP 20XX Example Wing Winter Encampment', guest: 'CAP 20XX Example Wing Winter Encampment graduation' },

  // ---- the installation
  INSTALLATION: 'Camp Example',             // the name people know (used everywhere on the form and in emails)
  INSTALLATION_FORMAL: 'Example RFTA',      // the official name (used in the legal text on the last page)

  // ---- who runs it (people email this person; it signs every email)
  CONTACT_NAME: 'Capt Jane Doe',
  CONTACT_TITLE: 'Encampment Administration Officer',
  CONTACT_EMAIL: 'jane.doe@example.org',

  // ---- email limits
  EMAIL_RESERVE: 20,                        // daily email recipients kept free for confirmation and help emails

  // ---- tab names in this spreadsheet
  ROSTER_SHEET: 'Roster',
  SECFOR_SHEET: 'SECFOR Roster (masked)',
  SUBMISSIONS_SHEET: 'Submissions',
  HELP_SHEET: 'Help Requests',
  EXPORT_LOG_SHEET: 'Export log',

  // ---- roster import from CAP Registration Zone (optional). Put the Reg Zone event exports (.xlsx) in one Drive folder, with
  // these file names, and give the folder's ID. Event group names (the keys) become the "Event" column on the Roster tab.
  // Leave REGZONE_FOLDER_ID blank to type or paste the Roster tab yourself (see README).
  REGZONE_FOLDER_ID: '',
  REGZONE_FILES: { Staff: 'STAFF_latest.xlsx', Encampment: 'STUDENTS_latest.xlsx' },
  STAFF_GROUP: 'Staff',                     // which group above is staff (everyone else is a student: slotted, alternate, ...)

  // ---- SSN rule. false: SSN optional (but strongly recommended) with a REAL ID or passport, and required (plus a second ID)
  // with a license that has no star or an "Other" ID. true: SSN required from everyone who gives an ID.
  SSN_REQUIRED_FOR_ALL: false,

  // ---- set to true once INSTALLATION_PRIVACY_ACT, INSTALLATION_RULES and CONSENTS (below) match the host installation's form
  LEGAL_TEXT_REVIEWED: false
};
// People whose base access form is NOT needed (for example their ID is already on file from an earlier activity at the same
// installation, or they have a military / DoD ID card). Their Roster row gets this text in "Base access override" (only if that
// cell is blank; you can also type your own reason there). They are not emailed reminders and count as done.
// Text that mentions military / DoD / CAC sends them a personal link that skips straight to the guest questions.
var ACCESS_OVERRIDES = {
  // '123456': 'ID on file from the fall activity',
  // '234567': 'Not needed: retired military (DoD ID)',
};

var INST = CONFIG.INSTALLATION;
var INST_FORMAL = CONFIG.INSTALLATION_FORMAL;
var CONTACT = CONFIG.CONTACT_NAME + ' (' + CONFIG.CONTACT_EMAIL + ')';
var STUCK = '\n\nStuck? Email ' + CONTACT + '.';
var SIGNATURE = CONFIG.CONTACT_NAME + '\n' + CONFIG.CONTACT_TITLE + ', ' + CONFIG.EVENT_NAME;
// ================================================================ INSTALLATION LEGAL TEXT: REPLACE THIS FOR YOUR INSTALLATION
// Shown on the last page of the form. The text below is an EXAMPLE adapted from one Army installation's visitor access request
// form (Privacy Act statement, numbered rules, and the vehicle note). Copy the host installation's own wording from ITS access
// request form, and ask its security office which IDs it accepts and whether a Social Security number is required.
var INSTALLATION_PRIVACY_ACT =
  'AUTHORITY: Public Law 110-181, The 2008 National Defense Authorization Act, Section 109.\n\n' +
  'PRINCIPAL PURPOSE: Provide necessary information to determine if applicant meets the installation access requirements in accordance ' +
  'with ' + INST_FORMAL + ' Installation Access Procedures for Visitors. Use of Date of Birth, Driver License Number and Social Security Number (SSN) ' +
  'are necessary for positive identification of the applicant in the Department of Justice database.\n\n' +
  'ROUTINE USE: National Crime Information Center Interstate Identification Index (NCIC-III) criminal history check using the Department of Justice database.\n\n' +
  'DISCLOSURE IS VOLUNTARY: Disclosure of SSN and other information is voluntary. However, applicant\'s failure to provide requested ' +
  'information will result in a delay or denial of installation access.\n\n' +
  'Information collected in this form is Controlled Unclassified Information (CUI). Sponsors will collect the required applicant ' +
  'information individually and not provide applicant access to each other\'s protected personal information.\n\n' +
  'On this form, a Social Security number is optional when a REAL ID-compliant ID is provided, but it is HIGHLY ' +
  'RECOMMENDED: failure to provide it may result in a delay or denial of installation access. It is required with an ID that is not ' +
  'REAL ID-compliant.';
var INSTALLATION_RULES =
    '1. I understand the information on this form is being collected in accordance with 50 U.S.C., Section 797, and DoD 5200.8, and federal ' +
  'laws permitting the installation commander to limit access to the installation for security reasons. I have voluntarily completed this ' +
  'form and shall provide the Army a specimen of my fingerprint(s) if requested. The information I have provided on this application is ' +
  'true, complete, and correct to the best of my knowledge, and is provided in good faith. I understand that a knowingly and willfully false ' +
  'statement on this application can be punished by fine or imprisonment or both (18 U.S.C. Section 1001).\n\n' +
  '2. The purpose for requesting information is to assist Access Control personnel in documenting visitor suitability for access to ' + INST_FORMAL + '. ' +
  'Social Security Number and date of birth are necessary to identify the person and records. This information may be used to determine ' +
  'suitability of person desiring access to ' + INST_FORMAL + '; as well as for lawful purposes including law enforcement and litigation. This ' +
  'information will be used to generate state and federal criminal history records checks.\n\n' +
  '3. Persons who do not pass the security vetting will be denied unescorted access and provided with the process for submitting an appeal.\n\n' +
  '4. Any conduct which is detrimental to the maintenance of good order and discipline aboard ' + INST_FORMAL + ' will result in revocation of ' +
  'visitor installation unescorted access.\n\n' +
  '5. Visitors will confine themselves only to those areas related to issuance for which access was granted. Visitors will not enter any ' +
  'marked restricted and/or controlled areas.\n\n' +
  '6. Vehicles and packages may be inspected at any time during a Random Antiterrorism Measure (RAM) while on the installation.\n\n' +
  '7. Any firearm brought onto ' + INST_FORMAL + ', for legitimate activity is required to be registered with the Directorate of Emergency Services (DES).\n\n' +
  '8. Concealed weapons permits are not valid on any Federal Military Installation.\n\n' +
  '9. Medical Marijuana Cards are not recognized by the Federal Government. Possession of Marijuana or any other federally recognized ' +
  'illegal substance is strictly prohibited.\n\n' +
  'If operating a vehicle you must present a valid Driver\'s License, current vehicle registration and proof of valid insurance upon ' +
  'arriving at the Access Control Point.';
// Boxes the person signing must check (all are required).
var CONSENTS = [
    'I have read the Privacy Act Statement.',
  'I understand that every adult and every driver listed will have an NCIC-III criminal history check, and anyone who does not pass will be denied unescorted access.',
  'I understand that vehicles and packages may be inspected at any time while on ' + INST_FORMAL + ', and every driver listed will bring a valid driver\'s license, current registration, and proof of insurance.',
  'Everyone I listed will follow the Rules for Visitors above (restricted areas, firearms, concealed weapons, marijuana).',
  'I swear or affirm that the information I provided is true, complete, and correct to the best of my knowledge, and I understand statement 1 above (including 18 U.S.C. Section 1001).',
  'If I listed anyone other than myself, they have agreed to the background check and to my giving their information, and I have shown them this page.'
];


// ID types: the list of acceptable identity documents on a typical installation access request form (REAL ID-compliant
// license or ID, passports, green card, naturalization certificate, TWIC, etc.), plus DoD ID cards. CHECK THIS AGAINST THE
// HOST INSTALLATION'S ACCESS FORM and edit ALL_TYPES / DRIVER_TYPES if it differs.
var DL_REAL = 'Driver\'s license WITH a star (REAL ID)';
var DL_EDL = 'Enhanced driver\'s license (EDL)';
var DL_NOSTAR_DRV = 'Driver\'s license WITHOUT a star (says "Federal Limits Apply")';
var NOSTAR = 'Driver\'s license or state ID WITHOUT a star (says "Federal Limits Apply")';
var ID_OTHER = 'Other government photo ID (explain in notes)';
var DOD = 'Military / DoD ID card (CAC, retiree, dependent)';   // holders do not need an access request
// Federal PIV card (a federal employee's ID badge): at many installations PIV holders get unescorted access without an NCIC check,
// like DoD ID card holders. A contractor's PIV-I usually does NOT count. CONFIRM with the host installation.
var PIV = 'Federal PIV card (federal employee ID badge)';
function isExemptId_(t) { return t === DOD || t === PIV; }
var DRIVER_TYPES = [DL_REAL, DL_EDL, DL_NOSTAR_DRV, 'U.S. Department of State driver\'s license'];
var ALL_TYPES = [DL_REAL, 'State ID card WITH a star (REAL ID)', DL_EDL, NOSTAR,
  'U.S. passport', 'U.S. passport card', 'Foreign passport', 'Border Crossing Card (DSP-150)',
  'Trusted Traveler card (Global Entry, NEXUS, SENTRI, FAST)', 'Veteran Health Identification Card (VHIC)',
  'Permanent Resident Card ("green card", I-551)', 'Certificate of Naturalization or Citizenship (N-550 or N-570)',
  'Employment Authorization Document (I-766)', 'Refugee Travel Document or Re-entry Permit (I-571 or I-327)',
  'TWIC card', 'Merchant Mariner Credential', 'Federal PIV or PIV-I card', 'U.S. Department of State driver\'s license',
  'Tribal photo ID', DOD, ID_OTHER];
var SSN_REQUIRED_TYPES = [DL_NOSTAR_DRV, NOSTAR, ID_OTHER];
var STATE_ID_TYPES = [DL_REAL, 'State ID card WITH a star (REAL ID)', DL_EDL, DL_NOSTAR_DRV, NOSTAR];   // issued by a state
var US_FED_TYPES = ['U.S. passport', 'U.S. passport card', 'Border Crossing Card (DSP-150)',
  'Trusted Traveler card (Global Entry, NEXUS, SENTRI, FAST)', 'Veteran Health Identification Card (VHIC)',
  'Permanent Resident Card ("green card", I-551)', 'Certificate of Naturalization or Citizenship (N-550 or N-570)',
  'Employment Authorization Document (I-766)', 'Refugee Travel Document or Re-entry Permit (I-571 or I-327)',
  'TWIC card', 'Merchant Mariner Credential', 'U.S. Department of State driver\'s license'];         // issued by the U.S., not a state
var US_STATES = ('AL Alabama|AK Alaska|AZ Arizona|AR Arkansas|CA California|CO Colorado|CT Connecticut|DE Delaware|' +
  'DC District of Columbia|FL Florida|GA Georgia|HI Hawaii|ID Idaho|IL Illinois|IN Indiana|IA Iowa|KS Kansas|KY Kentucky|' +
  'LA Louisiana|ME Maine|MD Maryland|MA Massachusetts|MI Michigan|MN Minnesota|MS Mississippi|MO Missouri|MT Montana|' +
  'NE Nebraska|NV Nevada|NH New Hampshire|NJ New Jersey|NM New Mexico|NY New York|NC North Carolina|ND North Dakota|OH Ohio|' +
  'OK Oklahoma|OR Oregon|PA Pennsylvania|RI Rhode Island|SC South Carolina|SD South Dakota|TN Tennessee|TX Texas|UT Utah|' +
  'VT Vermont|VA Virginia|WA Washington|WV West Virginia|WI Wisconsin|WY Wyoming|PR Puerto Rico|GU Guam|VI Virgin Islands|' +
  'AS American Samoa|MP Northern Mariana Islands').split('|');
// "CA", "Calif.", "California DMV", "State of California", "CA Department of Motor Vehicles" -> true
function isUSState_(v) { return !!usStateCode_(v); }
// Same parsing, returning the 2-letter code ('' if no state): "California DMV" -> "CA", "Sacramento, CA" -> "CA"
function usStateCode_(v) {
  var raw = String(v || '').replace(/[^A-Za-z]/g, '').toLowerCase();   // bare abbreviation first: "ID" (Idaho), "OR", "IN", "ME"
  var hit = '';
  if (raw.length === 2) US_STATES.some(function (x) { if (x.slice(0, 2).toLowerCase() === raw) { hit = x.slice(0, 2); return true; } });
  if (hit) return hit;
  var tail = /,\s*([A-Za-z]{2})\.?\s*$/.exec(String(v || ''));   // "Springfield, ID", "Boise, Id."
  if (tail) US_STATES.some(function (x) { if (x.slice(0, 2).toLowerCase() === tail[1].toLowerCase()) { hit = x.slice(0, 2); return true; } });
  if (hit) return hit;
  var t = String(v || '').replace(/\b(n|no)\.?\s+(dakota|carolina|mariana)/gi, 'north $2').replace(/\bs\.?\s+(dakota|carolina)/gi, 'south $1')
    .replace(/\bw\.?\s+virginia/gi, 'west virginia').replace(/\bw\.?\s*va\b\.?/gi, 'west virginia').replace(/\b(state|of|the|dmv|department|dept|motor|vehicles?|driver'?s?|licen[cs]es?|id|card|real|services?|division|bureau|office|agency)\b/gi, ' ')
    .replace(/[^A-Za-z ]/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase()
    .replace(/\s+(usa|u s a|us|u s|united states( of america)?|america)$/, '');   // "California, USA" -> "california"
  if (!t) return '';
  if (/^(calif|cal|cali)$/.test(t)) return 'CA';
  var best = 0;
  US_STATES.forEach(function (x) {
    var ab = x.slice(0, 2).toLowerCase(), nm = x.slice(3).toLowerCase().replace(/\b(of|the)\b/g, ' ').replace(/\s+/g, ' ').trim();
    var re = new RegExp('(^|\\s)(' + nm + '|' + ab + ')$');   // "Sacramento, CA", "City of Austin Texas"
    var score = (t === ab || t === nm) ? 1000 : (t.indexOf(' ') > 0 && re.test(t)) ? (re.exec(t)[2].length) : 0;
    if (score > best) { best = score; hit = x.slice(0, 2); }   // longest match wins: "West Virginia" -> WV, not VA
  });
  return hit;
}

// The lists for installation security show "ID issued by" and citizenship in one standard form:
// a 2-letter state for licenses and state IDs, "United States" for U.S. federal IDs, and a standard country name otherwise.
// Flags and fix emails still check what the person actually typed.
var COUNTRY_ALIASES = {
  'United States': ['us', 'usa', 'u s', 'u s a', 'united states', 'united states of america', 'america', 'american', 'us government', 'us department of state', 'department of state', 'state department', 'uscis', 'dhs', 'cbp', 'tsa', 'homeland security', 'us dept of state', 'dept of state'],
  'Mexico': ['mexico', 'mx', 'mexican', 'estados unidos mexicanos', 'mexico sre', 'sre'],
  'Canada': ['canada', 'canadian'],
  'China': ['china', 'prc', 'peoples republic of china', 'people s republic of china', 'chinese', 'cn'],
  'Taiwan': ['taiwan', 'roc', 'republic of china', 'taiwan roc', 'taiwanese'],
  'South Korea': ['south korea', 'korea', 'republic of korea', 'rok', 'korea south', 'korean', 'korea republic of', 'kr'],
  'India': ['india', 'indian', 'republic of india', 'in'],
  'Philippines': ['philippines', 'the philippines', 'republic of the philippines', 'filipino', 'ph'],
  'Vietnam': ['vietnam', 'viet nam', 'vietnamese', 'socialist republic of vietnam', 'vn'],
  'Japan': ['japan', 'japanese', 'jp'],
  'United Kingdom': ['united kingdom', 'uk', 'u k', 'great britain', 'britain', 'england', 'scotland', 'wales', 'british', 'gb'],
  'Germany': ['germany', 'deutschland', 'german'], 'France': ['france', 'french'], 'Italy': ['italy', 'italian'], 'Spain': ['spain', 'espana', 'spanish'],
  'Ireland': ['ireland', 'irish', 'republic of ireland'], 'Netherlands': ['netherlands', 'holland', 'the netherlands', 'dutch'],
  'Russia': ['russia', 'russian federation', 'russian'], 'Ukraine': ['ukraine', 'ukrainian'], 'Poland': ['poland', 'polish'],
  'Israel': ['israel', 'israeli'], 'Iran': ['iran', 'islamic republic of iran', 'iranian'], 'Turkey': ['turkey', 'turkiye', 'turkish'],
  'Pakistan': ['pakistan', 'pakistani'], 'Bangladesh': ['bangladesh'], 'Sri Lanka': ['sri lanka', 'srilanka'], 'Nepal': ['nepal'],
  'Afghanistan': ['afghanistan'], 'Iraq': ['iraq'], 'Saudi Arabia': ['saudi arabia', 'ksa'], 'United Arab Emirates': ['united arab emirates', 'uae'],
  'Egypt': ['egypt'], 'Nigeria': ['nigeria', 'nigerian'], 'Ethiopia': ['ethiopia'], 'Kenya': ['kenya'], 'Ghana': ['ghana'], 'South Africa': ['south africa'],
  'Australia': ['australia', 'australian'], 'New Zealand': ['new zealand', 'nz'], 'Singapore': ['singapore'], 'Malaysia': ['malaysia'],
  'Indonesia': ['indonesia'], 'Thailand': ['thailand', 'thai'], 'Cambodia': ['cambodia'], 'Laos': ['laos', 'lao pdr'], 'Myanmar': ['myanmar', 'burma'],
  'Hong Kong': ['hong kong', 'hong kong sar', 'hksar'], 'Mongolia': ['mongolia'], 'Armenia': ['armenia'],
  'El Salvador': ['el salvador', 'salvador'], 'Guatemala': ['guatemala'], 'Honduras': ['honduras'], 'Nicaragua': ['nicaragua'], 'Costa Rica': ['costa rica'],
  'Panama': ['panama'], 'Colombia': ['colombia', 'columbia'], 'Peru': ['peru'], 'Venezuela': ['venezuela'], 'Ecuador': ['ecuador'], 'Brazil': ['brazil', 'brasil'],
  'Argentina': ['argentina'], 'Chile': ['chile'], 'Cuba': ['cuba'], 'Dominican Republic': ['dominican republic'], 'Jamaica': ['jamaica'], 'Haiti': ['haiti']
};
var COUNTRY_LOOKUP_ = null;
function countryStd_(v) {
  var t = String(v || '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase()
    .replace(/\b(passport|passports|government|gov'?t|of the|ministry of foreign affairs|foreign affairs|citizen(ship)?|national|issued|by)\b/g, ' ')
    .replace(/[^a-z ]/g, ' ').replace(/\s+/g, ' ').trim();
  if (!t) return String(v || '').trim();
  if (!COUNTRY_LOOKUP_) {
    COUNTRY_LOOKUP_ = {};
    Object.keys(COUNTRY_ALIASES).forEach(function (k) { COUNTRY_ALIASES[k].forEach(function (a) { COUNTRY_LOOKUP_[a] = k; }); });
  }
  if (COUNTRY_LOOKUP_[t]) return COUNTRY_LOOKUP_[t];
  if (COUNTRY_LOOKUP_[t.replace(/^the /, '')]) return COUNTRY_LOOKUP_[t.replace(/^the /, '')];
  return String(v || '').trim();   // unknown: leave exactly as typed
}
function issuerStd_(v, idType) {
  var raw = String(v || '').trim();
  if (!raw || isExemptId_(idType)) return raw;
  if (STATE_ID_TYPES.indexOf(idType) >= 0) return usStateCode_(raw) || raw;
  if (US_FED_TYPES.indexOf(idType) >= 0 || idType === 'Federal PIV or PIV-I card') return countryStd_(raw) === 'United States' ? 'United States' : raw;
  if (idType === 'Foreign passport') return /^georgia$/i.test(raw) ? 'Georgia (country)' : countryStd_(raw);
  if (idType === 'Tribal photo ID') return raw;
  return usStateCode_(raw) || (countryStd_(raw) !== raw ? countryStd_(raw) : raw);
}
function citizenStd_(v) {
  var raw = String(v || '').trim();
  if (!raw) return raw;
  return isUSCitizenText_(raw) ? 'United States' : countryStd_(raw);
}
// "United States", "USA", "U.S.", "American", "US citizen", "United States of America" -> true
function isUSCitizenText_(v) {
  var t = String(v || '').replace(/\b(citizen(ship)?|national|born|native)\b/gi, '').replace(/[,.]+$/, '').replace(/\s+/g, ' ').trim()
    .replace(/^of\s+/i, '');   // "Citizen of the United States" -> "the United States
  return isUSA_(t) || /^american$/i.test(t);
}
function isUSA_(v) {
  return /^(u\.?\s?s\.?\s?a?\.?|united states( of america)?|america)$/i.test(String(v || '').replace(/\b(the|government|gov'?t)\b/gi, '').trim());
}
var SAME_ADDR = ['Same as the household address in Step 1', 'Different address (type it in the next box)'];
var SAME_PHONE = ['Same as YOUR phone number in Step 1', 'Different number (type it in the next box)'];   // must start with "Same"

var AGE_ATT = { ADULT: '18 or OLDER on the first day of encampment', MINOR: 'UNDER 18 on the first day of encampment' };
// Step 1 question: military / DoD ID and federal PIV card holders skip the ID questions (no access request needed).
// (Replaces the old "ALREADY ON FILE" answer on the age question, which is still understood in older responses.)
var COVER = {
  NONE: 'NO: neither applies (most people)',
  DOD: 'The attendee has a MILITARY / DoD ID card (CAC, retiree, or dependent ID)',
  PIV: 'The attendee has a FEDERAL PIV card (a federal employee ID badge; a contractor PIV-I does not count)'
};
var COVER_TITLE = 'Does the ATTENDEE (the person going to encampment) have a military ID card or a federal PIV card?';
var COVER_HELP = 'This question is about the ATTENDEE, not you.\n' +
  '- MILITARY / DoD ID card: a CAC, retiree, or dependent ID. A veteran ID card or VA health card does NOT count.\n' +
  '- FEDERAL PIV card: a federal employee ID badge. A contractor badge (PIV-I) does NOT count.\n' +
  '- Everyone else (most people): pick NO.\n\n' +
  'If you pick the military or PIV answer, the form skips the attendee\'s ID questions.';
var AGE_PAGE_TITLE = 'Step 1b: Attendee\'s age';
function ageHelp_() {
  return 'About the ATTENDEE, not you. First day of encampment: staff ' + CONFIG.STAFF_START + ', students ' + CONFIG.STUDENT_START +
    '. If you used your personal link, this may already be filled in: check it.';
}
var AGE_GUEST = { ADULT: '18 or OLDER on ' + CONFIG.GRADUATION, MINOR: 'UNDER 18 on ' + CONFIG.GRADUATION };
// Driver wording ("drive onto" confused parents who drive their cadet): the question is whether the person
// is the one BEHIND THE WHEEL. Answers must keep starting with YES / NO (that is how responses are read).
var DRV_ATT = { YES: 'YES: the attendee will drive THEMSELVES onto ' + INST + ' (the attendee is behind the wheel)',
                NO: 'NO: someone else drives the attendee, or the attendee is dropped off, takes the shuttle, or is not coming by car' };
var DRV_GUEST = { YES: 'YES: this guest will be BEHIND THE WHEEL, driving a vehicle onto ' + INST,
                  NO: 'NO: this guest will ride with someone else, or is not coming by car' };
var DRV_DOD = 'MILITARY / DoD ID card holder (CAC, retiree, or dependent ID): nothing else is needed, driving or not';
var DRV_PIV = 'FEDERAL PIV card holder (federal employee ID badge; not a contractor PIV-I): nothing else is needed, driving or not';
var DRV_TITLE = { att: 'Will the ATTENDEE drive THEMSELVES onto ' + INST + '?',
                  guest: 'Will THIS GUEST be the DRIVER (behind the wheel) going onto ' + INST + '?' };
var DRV_Q = { att: 'Will the attendee drive THEMSELVES onto ' + INST + ' (be the one behind the wheel)?',
              guest: 'Will THIS GUEST be the one BEHIND THE WHEEL, driving a vehicle onto ' + INST + ' for graduation?' };
var DRIVER_HELP_ATT =
  'Answer YES only if the attendee will drive THEMSELVES: the attendee is BEHIND THE WHEEL, driving a vehicle through the ' + INST +
  ' gate.\n\n' +
  'Answer NO for everything else, for example:\n' +
  '- riding in a vehicle someone else is driving\n' +
  '- being dropped off or picked up\n' +
  '- taking the shuttle from the park\n' +
  '- not coming by car at all\n' +
  '- the attendee is under 16 (too young to drive)\n\n' +
  'PARENTS: answer about your CADET, not you. If you or anyone else drives your cadet, answer NO.';
var DRIVER_HELP_GUEST =
  'Answer YES only if this guest will be BEHIND THE WHEEL, driving a vehicle through the ' + INST + ' gate for graduation.\n\n' +
  'Only ONE person in each vehicle is the driver. Everyone else in that vehicle answers NO.\n' +
  'Example: Dad drives, Mom and Grandma ride along. Dad = YES. Mom = NO. Grandma = NO.\n' +
  'Guests under 16: answer NO.\n\n' +
  'MILITARY / DoD ID or FEDERAL PIV card holders: pick that answer instead, whether or not they drive.';

var RE_SSN = '^[0-9]{3}-?[0-9]{2}-?[0-9]{4}$';
var RE_ID = '^[A-Za-z0-9][A-Za-z0-9 -]{3,}[A-Za-z0-9]$';
var RE_PHONE = '^[0-9()+. -]{10,}$';
var RE_PLACE = '^[A-Za-z][A-Za-z .,-]{1,}$';
var RE_EMAIL = '^[^@\\s]+@[^@\\s]+\\.[^@\\s]+$';

// ---------------------------------------------------------------- setup (first function, so it is the default in the Run menu)
function setup() {
  // Google stops a script after 6 minutes, and building this form takes longer than that. So setup works in rounds:
  // each run builds for about 4.5 minutes, then stops and asks you to click Run again. It always picks up where it stopped.
  var started = Date.now(), deadline = started + 270000;
  var props = PropertiesService.getScriptProperties();
  var ss = SpreadsheetApp.getActiveSpreadsheet(), ui = SpreadsheetApp.getUi();
  var bad = configProblems_();
  if (bad.length) { ui.alert('Fix these in the SETTINGS at the top of the script first:\n\n- ' + bad.join('\n- ')); return; }
  props.setProperty('SS_ID', ss.getId());
  ensureRoster_(ss);
  var formId = props.getProperty('FORM_ID'), form;
  if (formId) {                                           // form deleted or trashed since the last run: start fresh
    var gone = false;
    try { gone = DriveApp.getFileById(formId).isTrashed(); } catch (e) { gone = true; }
    if (gone) { ['FORM_ID', 'ITEM_MAP', 'SETUP_DONE'].forEach(function (k) { props.deleteProperty(k); }); formId = null; }
  }
  if (!formId) {
    form = FormApp.create(CONFIG.FORM_TITLE);
    props.setProperty('FORM_ID', form.getId());
    var parents = DriveApp.getFileById(ss.getId()).getParents();
    if (parents.hasNext()) DriveApp.getFileById(form.getId()).moveTo(parents.next());
  } else {
    form = FormApp.openById(formId);
  }
  if (!props.getProperty('ITEM_MAP')) {
    try {
      var map = buildForm_(form, deadline);
    } catch (err) {
      if (String(err.message || err) !== 'PAUSE') throw err;
      note_(ss, 'NOT FINISHED YET: ' + form.getItems().length + ' of about 425 questions built. ' +
        'Click Run again (with "setup" selected) to continue. It picks up where it stopped.');
      return;
    }
    props.setProperty('ITEM_MAP', JSON.stringify(map));
    // No form.setDestination(): full answers (SSN, DOB) are deliberately NOT copied into this spreadsheet.
    if (Date.now() - started > 180000) {
      note_(ss, 'NOT FINISHED YET: the form is built. Click Run again (with "setup" selected) once more to finish the last steps.');
      return;
    }
  }
  var msg = updateCapidValidation();
  writePersonalLinks();
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'onFormSubmitHandler') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onFormSubmitHandler').forForm(form).onFormSubmit().create();
  rebuild();
  props.setProperty('SETUP_DONE', new Date().toISOString());
  var done = 'SETUP COMPLETE. Edit the form: ' + form.getEditUrl() + '  General link: ' + form.getPublishedUrl() + '  ' + msg +
    '  Personal links are on the Roster tab. Test the form yourself before sending links.';
  note_(ss, done);
  props.setProperty('FORM_LINKS', done);
}

// Catches settings that were not filled in (the example values), so the form is never built with placeholders.
function configProblems_() {
  var out = [];
  JSON.stringify(CONFIG, function (k, v) {
    if (typeof v === 'string' && /20XX|example\.org|Jane Doe|Camp Example|Example RFTA|Example Wing/.test(v)) out.push(k + ' still has the example value "' + v + '"');
    return v;
  });
  if (!/^\d{4}-\d{2}-\d{2}$/.test(CONFIG.EVENT_END_ISO) || isNaN(new Date(CONFIG.EVENT_END_ISO + 'T00:00:00')))
    out.push('EVENT_END_ISO must be a date written as YYYY-MM-DD, for example 2027-01-03');
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(CONFIG.CONTACT_EMAIL)) out.push('CONTACT_EMAIL is not an email address');
  if (!CONFIG.LEGAL_TEXT_REVIEWED)
    out.push('Replace the example installation legal text (INSTALLATION_PRIVACY_ACT, INSTALLATION_RULES, CONSENTS) with the host ' +
      'installation\'s own wording, then set LEGAL_TEXT_REVIEWED: true');
  return out;
}

// Progress message that does not wait for a click (a pop-up would keep the script running into Google's 6-minute limit).
// Shows in the Apps Script execution log and as a small notice in the spreadsheet.
function note_(ss, msg) {
  console.log(msg);
  try { ss.toast(msg, 'Base Access setup', 60); } catch (e) {}
}

// Wraps the form so building can resume: the first N "add item" calls reuse the N items an earlier run already made
// (their settings are already applied, so setters are skipped). New items are created normally until the deadline.
function resumableForm_(form, deadline) {
  var existing = form.getItems(), k = 0;
  var AS = { addTextItem: 'asTextItem', addParagraphTextItem: 'asParagraphTextItem', addDateItem: 'asDateItem',
    addListItem: 'asListItem', addMultipleChoiceItem: 'asMultipleChoiceItem', addCheckboxItem: 'asCheckboxItem',
    addPageBreakItem: 'asPageBreakItem', addSectionHeaderItem: 'asSectionHeaderItem' };
  var w = {};
  Object.keys(AS).forEach(function (add) {
    w[add] = function () {
      if (k < existing.length) return wrapItem_(existing[k++][AS[add]](), false);
      if (Date.now() > deadline) throw new Error('PAUSE');
      k++;
      return wrapItem_(form[add](), true);
    };
  });
  ['setDescription', 'setCollectEmail', 'setAllowResponseEdits', 'setProgressBar', 'setConfirmationMessage'].forEach(function (m) {
    w[m] = function () { return form[m].apply(form, arguments); };
  });
  return w;
}
function unwrap_(x) { return x && x.__real ? x.__real : x; }
function wrapItem_(real, isNew) {
  var p = { __real: real };
  ['setTitle', 'setRequired', 'setHelpText', 'setValidation', 'setIncludesYear', 'setChoiceValues'].forEach(function (m) {
    p[m] = function () { if (isNew) real[m].apply(real, arguments); return p; };
  });
  p.getId = function () { return real.getId(); };
  p.createChoice = function (v, nav) { return nav === undefined ? real.createChoice(v) : real.createChoice(v, unwrap_(nav)); };
  p.setChoices = function (c) { real.setChoices(c); return p; };
  p.setGoToPage = function (x) { real.setGoToPage(unwrap_(x)); return p; };
  return p;
}

// ---------------------------------------------------------------- one-time fix for an already-built form
// Rewords the driver pages and questions in place (keeps where each answer leads). Safe to run more than once.
// Shown on screen right after Submit.
var CONFIRM_MSG = 'Submitted. Thank you!\n\n' +
  'NOW CHECK YOUR EMAIL. Within a few minutes we will email you a summary and a link to change your answers.\n\n' +
  '- Subject starts with "ACTION NEEDED": something is missing or wrong. The email says exactly what to fix. Use the link in THAT ' +
  'email to fix it. Do NOT fill out a new form.\n' +
  '- Subject starts with "Received" and says COMPLETE: you are done. (If you picked "Not sure yet" for guests, use the link in that ' +
  'email to add them later.)\n\n' +
  'Check that the email names the right attendee. No email within an hour? Check your spam folder, then email ' + CONFIG.CONTACT_EMAIL + '.\n\n' +
  '(If you used this form to ask for HELP, you will get a different email saying we got your message. We will write back.)';

function updateConfirmationMessage() {
  form_().setConfirmationMessage(CONFIRM_MSG);
  console.log('Confirmation message updated.');
  try { ss_().toast('Confirmation message updated.', 'Base Access', 30); } catch (e) {}
}

// ID expiration is REQUIRED. For an ID that truly has no expiration date, people enter 12/31/2099,
// which reports show as "No expiration".
var EXP_HELP = 'Required. The ID must still be valid on ' + CONFIG.GRADUATION + '. If NO expiration date is printed on the ID, enter 12/31/2099.';
function noExpiry_(v) { return /^2099-/.test(String(v || '')); }
function expShow_(v) { return noExpiry_(v) ? 'No expiration' : (v || ''); }
// Sex as printed on the ID; some states print X.
var SEX_CHOICES = ['Male', 'Female', 'X (as shown on the ID)'];
var SEX_HELP = 'Choose what is printed on the ID.';


// Installations generally prefer a REAL ID (star) or a passport. Other IDs work but can be slower at the gate.
var PREFER = 'BEST: a REAL ID (a license or state ID with a star) or a passport. Use one of these if you have it. ' + INST +
  ' security prefers them. The other IDs on this list still work, but they slow down entry.';
var STAR_HOW = 'HOW TO TELL: look at the top corner of the license. A star = REAL ID. No star, or it says "Federal Limits Apply" = NOT ' +
  'REAL ID. A license WITHOUT a star still works, but then you must also give a Social Security number AND a second ID, and bring the ' +
  'second ID to the gate.';
var DL_HELP = 'Pick the driver\'s license this person will use to drive through the gate.\n\n' + STAR_HOW;
var PHOTO_ID_HELP = 'Pick the ONE photo ID this person will show at the gate.\n\n' + PREFER + '\n\n' + STAR_HOW;

// Non-citizens: installation access forms usually ask for the immigration document (type, number, expiration) and the passport
// number. CHECK the host installation's form.
var IMM_TITLE = 'NOT a U.S. citizen only: immigration document type, number, and expiration date, and passport number';
var IMM_HELP = 'REQUIRED if the country of citizenship is NOT the United States. Leave blank for U.S. citizens.\n' +
  'Type all four:\n' +
  '1. the document type, for example: green card, U.S. visa, ESTA approval, or work permit (EAD)\n' +
  '2. its number\n' +
  '3. its expiration date (or "none")\n' +
  '4. the passport number';
// Other plain-language help texts (write for the lowest common denominator). Used by buildForm_ and
// by updateAllWording() on the live form.
var HELP = {
  step1: '"The attendee" means the person registered for encampment (staff or student). Parents: the attendee is YOUR CADET, not you.',
  capid: 'The CAPID of the person GOING to encampment. Parents: your CADET\'s CAPID, not yours. If you opened this form from your ' +
    'personal link, it is already filled in.',
  rosterName: 'Your personal link fills this in. Check that it is the right person. If it is blank, type the attendee\'s rank and ' +
    'name, for example C/Amn Jane Smith.',
  phone: 'At least 10 digits, for example 925-555-0123.',
  addr: 'Where the attendee lives. Your personal link fills this in from CAP records: fix it if it is wrong.',
  ssnOptPage: 'With the ID you picked, the Social Security number is optional, but HIGHLY RECOMMENDED. ' + INST + ' security uses it to ' +
    'identify people for the background check. Leaving it blank may delay or prevent access.',
  ssnOpt: 'Format: 123-45-6789. Leaving it blank may delay or prevent access to ' + INST + '.',
  ssnReqPage: 'REQUIRED: the ID you picked has NO star ("Federal Limits Apply") or is an "Other" ID. ' + INST + ' accepts it ONLY with a ' +
    'Social Security number AND a second ID. Bring the second ID to the gate too.\n\n' +
    'Have a U.S. passport or a REAL ID (star)? Click Back and pick that as the ID instead. Then none of this is needed.',
  doc2Type: 'Type what the second ID is, for example: U.S. passport card, green card, or birth certificate.',
  doc2Number: 'The number printed on the second ID.',
  dlPage: 'The DRIVER (the person behind the wheel at the ' + INST + ' gate) must give their driver\'s license here.',
  detailsPage: 'Copy everything EXACTLY as printed on the ID you picked earlier.',
  idNumber: 'The number printed on the ID, exactly as shown (letters and numbers). This is NOT the CAPID.',
  issuer: 'WHO issued the ID:\n' +
    '- Driver\'s license or state ID: the STATE, for example CA\n' +
    '- U.S. passport, U.S. passport card, green card, or another U.S. government card: United States\n' +
    '- Foreign passport: the COUNTRY, for example Canada\n' +
    'Do NOT type "DMV" or "Passport".',
  citizen: 'For example: United States. If this is NOT the United States, you MUST also fill in the next box.',
  guestPage: 'One guest per page. Parents: if you are coming to graduation, enter yourself as one of the guests.',
  guestAge: 'Guests UNDER 18 who are not driving only need their name and date of birth.'
};
var GUEST_NAME_TITLE = { first: 'Guest\'s first name (as on their ID or birth certificate)',
                         last: 'Guest\'s last name (as on their ID or birth certificate)' };
function guestPageTitle_(g) { return g === 1 ? 'Guest: this is the LAST guest to enter' : 'Guest: ' + g + ' guests left to enter, including this one'; }
function issuerTitle_(Whos) { return Whos + ' ID issued by: which STATE or COUNTRY?'; }

// Wording-only update of the LIVE form: intro, confirmation message, every page's and question's help
// text, and the Step 1 question's title and help. Answer choices keep their wording, so personal links and older responses
// still work. No questions are added, removed or moved. Takes about a minute; safe to repeat.
// Best run when few people are filling out the form (late evening or early morning).
function updateAllWording() {
  updateDriverWording();   // driver pages, questions and answers ("drive THEMSELVES" / "behind the wheel")
  var form = form_(), map = map_(), n = 0;
  form.setDescription(formDescription_());
  form.setConfirmationMessage(CONFIRM_MSG);
  var byField = {
    capid: HELP.capid, rosterName: HELP.rosterName, filledByPhone: HELP.phone, householdAddr: HELP.addr, age: ageHelp_(),
    doc2Type: HELP.doc2Type, doc2Number: HELP.doc2Number, idNumber: HELP.idNumber, idIssuer: HELP.issuer, idExpires: EXP_HELP,
    citizen: HELP.citizen, immDoc: IMM_HELP
  };
  Object.keys(map).forEach(function (id) {
    var m = map[id];
    if (!m || typeof m !== 'object' || !m.f) return;
    var it = form.getItemById(Number(id));
    if (!it) return;
    var t = it.getTitle();
    if (m.f === 'covered') {
      var q = it.asMultipleChoiceItem();
      q.setTitle(COVER_TITLE).setHelpText(COVER_HELP);
      n++; return;
    }
    if (m.f === 'idType') {
      var isDL = /driver's license type/.test(t);
      it.setHelpText(isDL ? DL_HELP : PHOTO_ID_HELP); n++;
      return;
    }
    if (m.f === 'ssn') { if (/highly recommended/i.test(t)) { it.setHelpText(HELP.ssnOpt); n++; } return; }
    if (m.f === 'age' && m.s === 'guest') { it.setHelpText(HELP.guestAge); n++; return; }
    if (byField[m.f] && m.s !== 'help') { it.setHelpText(byField[m.f]); n++; }
  });
  form.getItems(FormApp.ItemType.PAGE_BREAK).forEach(function (it) {
    var t = it.getTitle(), pb = it.asPageBreakItem(), h = null;
    if (t === 'Step 1: Who is this form for?') h = HELP.step1;
    else if (t === AGE_PAGE_TITLE) h = 'About the ATTENDEE, not you.';
    else if (t === 'Graduation guests') h = guestCountHelp_();
    else if (/^Guest: /.test(t)) h = HELP.guestPage;
    else if (/Social Security number \(HIGHLY RECOMMENDED\)$/.test(t)) h = HELP.ssnOptPage;
    else if (/Social Security number and second ID \(REQUIRED\)$/.test(t)) h = HELP.ssnReqPage;
    else if (/ ID details$/.test(t)) h = HELP.detailsPage;
    else if (/^(Attendee|Guest)'s driver's license$/.test(t)) h = HELP.dlPage;
    else if (/military \/ DoD or federal PIV card$/.test(t)) h = DOD_HELP;
    if (h !== null) { pb.setHelpText(h + STUCK); n++; }
  });
  writePersonalLinks();
  var msg = 'Wording updated in ' + n + ' places (intro, confirmation message, pages and questions). Personal links regenerated.';
  console.log(msg); try { ss_().toast(msg, 'Base Access', 30); } catch (e) {}
}

// Quick: rewrites ONLY the form's intro text (no questions or pages change). Safe to run any time.
function updateIntroText() {
  form_().setDescription(formDescription_());
  console.log('Form intro text updated.');
  try { ss_().toast('Form intro text updated.', 'Base Access', 30); } catch (e) {}
}


function updateDriverWording() {
  var form = form_(), map = map_(), n = 0;
  form.getItems(FormApp.ItemType.PAGE_BREAK).forEach(function (it) {
    var t = it.getTitle(), pb = it.asPageBreakItem();
    if (t === DRV_TITLE.att) { pb.setHelpText(DRIVER_HELP_ATT + STUCK); n++; }
    if (t === DRV_TITLE.guest) { pb.setHelpText(DRIVER_HELP_GUEST + STUCK); n++; }
  });
  Object.keys(map).forEach(function (id) {
    var m = map[id];
    if (!m || m.f !== 'drives') return;
    var q = form.getItemById(Number(id)).asMultipleChoiceItem(), TXT = m.s === 'att' ? DRV_ATT : DRV_GUEST;
    var choices = q.getChoices().map(function (c) {
      var v = String(c.getValue()), dest = c.getGotoPage();
      var val = v.indexOf('YES') === 0 ? TXT.YES : v.indexOf('NO') === 0 ? TXT.NO : v;   // keep any other answer (military/DoD) as is
      return dest ? q.createChoice(val, dest) : q.createChoice(val, c.getPageNavigationType());
    });
    q.setTitle(m.s === 'att' ? DRV_Q.att : DRV_Q.guest).setChoices(choices);
    n++;
  });
  console.log('Updated ' + n + ' driver pages and questions.');
  try { ss_().toast('Updated ' + n + ' driver pages and questions.', 'Base Access', 30); } catch (e) {}
}

// ---------------------------------------------------------------- after editing wording in this script
// Pushes the current wording (intro, help texts, confirmation message, driver questions) onto the existing form, then rebuilds
// the tabs. It does not add, remove or move questions. Takes about a minute; run it when few people are filling out the form.
function applyWordingChanges() {
  updateAllWording();
  rebuild();
  try { SpreadsheetApp.getUi().alert('Done: form wording updated and tabs rebuilt.'); } catch (e) {}
}

// Menu version of rebuild(): says when it is finished.
function rebuildNow() {
  var ss = ss_();
  try { ss.toast('Rebuilding tabs...', 'Base Access', 60); } catch (e) {}
  rebuild();
  var n = (ss.getSheetByName(CONFIG.SUBMISSIONS_SHEET) || { getLastRow: function () { return 1; } }).getLastRow() - 1;
  var msg = 'Done: tabs rebuilt at ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'h:mm a') + ' (' + n + ' submissions).';
  try { SpreadsheetApp.getUi().alert(msg); } catch (e) { console.log(msg); }
}

// ---------------------------------------------------------------- menu
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Base Access')
    .addItem('Rebuild tabs now', 'rebuildNow')
    .addItem('Pull new people from Reg Zone exports', 'syncRosterFromRegZone')
    .addItem('Turn on daily Reg Zone pull (6 AM)', 'turnOnDailySync')
    .addItem('Check automation status', 'showAutomationStatus')
    .addSeparator()
    .addItem('Run emails and automation from MY Google account', 'takeOverAutomation')
    .addItem('Email personal links (people not emailed yet)', 'emailPersonalLinks')
    .addItem('Send reminders (staff + slotted students who have not submitted)', 'emailReminders')
    .addItem('Email "ID on file" people about guests', 'emailOnFileAboutGuests')
    .addItem('Email fix requests (submitted forms with problems)', 'emailFixRequests')
    .addItem('Preview follow-ups (staff + slotted; sends nothing)', 'previewFollowUps')
    .addItem('Send follow-ups now (staff + slotted)', 'sendFollowUpsNow')
    .addItem('Turn on automatic follow-ups (Mon + Thu ~9 AM)', 'turnOnAutoFollowUps')
    .addItem('Turn off automatic follow-ups', 'turnOffAutoFollowUps')
    .addItem('Create full SECFOR export (SSN, DOB, ID numbers)', 'createSecforExport')
    .addSeparator()
    .addItem('Apply wording changes (after editing texts in the script)', 'applyWordingChanges')
    .addItem('Update CAPID list from Roster tab', 'updateCapidValidation')
    .addItem('Regenerate personal links', 'writePersonalLinks')
    .addSeparator()
    .addItem('Delete ALL form data (after the event)', 'deleteAllData')
    .addToUi();
  SpreadsheetApp.getUi().createMenu('Registration reminders')
    .addItem('Preview approval reminders (sends nothing)', 'previewApprovalReminders')
    .addItem('Send approval reminders now', 'sendApprovalReminders')
    .addItem('Send a test approval email to me', 'testApprovalReminder')
    .addSeparator()
    .addItem('Turn on weekly approval reminders (Mondays ~8 AM)', 'turnOnWeeklyApprovalReminders')
    .addItem('Turn off weekly approval reminders', 'turnOffWeeklyApprovalReminders')
    .addToUi();
}

// ---------------------------------------------------------------- form builder
function guestBy_() {
  return CONFIG.GUEST_DEADLINE ? ' as soon as you can, and no later than ' + CONFIG.GUEST_DEADLINE : ' as soon as you know';
}
var GUEST_LATER = 'Not sure yet: I will add guests later with the link in my "Received" email';

function formDescription_() {
  return (
    'WHO: EVERYONE registered for the ' + CONFIG.EVENT_NAME + ' (staff and students, including ALTERNATES) must fill out this ' +
    'form ONE time by ' + CONFIG.DUE_DATE + '.\n\n' +
    'WHY: ' + INST + ' is a U.S. Army base. Its security office runs a criminal background check on every adult and every ' +
    'driver before letting them on. This form collects what they need, and asks who is coming to graduation (' + CONFIG.GRADUATION + ').\n\n' +
    'FILL IT OUT EVEN IF:\n' +
    '- nobody is coming to graduation and nobody is driving (it takes about 1 minute)\n' +
    '- the attendee has a MILITARY / DoD ID card or a FEDERAL PIV card (about 1 minute: pick that answer in Step 1 and the form ' +
    'skips the ID questions)\n' +
    '- the attendee is an ALTERNATE or not slotted yet. Background checks take time. Filling out this form does NOT mean the ' +
    'attendee has been slotted or accepted. ' + CONFIG.SLOTTING_NOTE + '\n' +
    '- you do not know your graduation guests yet. Pick "Not sure yet". We will email you a link. Use it to add your guests' + guestBy_() + '.\n\n' +
    'PARENTS: you may fill this out for your cadet.\n' +
    '- "The attendee" always means YOUR CADET, not you.\n' +
    '- If YOU are coming to graduation, you are a GUEST. Enter your own information on the guest pages at the end.\n\n' +
    'DRIVER means the person BEHIND THE WHEEL going through the ' + INST + ' gate. Passengers are NOT drivers. Cadets being dropped off ' +
    'are NOT drivers.\n\n' +
    'HAVE THESE READY:\n' +
    '- The attendee\'s CAPID\n' +
    '- The attendee\'s photo ID, ONLY if the attendee is 18 or older or will drive THEMSELVES onto ' + INST + '\n' +
    '- A photo ID for each guest who is 18 or older or will drive. Best: a REAL ID (star) or a passport.\n' +
    '- Social Security numbers: HIGHLY RECOMMENDED for everyone giving an ID. REQUIRED, together with a second ID, for a license ' +
    'WITHOUT a star ("Federal Limits Apply").\n' +
    '- Drivers: bring a driver\'s license, current vehicle registration, and proof of insurance to the gate.\n\n' +
    'Only list people who agree to the background check. Vehicles and packages may be inspected on the installation. The last page ' +
    'has the full Privacy Act statement and rules.\n\n' +
    'PRIVACY: We use this only to get people onto ' + INST + ' for encampment and graduation. Only the encampment administration director ' +
    'and ' + INST + ' security see the full information, and we delete it after the event.');
}

function buildForm_(realForm, deadline) {
  var form = resumableForm_(realForm, deadline || Date.now() + 3600000);
  var map = {};
  var B = {
    reg: function (item, s, f, g) { map[item.getId()] = { s: s, f: f, g: g || 0 }; return item; },
    txt: function (title, req, re, help, s, f, g) {
      var it = form.addTextItem().setTitle(title).setRequired(!!req);
      if (help) it.setHelpText(help);
      if (re) it.setValidation(FormApp.createTextValidation().setHelpText(help || 'Check the format.')
        .requireTextMatchesPattern(re).build());
      return B.reg(it, s, f, g);
    },
    para: function (title, req, help, s, f, g) {
      var it = form.addParagraphTextItem().setTitle(title).setRequired(!!req);
      if (help) it.setHelpText(help);
      return B.reg(it, s, f, g);
    },
    date: function (title, req, help, s, f, g) {
      var it = form.addDateItem().setTitle(title).setRequired(!!req).setIncludesYear(true);
      if (help) it.setHelpText(help);
      return B.reg(it, s, f, g);
    },
    list: function (title, req, choices, help, s, f, g) {
      var it = form.addListItem().setTitle(title).setRequired(!!req).setChoiceValues(choices);
      if (help) it.setHelpText(help);
      return B.reg(it, s, f, g);
    },
    mc: function (title, req, help, s, f, g) {
      var it = form.addMultipleChoiceItem().setTitle(title).setRequired(!!req).setChoiceValues(['x']);
      if (help) it.setHelpText(help);
      return B.reg(it, s, f, g);
    },
    page: function (title, help) { B.lastPage = form.addPageBreakItem().setTitle(title).setHelpText((help || '') + STUCK); return B.lastPage; }
  };

  form.setDescription(formDescription_());
  form.setCollectEmail(true);
  form.setAllowResponseEdits(true);
  form.setProgressBar(true);
  form.setConfirmationMessage(CONFIRM_MSG);

  // ---- page 1: start
  var start = B.mc('What do you want to do?', true, null, 'start', 'what');

  // ---- help page
  var pHelp = B.page('Get help or report a problem',
    'Tell us what is wrong and we will email you back. Do NOT type Social Security numbers or ID numbers on this page.');
  B.txt('Your name', true, null, null, 'help', 'name');
  B.txt('Your email address', true, RE_EMAIL, 'We will answer by email.', 'help', 'email');
  B.txt('Your phone number (optional)', false, RE_PHONE, 'At least 10 digits.', 'help', 'phone');
  B.txt('The attendee\'s CAPID, if you know it', false, null, 'The CAP ID of the person going to encampment.', 'help', 'capid');
  B.para('What is the problem or question?', true,
    'Tell us what you were trying to do, what page you were on, and any error message you saw.', 'help', 'problem');

  // ---- step 1: who
  var pMain = B.page('Step 1: Who is this form for?', HELP.step1);
  pMain.setGoToPage(FormApp.PageNavigationType.SUBMIT);   // ends the help page (the page before this one)
  var capid = B.txt('Attendee\'s CAPID', true, null, HELP.capid, 'main', 'capid');
  map.CAPID_ITEM = capid.getId();
  var nameItem = B.txt('Attendee\'s rank and name', false, null, HELP.rosterName, 'main', 'rosterName');
  map.NAME_ITEM = nameItem.getId();
  B.list('Who is filling out this form?', true,
    ['I AM the attendee (I am going to encampment)', 'I am the attendee\'s parent or guardian', 'Someone else helping the attendee'],
    null, 'main', 'filledBy');
  B.txt('YOUR full name (the person filling out this form)', true, null, null, 'main', 'filledByName');
  B.txt('YOUR phone number', true, RE_PHONE, HELP.phone, 'main', 'filledByPhone');
  var addr = B.txt('Household home address (street, city, state, ZIP)', true, null, HELP.addr, 'main', 'householdAddr');
  map.ADDR_ITEM = addr.getId();
  var cover = B.mc(COVER_TITLE, true, COVER_HELP, 'main', 'covered');
  map.COVER_ITEM = cover.getId();
  var pAge = B.page(AGE_PAGE_TITLE, 'About the ATTENDEE, not you.');
  map.AGE_PAGE = pAge.getId();
  var attAge = B.mc('How old will the ATTENDEE be on the first day of encampment?', true, ageHelp_(), 'main', 'age');
  map.AGE_ITEM = attAge.getId();

  // ---- is the attendee the driver? (one page for 18+, one for under 18, so the answer can route correctly)
  var attDrv = driverPages_(B, 'att', 0, DRV_TITLE.att, DRV_Q.att, DRIVER_HELP_ATT);

  // ---- attendee ID pages
  var att = personBlock_(B, 'att', 0, 'the attendee', 'Attendee\'s');

  // ---- guest count. Two copies of this page: the first is for paths where NO ID information was given (already on file,
  // military/DoD ID, under 18 and not driving): there "0 guests" submits right away, with no Privacy Act page or attestation.
  // The second is for everyone who gave ID information: "0 guests" goes to the attestation page.
  var pCount0 = B.page('Graduation guests', guestCountHelp_());
  var gcount0 = B.list('How many guests are coming to graduation for this attendee?', true, guestCountChoices_(), null, 'gcount', 'count');
  var pCount = B.page('Graduation guests', guestCountHelp_());
  var countChoices = guestCountChoices_();
  var gcount = B.list('How many guests are coming to graduation for this attendee?', true, countChoices, null, 'gcount', 'count');
  var k;

  // ---- guest slots: slot g means "g guests still to enter, counting this one"
  var basics = {}, ages = {}, drvs = {}, blocks = {};
  for (var g = CONFIG.MAX_GUESTS; g >= 1; g--) {
    basics[g] = B.page(guestPageTitle_(g), HELP.guestPage);
    B.txt(GUEST_NAME_TITLE.first, true, null, null, 'guest', 'first', g);
    B.txt('Guest\'s FULL middle name', false, null, 'Spell out the whole middle name, not just an initial. Leave blank if none.', 'guest', 'middle', g);
    B.txt(GUEST_NAME_TITLE.last, true, null, null, 'guest', 'last', g);
    B.txt('Guest\'s suffix (Jr., III)', false, null, 'Leave blank if none.', 'guest', 'suffix', g);
    B.date('Guest\'s date of birth', true, null, 'guest', 'dob', g);
    B.txt('How is this guest related to the attendee?', true, null, 'For example: mother, grandfather, brother, family friend.', 'guest', 'relationship', g);
    ages[g] = B.mc('How old will THIS GUEST be on ' + CONFIG.GRADUATION + '?', true, HELP.guestAge, 'guest', 'age', g);
    drvs[g] = driverPages_(B, 'guest', g, DRV_TITLE.guest, DRV_Q.guest, DRIVER_HELP_GUEST);
    blocks[g] = personBlock_(B, 'guest', g, 'this guest', 'Guest\'s');
  }

  // ---- final page: Privacy Act statement, notices, consent, attestation (every path except help ends here)
  var pFinal = attestationPage_(form, B);

  // ---- navigation
  start.setChoices([start.createChoice('Fill out the form', pMain),
                    start.createChoice('I need help or want to report a problem with this form', pHelp)]);
  cover.setChoices([cover.createChoice(COVER.NONE, pAge), cover.createChoice(COVER.DOD, att.dod), cover.createChoice(COVER.PIV, att.dod)]);
  attAge.setChoices([attAge.createChoice(AGE_ATT.ADULT, attDrv.adult), attAge.createChoice(AGE_ATT.MINOR, attDrv.minor)]);
  routeDriver_(attDrv, DRV_ATT, att.vehicle, att.idType, pCount0);   // attendee's military/DoD answer is in Step 1
  att.dod.setGoToPage(pCount);                       // after the details page, skip the DoD page
  map['DOD_att0'] = att.dod.getId();
  var S = FormApp.PageNavigationType.SUBMIT;
  gcount.setChoices(countChoices.map(function (v) { return gcount.createChoice(v, /^[1-9]$/.test(v) ? basics[Number(v)] : pFinal); }));
  gcount0.setChoices(countChoices.map(function (v) { return gcount0.createChoice(v, /^[1-9]$/.test(v) ? basics[Number(v)] : S); }));
  map.GCOUNT2_PAGE = pCount0.getId(); map.GCOUNT2_ITEM = gcount0.getId();
  for (g = CONFIG.MAX_GUESTS; g >= 1; g--) {
    var q = ages[g], nextOrSubmit = g > 1 ? basics[g - 1] : pFinal;
    q.setChoices([q.createChoice(AGE_GUEST.ADULT, drvs[g].adult), q.createChoice(AGE_GUEST.MINOR, drvs[g].minor)]);
    routeDriver_(drvs[g], DRV_GUEST, blocks[g].vehicle, blocks[g].idType, nextOrSubmit, blocks[g].dod);
    blocks[g].dod.setGoToPage(nextOrSubmit);
    map['DOD_guest' + g] = blocks[g].dod.getId();
  }
  return map;
}

// Final page: the installation's Privacy Act statement and visitor rules (INSTALLATION_PRIVACY_ACT and INSTALLATION_RULES below).
function guestCountHelp_() {
  return 'Graduation is ' + CONFIG.GRADUATION + '. A GUEST is anyone coming onto ' + INST + ' for graduation who is NOT registered ' +
    'for encampment.\n\n' +
    'HOW TO COUNT:\n' +
    '- Count everyone coming for this attendee, including children and babies.\n' +
    '- Parents: if you are coming, count yourself.\n' +
    '- Example: Mom, Dad, and little brother = 3 guests.\n\n' +
    'WHICH ANSWER TO PICK:\n' +
    '- Do not know yet? Pick "Not sure yet". After you submit, we email you a link. Use that link to add guests' + guestBy_() + '.\n' +
    '- Brothers or sisters also at encampment, or a parent on staff? List the guests on ONE form only. On the other forms, pick ' +
    '"Someone already sent in the guest list for this attendee on another form".\n' +
    '- Nobody coming? Pick 0.\n' +
    '- More than ' + CONFIG.MAX_GUESTS + '? Pick ' + CONFIG.MAX_GUESTS + ', then email ' + CONTACT + ' about the rest.\n\n' +
    'DRIVERS: anyone DRIVING onto ' + INST + ' must bring a driver\'s license, current vehicle registration, and proof of insurance.';
}
// Order: "Not sure yet", "already sent on another form", then 0, 1 ... 8.
var GUEST_ZERO = '0: nobody is coming to graduation for this attendee';
var GUEST_SENT = 'Someone already sent in the guest list for this attendee on another form';
function guestCountChoices_() {
  var c = [GUEST_LATER, GUEST_SENT, GUEST_ZERO];
  for (var k = 1; k <= CONFIG.MAX_GUESTS; k++) c.push(String(k));
  return c;
}


var FINAL_TITLE = 'Last step: read this page, check every box, and type your name';   // must start with "Last step:"
function attestationPage_(form, B) {
  var p = B.page(FINAL_TITLE,
    'Read this page, check every box, and type your name. This is the same information ' + INST_FORMAL + '\'s Installation Access Request Form requires.');
  form.addSectionHeaderItem().setTitle('PRIVACY ACT STATEMENT').setHelpText(
    INSTALLATION_PRIVACY_ACT);
  form.addSectionHeaderItem().setTitle('RULES FOR VISITORS TO ' + INST_FORMAL.toUpperCase()).setHelpText(
    INSTALLATION_RULES);
  B.reg(form.addCheckboxItem().setTitle('Check every box').setRequired(true).setChoiceValues(CONSENTS).setValidation(FormApp.createCheckboxValidation().requireSelectExactly(CONSENTS.length).build()), 'final', 'consents');
  B.txt('Type your full name as your signature', true, null, 'Typing your name here counts as signing this form. Today\'s date is recorded automatically.', 'final', 'signature');
  return p;
}

// Two copies of the driver question: one reached by people 18+, one by people under 18.
function driverPages_(B, s, g, title, question, help) {
  var d = {};
  d.adult = B.page(title, help);
  d.qAdult = B.mc(question, true, null, s, 'drives', g);
  d.minor = B.page(title, help);
  d.qMinor = B.mc(question, true, null, s, 'drives', g);
  return d;
}
// Driver -> vehicle page. Not driving: 18+ -> photo ID page; under 18 -> skip ID (go to notDrivingMinorNext).
function routeDriver_(d, TXT, vehiclePage, idTypePage, notDrivingMinorNext, dodPage) {
  d.qAdult.setChoices([d.qAdult.createChoice(TXT.YES, vehiclePage), d.qAdult.createChoice(TXT.NO, idTypePage)]
    .concat(dodPage ? [d.qAdult.createChoice(DRV_DOD, dodPage), d.qAdult.createChoice(DRV_PIV, dodPage)] : []));
  d.qMinor.setChoices([d.qMinor.createChoice(TXT.YES, vehiclePage), d.qMinor.createChoice(TXT.NO, notDrivingMinorNext)]
    .concat(dodPage ? [d.qMinor.createChoice(DRV_DOD, dodPage), d.qMinor.createChoice(DRV_PIV, dodPage)] : []));
}

/**
 * One person's ID pages, in this order:
 *   vehicle -> driver's license type -> | photo ID type (non-drivers start here) -> SSN optional | SSN + second ID required -> details
 * REAL ID-compliant IDs go to "SSN optional". A license/ID without a star, or an "Other" ID, goes to "SSN + second ID required"
 * (many installations accept a non-REAL ID license only with an SSN and a second ID).
 * After details the form continues to whatever comes next (guest count, next guest, or the final attestation page).
 */
function personBlock_(B, s, g, who, Whos) {
  var b = {};
  b.vehicle = B.page(vehicleTitle_(who),
    'You said ' + who + ' will be the DRIVER (behind the wheel). If that is wrong, click Back and answer NO.\n\n' +
    'Every DRIVER must have a valid driver\'s license, current vehicle registration, and proof of valid insurance in the vehicle.');
  B.list('Registration and insurance', true,
    ['Yes: the vehicle has CURRENT registration and VALID insurance, and proof of both will be in the vehicle'], null, s, 'insConfirm', g);

  b.drvType = B.page(Whos + ' driver\'s license', HELP.dlPage);
  var drv = B.list(Whos + ' driver\'s license type', true, ['x'], DL_HELP, s, 'idType', g);

  b.idType = B.page(Whos + ' photo ID', 'Pick the photo ID ' + who + ' will show at the gate.');
  var idt = B.list(Whos + ' type of photo ID', true, ['x'], PHOTO_ID_HELP, s, 'idType', g);

  if (!CONFIG.SSN_REQUIRED_FOR_ALL) {
    b.ssnOpt = B.page(Whos + ' Social Security number (HIGHLY RECOMMENDED)', HELP.ssnOptPage);
    B.txt(Whos + ' Social Security number (highly recommended)', false, RE_SSN, HELP.ssnOpt, s, 'ssn', g);
  }
  b.ssnReq = B.page(Whos + ' Social Security number and second ID (REQUIRED)', HELP.ssnReqPage);
  B.txt(Whos + ' Social Security number (required)', true, RE_SSN, 'Format: 123-45-6789.', s, 'ssn', g);
  B.txt(Whos + ' second ID: what kind of ID is it?', true, null, HELP.doc2Type, s, 'doc2Type', g);
  B.txt(Whos + ' second ID number', true, null, HELP.doc2Number, s, 'doc2Number', g);

  b.details = B.page(Whos + ' ID details', HELP.detailsPage);
  // (details questions follow; the DoD page is added after them)
  if (s === 'att') {
    B.txt('Attendee\'s first name (exactly as on the ID)', true, null, null, s, 'first', g);
    B.txt('Attendee\'s FULL middle name', false, null, 'Spell out the whole middle name, not just an initial. Leave blank if none.', s, 'middle', g);
    B.txt('Attendee\'s last name (exactly as on the ID)', true, null, null, s, 'last', g);
    B.txt('Attendee\'s suffix (Jr., III)', false, null, 'Leave blank if none.', s, 'suffix', g);
    B.date('Attendee\'s date of birth', true, null, s, 'dob', g);
  }
  B.txt(Whos + ' ID number', true, RE_ID, HELP.idNumber, s, 'idNumber', g);
  B.txt(issuerTitle_(Whos), true, RE_PLACE, HELP.issuer, s, 'idIssuer', g);
  B.date(Whos + ' ID expiration date', true, EXP_HELP, s, 'idExpires', g);
  B.list(Whos + ' sex', true, SEX_CHOICES, SEX_HELP, s, 'sex', g);
  B.txt(Whos + ' country of citizenship', true, RE_PLACE, HELP.citizen, s, 'citizen', g);
  B.txt(IMM_TITLE, false, null, IMM_HELP, s, 'immDoc', g);
  B.list(Whos + ' home address', true, SAME_ADDR, null, s, 'addrSame', g);
  B.txt(Whos + ' home address, ONLY if different (street, city, state, ZIP)', false, null, 'Leave blank if you picked "Same".', s, 'addressOther', g);
  if (s === 'att') {   // guests' phone numbers are not collected
    B.list(Whos + ' home or cell phone', true, SAME_PHONE, null, s, 'phoneSame', g);
    B.txt(Whos + ' home or cell phone, ONLY if different', false, RE_PHONE, 'At least 10 digits. Leave blank if you picked "Same".', s, 'phoneOther', g);
  }
  B.txt(Whos + ' notes for ' + INST + ' security (optional)', false, null,
    'For example: explain an "Other" ID, or a name that is different on two documents.', s, 'notes', g);

  dodQuestions_(B, s, g, Whos);   // b.dod page + its questions, right after the details page
  b.dod = B.lastPage;

  function ssnPage(t) {
    if (t === DOD) return b.dod;
    return (CONFIG.SSN_REQUIRED_FOR_ALL || SSN_REQUIRED_TYPES.indexOf(t) >= 0) ? b.ssnReq : b.ssnOpt;
  }
  drv.setChoices(DRIVER_TYPES.map(function (t) { return drv.createChoice(t, ssnPage(t)); }));
  idt.setChoices(ALL_TYPES.map(function (t) { return idt.createChoice(t, ssnPage(t)); }));
  if (b.ssnOpt) b.ssnReq.setGoToPage(b.details);   // the page before ssnReq is ssnOpt: after "SSN optional", skip "SSN required"
  return b;
}

function vehicleTitle_(who) {
  return who === 'the attendee' ? 'The vehicle the attendee will drive (the attendee is the DRIVER)'
                                : 'The vehicle this guest will drive (this guest is the DRIVER)';
}
// Short page for military / DoD ID and federal PIV card holders: they do not need an access request.
var DRIVER_DOCS = 'If this person will be the DRIVER (behind the wheel) at the ' + INST + ' gate, they must also bring their driver\'s ' +
  'license, current vehicle registration, and proof of insurance: the gate can ask any driver for them.';
var DOD_PAGE_TITLE = ' military / DoD or federal PIV card';
var DOD_HELP = 'Military / DoD ID card holders (CAC, retiree, or dependent ID) and federal PIV card holders (federal employee ID badge) ' +
  'do NOT need a ' + INST + ' access request. Nothing else is needed on this page, even if driving. Just show the card at the gate.\n\n' +
  DRIVER_DOCS + '\n\nClick Next to continue.';
function dodQuestions_(B, s, g, Whos) {
  B.page(Whos + DOD_PAGE_TITLE, DOD_HELP);       // information only: no questions
}






// ---------------------------------------------------------------- helpers
function props_() { return PropertiesService.getScriptProperties(); }
function ss_() { return SpreadsheetApp.openById(props_().getProperty('SS_ID')); }
function form_() { return FormApp.openById(props_().getProperty('FORM_ID')); }
function map_() { return JSON.parse(props_().getProperty('ITEM_MAP')); }

// The Roster tab: one row per registered attendee. Filled by "Pull new people from Reg Zone exports", or typed / pasted by hand.
var ROSTER_COLUMNS = ['CAPID', 'Rank', 'Last name', 'First name', 'Event', 'Member type', 'Birth month/year (Reg Zone)',
  'Age at event start', 'Must submit own ID (18+ at start)', 'Email', 'Parent email', 'Unit', 'Home address (Reg Zone)', 'Slot status'];
function ensureRoster_(ss) {
  if (ss.getSheetByName(CONFIG.ROSTER_SHEET)) return;
  var sh = ss.insertSheet(CONFIG.ROSTER_SHEET);
  sh.getRange(1, 1, 1, ROSTER_COLUMNS.length).setValues([ROSTER_COLUMNS]).setFontWeight('bold');
}
function rosterRows_() {
  var sh = ss_().getSheetByName(CONFIG.ROSTER_SHEET);
  var v = sh.getDataRange().getValues();
  var h = v.shift().map(String);
  return { sheet: sh, header: h, rows: v };
}
function ensureCol_(sheet, header, name) {
  var i = header.indexOf(name);
  if (i >= 0) return i + 1;
  var c = header.length + 1;
  sheet.getRange(1, c).setValue(name).setFontWeight('bold');
  header.push(name);
  return c;
}
function rowObj_(h, x) { var o = {}; h.forEach(function (k, j) { o[k] = x[j]; }); return o; }
function rosterObj_(r) {
  var out = {};
  r.rows.forEach(function (x, i) { var o = rowObj_(r.header, x); o._row = i; out[String(o.CAPID).trim()] = o; });
  return out;
}
function isAdult_(o) { return String(o['Must submit own ID (18+ at start)']) === 'Yes'; }
function rosterName_(o) { return o ? [o.Rank, nameCase_(o['First name']), nameCase_(o['Last name'])].join(' ') : ''; }
function last4_(v) { var d = String(v || '').replace(/[^A-Za-z0-9]/g, ''); return d ? '...' + d.slice(-4) : ''; }
function ageOn_(dobIso) {
  if (!dobIso) return '';
  var d = new Date(dobIso + 'T00:00:00'), g = new Date(CONFIG.EVENT_END_ISO + 'T00:00:00');
  return g.getFullYear() - d.getFullYear() - ((g.getMonth() < d.getMonth() || (g.getMonth() === d.getMonth() && g.getDate() < d.getDate())) ? 1 : 0);
}

// ---------------------------------------------------------------- CAPID validation + personal links
function updateCapidValidation() {
  var map = map_();
  var r = rosterRows_();
  var ci = r.header.indexOf('CAPID');
  var ids = r.rows.map(function (x) { return String(x[ci]).trim(); }).filter(function (x) { return /^[0-9]{5,7}$/.test(x); });
  var item = form_().getItemById(map.CAPID_ITEM).asTextItem();
  if (!ids.length) {
    item.setValidation(FormApp.createTextValidation().setHelpText('Enter a 6-digit CAPID.').requireTextMatchesPattern('^[0-9]{5,7}$').build());
    return 'The Roster tab has no CAPIDs yet, so the CAPID question accepts any CAPID. Fill in the roster, then run "Update CAPID list from Roster tab".';
  }
  try {
    item.setValidation(FormApp.createTextValidation()
      .setHelpText('This CAPID is not on the list of registered attendees. Check the number. Parents: enter your CADET\'s CAPID. Still stuck? Email ' + CONTACT + '.')
      .requireTextMatchesPattern('^(' + ids.join('|') + ')$').build());
    return 'CAPID question locked to the ' + ids.length + ' CAPIDs on the Roster tab.';
  } catch (err) {
    item.setValidation(FormApp.createTextValidation().setHelpText('Enter a 6-digit CAPID.')
      .requireTextMatchesPattern('^[0-9]{6}$').build());
    return 'WARNING: Google would not accept the full CAPID list (' + err + '). The question now accepts any 6-digit CAPID; ' +
      'unknown CAPIDs are flagged on the Submissions tab.';
  }
}

function writePersonalLinks() {
  // Builds two template links (18+ and under 18) once, then fills in each person's values. Much faster than one API call per person.
  var form = form_(), map = map_();
  var capItem = form.getItemById(map.CAPID_ITEM).asTextItem();
  var ageItem = form.getItemById(map.AGE_ITEM).asMultipleChoiceItem();
  var nameItem = form.getItemById(map.NAME_ITEM).asTextItem();
  var addrItem = form.getItemById(map.ADDR_ITEM).asTextItem();
  var P = { cap: 'ZZCAPIDZZ', name: 'ZZNAMEZZ', addr: 'ZZADDRZZ' };
  var covItem = map.COVER_ITEM ? form.getItemById(Number(map.COVER_ITEM)).asMultipleChoiceItem() : null;
  function template(age, cover) {
    var r = form.createResponse().withItemResponse(capItem.createResponse(P.cap)).withItemResponse(nameItem.createResponse(P.name))
      .withItemResponse(addrItem.createResponse(P.addr));
    if (covItem && cover) r = r.withItemResponse(covItem.createResponse(cover));
    if (age) r = r.withItemResponse(ageItem.createResponse(age));
    return r.toPrefilledUrl();
  }
  var T = { adult: template(AGE_ATT.ADULT, COVER.NONE), minor: template(AGE_ATT.MINOR, COVER.NONE), dod: template(null, COVER.DOD) };
  function fill(url, vals) {
    var q = url.indexOf('?'), base = url.slice(0, q + 1);
    var parts = url.slice(q + 1).split('&').map(function (kv) {
      var k = kv.split('=')[0], v = kv.slice(k.length + 1);
      if (v === P.cap) return k + '=' + encodeURIComponent(vals.cap);
      if (v === P.name) return k + '=' + encodeURIComponent(vals.name);
      if (v === P.addr) return vals.addr ? k + '=' + encodeURIComponent(vals.addr) : null;
      return kv;
    }).filter(function (x) { return x !== null; });
    return base + parts.join('&');
  }
  var r = rosterRows_(), h = r.header;
  var col = ensureCol_(r.sheet, h, 'Personal form link');
  var out = r.rows.map(function (x) {
    var o = rowObj_(h, x);
    var ov = String(o['Base access override'] || '').trim();
    var t = /military|DoD|CAC/i.test(ov) ? T.dod : !isAdult_(o) ? T.minor : T.adult;
    return [fill(t, { cap: String(o.CAPID).trim(), name: rosterName_(o) + ' (' + o.Event + ')',
      addr: String(o['Home address (Reg Zone)'] || '').trim() })];
  });
  if (out.length) r.sheet.getRange(2, col, out.length, 1).setValues(out);
}

function linkEmailBody_(o) {
  var who = rosterName_(o);
  return 'Hello,\n\n' + who + ' is registered for the ' + CONFIG.EVENT_NAME + ' (' + o.Event + ') at ' + INST + '.\n\n' +
    'WHAT YOU NEED TO DO: fill out the ' + INST + ' base access form ONE time, by ' + CONFIG.DUE_DATE + '. Use this link. It already ' +
    'has ' + who + '\'s CAPID filled in:\n' + o['Personal form link'] + '\n\n' +
    'EVERYONE must fill it out, even if:\n' +
    '- nobody is coming to graduation and nobody is driving (it takes about 1 minute)\n' +
    '- the attendee has a MILITARY / DoD ID card or a FEDERAL PIV card (about 1 minute: pick that answer in Step 1 and the form ' +
    'skips the ID questions)\n' +
    '- the attendee is an ALTERNATE or not slotted yet. ' + INST + ' security needs time for background checks. Filling out the form ' +
    'does NOT mean the attendee has been slotted or accepted. ' + CONFIG.SLOTTING_NOTE + '\n\n' +
    'PARENTS: you may fill it out for your cadet. On the form, "the attendee" means your cadet. If you are coming to graduation (' +
    CONFIG.GRADUATION + '), you are a GUEST: enter your own information on the guest pages at the end.\n\n' +
    'DO NOT KNOW YOUR GRADUATION GUESTS YET? Fill out the form now and pick "Not sure yet". We will email you a link. Use that link ' +
    'to add your guests' + guestBy_() + '. Do NOT fill out a second form.\n\n' +
    'BROTHERS OR SISTERS ALSO AT ENCAMPMENT, OR A PARENT ON STAFF? Each attendee needs their own form, but list the graduation ' +
    'guests on ONE form only. On the other forms, pick "Someone already sent in the guest list for this attendee on another form".\n\n' +
    'Questions? Reply to this email.\n\n' + SIGNATURE;
}

function emailPersonalLinks() { sendLinks_(false); }

// One email to everyone marked "not needed" (Base access override) who has not submitted: their ID is covered,
// but they should still use their link if they are bringing graduation guests.
function emailOnFileAboutGuests() {
  var ui = SpreadsheetApp.getUi();
  rebuild();
  writePersonalLinks();
  var r = rosterRows_(), h = r.header;
  var cOv = h.indexOf('Base access override'), cSub = h.indexOf('Form submitted'), cLink = h.indexOf('Personal form link');
  var cSent = ensureCol_(r.sheet, h, 'On-file email sent');
  var todo = r.rows.map(function (x, i) { return { x: x, i: i }; }).filter(function (t) {
    return String(t.x[cOv] || '').trim() && t.x[cSub] === 'Not needed' && !t.x[cSent - 1];
  });
  if (ui.alert('Email "ID on file" people', todo.length + ' people are marked not needed and have not submitted. Email them about guests?',
      ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  var sent = 0;
  todo.forEach(function (t) {
    var o = rowObj_(h, t.x), to = String(o.Email || '').trim(), cc = String(o['Parent email'] || '').trim();
    if (!to && !cc) return;
    var who = rosterName_(o);
    var dodOv = /military|DoD|CAC/i.test(String(o['Base access override']));
    var body = 'Hello,\n\nGood news: we do NOT need ' + who + '\'s own ID information for ' + INST + ' (' +
      String(o['Base access override']).replace(/^Not needed:\s*/, '') + ').\n\n' +
      'WHAT YOU NEED TO DO:\n' +
      '- If NOBODY is coming to graduation (' + CONFIG.GRADUATION + ') for you: nothing. You are done.\n' +
      '- If someone IS coming to graduation for you: list them on the base access form, using this link, no later than ' +
      (CONFIG.GUEST_DEADLINE || CONFIG.DUE_DATE) + ':\n' + o['Personal form link'] + '\n' +
      (dodOv ? '  (The link skips your own information and goes straight to the guest questions.)\n'
             : '  (The form will also ask for your own ID. Please fill it in: it only takes a minute.)\n') +
      '\nQuestions? Reply to this email.\n\n' + SIGNATURE;
    var opts = { replyTo: CONFIG.CONTACT_EMAIL, name: CONFIG.CONTACT_NAME };
    if (to && cc && cc.toLowerCase() !== to.toLowerCase()) opts.cc = cc;
    if (safeSend_(to || cc, 'Base access: your ID is on file (guests only if needed)', body, opts)) sent++;
    r.sheet.getRange(t.i + 2, cSent).setValue(new Date());
  });
  ui.alert('Sent ' + sent + ' emails' + (sent < todo.length ? ' (the rest are waiting on the Email queue tab)' : '') + '.');
}      // first email: only people never emailed and not submitted
function emailReminders() { sendLinks_(true); }           // reminder: everyone who has not submitted, even if emailed before

function sendLinks_(reminder) {
  var ui = SpreadsheetApp.getUi();
  rebuild();
  var r = rosterRows_(), h = r.header;
  var cLink = h.indexOf('Personal form link'), cSub = h.indexOf('Form submitted'), cRZ = h.indexOf('Reg Zone status');
  var cSent = ensureCol_(r.sheet, h, 'Link emailed');
  var todo = r.rows.map(function (x, i) { return { x: x, i: i }; })
    .filter(function (t) { return t.x[cSub] !== 'Yes' && t.x[cSub] !== 'Not needed' && t.x[cLink] && (reminder || !t.x[cSent - 1]); })
    .filter(function (t) { return cRZ < 0 || t.x[cRZ] !== 'Not in latest Reg Zone pull'; });   // withdrawn: do not email
  var cSlot = h.indexOf('Slot status');
  if (reminder) {
    // Reminders go only to staff and SLOTTED students. Alternates are welcome to submit but are not chased.
    if (cSlot < 0) { ui.alert('Run "Pull new people from Reg Zone exports" first, so the roster knows who is slotted.'); return; }
    todo = todo.filter(function (t) { var st = String(t.x[cSlot] || ''); return st === 'Staff' || st === 'Slotted'; });
  }
  // Each email may go to 2 recipients. Keep RESERVE recipients free for "Received" receipts and help emails.
  var quota = Math.max(0, Math.floor((MailApp.getRemainingDailyQuota() - CONFIG.EMAIL_RESERVE) / 2));
  if (!quota) { ui.alert('Not enough email allowance left today (keeping ' + CONFIG.EMAIL_RESERVE + ' for receipts). Try again tomorrow.'); return; }
  if (ui.alert(reminder ? 'Send reminders' : 'Email personal links', todo.length + (reminder ? ' people have not submitted yet.' :
      ' people have not been emailed a link yet.') + (reminder ? ' (staff and slotted students only; alternates are skipped)' : '') +
      ' You can send ' + quota +
      ' emails now (attendee + parent each), keeping ' + CONFIG.EMAIL_RESERVE + ' for receipts.\nSend now?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  var sent = 0;
  todo.slice(0, quota).forEach(function (t) {
    var o = rowObj_(h, t.x);
    var to = String(o.Email || '').trim(), cc = String(o['Parent email'] || '').trim();
    if (!to && !cc) return;
    var who = rosterName_(o);
    var body = linkEmailBody_(o);
    var opts = { replyTo: CONFIG.CONTACT_EMAIL, name: CONFIG.CONTACT_NAME };
    if (to && cc && cc.toLowerCase() !== to.toLowerCase()) opts.cc = cc;
    safeSend_(to || cc, (reminder ? 'REMINDER - ' : '') + 'ACTION NEEDED by ' + CONFIG.DUE_DATE +
      ': base access form for ' + who, body, opts);
    r.sheet.getRange(t.i + 2, cSent).setValue(new Date());
    sent++;
  });
  ui.alert('Sent ' + sent + ' emails.' + (todo.length > quota ? ' ' + (todo.length - quota) + ' left; run again tomorrow.' : ''));
}

// ---------------------------------------------------------------- submit trigger
function onFormSubmitHandler(e) {
  if (!isAutomationOwner_()) return;           // a leftover trigger under another account: do nothing (prevents double emails)
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try { rebuild(); } finally { lock.releaseLock(); }
  if (!e || !e.response) return;
  var p = parseResponse_(e.response, map_());
  try {
    if (p.help.problem) handleHelp_(p);
    else sendReceipt_(p);
  } catch (err) { console.log('email failed: ' + err); }
}

function parseResponse_(resp, map) {
  var p = { id: resp.getId(), ts: resp.getTimestamp(), email: resp.getRespondentEmail(), edit: resp.getEditResponseUrl(),
            start: {}, help: {}, main: {}, att: {}, gcount: {}, guests: {}, final: {} };
  resp.getItemResponses().forEach(function (ir) {
    var m = map[ir.getItem().getId()];
    if (!m) return;
    var val = ir.getResponse();
    if (Array.isArray(val)) val = val.join('; ');
    if (m.s === 'guest') (p.guests[m.g] = p.guests[m.g] || {})[m.f] = val;
    else p[m.s][m.f] = val;
  });
  p.capid = String(p.main.capid || '').trim();
  [p.att].concat(Object.keys(p.guests).map(function (k) { return p.guests[k]; })).forEach(function (q) {
    if (q.drives === DRV_DOD) q.idType = DOD;
    if (q.drives === DRV_PIV) q.idType = PIV;
  });
  if (p.main.covered === COVER.DOD && !p.att.idType) p.att.idType = DOD;
  if (p.main.covered === COVER.PIV && !p.att.idType) p.att.idType = PIV;
  return p;
}

// Plain-language list of problems in ONE response, for the receipt email. Empty list = looks complete.
function receiptIssues_(p, ro) {
  var END = new Date(CONFIG.EVENT_END_ISO + 'T00:00:00'), out = [];
  var ovr = ro ? String(ro['Base access override'] || '').trim() : '';
  var ageAns = String(p.main.age || '');
  // Plain, specific fix text: say what they entered, what is wrong, and exactly what to change.
  function nice(who, flags, x) {
    x = x || {};
    var iss = String(x.idIssuer || '').trim(), type = String(x.idType || ''), issuerDone = false;
    flags.forEach(function (f) {
      var m;
      if (f === 'no SSN (optional)') return;
      if (f === 'SSN REQUIRED but missing') out.push(who + ': the ID type picked is a license or ID WITHOUT a star ("Federal Limits Apply"). ' +
        'With that kind of ID, a Social Security number is REQUIRED. FIX: enter the Social Security number. (If the license actually HAS ' +
        'a star, change the ID type to "' + DL_REAL + '" instead.)');
      else if (f === 'SSN looks invalid') out.push(who + ': the Social Security number entered is not a valid 9-digit number. FIX: re-enter it ' +
        'exactly as it appears on the Social Security card.');
      else if (f === 'non-REAL ID without a second ID') out.push(who + ': the ID type picked is a license or ID WITHOUT a star ("Federal Limits ' +
        'Apply"). With that kind of ID, a SECOND ID is also required. FIX: fill in the second ID questions. (If the license actually HAS a ' +
        'star, change the ID type to "' + DL_REAL + '" instead.)');
      else if (f.indexOf('second ID not usable') === 0) out.push(who + ': the second ID entered ("' + String(x.doc2Type || '') + '") ' +
        (/\d/.test(String(x.doc2Number || '')) ? 'is not an ID this person has right now.' : 'has no ID number.') +
        ' The second ID must be an ID this person ALREADY HAS, for example a U.S. passport or passport card, a green card, or a birth ' +
        'certificate, with the number printed on it. FIX: enter that second ID and its number. (If this person gets a license WITH a star ' +
        'before the form is due, you can instead change the ID type to "' + DL_REAL + '" and enter that license.)');
      else if (f.indexOf('driving but no driver') === 0) out.push(who + ': marked as the DRIVER (behind the wheel), but the ID entered ("' +
        type + '") is not a driver\'s license. FIX: if this person WILL be behind the wheel at the ' + INST + ' gate, change the ID to ' +
        'their driver\'s license. If someone else is driving, change the driver question to NO.');
      else if (f.indexOf('too young for a license') >= 0) out.push(who + ': marked as the DRIVER (behind the wheel), but is under 16. ' +
        'FIX: change the driver question to NO.');
      else if (f.indexOf('names a kind of ID') >= 0 || f.indexOf('should be the STATE') >= 0 || f.indexOf('should be the United States') >= 0 ||
               f.indexOf('should be the foreign COUNTRY') >= 0) { if (!issuerDone) out.push(issuerFix_(who, type, iss)); issuerDone = true; }
      else if (f.indexOf('placeholder') >= 0) out.push(who + ': the ID number looks incomplete. FIX: enter the ID number exactly as printed on the ID.');
      else if (f.indexOf('looks like the CAPID') >= 0) out.push(who + ': the ID number entered looks like a CAP ID, not the number on the ID. FIX: enter the ' +
        'number printed on the ID.');
      else if (f.indexOf('issuing state or country missing') >= 0) {
        if (!issuerDone) out.push(who + ': ' + (iss ? '"ID issued by" says "' + iss + '", which does not say WHICH state or country.'
          : '"ID issued by" is blank.') + ' FIX: ' + issuerWhat_(type) + '.');
        issuerDone = true;
      }
      else if (f === 'no ID expiration date') out.push(who + ': the ID expiration date is blank. FIX: enter the expiration date printed on the ' +
        'ID. If the ID has NO expiration date at all, enter 12/31/2099.');
      else if ((m = /^(DoD )?ID expires (\d{4}-\d{2}-\d{2})/.exec(f))) out.push(who + ': the ID expires ' + niceDate_(m[2]) + ', which is before ' +
        CONFIG.GRADUATION + '. The ID must still be valid then. FIX: use a different ID that is still valid on ' + CONFIG.GRADUATION +
        ', and change the ID type, number, "ID issued by" and expiration date to match that ID.');
      else if (f.indexOf('"Other" ID type') === 0) out.push(who + ': the ID type is "Other", but the notes box does not say what the ID is. ' +
        'FIX: in the notes box, type what the ID is (for example "Tribal ID").');
      else if (f === 'home address missing') out.push(who + ': the home address is blank. FIX: pick "Same as the household address" or type the address.');
      else if (f === 'home/cell phone missing') out.push(who + ': the phone number is blank. FIX: pick "Same as YOUR phone number in Step 1" or type the number.');
      else if (f.indexOf('immigration document') >= 0) out.push(who + ': the country of citizenship is "' + String(x.citizen || '') + '", not ' +
        'the United States, so the immigration document is required, but that box is blank. FIX: in the box "NOT a U.S. citizen only", ' +
        'type the document type (for example green card, U.S. visa, ESTA approval, or work permit), its number, its expiration date, and ' +
        'the passport number. (If this person IS a U.S. citizen, change the country of citizenship to United States instead.)');
      else out.push(who + ': ' + f + '.');
    });
  }
  var adult = ro && isAdult_(ro);
  if (ro && adult && ageAns === AGE_ATT.MINOR) out.push('The form says the attendee is UNDER 18, but CAP records show the attendee is 18 or older on the first day of encampment. FIX: in Step 1, pick "18 or OLDER" and fill in the attendee\'s own ID.');
  if (p.att.idType) {
    var a = resolve_(p.att, p.main), af = personFlags_(a, END);
    if (ro && a.dob && ro['Birth month/year (Reg Zone)']) {
      var rz = regZoneMMYY_(ro['Birth month/year (Reg Zone)']);
      if (rz && a.dob.slice(5, 7) + '/' + a.dob.slice(2, 4) !== rz) out.push('Attendee: the date of birth entered does not match CAP records (CAP shows month/year ' + rz + '). FIX: check the date of birth and correct it if it is wrong.');
    }
    nice('Attendee', af, a);
  }
  else if (((adult && !ovr) || drivesOf_(p.att))) out.push('The attendee\'s own ID is required (the attendee is 18 or older, or is driving), but it was not filled in. FIX: in Step 1, pick NO, then fill in the attendee\'s ID.');
  Object.keys(p.guests).map(Number).sort(function (a, b) { return b - a; }).forEach(function (k) {
    var q = resolve_(p.guests[k], p.main);
    if (!q.first && !q.last) return;
    var who = 'Guest ' + [q.first, q.last].join(' ').trim(), needs = q.driving || q.age === AGE_GUEST.ADULT;
    if (!needs) return;
    if (!q.idType) out.push(who + ': ID is required (18 or older, or driving), but it was not filled in. FIX: fill in this guest\'s ID.');
    else nice(who, personFlags_(q, END), q);
  });
  if (!p.final.signature && needsSignature_(p)) out.push('The last page was not signed. FIX: on the last page, check every box, type your full name in the box at the bottom, and click Submit.');
  if (!p.main.covered && !p.main.age) out.push('Step 1 is not finished: the question about the attendee\'s ID and age was not answered. FIX: answer that question in Step 1.');
  if (!p.gcount.count) out.push('The graduation guest question ("How many guests are coming?") was not answered. FIX: answer it. If nobody is coming, pick 0. If you do not know yet, pick "Not sure yet".');
  return out;
}

// "2026-06-30" -> "30 Jun 2026"
function niceDate_(iso) {
  var m = /^(\d{4})-(\d{2})-(\d{2})/.exec(String(iso || ''));
  return m ? Number(m[3]) + ' ' + 'Jan Feb Mar Apr May Jun Jul Aug Sep Oct Nov Dec'.split(' ')[Number(m[2]) - 1] + ' ' + m[1] : String(iso || '');
}
// What "ID issued by" should say for this ID type, in plain words.
function issuerWhat_(type) {
  if (STATE_ID_TYPES.indexOf(type) >= 0) return 'enter the STATE that issued the license or state ID, for example CA';
  if (US_FED_TYPES.indexOf(type) >= 0) return 'enter United States';
  if (type === 'Foreign passport') return 'enter the COUNTRY that issued the passport, for example Canada';
  return 'enter the state, country or federal agency that issued the ID';
}
function issuerFix_(who, type, iss) {
  var said = '"ID issued by" says "' + iss + '"';
  if (/^(state|country|n\/?a|none|unknown)$/i.test(iss)) return who + ': ' + said + ', which does not say WHICH state or country. FIX: ' +
    issuerWhat_(type) + '.';
  if (STATE_ID_TYPES.indexOf(type) >= 0) return who + ': ' + said + ', but a driver\'s license or state ID is issued by a STATE. ' +
    'FIX: change "ID issued by" to the state, for example CA. (If the ID number you entered is NOT from a driver\'s license or state ' +
    'ID, for example it is from a U.S. passport, change the ID type to match that ID instead.)';
  if (US_FED_TYPES.indexOf(type) >= 0) return who + ': ' + said + ', but a "' + type + '" is issued by the United States, not a state. ' +
    'FIX: change "ID issued by" to United States. (If the ID number you entered is actually from a driver\'s license, change the ID ' +
    'type to the driver\'s license instead.)';
  if (type === 'Foreign passport') return who + ': ' + said + ', but a foreign passport is issued by a foreign COUNTRY. FIX: change ' +
    '"ID issued by" to the country on the passport, for example Canada. (If it is a U.S. passport, change the ID type to "U.S. passport" ' +
    'and "ID issued by" to United States.)';
  return who + ': ' + said + '. That is a kind of ID or an office, not who issued the ID. FIX: ' + issuerWhat_(type) + '.';
}

function sendReceipt_(p, reminder) {
  var ro = rosterObj_(rosterRows_())[p.capid];
  var lines = ['Attendee: ' + (ro ? rosterName_(ro) + ' (CAPID ' + p.capid + ', ' + ro.Event + ')' : 'CAPID ' + p.capid)];
  var exempt = isExemptId_(p.att.idType);
  lines.push('Attendee ID: ' + (exempt ? 'not needed (' + p.att.idType + ')' : p.att.idType ? 'given (' + p.att.idType + ')' :
    'not needed (under 18 and not driving)'));
  if (p.att.drives) lines.push('Attendee drives THEMSELVES onto ' + INST + ': ' + (drivesOf_(p.att) ? 'YES' : 'no'));
  var gs = Object.keys(p.guests).map(Number).sort(function (a, b) { return b - a; })
    .map(function (k) { return p.guests[k]; }).filter(function (q) { return q.first || q.last; })
    .map(function (q) { return nameCase_(q.first) + ' ' + nameCase_(q.last) + (drivesOf_(q) ? ' (DRIVER)' : isExemptId_(q.idType) ? ' (military/DoD or PIV card)' : ''); });
  lines.push('Graduation guests: ' + (gs.length ? gs.length + ' (' + gs.join(', ') + ')' : (p.gcount.count || '0')));
  // A second form for the same attendee: its guests are ADDED to the earlier ones. Show the full combined list.
  var all = GUESTS_BY_CAP_[p.capid] || [];
  if (gs.length && all.length > gs.length) {
    lines.push('');
    lines.push('We already had an earlier form for this attendee, so these guests were ADDED to the ones listed before. ' +
      'Full guest list now (' + all.length + '): ' + all.map(function (q) { return q.first + ' ' + q.last; }).join(', ') + '.');
    lines.push('Nobody was removed. To remove a guest, reply to this email.');
  } else if (!gs.length && all.length) {
    lines.push('');
    lines.push('An earlier form for this attendee already lists ' + all.length + ' graduation guest' + (all.length > 1 ? 's' : '') + ': ' +
      all.map(function (q) { return q.first + ' ' + q.last; }).join(', ') + '. They are still on the list. To remove a guest, reply to this email.');
  }
  var issues = receiptIssues_(p, ro);
  var notSure = String(p.gcount.count || '') === GUEST_LATER && !all.length;
  var fix = issues.length
    ? 'STATUS: NOT COMPLETE. ' + (issues.length === 1 ? '1 item needs' : issues.length + ' items need') + ' fixing.\n\n' +
      'WHAT YOU NEED TO DO: click the link at the bottom of this email, change the answer' + (issues.length === 1 ? '' : 's') +
      ' below, and click Submit again. Do NOT fill out a new form.\n\n' +
      issues.map(function (x, i) { return (i + 1) + '. ' + x; }).join('\n\n') + '\n\n'
    : notSure
      ? 'STATUS: COMPLETE for now. You picked "Not sure yet" for graduation guests. When you know who is coming, click the link at the ' +
        'bottom of this email and add them' + guestBy_() + '. Do NOT fill out a new form.\n\n'
      : 'STATUS: COMPLETE. Nothing else is needed. Thank you!\n\n';
  if (issues.length) {
    props_().setProperty('FIX_' + p.id, String(p.ts.getTime()));   // so "Email fix requests" won't repeat it
    props_().setProperty('FIXAT_' + p.id, String(Date.now()));     // when they were last told (automatic follow-ups)
  }
  // Send to whoever filled out the form, CC the attendee (CAP email on the roster) and, for attendees under 18, the parent.
  // Each address only once.
  var seen = {}, list = [];
  [p.email, ro && ro.Email, ro && !isAdult_(ro) && ro['Parent email']].forEach(function (a) {
    a = String(a || '').trim();
    if (a && a.indexOf('@') > 0 && !seen[a.toLowerCase()]) { seen[a.toLowerCase()] = 1; list.push(a); }
  });
  if (!list.length) return;
  var opts = { replyTo: CONFIG.CONTACT_EMAIL, name: CONFIG.CONTACT_NAME };
  if (list.length > 1) opts.cc = list.slice(1).join(',');
  safeSend_(list[0], (reminder && issues.length ? 'REMINDER - ' : '') +
      (issues.length ? 'ACTION NEEDED: fix your base access form for ' : 'Received: base access form for ') +
      (ro ? rosterName_(ro) : 'CAPID ' + p.capid),
    'Hello,\n\n' + (reminder && issues.length ? 'REMINDER: the ' + INST + ' base access form for ' + (ro ? rosterName_(ro) : 'CAPID ' + p.capid) +
      ' still has a problem that must be fixed before ' + INST + ' security can process it.\n\n'
      : 'We received the ' + INST + ' base access form for ' + (ro ? rosterName_(ro) : 'CAPID ' + p.capid) + '.\n\n') +
    fix +
    'WHAT WE RECEIVED:\n' + lines.join('\n') + '\n\n' +
    'Wrong attendee name above? Then the CAPID was typed wrong. Use the link below to fix it.\n\n' +
    'DRIVERS (anyone behind the wheel at the ' + INST + ' gate): bring a driver\'s license, current vehicle registration, and proof of ' +
    'insurance.\n\n' +
    'This form does NOT mean the attendee has been slotted or accepted. ' + CONFIG.SLOTTING_NOTE + '\n\n' +
    'KEEP THIS EMAIL. To change your answers or add graduation guests later, use this link (do NOT fill out a new form):\n' + p.edit +
    '\n\nQuestions? Reply to this email.\n\n' + SIGNATURE, opts);
}

// ---------------------------------------------------------------- automatic follow-ups
// Staff and SLOTTED students are followed up until they are done. Alternates are not chased (they still get the automatic
// confirmation / fix email whenever they submit or edit). Twice a week (Mon + Thu, about 9 AM) when turned on:
//   1. NOT SUBMITTED yet: the personal-link reminder, if not emailed in the last CHASE.EVERY_DAYS days.
//   2. SUBMITTED WITH PROBLEMS: the "REMINDER - ACTION NEEDED" fix email, if not told in the last CHASE.EVERY_DAYS days.
// People with 2+ forms are skipped (check by hand: one edit link may be the wrong one). You get a summary email each run.
var CHASE = { EVERY_DAYS: 2, HOUR: 9 };   // at least 2 days since they were last emailed; the Mon + Thu schedule makes it twice a week
function chasePlan_() {
  var r = rosterRows_(), h = r.header, now = Date.now(), gap = CHASE.EVERY_DAYS * 86400000 - 3 * 3600000;
  var col = function (n) { return h.indexOf(n); };
  var cSub = col('Form submitted'), cSlot = col('Slot status'), cRZ = col('Reg Zone status'), cLink = col('Personal form link');
  var cSent = col('Link emailed');
  var chased = function (x) { var st = String(x[cSlot] || ''); return st === 'Staff' || st === 'Slotted'; };
  var plan = { links: [], fixes: [], multi: [], noSlot: cSlot < 0, waitLinks: [], waitFixes: [] };
  if (cSlot < 0) return plan;
  r.rows.forEach(function (x, i) {
    if (!chased(x) || x[cSub] === 'Yes' || x[cSub] === 'Not needed' || !x[cLink]) return;
    if (cRZ >= 0 && x[cRZ] === 'Not in latest Reg Zone pull') return;
    var last = cSent >= 0 && x[cSent] ? new Date(x[cSent]).getTime() : 0;
    if (now - last >= gap) plan.links.push({ i: i, o: rowObj_(h, x) });
    else plan.waitLinks.push(rowObj_(h, x));
  });
  var map = map_(), roster = rosterObj_(r), byCap = {};
  form_().getResponses().forEach(function (resp) {
    var p = parseResponse_(resp, map);
    if (!p.capid || p.help.problem) return;
    (byCap[p.capid] = byCap[p.capid] || []).push(p);
  });
  Object.keys(byCap).forEach(function (cap) {
    var ro = roster[cap];
    if (!ro || !(ro['Slot status'] === 'Staff' || ro['Slot status'] === 'Slotted')) return;
    var list = byCap[cap].sort(function (a, b) { return a.ts - b.ts; }), p = list[list.length - 1];
    if (!receiptIssues_(p, ro).length) return;
    if (list.length > 1) { plan.multi.push(rosterName_(ro)); return; }
    var told = Math.max(Number(props_().getProperty('FIXAT_' + p.id) || 0), p.ts.getTime());
    if (now - told >= gap) plan.fixes.push({ p: p, ro: ro });
    else plan.waitFixes.push(ro);
  });
  return plan;
}
function chaseSend_(plan) {
  var r = rosterRows_(), h = r.header, cSent = ensureCol_(r.sheet, h, 'Link emailed'), links = 0, fixes = 0, waiting = 0;
  plan.links.forEach(function (t) {
    var o = t.o, to = String(o.Email || '').trim(), cc = String(o['Parent email'] || '').trim();
    if (!to && !cc) return;
    if (MailApp.getRemainingDailyQuota() <= CONFIG.EMAIL_RESERVE) { waiting++; return; }   // keep room for receipts; next run
    var opts = { replyTo: CONFIG.CONTACT_EMAIL, name: CONFIG.CONTACT_NAME };
    if (to && cc && cc.toLowerCase() !== to.toLowerCase()) opts.cc = cc;
    safeSend_(to || cc, 'REMINDER - ACTION NEEDED by ' + CONFIG.DUE_DATE + ': base access form for ' + rosterName_(o),
      linkEmailBody_(o), opts);
    r.sheet.getRange(t.i + 2, cSent).setValue(new Date(Date.now()));
    links++;
  });
  plan.fixes.forEach(function (t) {
    if (MailApp.getRemainingDailyQuota() <= CONFIG.EMAIL_RESERVE) { waiting++; return; }
    sendReceipt_(t.p, true); fixes++;
  });
  return { links: links, fixes: fixes, waiting: waiting };
}
function chaseSummary_(plan, res) {
  return 'Base access follow-ups (staff and slotted students only):\n' +
    '- Not submitted yet, reminded: ' + res.links + '\n' +
    '- Submitted with problems, reminded: ' + res.fixes + '\n' +
    (res.waiting ? '- Not sent (daily email limit), will go next run: ' + res.waiting + '\n' : '') +
    (plan.multi.length ? '- NOT emailed, 2+ forms with problems (check by hand): ' + plan.multi.join(', ') + '\n' : '') +
    '\nDetails: Submissions tab and Email log tab of the base access sheet.';
}
function previewFollowUps() {
  var ui = SpreadsheetApp.getUi();
  rebuild();
  var plan = chasePlan_();
  if (plan.noSlot) { ui.alert('Run "Pull new people from Reg Zone exports" first, so the roster knows who is slotted.'); return; }
  var nm = function (x) { return x.slice(0, 40).join(', ') + (x.length > 40 ? ' ...' : ''); };
  ui.alert('Follow-ups due now (staff and slotted students only; nothing is sent):\n\n' +
    'NOT SUBMITTED (' + plan.links.length + '): ' + nm(plan.links.map(function (t) { return rosterName_(t.o); })) + '\n\n' +
    'SUBMITTED WITH PROBLEMS (' + plan.fixes.length + '): ' + nm(plan.fixes.map(function (t) { return rosterName_(t.ro); })) + '\n\n' +
    (plan.multi.length ? '2+ FORMS, check by hand (' + plan.multi.length + '): ' + nm(plan.multi) + '\n\n' : '') +
    'NOT DUE YET (emailed in the last ' + CHASE.EVERY_DAYS + ' days; they will be included in a later run): ' +
    plan.waitLinks.length + ' not submitted, ' + plan.waitFixes.length + ' with problems' +
    (plan.waitFixes.length ? ' (' + nm(plan.waitFixes.map(function (o) { return rosterName_(o); })) + ')' : '') + '.\n\n' +
    'Automatic runs: Mondays and Thursdays around ' + CHASE.HOUR + ' AM. Nobody is emailed twice within ' + CHASE.EVERY_DAYS + ' days.');
}
function sendFollowUpsNow() {
  var ui = SpreadsheetApp.getUi();
  rebuild();
  var plan = chasePlan_();
  if (plan.noSlot) { ui.alert('Run "Pull new people from Reg Zone exports" first, so the roster knows who is slotted.'); return; }
  if (ui.alert('Send follow-ups now', plan.links.length + ' not submitted + ' + plan.fixes.length + ' with problems (staff and slotted ' +
      'students only). Send now?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  var res = chaseSend_(plan);
  ui.alert(chaseSummary_(plan, res));
}
function autoFollowUps() {
  if (!isAutomationOwner_()) return;
  rebuild();
  var plan = chasePlan_();
  if (plan.noSlot) { sendMail_(CONFIG.CONTACT_EMAIL, 'Base access follow-ups NOT sent', 'The Roster tab has no "Slot status" column. Run "Pull new people from Reg Zone exports".'); return; }
  var res = chaseSend_(plan);
  sendMail_(CONFIG.CONTACT_EMAIL, 'Base access follow-ups: ' + (res.links + res.fixes) + ' sent', chaseSummary_(plan, res));
}
function turnOnAutoFollowUps() {
  turnOffAutoFollowUps(true);
  [ScriptApp.WeekDay.MONDAY, ScriptApp.WeekDay.THURSDAY].forEach(function (d) {
    ScriptApp.newTrigger('autoFollowUps').timeBased().onWeekDay(d).atHour(CHASE.HOUR).create();
  });
  props_().setProperty('AUTO_FOLLOWUPS', 'on');
  SpreadsheetApp.getUi().alert('Automatic follow-ups are ON: every Monday and Thursday around ' + CHASE.HOUR + ' AM, staff and slotted ' +
    'students who have not submitted, or whose form still has problems, get a reminder (never twice within ' + CHASE.EVERY_DAYS +
    ' days). Alternates are not chased. You get a summary email each time.');
}
function turnOffAutoFollowUps(quiet) {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'autoFollowUps') ScriptApp.deleteTrigger(t); });
  if (quiet !== true) { props_().deleteProperty('AUTO_FOLLOWUPS'); SpreadsheetApp.getUi().alert('Automatic follow-ups are off.'); }
}

// Menu: email the same "ACTION NEEDED" receipt to everyone whose CURRENT answers still have problems.
// Skips anyone already sent a fix request for this exact version of their answers (it re-sends after they edit, if problems remain).
function emailFixRequests() {
  var ui = SpreadsheetApp.getUi();
  rebuild();
  var map = map_(), roster = rosterObj_(rosterRows_()), byCap = {};
  form_().getResponses().forEach(function (resp) {
    var p = parseResponse_(resp, map);
    if (!p.capid || p.help.problem) return;
    (byCap[p.capid] = byCap[p.capid] || []).push(p);
  });
  var todo = [], multi = [], noEmail = [];
  Object.keys(byCap).forEach(function (cap) {
    var list = byCap[cap].sort(function (a, b) { return a.ts - b.ts; }), p = list[list.length - 1], ro = roster[cap];
    if (!receiptIssues_(p, ro).length) return;
    var who = ro ? rosterName_(ro) : 'CAPID ' + cap;
    if (list.length > 1) { multi.push(who); return; }        // answers split across 2+ forms: check by hand
    if (!p.email && !(ro && String(ro.Email || '').trim())) { noEmail.push(who); return; }
    if (props_().getProperty('FIX_' + p.id) === String(p.ts.getTime())) return;   // already told about this version
    todo.push({ p: p, who: who });
  });
  var extra = (multi.length ? '\n\nNot emailed (more than one form, check by hand): ' + multi.join(', ') : '') +
              (noEmail.length ? '\n\nNot emailed (no email on the response): ' + noEmail.join(', ') : '');
  if (!todo.length) { ui.alert('No new fix requests to send.' + extra); return; }
  if (ui.alert('Email fix requests', todo.length + ' submitted forms still have problems and have not been sent a fix request for their current answers:\n' +
      todo.map(function (t) { return t.who; }).join(', ') + extra + '\n\nSend each one the "ACTION NEEDED" email with their edit link?',
      ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  todo.forEach(function (t) { sendReceipt_(t.p); });
  ui.alert('Sent ' + todo.length + ' fix requests.' + (MailApp.getRemainingDailyQuota() <= 0 ? ' Some may be waiting on the Email queue tab.' : ''));
}

// ---------------------------------------------------------------- email with a waiting line
// Gmail limits how many emails a script can send per day. If the limit is reached, the email waits on the "Email queue"
// tab and an hourly check sends it as soon as the limit frees up. Nothing is lost.
// Plain-text emails can show ragged line breaks in some mail apps (plain text gets re-wrapped). Every email also
// carries an HTML version built from the same text: paragraphs, real line breaks, clickable links, and bold section labels.
// The plain text stays as the fallback for mail apps that do not show HTML.
function textToHtml_(body) {
  var esc = function (t) { return String(t).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;'); };
  var link = function (t) {
    return t.replace(/(https?:\/\/[^\s<]+)/g, function (u) {
      var trail = (u.match(/[.,);]+$/) || [''])[0], url = u.slice(0, u.length - trail.length);
      return '<a href="' + url + '" style="color:#1a56c4;word-break:break-all">' + url + '</a>' + trail;
    });
  };
  var label = function (t) {   // "STATUS: COMPLETE", "WHAT YOU NEED TO DO:", "FIX:" -> bold
    t = t.replace(/^((?:[A-Z0-9][A-Z0-9 ,'()\/&-]{2,}?)(?::|(?= -)))/, '<b>$1</b>');
    if (t.indexOf('<b>') !== 0) t = t.replace(/^([A-Z][A-Z]+(?: [A-Z]+)*\.?)(?=[ .]|$)/, function (m0) { return m0.replace(/[^A-Z]/g, '').length >= 5 && m0 !== 'CAPID' ? '<b>' + m0 + '</b>' : m0; });   // "KEEP THIS EMAIL.", "DRIVERS"
    t = t.replace(/\bFIX:/g, '<b>FIX:</b>');
    t = t.replace(/^<b>STATUS:<\/b> (COMPLETE[^.]*\.)/, '<b>STATUS:</b> <b style="color:#17795a">$1</b>')
         .replace(/^<b>STATUS:<\/b> (NOT COMPLETE\.)/, '<b>STATUS:</b> <b style="color:#b42318">$1</b>');
    return t;
  };
  var blocks = String(body || '').replace(/\r/g, '').split(/\n\s*\n/);
  var html = blocks.map(function (b) {
    var lines = b.split('\n');
    if (lines.every(function (l) { return /^\s*- /.test(l); })) {
      return '<ul style="margin:0 0 14px 0;padding-left:22px">' + lines.map(function (l) {
        return '<li style="margin:0 0 4px 0">' + label(link(esc(l.replace(/^\s*- /, '')))) + '</li>'; }).join('') + '</ul>';
    }
    var m = /^(\d+)\. ([\s\S]*)$/.exec(b);
    if (m) return '<p style="margin:0 0 14px 0;padding-left:26px;text-indent:-26px"><b style="display:inline-block;width:26px;text-indent:0">' +
      m[1] + '.</b>' + lines.map(function (l, i) { return label(link(esc(i ? l : l.replace(/^\d+\. /, '')))); }).join('<br>') + '</p>';
    return '<p style="margin:0 0 14px 0">' + lines.map(function (l) { return label(link(esc(l))); }).join('<br>') + '</p>';
  }).join('');
  return '<div style="font-family:Arial,Helvetica,sans-serif;font-size:15px;line-height:1.5;color:#1b1f27;max-width:640px">' + html + '</div>';
}
function sendMail_(to, subject, body, opts) {
  var o = {}; Object.keys(opts || {}).forEach(function (k) { o[k] = opts[k]; });
  if (!o.htmlBody) o.htmlBody = textToHtml_(body);
  MailApp.sendEmail(to, subject, body, o);
}
function safeSend_(to, subject, body, opts) {
  try {
    if (MailApp.getRemainingDailyQuota() > 0) {
      sendMail_(to, subject, body, opts || {});
      logEmail_(to, opts, subject, 'Sent');
      return true;
    }
  } catch (err) { console.log('send failed, queuing: ' + err); }
  var sh = ss_().getSheetByName('Email queue') || ss_().insertSheet('Email queue');
  if (sh.getLastRow() === 0) sh.appendRow(['Queued', 'To', 'Subject', 'Body', 'Options', 'Sent']);
  sh.appendRow([new Date(), to, subject, body, JSON.stringify(opts || {}), '']);
  ensureRetryTrigger_();
  logEmail_(to, opts, subject, 'Queued (daily limit)');
  return false;
}
// "Email log" tab: one row per email the script sends (time, to, cc, subject, status). No message bodies.
function logEmail_(to, opts, subject, status) {
  try {
    var ss = ss_(), sh = ss.getSheetByName('Email log');
    if (!sh) { sh = ss.insertSheet('Email log'); sh.appendRow(['Time', 'To', 'CC', 'Subject', 'Status']); sh.getRange(1, 1, 1, 5).setFontWeight('bold'); }
    sh.appendRow([Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'), to, (opts && opts.cc) || '', subject, status]);
  } catch (e) { console.log('email log failed: ' + e); }
}
function ensureRetryTrigger_() {
  var has = ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'sendQueuedEmails'; });
  if (!has) ScriptApp.newTrigger('sendQueuedEmails').timeBased().everyHours(1).create();
}
function sendQueuedEmails() {
  if (!isAutomationOwner_()) return;
  var sh = ss_().getSheetByName('Email queue');
  if (!sh || sh.getLastRow() < 2) return;
  var v = sh.getRange(2, 1, sh.getLastRow() - 1, 6).getValues(), sent = 0;
  for (var i = 0; i < v.length; i++) {
    if (v[i][5]) continue;
    if (MailApp.getRemainingDailyQuota() < 1) break;
    try {
      var o = JSON.parse(v[i][4] || '{}');
      sendMail_(v[i][1], v[i][2], v[i][3], o);
      sh.getRange(i + 2, 6).setValue(new Date()); sent++;
      logEmail_(v[i][1], o, v[i][2], 'Sent from queue');
    } catch (err) { console.log('queued send failed: ' + err); break; }
  }
  console.log('Sent ' + sent + ' queued emails.');
}

// ---------------------------------------------------------------- help requests (logged + emailed; no AI)
function handleHelp_(p) {
  var ss = ss_();
  var sh = ss.getSheetByName(CONFIG.HELP_SHEET);
  if (!sh) {
    sh = ss.insertSheet(CONFIG.HELP_SHEET);
    sh.appendRow(['Received', 'Name', 'Email', 'Phone', 'CAPID', 'Problem', 'What our records show', 'Resolved', 'Notes']);
    sh.getRange(1, 1, 1, 9).setFontWeight('bold');
  }
  var h = p.help, cap = String(h.capid || '').replace(/\D/g, '');
  var context = helpContext_(cap);
  sh.appendRow([p.ts, h.name || '', h.email || '', h.phone || '', cap, h.problem || '', context, '', '']);
  safeSend_(CONFIG.CONTACT_EMAIL, 'Base access form: help request from ' + (h.name || h.email),
    'From: ' + h.name + ' <' + h.email + '> ' + (h.phone || '') + '\n' + context + '\n\nProblem:\n' + h.problem +
    '\n\nReply to this email to answer them. All requests: ' + ss.getUrl(), { replyTo: h.email || CONFIG.CONTACT_EMAIL });
  if (h.email) safeSend_(h.email, 'We got your message about the base access form',
    'Hello ' + (h.name || '') + ',\n\nWe got your message and ' + CONFIG.CONTACT_NAME + ' will email you back.\n\n' +
    'In the meantime, these answers fix most problems:\n\n' +
    '1. "CAPID not accepted": enter the CAPID of the person GOING to encampment. Parents: your CADET\'s CAPID, not yours. Check every digit.\n\n' +
    '2. Parents coming to graduation are GUESTS. Enter your own information on the guest pages at the end, not on the attendee pages.\n\n' +
    '3. An attendee UNDER 18 who is NOT driving does not need an ID. Just answer the guest question.\n\n' +
    '4. Military / DoD ID card or federal PIV card? Pick that answer in Step 1. The form then skips the ID questions.\n\n' +
    '5. No star on the license ("Federal Limits Apply")? It still works, but you must also give a Social Security number AND a second ID, and bring the second ID to the gate.\n\n' +
    '6. To change answers you already sent, use the link in the email we sent you after you submitted. Do NOT fill out a new form.\n\n' +
    '7. "Driver" means the person BEHIND THE WHEEL. Passengers, people being dropped off, and shuttle riders are NOT drivers.\n\n' +
    '8. Filling out this form does NOT mean the attendee has been slotted. ' + CONFIG.SLOTTING_NOTE + '\n\n' +
    SIGNATURE, { replyTo: CONFIG.CONTACT_EMAIL, name: CONFIG.CONTACT_NAME });
}

function helpContext_(cap) {
  if (!cap) return 'CAPID: not given.';
  var ro = rosterObj_(rosterRows_())[cap];
  if (!ro) return 'CAPID ' + cap + ' is NOT on the encampment list (typo, or registered after the list was made).';
  return 'CAPID ' + cap + ' = ' + rosterName_(ro) + ', ' + ro.Event + ', ' + (isAdult_(ro) ? '18 or older' : 'under 18') +
    '. Form submitted: ' + (ro['Form submitted'] || 'No') + '. Own ID on file: ' + (ro['Own ID on file'] || '-') +
    '. Problems on their form: ' + (ro['Form flags'] || 'none') + '.';
}

// ---------------------------------------------------------------- collect people from all responses
function collect_() {
  var form = form_(), map = map_();
  var r = rosterRows_(), roster = rosterObj_(r);
  var END = new Date(CONFIG.EVENT_END_ISO + 'T00:00:00');
  var parsed = form.getResponses().map(function (resp) { return parseResponse_(resp, map); })
    .filter(function (p) { return p.capid; })
    .sort(function (a, b) { return a.ts - b.ts; });
  var byCap = {};
  parsed.forEach(function (p) {
    var b = byCap[p.capid] = byCap[p.capid] || { all: [], att: null, guests: null, gResps: [] };
    b.all.push(p);
    if (p.att.idType) b.att = p;
    var c = p.gcount.count || '';
    var named = Object.keys(p.guests).some(function (k) { return p.guests[k].first || p.guests[k].last; });
    if ((c && c.indexOf('already sent') < 0) || named) b.guests = p;    // guests typed in still count even if the count answer was lost
    if (named) b.gResps.push(p);    // every form that names guests: a second form ADDS guests, it never wipes the earlier ones
  });
  var people = [], subs = [], status = {};
  Object.keys(byCap).forEach(function (cap) {
    var b = byCap[cap], last = b.all[b.all.length - 1], ro = roster[cap];
    var sponsor = ro ? rosterName_(ro) : '(not on roster)', event = ro ? ro.Event : '';
    var ageAns = String(last.main.age || ''), attDrives = drivesOf_(last.att);
    var kind = (last.main.covered && last.main.covered !== COVER.NONE ? last.main.covered : ageAns) + (attDrives ? '; DRIVER' : '');
    var flags = [];
    if (!ro) flags.push('CAPID not on roster');
    var typed = String(last.main.rosterName || '').toLowerCase();
    // Flag only if they typed an actual name that is not this person's (typing just a rank, e.g. "Cadet Airman", is harmless)
    var nameWords = typed.replace(/[^a-z\s]/g, ' ').split(/\s+/).filter(function (w) { return w.length > 1 && !RANK_WORDS.test(w); });
    if (ro && nameWords.length && typed.indexOf(String(ro['Last name']).toLowerCase()) < 0) flags.push('name on form ("' + last.main.rosterName + '") does not match CAPID');
    if (!last.final.signature && needsSignature_(last)) flags.push('attestation not signed');
    if (!last.main.covered && !last.main.age) flags.push('Step 1 ID/age question not answered');
    if (!b.all.some(function (x) { return x.gcount.count; })) flags.push('guest question not answered');
    if (b.all.length > 1) flags.push(b.all.length + ' responses (using the latest of each part)');
    var ovr = ro ? String(ro['Base access override'] || '').trim() : '';
    var must = (ro && isAdult_(ro) && !ovr) || attDrives;
    if (must && !b.att) flags.push('needs own ID (18+ or driving) but none given');
    if (ro && isAdult_(ro) && ageAns === AGE_ATT.MINOR) flags.push('form says under 18, Reg Zone says 18+');
    if (ro && !isAdult_(ro) && ageAns === AGE_ATT.ADULT) flags.push('form says 18+, Reg Zone says under 18');
    if (b.att) {
      var a = resolve_(b.att.att, b.att.main), pf = personFlags_(a, END);
      if (ro && a.dob && ro['Birth month/year (Reg Zone)']) {
        var mmyy = a.dob.slice(5, 7) + '/' + a.dob.slice(2, 4);
        var rz = regZoneMMYY_(ro['Birth month/year (Reg Zone)']);
        if (rz && mmyy !== rz) pf.push('date of birth month/year differs from Reg Zone (' + rz + ')');
      }
      if (isExemptId_(a.idType) && !a.last) a.last = a.dodName || sponsor;
      people.push({ cat: a.idType === DOD ? 'Attendee (DoD ID holder)' : a.idType === PIV ? 'Attendee (federal PIV holder)' : 'Attendee', event: event, cap: cap, sponsor: sponsor, p: a, rel: '', flags: pf, ts: b.att.ts, signed: signed_(b.att) });
      if (pf.length) flags.push('Attendee: ' + pf.join('; '));
      if (pf.length && b.all.some(function (x) { return x !== b.att && x.att.idType && !personFlags_(resolve_(x.att, x.main), END).length; }))
        flags.push('Attendee: an EARLIER form had this ID with no problems; the later form replaced it: check which is right');
    }
    var stated = last.gcount.count || '', entered = 0;
    if (b.guests) stated = b.guests.gcount.count;
    // Guests from EVERY form for this CAPID, combined (someone may send a second form with just a new
    // guest). The same name on a later form replaces the earlier entry for that person; nobody is dropped.
    var gList = guestsCombined_(b.gResps);
    GUESTS_BY_CAP_[cap] = gList.map(function (g) { return g.q; });
    if (b.gResps.length > 1) flags.push('guests combined from ' + b.gResps.length + ' forms (' + gList.length + ' people): check');
    if (b.gResps.length && b.guests && b.gResps.indexOf(b.guests) < 0 && /^0/.test(String(stated)))
      flags.push('latest form says 0 guests, but an earlier form lists guests (kept): check');
    if (gList.length) {
      gList.forEach(function (g) {
        var gq = g.q;
        entered++;
        var drives = gq.driving, needsId = drives || gq.age === AGE_GUEST.ADULT;
        var gf = needsId ? personFlags_(gq, END) : [];
        if (needsId && !gq.idType) gf.push('needs ID (adult or driver) but none given');
        if (!needsId && ageOn_(gq.dob) !== '' && ageOn_(gq.dob) >= 18) gf.push('marked under 18 but date of birth says ' + ageOn_(gq.dob));
        people.push({ cat: !needsId ? 'Guest (under 18, not driving)' : gq.idType === DOD ? 'Guest (DoD ID holder)' : gq.idType === PIV ? 'Guest (federal PIV holder)' : (drives ? 'Guest (driver)' : 'Guest'),
                      event: event, cap: cap,
                      sponsor: sponsor, p: gq, rel: gq.relationship || '', flags: gf, ts: g.from.ts, signed: signed_(g.from) });
        if (gf.length) flags.push('Guest ' + [gq.first, gq.last].join(' ') + ': ' + gf.join('; '));
        if (gf.length && needsId && g.earlier.some(function (e) { return e.idType && !personFlags_(e, END).length; }))
          flags.push('Guest ' + [gq.first, gq.last].join(' ') + ': an EARLIER form had this guest with no problems; the later form replaced it: check which is right');
      });
    }
    var n = parseInt(stated, 10);
    if (b.gResps.length <= 1 && b.guests && !isNaN(n) && n !== entered) flags.push('guest count says ' + n + ' but ' + entered + ' entered');
    var ownId = !b.att ? (must ? 'MISSING' : 'Not required') : b.att.att.idType === DOD ? 'Not needed (military/DoD ID)' :
      b.att.att.idType === PIV ? 'Not needed (federal PIV)' : 'Yes';
    subs.push([cap, sponsor, event, ro ? ro['Age at event start'] : '', kind, b.all.length, last.ts, last.main.filledBy || '',
      last.email || '', last.main.filledByPhone || '', ownId, stated, entered,
      shortFlags_(flags), last.edit]);
    status[cap] = { att: ownId, guests: entered, flags: shortFlags_(flags) };
  });
  people = mergeSamePerson_(people);
  people.sort(function (x, y) { return (x.cap + x.cat).localeCompare(y.cap + y.cat); });
  return { r: r, people: people, subs: subs, status: status };
}

// Guests named on any of these responses (oldest first). Every guest on a single form is kept, even if two have the same name
// (for example a father and son). Across forms, a guest on a LATER form with the same name (first + last + suffix) as a guest on an
// EARLIER form replaces that earlier entry (it is a correction); everyone else is added. Names in any alphabet are compared as typed
// (only spaces, periods, hyphens and apostrophes are ignored). Returns [{ q: resolved guest, from: response, earlier: [older q] }].
var GUESTS_BY_CAP_ = {};
function guestKey_(gq) {
  var k = [gq.first, gq.last, gq.suffix].map(function (x) { return String(x || '').toLowerCase().replace(/[\s.'\u2019-]+/g, ''); }).join('|');
  return k === '||' ? null : k;
}
function guestsCombined_(resps) {
  var list = [];
  resps.slice().sort(function (a, b) { return a.ts - b.ts; }).forEach(function (r) {
    var before = list.length, used = {};
    Object.keys(r.guests).map(Number).sort(function (x, y) { return y - x; }).forEach(function (sl) {
      var gq = resolve_(r.guests[sl], r.main);
      if (!gq.first && !gq.last) return;
      var key = guestKey_(gq), hit = -1;
      for (var i = 0; i < before; i++) if (!used[i] && key && guestKey_(list[i].q) === key) { hit = i; break; }   // earlier forms only
      if (hit >= 0) { used[hit] = 1; list[hit] = { q: gq, from: r, earlier: list[hit].earlier.concat([list[hit].q]) }; }
      else list.push({ q: gq, from: r, earlier: [] });
    });
  });
  return list;
}

// The Privacy Act page / attestation is only needed when the response gives someone's ID information:
// the attendee's own ID (other than a military/DoD card) or any graduation guest.
function needsSignature_(p) {
  if (p.att && p.att.idType && !isExemptId_(p.att.idType)) return true;
  return Object.keys(p.guests || {}).some(function (k) { var q = p.guests[k]; return q && (q.first || q.last); });
}

function signed_(p) {
  return p.final && p.final.signature ? p.final.signature + ', ' + Utilities.formatDate(p.ts, Session.getScriptTimeZone(), 'yyyy-MM-dd') : 'NOT SIGNED';
}
// Reg Zone gives birth month/year as "MM/YY". Google Sheets sometimes turns values like "08/03" into a date (3 Aug of
// this year), so read it back from the date: month = MM, day of month = YY.
function regZoneMMYY_(v) {
  if (v instanceof Date) return ('0' + (v.getMonth() + 1)).slice(-2) + '/' + ('0' + v.getDate()).slice(-2);
  var m = /^(\d{1,2})\/(\d{2})$/.exec(String(v || '').trim());
  return m ? ('0' + m[1]).slice(-2) + '/' + m[2] : '';
}
function isYes_(v) { return String(v || '').indexOf('YES') === 0; }
function drivesOf_(q) { return isYes_(q.drives); }
// Names typed all lowercase or ALL CAPS become "Title Case". Names already in mixed case are kept exactly
// as typed (e.g. "McGrath", "de la Torre"). Handles hyphens and apostrophes (O'Neil), Mc (McGrath) and suffixes (Jr, III).
var NAME_PARTICLES = /^(van|von|der|den|de|del|della|degli|di|da|das|do|dos|du|la|le|ten|ter|y|e|bin|binti|bint|al|el|ibn)$/;
function nameCase_(v) {
  var t = String(v || '');
  if (!/[A-Za-z\u00C0-\u024F]/.test(t)) return t;
  if (t !== t.toLowerCase() && t !== t.toUpperCase()) {
    // Mixed case ("Suraj krishna", "Yeung salazar"): capitalize each word that starts lowercase, except particles
    // like "van" or "de" ("Quinton Heath van Horn" stays). Other letters are left as typed (McDonald, DeLeon).
    return t.split(/(\s+)/).map(function (w) {
      if (!w || /^\s+$/.test(w)) return w;
      if (NAME_PARTICLES.test(w)) return w;   // "van Horn" as a last name stays as entered
      var c = w.charAt(0);
      return c !== c.toLowerCase() || c === c.toUpperCase() ? w : c.toUpperCase() + w.slice(1);
    }).join('');
  }
  var parts = t.toLowerCase().split(/([\s\-'\u2019.]+)/), wordNo = 0;
  return parts.map(function (w) {
    if (!w || /^[\s\-'\u2019.]+$/.test(w)) return w;
    wordNo++;
    if (/^(ii|iii|iv)$/.test(w) || (wordNo > 1 && /^(v|vi|vii|viii|ix|x)$/.test(w))) return w.toUpperCase();   // II, III, IV, VI...
    if (/^mc.{2,}/.test(w)) return 'Mc' + w.charAt(2).toUpperCase() + w.slice(3);
    return w.charAt(0).toUpperCase() + w.slice(1);
  }).join('');
}
// fill in "Same as Step 1" address/phone and the driving flag; returns a copy
function resolve_(q, main) {
  var o = {}; Object.keys(q).forEach(function (k) { o[k] = q[k]; });
  ['first', 'middle', 'last', 'suffix', 'idIssuer', 'relationship'].forEach(function (k) {   // tidy stray spaces
    if (typeof o[k] === 'string') o[k] = o[k].replace(/\s+/g, ' ').trim();
  });
  ['first', 'middle', 'last', 'suffix'].forEach(function (k) { if (typeof o[k] === 'string') o[k] = nameCase_(o[k]); });
  o.address = String(q.addrSame || '').indexOf('Same') === 0 ? (main.householdAddr || '') : (q.addressOther || '');
  o.phone = String(q.phoneSame || '').indexOf('Same') === 0 ? (main.filledByPhone || '') : (q.phoneOther || '');
  o.driving = drivesOf_(q);
  o.__capid = String(main.capid || '').trim();
  if (!o.driving) { ['plate', 'plateState', 'year', 'make', 'model', 'color', 'owner'].forEach(function (k) { delete o[k]; }); }
  return o;
}

function vehicle_(p) {   // vehicle details (plate, make, model) are not collected; this just marks drivers
  if (p.plate) return [p.plate, p.plateState, [p.year, p.make, p.model].join(' ').trim(), p.color, p.owner ? 'owner: ' + p.owner : ''].filter(Boolean).join(', ');
  return p.driving ? 'Driving' : '';
}

// ---------------------------------------------------------------- rebuild masked tabs
function rebuild() {
  var c = collect_(), ss = ss_();
  var H = ['Category', 'Event', 'Sponsor CAPID', 'Sponsor', 'Last name', 'First name', 'Middle name', 'Suffix',
    'Age on ' + CONFIG.GRADUATION, 'SSN', 'ID type', 'ID number', 'ID issued by', 'ID expiration', 'Second ID (no-star license)',
    'Vehicle', 'Relationship', 'Flags', 'Response submitted'];
  var rows = c.people.map(function (x) {
    var p = x.p;
    return [x.cat, x.event, x.cap, x.sponsor, p.last || '', p.first || '', p.middle || '', p.suffix || '', ageOn_(p.dob),
      p.ssn ? 'on file (' + last4_(p.ssn) + ')' : '', p.idType || '', last4_(p.idNumber), issuerStd_(p.idIssuer, p.idType), expShow_(p.idExpires),
      p.doc2Type ? p.doc2Type + (p.doc2Number ? ' (' + last4_(p.doc2Number) + ')' : '') : '',
      p.driving ? 'driving' : '', x.rel, x.flags.join('; '), x.ts];
  });
  writeTab_(ss, CONFIG.SECFOR_SHEET, H, rows);
  writeTab_(ss, CONFIG.SUBMISSIONS_SHEET, ['CAPID', 'Attendee', 'Event', 'Age at event start', 'Attendee (as answered)', 'Responses',
    'Last response', 'Filled in by', 'Respondent email', 'Respondent phone', 'Own ID on file', 'Guests stated', 'Guests entered',
    'Flags', 'Edit link'], c.subs.sort(function (a, b) { return new Date(b[6]) - new Date(a[6]); }));   // newest first
  var r = c.r, h = r.header;
  var cOv = ensureCol_(r.sheet, h, 'Base access override') - 1;
  var ovVals = r.rows.map(function (x) {
    var cur = String(x[cOv] === undefined ? '' : x[cOv]).trim();
    return [cur || ACCESS_OVERRIDES[String(x[h.indexOf('CAPID')]).trim()] || ''];
  });
  if (ovVals.length) r.sheet.getRange(2, cOv + 1, ovVals.length, 1).setValues(ovVals);
  var cols = ['Form submitted', 'Own ID on file', 'Guests entered', 'Form flags'].map(function (n) { return ensureCol_(r.sheet, h, n); });
  var vals = r.rows.map(function (x, i) {
    var o = rowObj_(h, x), st = c.status[String(o.CAPID).trim()], ov = ovVals[i][0];
    if (!st) return ov ? ['Not needed', ov, '', ''] : ['No', isAdult_(o) ? 'MISSING' : 'Not required', '', ''];
    return ['Yes', (st.att === 'MISSING' && ov) ? ov : st.att, st.guests, st.flags];
  });
  if (vals.length) cols.forEach(function (col, k) {
    r.sheet.getRange(2, col, vals.length, 1).setValues(vals.map(function (v) { return [v[k]]; }));
  });
}

// ---------------------------------------------------------------- full export for installation security
function createSecforExport() {
  var ui = SpreadsheetApp.getUi();
  if (ui.alert('Full SECFOR export', 'This creates a NEW private spreadsheet in your My Drive with full SSNs, dates of birth and ID numbers.\n' +
      'Send it to ' + INST + ' security the way they tell you to, then delete it. Continue?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  var c = collect_();
  var H = ['Category', 'Event', 'Sponsor CAPID', 'Sponsor', 'Last name', 'First name', 'Middle name', 'Suffix', 'Date of birth',
    'Age on ' + CONFIG.GRADUATION, 'SSN', 'ID type', 'ID number', 'ID issued by', 'ID expiration',
    'Second ID', 'Second ID number', 'Home address', 'Home or cell phone', 'Sex', 'Country of citizenship', 'Immigration document (non-citizens)',
    'Vehicle (driver)', 'Date(s) of visit', 'Purpose', 'Relationship', 'Notes', 'Signed by (respondent)', 'Flags'];
  var isDoD = function (x) { return / holder\)$/.test(x.cat); };   // military/DoD and federal PIV holders
  var dods = c.people.filter(isDoD).map(function (x) {
    return [x.cat, x.event, x.cap, x.sponsor, x.p.last || '', x.p.first || '', x.p.idExpires || '', vehicle_(x.p), x.rel, x.flags.join('; ')];
  });
  var rows = c.people.filter(function (x) { return x.cat.indexOf('under 18, not driving') < 0 && !isDoD(x); }).map(function (x) {
    var p = x.p;
    var guest = x.cat.indexOf('Guest') === 0;
    return [x.cat, x.event, x.cap, x.sponsor, p.last || '', p.first || '', p.middle || '', p.suffix || '', p.dob || '', ageOn_(p.dob),
      p.ssn || '', p.idType || '', p.idNumber || '', issuerStd_(p.idIssuer, p.idType), expShow_(p.idExpires),
      p.doc2Type || '', p.doc2Number || '', p.address || '', p.phone || '', p.sex || '', citizenStd_(p.citizen), p.immDoc || '',
      vehicle_(p), guest ? CONFIG.VISIT_DATES.Guest : (CONFIG.VISIT_DATES[x.event] || ''),
      guest ? CONFIG.PURPOSE.guest : CONFIG.PURPOSE.attendee, x.rel, p.notes || '', x.signed || '', x.flags.join('; ')];
  });
  var minors = c.people.filter(function (x) { return x.cat.indexOf('under 18, not driving') >= 0; }).map(function (x) {
    return [x.event, x.cap, x.sponsor, x.p.last || '', x.p.first || '', x.p.dob || '', ageOn_(x.p.dob), x.rel, CONFIG.VISIT_DATES.Guest];
  });
  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HHmm');
  var out = SpreadsheetApp.create('CUI - SECFOR export ' + stamp + ' (CONTAINS SSNs - delete after sending)');
  writeTab_(out, 'Access roster', H, rows);
  writeTab_(out, 'Guests under 18 (not driving)', ['Event', 'Sponsor CAPID', 'Sponsor', 'Last name', 'First name', 'Date of birth',
    'Age on ' + CONFIG.GRADUATION, 'Relationship', 'Date of visit'], minors);
  writeTab_(out, 'DoD and PIV holders (no access request)', ['Category', 'Event', 'Sponsor CAPID', 'Sponsor', 'Last name (or full name)',
    'First name', 'DoD ID expiration', 'Vehicle (driver)', 'Relationship', 'Flags'], dods);
  var first = out.getSheets()[0];
  if (first.getName() !== 'Access roster') out.deleteSheet(first);
  var log = ss_().getSheetByName(CONFIG.EXPORT_LOG_SHEET) || ss_().insertSheet(CONFIG.EXPORT_LOG_SHEET);
  if (log.getLastRow() === 0) log.appendRow(['Created', 'By', 'People', 'File', 'Deleted?']);
  log.appendRow([new Date(), Session.getActiveUser().getEmail(), rows.length + minors.length + dods.length, out.getUrl(), '']);
  ui.alert('Export created (' + rows.length + ' with ID, ' + minors.length + ' minors, ' + dods.length + ' DoD / PIV holders):\n' + out.getUrl() +
    '\n\nIt is private to you. Send it as ' + INST + ' security directs (for example DoD SAFE), then delete it and mark it deleted on the Export log tab.');
}

// ---------------------------------------------------------------- roster sync from the Reg Zone exports
// Adds anyone new, refreshes rank/name/age/email/address for everyone, and marks people no longer in any export.
// Never deletes rows: people added by hand stay (they show "Not in latest Reg Zone pull").
function syncRosterFromRegZone(e) {
  var triggered = !!(e && e.triggerUid);
  if (triggered && !isAutomationOwner_()) {
    // A daily-pull trigger set up from a different Google account than the one running the automation: it used to stop
    // here silently every morning. Now it removes itself and says so once.
    var me = Session.getEffectiveUser().getEmail();
    ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'syncRosterFromRegZone') ScriptApp.deleteTrigger(t); });
    sendMail_(CONFIG.CONTACT_EMAIL, 'Base access roster: daily Reg Zone pull did NOT run',
      'The daily Reg Zone pull was turned on from ' + me + ', but the emails and automation run from ' + automationOwner_() + '.\n\n' +
      'So the pull did not run, and it has been turned off for ' + me + '.\n\n' +
      'To fix: open the spreadsheet signed in as ' + automationOwner_() + ' and pick Base Access > Turn on daily Reg Zone pull (6 AM).');
    return;
  }
  if (!CONFIG.REGZONE_FOLDER_ID) {
    var m0 = 'No Reg Zone folder is set (CONFIG.REGZONE_FOLDER_ID). Fill in the Roster tab yourself (see README), or set the folder.';
    try { SpreadsheetApp.getUi().alert(m0); } catch (e) { console.log(m0); }
    return;
  }
  var folder;
  try { folder = DriveApp.getFolderById(CONFIG.REGZONE_FOLDER_ID); } catch (err) {
    var fmsg = 'The Reg Zone pull did NOT run: ' + Session.getEffectiveUser().getEmail() + ' cannot open the Reg Zone folder. Share the folder with this account.';
    try { SpreadsheetApp.getUi().alert(fmsg); } catch (x) { sendMail_(CONFIG.CONTACT_EMAIL, 'Base access roster: Reg Zone pull did NOT run', fmsg); }
    return;
  }
  var recs = {}, missing = [];
  Object.keys(CONFIG.REGZONE_FILES).forEach(function (ev) {
    var it = folder.getFilesByName(CONFIG.REGZONE_FILES[ev]);
    if (!it.hasNext()) { missing.push(CONFIG.REGZONE_FILES[ev]); return; }
    var rows = xlsxRows_(it.next().getBlob());
    var H = rows.shift(), ix = {};
    H.forEach(function (k, j) { ix[k] = j; });
    rows.forEach(function (x) {
      var v = function (k) { return ix[k] === undefined ? '' : String(x[ix[k]] === undefined ? '' : x[ix[k]]).trim(); };
      var cap = v('RegistrantsCAPID');
      if (!/^[0-9]{5,7}$/.test(cap)) return;
      var age = Number(v('AgeAtEventStart'));
      var zip = v('Zip').replace(/\.0$/, ''); if (/^[0-9]{9}$/.test(zip)) zip = zip.slice(0, 5) + '-' + zip.slice(5);
      var street = [v('Addr1'), v('Addr2')].filter(String).join(', ');
      recs[cap] = {
        'CAPID': cap, 'Rank': RANK_FIX[v('Rank')] || v('Rank'), 'Last name': nameCase_(v('NameLast')), 'First name': nameCase_(v('NameFirst')), 'Event': ev,
        'Member type': v('MbrType').toUpperCase() === 'CADET' ? 'Cadet' : 'Senior', 'Birth month/year (Reg Zone)': v('DOB'),
        'Age at event start': isNaN(age) ? '' : age, 'Must submit own ID (18+ at start)': age >= 18 ? 'Yes' : 'No',
        'Email': v('Email'), 'Parent email': v('CadetParentEmailPrimary'),
        'Unit': v('Unit') ? v('Region') + '-' + v('Wing') + '-' + ('000' + v('Unit').replace(/\.0$/, '')).slice(-3) : '',
        'Home address (Reg Zone)': street ? street + ', ' + v('City') + ', ' + v('State') + ' ' + zip : '',
        'Slot status': ev === CONFIG.STAFF_GROUP ? 'Staff' : v('RegistrationStatus')
      };
    });
  });
  if (missing.length) {
    var msg = 'Could not find ' + missing.join(', ') + ' in the Reg Zone folder. Nothing was changed. ' +
      'Make sure this Google account can open the folder.';
    try { SpreadsheetApp.getUi().alert(msg); } catch (x) { sendMail_(CONFIG.CONTACT_EMAIL, 'Base access roster: Reg Zone pull did NOT run', msg); }
    return;
  }
  var r = rosterRows_(), h = r.header;
  var cStatus = ensureCol_(r.sheet, h, 'Reg Zone status') - 1;
  var fields = ['CAPID', 'Rank', 'Last name', 'First name', 'Event', 'Member type', 'Birth month/year (Reg Zone)', 'Age at event start',
    'Must submit own ID (18+ at start)', 'Email', 'Parent email', 'Unit', 'Home address (Reg Zone)', 'Slot status'];
  fields.forEach(function (f) { ensureCol_(r.sheet, h, f); });
  var seen = {}, added = [], gone = [];
  var rows = r.rows.map(function (x) {
    var row = h.map(function (k, j) { return x[j] === undefined ? '' : x[j]; });
    var cap = String(row[h.indexOf('CAPID')]).trim(), rec = recs[cap];
    seen[cap] = true;
    if (rec) {
      fields.forEach(function (f) { if (rec[f] !== '') row[h.indexOf(f)] = rec[f]; });
      row[cStatus] = 'Registered';
    } else {
      if (row[cStatus] !== 'Not in latest Reg Zone pull') gone.push(rosterName_(rowObj_(h, row)));
      row[cStatus] = 'Not in latest Reg Zone pull';
    }
    return row;
  });
  Object.keys(recs).forEach(function (cap) {
    if (seen[cap]) return;
    var rec = recs[cap], row = h.map(function () { return ''; });
    fields.forEach(function (f) { row[h.indexOf(f)] = rec[f]; });
    row[cStatus] = 'Registered';
    rows.push(row);
    added.push(rec.Rank + ' ' + rec['First name'] + ' ' + rec['Last name'] + ' (' + rec.Event + ')');
  });
  r.sheet.getRange(2, h.indexOf('CAPID') + 1, rows.length, 1).setNumberFormat('@');
  r.sheet.getRange(2, h.indexOf('Birth month/year (Reg Zone)') + 1, rows.length, 1).setNumberFormat('@');
  rows.forEach(function (row) {                       // write back as text, not a date
    var c = h.indexOf('Birth month/year (Reg Zone)'); row[c] = regZoneMMYY_(row[c]) || row[c];
  });
  r.sheet.getRange(2, 1, rows.length, h.length).setValues(rows);
  var v = updateCapidValidation();
  writePersonalLinks();
  rebuild();
  props_().setProperty('LAST_SYNC', new Date().toISOString());
  var summary = 'Reg Zone pull done.\n\nAdded ' + added.length + (added.length ? ':\n- ' + added.join('\n- ') : '') +
    '\n\nNo longer in Reg Zone (kept on the roster, marked): ' + gone.length + (gone.length ? '\n- ' + gone.join('\n- ') : '') +
    '\n\n' + v + '\nNew people have personal links on the Roster tab. Use "Email personal links" to send them.';
  try { SpreadsheetApp.getUi().alert(summary); } catch (e) {
    if (added.length || gone.length) sendMail_(CONFIG.CONTACT_EMAIL, 'Base access roster: daily Reg Zone pull', summary);
  }
}

function showAutomationStatus() {
  var me = Session.getEffectiveUser().getEmail(), owner = automationOwner_() || '(not set: this account)';
  var mine = ScriptApp.getProjectTriggers().map(function (t) { return t.getHandlerFunction(); });
  var has = function (f) { return mine.indexOf(f) >= 0 ? 'ON' : 'off'; };
  var last = props_().getProperty('LAST_SYNC');
  var tz = Session.getScriptTimeZone();
  SpreadsheetApp.getUi().alert('Signed in as: ' + me + '\nEmails and automation run from: ' + owner + '\n\n' +
    'Scheduled jobs set up by THIS account (Google only shows your own):\n' +
    '- Daily Reg Zone pull: ' + has('syncRosterFromRegZone') + '\n' +
    '- Confirmation emails on submit: ' + has('onFormSubmitHandler') + '\n' +
    '- Automatic follow-ups (Mon + Thu): ' + has('autoFollowUps') + '\n' +
    '- Email queue: ' + has('sendQueuedEmails') + '\n' +
    '- Weekly approval reminders (Registration reminders menu, Mondays): ' + has('weeklyApprovalReminders') + '\n\n' +
    'Last Reg Zone pull that finished: ' + (last ? Utilities.formatDate(new Date(last), tz, 'EEE d MMM, h:mm a') : 'not recorded yet') + '\n' +
    'Script time zone (the "6 AM" is in this zone): ' + tz +
    (owner.toLowerCase() !== me.toLowerCase() && automationOwner_() ? '\n\nYou are NOT the automation account, so the list above is not the one that matters. Check from ' + owner + '.' : ''));
}

function turnOnDailySync() {
  if (!isAutomationOwner_()) {
    SpreadsheetApp.getUi().alert('Emails and automation run from ' + automationOwner_() + ', not from this account (' +
      Session.getEffectiveUser().getEmail() + ').\n\nSign in as ' + automationOwner_() + ' and turn on the daily pull from there. ' +
      '(Or pick "Run emails and automation from MY Google account" first, to move everything to this account.)');
    return;
  }
  if (!CONFIG.REGZONE_FOLDER_ID) { SpreadsheetApp.getUi().alert('Set CONFIG.REGZONE_FOLDER_ID first (see README).'); return; }
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'syncRosterFromRegZone') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('syncRosterFromRegZone').timeBased().everyDays(1).atHour(6).create();
  props_().setProperty('DAILY_SYNC', 'on');
  SpreadsheetApp.getUi().alert('The roster will pull from the Reg Zone exports every morning around 6 AM. ' +
    'You will get an email only when someone is added or drops off.');
}

var RANK_WORDS = /^(c|cadet|cadets|airman|amn|basic|ab|first|class|a1c|senior|sra|staff|ssgt|technical|tsgt|master|msgt|smsgt|chief|cmsgt|sergeant|sgt|second|2d|2nd|2dlt|first|1st|1stlt|lieutenant|lt|lieutenant|captain|capt|major|maj|colonel|col|ltcol|lt|sm|member|mr|mrs|ms|encampment|staff|ats|student)$/;
var RANK_FIX = { 'CADET': 'C/AB' };   // Reg Zone shows C/AB as "CADET". Cadet officer grades stay as CAP writes them (C/2dLt, C/1stLt, C/LtCol).

// Minimal .xlsx reader (first worksheet only): returns rows as arrays of strings, header row first.
function xlsxRows_(blob) {
  var parts = Utilities.unzip(blob.setContentType('application/zip')), files = {};
  parts.forEach(function (b) { files[b.getName()] = b; });
  var text = function (name) { return files[name] ? files[name].getDataAsString('UTF-8') : ''; };
  var dec = function (t) {
    return t.replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"').replace(/&apos;/g, "'")
      .replace(/&#(\d+);/g, function (m, n) { return String.fromCharCode(Number(n)); }).replace(/&amp;/g, '&');
  };
  var shared = [];
  var sst = text('xl/sharedStrings.xml'), m, re = /<si>([\s\S]*?)<\/si>/g;
  while ((m = re.exec(sst))) {
    var t = '', tm, tre = /<t[^>]*>([\s\S]*?)<\/t>/g;
    while ((tm = tre.exec(m[1]))) t += tm[1];
    shared.push(dec(t));
  }
  var sheetName = Object.keys(files).filter(function (n) { return /^xl\/worksheets\/sheet\d+\.xml$/.test(n); }).sort()[0];
  var xml = text(sheetName), out = [], rre = /<row[^>]*>([\s\S]*?)<\/row>/g, rm;
  var colNum = function (letters) { var n = 0; for (var i = 0; i < letters.length; i++) n = n * 26 + letters.charCodeAt(i) - 64; return n - 1; };
  while ((rm = rre.exec(xml))) {
    var row = [], cre = /<c r="([A-Z]+)\d+"([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g, cm;
    while ((cm = cre.exec(rm[1]))) {
      var type = (/ t="([^"]+)"/.exec(cm[2]) || [])[1], body = cm[3] || '', val = '';
      if (type === 'inlineStr') { var it = /<t[^>]*>([\s\S]*?)<\/t>/.exec(body); val = it ? dec(it[1]) : ''; }
      else { var vm = /<v>([\s\S]*?)<\/v>/.exec(body); val = vm ? (type === 's' ? shared[Number(vm[1])] : dec(vm[1])) : ''; }
      row[colNum(cm[1])] = val;
    }
    for (var i = 0; i < row.length; i++) if (row[i] === undefined) row[i] = '';
    out.push(row);
  }
  return out;
}

// ---------------------------------------------------------------- which Google account runs the automation
// Automatic triggers run as the account that created them, and emails come from that account (with its daily limit).
// takeOverAutomation(): run while signed in as the account that should send (for example your wing Google Workspace account). It removes this account's
// old triggers, creates fresh ones, and records this account as the owner. Triggers left under any other account
// then do nothing, so there are never double emails.
function automationOwner_() { return props_().getProperty('AUTOMATION_OWNER') || ''; }
function isAutomationOwner_() {
  var owner = automationOwner_(), me = '';
  try { me = Session.getEffectiveUser().getEmail(); } catch (e) {}
  return !owner || !me || owner.toLowerCase() === me.toLowerCase();
}
function takeOverAutomation() {
  var ui = SpreadsheetApp.getUi(), me = Session.getEffectiveUser().getEmail();
  var form;
  try { form = form_(); form.getTitle(); } catch (e) {
    ui.alert('This account (' + me + ') cannot open the form. Share the form with it as an EDITOR first ' +
      '(open the form > three dots > Add collaborators), then run this again.'); return;
  }
  ScriptApp.getProjectTriggers().forEach(function (t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('onFormSubmitHandler').forForm(form).onFormSubmit().create();
  ScriptApp.newTrigger('sendQueuedEmails').timeBased().everyHours(1).create();
  var syncNote = '';
  if (props_().getProperty('DAILY_SYNC') === 'on') {
    try { DriveApp.getFolderById(CONFIG.REGZONE_FOLDER_ID).getName();
      ScriptApp.newTrigger('syncRosterFromRegZone').timeBased().everyDays(1).atHour(6).create();
      syncNote = '\n- Daily Reg Zone pull: on.';
    } catch (e) { syncNote = '\n- Daily Reg Zone pull: OFF, because this account cannot open the Reg Zone folder. Share that folder with ' + me + ' and run this again.'; }
  }
  if (props_().getProperty('AUTO_FOLLOWUPS') === 'on') {
    [ScriptApp.WeekDay.MONDAY, ScriptApp.WeekDay.THURSDAY].forEach(function (d) {
      ScriptApp.newTrigger('autoFollowUps').timeBased().onWeekDay(d).atHour(CHASE.HOUR).create();
    });
    syncNote += '\n- Automatic follow-ups (Mon + Thu): on.';
  }
  if (props_().getProperty('WEEKLY_APPROVALS') === 'on') {
    ScriptApp.newTrigger('weeklyApprovalReminders').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).create();
    syncNote += '\n- Weekly approval reminders: on.';
  }
  props_().setProperty('AUTOMATION_OWNER', me);
  ui.alert('Done. Emails and automation now run from ' + me + '.\n' +
    '- Emails this account can still send today: ' + MailApp.getRemainingDailyQuota() + '.' + syncNote +
    '\n- Old triggers under any other account will now do nothing, so there will be no double emails.');
}

// ---------------------------------------------------------------- after the event
function deleteAllData() {
  var ui = SpreadsheetApp.getUi();
  if (ui.alert('Delete ALL form data?', 'This permanently deletes every form response (all SSNs, IDs, guests) and clears the tabs. ' +
      'It cannot be undone. Continue?', ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  if (ui.prompt('Type DELETE to confirm').getResponseText() !== 'DELETE') return;
  form_().deleteAllResponses();
  rebuild();
  ui.alert('All responses deleted. Also delete any SECFOR export files listed on the Export log tab, and empty your Drive trash.');
}

// ---------------------------------------------------------------- flags + tab writer
function shortFlags_(flags) {   // drop the informational 'no SSN' note from summary columns
  return flags.map(function (f) { return f.replace(/(; )?no SSN \(optional\)/g, '').replace(/: ; /, ': '); })
    .filter(function (f) { return !/:\s*$/.test(f); }).join(' | ');
}

// A parent listed as a guest on two siblings' forms (or a staff member also listed as a guest) must appear ONCE on the
// lists for installation security. Same first + last name + date of birth = same person. Keep the attendee row, else the first.
function mergeSamePerson_(people) {
  var rankOf = function (x) { return x.cat.indexOf('Attendee') === 0 ? 0 : x.cat === 'Guest (driver)' ? 1 : x.cat === 'Guest' ? 2 : 3; };
  var keyOf = function (x) {
    var n = function (v) { return String(v || '').toLowerCase().normalize('NFD').replace(/[^a-z]/g, ''); };
    var d = x.p.dob, ds = d instanceof Date ? d.getFullYear() + '-' + (d.getMonth() + 1) + '-' + d.getDate() : String(d || '').trim();
    return ds && n(x.p.first) && n(x.p.last) ? n(x.p.first) + '|' + n(x.p.last) + '|' + ds : '';
  };
  var byKey = {}, out = [];
  people.forEach(function (x) {
    var k = keyOf(x);
    if (!k) { out.push(x); return; }
    var y = byKey[k];
    if (!y) { byKey[k] = x; out.push(x); return; }
    var keep = rankOf(x) < rankOf(y) ? x : y, other = keep === x ? y : x;
    if (keep === x) out[out.indexOf(y)] = x;
    byKey[k] = keep;
    if (String(keep.sponsor).indexOf(other.sponsor) < 0) {
      keep.sponsor = keep.sponsor + ' / ' + other.sponsor;
      keep.cap = keep.cap + ' / ' + other.cap;
    }
    if (!keep.rel && other.rel && keep.cat.indexOf('Attendee') !== 0) keep.rel = other.rel;
    other.flags.forEach(function (f) { if (keep.flags.indexOf(f) < 0) keep.flags.push(f); });
  });
  return out;
}

function personFlags_(p, END) {
  var f = [];
  if (isExemptId_(p.idType)) {
    if (p.idExpires && new Date(p.idExpires + 'T00:00:00') < END) f.push('DoD ID expires ' + p.idExpires + ', before ' + CONFIG.GRADUATION);
    return f;                                         // DoD / PIV holders: no access request, so nothing else is needed
  }
  var ssn = String(p.ssn || '').replace(/\D/g, '');
  var needSsn = CONFIG.SSN_REQUIRED_FOR_ALL || SSN_REQUIRED_TYPES.indexOf(p.idType) >= 0;
  if (!ssn) f.push(needSsn ? 'SSN REQUIRED but missing' : 'no SSN (optional)');
  else if (/^(\d)\1{8}$/.test(ssn) || /^(000|666|9)/.test(ssn) || ssn.slice(3, 5) === '00' || ssn.slice(5) === '0000')
    f.push('SSN looks invalid');
  if (SSN_REQUIRED_TYPES.indexOf(p.idType) >= 0 && !p.doc2Type) f.push('non-REAL ID without a second ID');
  // A second ID must be one they HAVE now, with its number ("Will be obtaining my Real ID" or a blank number is not usable)
  else if (SSN_REQUIRED_TYPES.indexOf(p.idType) >= 0 &&
           (/\b(will|going to|plan|planning|obtain|obtaining|getting|apply|applying|applied|pending|in progress|don'?t have|do not have|none|n\/?a)\b/i.test(String(p.doc2Type)) ||
            !/\d/.test(String(p.doc2Number || ''))))
    f.push('second ID not usable (must be an ID they have now, with its number)');
  if (p.driving && DRIVER_TYPES.indexOf(p.idType) < 0) f.push('driving but no driver\'s license given');
  var age = ageOn_(p.dob);
  if (p.driving && age !== '' && age < 16) f.push('listed as a DRIVER but only ' + age + ' (too young for a license)');
  // "Issued by" must name a state, country or federal agency, not a kind of ID ("Passport", "ID card") or a nameless office ("DMV")
  var iss0 = String(p.idIssuer || '').trim();
  var issWhere = iss0.replace(/\b(u\.?s\.?\s+)?(passport|passports|driver'?s?|drivers|licen[cs]es?|dl|id|ids|card|cards|identification|real|star|state\s+id|government|gov'?t|photo|federal\s+limits\s+apply|dmv|department|dept|motor|vehicles?|bureau|office|agency|of|the)\b/gi, '').replace(/[^A-Za-z]/g, '');
  if (iss0 && !issWhere && !/^(u\.?s\.?a?\.?)$/i.test(iss0)) f.push('"ID issued by" names a kind of ID or office, not a state, country or federal agency');
  // The issuer must fit the ID type: a license or state ID comes from a STATE; a U.S. passport or federal
  // card from the United States (not a state); a foreign passport from a foreign COUNTRY.
  if (iss0 && issWhere && !/^(state|country|n\/?a|none|unknown)$/i.test(iss0)) {
    var st = isUSState_(iss0), us = isUSA_(iss0);
    if (STATE_ID_TYPES.indexOf(p.idType) >= 0 && !st) f.push('"ID issued by" should be the STATE for a driver\'s license or state ID');
    else if (US_FED_TYPES.indexOf(p.idType) >= 0 && st) f.push('"ID issued by" should be the United States (or the federal agency) for this ID, not a state');
    else if (p.idType === 'Foreign passport' && (us || (st && !/^georgia$/i.test(iss0)))) f.push('"ID issued by" should be the foreign COUNTRY for a foreign passport');
  }
  var idn = String(p.idNumber || '').replace(/[^A-Za-z0-9]/g, '');
  if (idn && /^(.)\1+$/.test(idn)) f.push('ID number looks like a placeholder');
  if (idn && p.__capid && idn.indexOf(p.__capid) >= 0) f.push('ID number looks like the CAPID, not the number on the ID');
  var iss = String(p.idIssuer || '').trim();
  if (p.idType && (!iss || /^(state|country|n\/?a|none|unknown)$/i.test(iss))) f.push('ID issuing state or country missing');
  if (p.idType && !p.idExpires) f.push('no ID expiration date');
  if (p.idExpires && !noExpiry_(p.idExpires) && new Date(p.idExpires + 'T00:00:00') < END) f.push('ID expires ' + p.idExpires + ', before ' + CONFIG.GRADUATION);
  if (p.idType === ID_OTHER && !p.notes) f.push('"Other" ID type not explained');
  if (p.idType && !p.address) f.push('home address missing');
  if (p.idType && p.phoneSame !== undefined && !p.phone) f.push('home/cell phone missing');
  if (p.citizen && !isUSCitizenText_(p.citizen) && !p.immDoc)
    f.push('citizenship "' + p.citizen + '" but no immigration document given');
  return f;
}

function writeTab_(ss, name, header, rows) {
  var sh = ss.getSheetByName(name) || ss.insertSheet(name);
  sh.clearContents();
  sh.getRange(1, 1, 1, header.length).setValues([header]).setFontWeight('bold').setWrap(true);
  if (!rows.length) return;
  rows = rows.map(function (row) { return row.map(function (v) { return v instanceof Date ?
    Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') : v; }); });
  var rg = sh.getRange(2, 1, rows.length, header.length);
  rg.setNumberFormat('@');   // plain text first, so SSNs and ID numbers keep leading zeros
  rg.setValues(rows);
}


// ================================================================ REGISTRATION (APPROVAL) REMINDERS
// Emails everyone whose Reg Zone registration is missing a unit approval or (under 18) a parent approval.
// Source: the Reg Zone exports in the Reg Zone folder (same files as the roster pull). Includes alternates (a missing
// approval keeps them from being slotted) and staff.
//   Missing parent approval -> To: parent + member. CC: unit commander (families often need help finding where to approve).
//   Missing unit approval   -> To: member (+ parent if under 18). CC: unit commander (the commander is the one who approves).
// Each person gets at most one reminder every REMIND_EVERY_DAYS days, however often this runs. Every run is logged on the
// "Approval reminders" tab. RST, CPPT and fees are NOT included (add your own if needed).
var APPROVALS = {
  REMIND_EVERY_DAYS: 6,
  COMBINE_AT: 3,                   // a commander with this many (or more) pending registrations gets ONE combined email, not a CC on each
  MAX_DATA_AGE_HOURS: 36,          // the weekly run skips (and emails the contact) if the exports are older than this
  TAB: 'Approval reminders',
  // Deadline sentence. Blank = no date given ("as soon as possible").
  DEADLINE: '',
  // Optional how-to text for parents (step by step). Blank = generic wording that points them to the commander and the admin contacts.
  PARENT_HOWTO: 'How to approve:\n' +
    '  1. Find the email from eServices saying this event needs Parent/Guardian approval, and click the link in it. ' +
    '(Or log in to eServices at https://www.capnhq.gov/eServices.aspx and open the event in Registration Zone.)\n' +
    '  2. In Event Details, click the red "Event Approval" button. (A message that the cadet is already registered for another event at that time does not block approval.)\n' +
    '  3. Choose Approve, then click Submit at the bottom of the screen.\n' +
    '  No eServices parent account yet? The cadet starts the link from their own eServices account; ' +
    'you then verify with the Parent Access Code and register a guest account. Step-by-step instructions and short videos: ' +
    'https://www.gocivilairpatrol.com/programs/cadets/parents/parent-guardian-portal\n' +
    '  Still stuck? The unit commander (copied on this email) can help, or see the end of this email.',
  UNIT_HOWTO: '',
  // Cadet admin leads: anyone stuck should email all of these.
  HELP_CONTACTS: CONFIG.CONTACT_NAME + ', ' + CONFIG.CONTACT_TITLE + ' (' + CONFIG.CONTACT_EMAIL + ')'   // add more people if you like
};

function approvalData_() {
  if (!CONFIG.REGZONE_FOLDER_ID) throw new Error('Approval reminders need the Reg Zone exports: set CONFIG.REGZONE_FOLDER_ID (see README).');
  var folder = DriveApp.getFolderById(CONFIG.REGZONE_FOLDER_ID), out = [], oldest = null, missing = [];
  Object.keys(CONFIG.REGZONE_FILES).forEach(function (ev) {
    var it = folder.getFilesByName(CONFIG.REGZONE_FILES[ev]);
    if (!it.hasNext()) { missing.push(CONFIG.REGZONE_FILES[ev]); return; }
    var f = it.next(), upd = f.getLastUpdated();
    if (!oldest || upd < oldest) oldest = upd;
    var rows = xlsxRows_(f.getBlob()), H = rows.shift(), ix = {};
    H.forEach(function (k, j) { ix[k] = j; });
    rows.forEach(function (x) {
      var v = function (k) { return ix[k] === undefined ? '' : String(x[ix[k]] === undefined ? '' : x[ix[k]]).trim(); };
      var cap = v('RegistrantsCAPID');
      if (!/^[0-9]{5,7}$/.test(cap)) return;
      var st = v('RegistrationStatus');
      if (/withdr|cancel|disapprov|denied/i.test(st)) return;
      var age = Number(v('AgeAtEventStart')), minor = !(age >= 18);
      var needUnit = v('UnitApproved') !== 'Yes';
      var needParent = minor && v('ParentApproved') !== 'Approved';
      if (!needUnit && !needParent) return;
      out.push({ ev: ev, cap: cap, rank: RANK_FIX[v('Rank')] || v('Rank'), first: nameCase_(v('NameFirst')), last: nameCase_(v('NameLast')), status: st,
        minor: minor, needUnit: needUnit, needParent: needParent, email: v('Email'), parent: v('CadetParentEmailPrimary'),
        ccName: v('UnitCCName'), ccEmail: v('UnitCCEmail'), unit: v('Unit').replace(/\.0$/, ''), wing: v('Wing') });
    });
  });
  return { list: out, asOf: oldest, missing: missing };
}

function approvalEmail_(p, asOf, digest) {
  var who = (p.rank + ' ' + p.first + ' ' + p.last).trim();
  var items = [], subjBits = [];
  var parentHow = APPROVALS.PARENT_HOWTO || 'If you are not sure where to approve, the unit commander (copied on this email) can help, or reply to this email.';
  var unitHow = APPROVALS.UNIT_HOWTO || 'WHAT TO DO: ask your unit commander to approve it in Reg Zone.';
  if (digest) {
    parentHow = parentHow.replace(' (copied on this email)', '');
    unitHow = unitHow + ' (Your unit commander was also sent a list of everyone in the unit who is waiting.)';
  }
  if (p.needParent) { subjBits.push('parent approval');
    items.push('- PARENT/GUARDIAN APPROVAL is missing. A parent or guardian must approve this registration in Reg Zone (CAP ' +
      'eServices).\n' + parentHow); }
  if (p.needUnit) { subjBits.push('unit commander approval');
    items.push('- UNIT COMMANDER APPROVAL is missing. The unit commander' + (p.ccName ? ', ' + rankLast_(p.ccName) + ',' : '') +
      ' must approve this registration in Reg Zone.\n  ' + unitHow); }
  var student = p.ev !== 'Staff';
  var when = APPROVALS.DEADLINE ? 'by ' + APPROVALS.DEADLINE : 'as soon as possible';
  var asOfTxt = asOf ? Utilities.formatDate(asOf, Session.getScriptTimeZone(), 'EEE d MMM h:mm a') : 'latest';
  var body = 'Hello,\n\n' + who + ' (CAPID ' + p.cap + ') is registered for the ' + CONFIG.EVENT_NAME + ' (' + p.ev + ') at ' + INST +
    ', ' + (CONFIG.VISIT_DATES[p.ev] || '') + '.\n\n' +
    'As of the Reg Zone report of ' + asOfTxt + ', the registration is still missing:\n\n' + items.join('\n\n') + '\n\n' +
    (student
      ? 'Please get this done ' + when + '. A registration that is missing an approval cannot be slotted. That includes ALTERNATES: ' +
        'if a slot opens, only cadets with every approval can be moved in.\n\n'
      : 'Please get this done ' + when + '. The staff registration is not complete until it is approved.\n\n') +
    (p.ccEmail && !digest ? 'Commander: you are copied because ' + (p.needUnit ? 'this registration is waiting for your approval in Reg Zone.'
      : 'families often need help finding where to give parent approval in Reg Zone.') + '\n\n' : '') +
    'If this is already done, thank you, and please ignore this email (Reg Zone reports can be a day behind).\n\n' +
    'Stuck or have questions? Email all three: ' + APPROVALS.HELP_CONTACTS + '.\n\n' +
    SIGNATURE;
  var subject = 'ACTION NEEDED: ' + subjBits.join(' and ') + ' for ' + who + ' (' + CONFIG.EVENT_NAME + ')';
  var seen = {}, to = [], cc = [];
  function add(list, a) { a = String(a || '').trim(); if (a && a.indexOf('@') > 0 && !seen[a.toLowerCase()]) { seen[a.toLowerCase()] = 1; list.push(a); } }
  if (p.needParent) add(to, p.parent);
  add(to, p.email);
  if (!p.needParent && p.minor) add(to, p.parent);
  if (!digest) add(cc, p.ccEmail);
  return { to: to, cc: cc, subject: subject, body: body };
}

// "Maj Christopher D. Devine" -> "Maj Devine" (CAP style: rank and last name).
function rankLast_(full) {
  var m = /^((?:C\/)?(?:Lt Col|1st Lt|2d Lt|LtCol|1stLt|2dLt|[A-Za-z]+))\s+.*?(\S+)$/.exec(String(full || '').trim());
  return m ? m[1] + ' ' + m[2] : String(full || '').trim();
}

// One email to a unit commander listing every registration from the unit that is missing an approval.
function commanderDigest_(ccEmail, people, asOf) {
  var name = people[0].ccName ? rankLast_(people[0].ccName) : 'Commander', unit = people[0].unit ? (people[0].wing ? people[0].wing + '-' : '') + ('000' + people[0].unit).slice(-3) : 'your unit';
  var asOfTxt = asOf ? Utilities.formatDate(asOf, Session.getScriptTimeZone(), 'EEE d MMM h:mm a') : 'latest';
  function line(p) {
    var self = String(p.email || '').toLowerCase() === ccEmail.toLowerCase();
    return '- ' + (p.rank + ' ' + p.first + ' ' + p.last).trim() + ' (' + p.cap + '), ' + p.ev +
      (p.ev !== 'Staff' && p.status ? ', ' + p.status : '') + (self ? '  <- your own registration' : '');
  }
  var unitList = people.filter(function (p) { return p.needUnit; }), parList = people.filter(function (p) { return p.needParent; });
  var body = 'Hello ' + name + ',\n\n' +
    'As of the Reg Zone report of ' + asOfTxt + ', these ' + CONFIG.EVENT_NAME + ' registrations from ' + unit +
    ' are missing an approval:\n\n' +
    (unitList.length ? 'WAITING ON YOUR UNIT APPROVAL (' + unitList.length + '):\n' + unitList.map(line).join('\n') + '\n\n' : '') +
    (parList.length ? 'WAITING ON PARENT/GUARDIAN APPROVAL (' + parList.length + '). The parents and cadets were emailed the steps; ' +
      'families often need a hand finding where to approve in Reg Zone:\n' + parList.map(line).join('\n') + '\n\n' : '') +
    (people.some(function (p) { return String(p.email || '').toLowerCase() === ccEmail.toLowerCase() && p.needUnit; })
      ? 'Your own registration is on the list. If Reg Zone will not let you approve your own, please ask whoever can approve for the unit.\n\n' : '') +
    'Students missing an approval cannot be slotted, including alternates. Staff registrations are not complete until approved. ' +
    'Each member (and the parent, for cadets under 18) was also emailed directly.\n\n' +
    'If some of these are already done, thank you (Reg Zone reports can be a day behind).\n\n' +
    'Stuck or have questions? Email all three: ' + APPROVALS.HELP_CONTACTS + '.\n\n' +
    SIGNATURE;
  return { to: [ccEmail], cc: [], subject: 'ACTION NEEDED: ' + people.length + ' registration' + (people.length === 1 ? '' : 's') + ' from ' + unit + ' missing approvals (' + CONFIG.EVENT_NAME + ')', body: body };
}

function approvalHistory_() { try { return JSON.parse(props_().getProperty('APPROVAL_SENT') || '{}'); } catch (e) { return {}; } }

// Builds the plan and writes it to the tab. Returns the plan.
function approvalPlan_() {
  var d = approvalData_(), hist = approvalHistory_(), now = new Date();
  var byCC = {};
  d.list.forEach(function (p) { var k = String(p.ccEmail || '').toLowerCase(); if (k) (byCC[k] = byCC[k] || []).push(p); });
  var combine = function (p) { var k = String(p.ccEmail || '').toLowerCase(); return k && byCC[k].length >= APPROVALS.COMBINE_AT; };
  var plan = d.list.map(function (p) {
    var key = p.ev + ':' + p.cap, h = hist[key], m = approvalEmail_(p, d.asOf, combine(p));
    var days = h ? (now - new Date(h.last)) / 86400000 : null;
    var send = !m.to.length ? 'No (no email address)' : (days !== null && days < APPROVALS.REMIND_EVERY_DAYS ? 'No (emailed ' + Math.floor(days) + ' days ago)' : 'Yes');
    return { p: p, key: key, m: m, send: send, h: h, combined: combine(p) };
  });
  var digests = Object.keys(byCC).filter(function (k) { return byCC[k].length >= APPROVALS.COMBINE_AT; }).map(function (k) {
    var key = 'CC:' + k, h = hist[key];
    // Commander was CC'd on individual emails before combined emails existed: count that as their last email.
    var last = h ? h.last : byCC[k].map(function (p) { var x = hist[p.ev + ':' + p.cap]; return x ? x.last : ''; }).sort().pop();
    var days = last ? (now - new Date(last)) / 86400000 : null;
    return { key: key, h: h, m: commanderDigest_(byCC[k][0].ccEmail, byCC[k], d.asOf), n: byCC[k].length, name: byCC[k][0].ccName,
             send: days !== null && days < APPROVALS.REMIND_EVERY_DAYS ? 'No (emailed ' + Math.floor(days) + ' days ago)' : 'Yes' };
  });
  var ss = ss_(), sh = ss.getSheetByName(APPROVALS.TAB) || ss.insertSheet(APPROVALS.TAB);
  sh.clearContents();
  var head = ['Event', 'CAPID', 'Name', 'Reg Zone status', 'Missing', 'To', 'CC (unit commander)', 'Times emailed', 'Last emailed', 'Send on next run?'];
  var rows = plan.map(function (x) {
    return [x.p.ev, x.p.cap, x.p.rank + ' ' + x.p.first + ' ' + x.p.last, x.p.status,
      [x.p.needParent ? 'Parent approval' : '', x.p.needUnit ? 'Unit approval' : ''].filter(String).join(' + '),
      x.m.to.join(', '), x.combined ? '(in commander\'s combined email)' : x.m.cc.join(', '), x.h ? x.h.n : 0, x.h ? Utilities.formatDate(new Date(x.h.last), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') : '', x.send];
  });
  sh.getRange(1, 1, 1, head.length).setValues([head]).setFontWeight('bold');
  if (rows.length) sh.getRange(2, 1, rows.length, head.length).setNumberFormat('@').setValues(rows);
  var r0 = rows.length + 3;
  if (digests.length) {
    sh.getRange(r0, 1).setValue('Combined emails to unit commanders (' + APPROVALS.COMBINE_AT + ' or more pending in their unit)').setFontWeight('bold');
    var drows = digests.map(function (g) { return ['Commander', '', g.name, '', g.n + ' pending', g.m.to.join(', '), '', g.h ? g.h.n : 0,
      g.h ? Utilities.formatDate(new Date(g.h.last), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm') : '', g.send]; });
    sh.getRange(r0 + 1, 1, drows.length, head.length).setNumberFormat('@').setValues(drows);
    r0 += drows.length + 2;
  }
  sh.getRange(r0, 1).setValue('Reg Zone data as of ' + (d.asOf ? Utilities.formatDate(d.asOf, Session.getScriptTimeZone(), 'EEE d MMM yyyy h:mm a') : '?') +
    ' (when the Reg Zone exports were last saved to Drive). Built ' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'EEE d MMM h:mm a') + '.');
  return { plan: plan, digests: digests, asOf: d.asOf, missing: d.missing };
}

function previewApprovalReminders() {
  var r = approvalPlan_(), ui = SpreadsheetApp.getUi();
  if (r.missing.length) { ui.alert('Could not find ' + r.missing.join(', ') + ' in the Reg Zone folder.'); return; }
  var go = r.plan.filter(function (x) { return x.send === 'Yes'; });
  try { ss_().getSheetByName(APPROVALS.TAB).activate(); } catch (e) {}
  ui.alert('Approval reminders (preview, nothing sent)', r.plan.length + ' registrations are missing an approval. ' + go.length +
    ' would be emailed now, plus ' + r.digests.filter(function (g) { return g.send === 'Yes'; }).length +
    ' combined emails to unit commanders. See the "' + APPROVALS.TAB + '" tab.', ui.ButtonSet.OK);
}

function testApprovalReminder() {
  var r = approvalPlan_(), ui = SpreadsheetApp.getUi();
  if (!r.plan.length) { ui.alert('Nobody is missing an approval right now.'); return; }
  var pick = r.plan.filter(function (x) { return x.p.needParent && x.p.needUnit; })[0] || r.plan[0];
  sendMail_(CONFIG.CONTACT_EMAIL, '[TEST] ' + pick.m.subject,
    '(TEST: the real email would go To: ' + pick.m.to.join(', ') + '  CC: ' + pick.m.cc.join(', ') + ')\n\n' + pick.m.body,
    { name: CONFIG.CONTACT_NAME });
  ui.alert('Sent a test to ' + CONFIG.CONTACT_EMAIL + ' using ' + pick.p.first + ' ' + pick.p.last + '\'s data. Nobody else was emailed.');
}

function sendApprovalReminders() { sendApprovals_(false); }
function weeklyApprovalReminders() { sendApprovals_(true); }

function sendApprovals_(auto) {
  if (auto && !isAutomationOwner_()) return;
  var r = approvalPlan_();
  var go = r.plan.filter(function (x) { return x.send === 'Yes'; });
  var goD = r.digests.filter(function (g) { return g.send === 'Yes'; });
  var ageH = r.asOf ? (new Date() - r.asOf) / 3600000 : 999;
  if (auto) {
    if (r.missing.length || ageH > APPROVALS.MAX_DATA_AGE_HOURS) {
      sendMail_(CONFIG.CONTACT_EMAIL, 'Approval reminders NOT sent: Reg Zone data is ' + Math.round(ageH / 24) + ' days old',
        'The weekly approval reminders did not go out because the Reg Zone exports in Drive are from ' + r.asOf +
        '. Get fresh Reg Zone exports saved to the Drive folder, then use Registration reminders > Send approval reminders now.');
      return;
    }
  } else {
    var ui = SpreadsheetApp.getUi();
    if (r.missing.length) { ui.alert('Could not find ' + r.missing.join(', ') + ' in the Reg Zone folder.'); return; }
    if (!go.length && !goD.length) { ui.alert('Nobody to email right now (see the "' + APPROVALS.TAB + '" tab).'); return; }
    var warn = ageH > APPROVALS.MAX_DATA_AGE_HOURS ? '\n\nWARNING: the Reg Zone data is ' + Math.round(ageH) + ' hours old. People who already finished may be emailed.' : '';
    if (ui.alert('Send approval reminders', 'Email ' + go.length + ' registrations missing an approval, plus ' + goD.length +
        ' combined emails to unit commanders (' + goD.map(function (g) { return g.name + ': ' + g.n; }).join(', ') + ')?\n\n' +
        go.map(function (x) { return x.p.first + ' ' + x.p.last + ' (' + x.p.ev + ')'; }).join(', ') + warn, ui.ButtonSet.YES_NO) !== ui.Button.YES) return;
  }
  var hist = approvalHistory_(), sent = 0, now = new Date().toISOString();
  go.forEach(function (x) {
    var opts = { replyTo: CONFIG.CONTACT_EMAIL, name: CONFIG.CONTACT_NAME };
    if (x.m.cc.length) opts.cc = x.m.cc.join(',');
    safeSend_(x.m.to.join(','), x.m.subject, x.m.body, opts);
    hist[x.key] = { last: now, n: (hist[x.key] ? hist[x.key].n : 0) + 1 };
    sent++;
  });
  goD.forEach(function (g) {
    safeSend_(g.m.to[0], g.m.subject, g.m.body, { replyTo: CONFIG.CONTACT_EMAIL, name: CONFIG.CONTACT_NAME });
    hist[g.key] = { last: now, n: (hist[g.key] ? hist[g.key].n : 0) + 1 };
  });
  props_().setProperty('APPROVAL_SENT', JSON.stringify(hist));
  approvalPlan_();   // refresh the tab with the new "last emailed" dates
  var msg = 'Sent ' + sent + ' approval reminders and ' + goD.length + ' combined commander emails.';
  if (auto) sendMail_(CONFIG.CONTACT_EMAIL, 'Weekly approval reminders: ' + sent + ' sent', msg + '\n\nDetails: "' + APPROVALS.TAB + '" tab of the base access sheet.');
  else SpreadsheetApp.getUi().alert(msg + (MailApp.getRemainingDailyQuota() <= 0 ? ' Some may be waiting on the Email queue tab.' : ''));
}

function turnOnWeeklyApprovalReminders() {
  turnOffWeeklyApprovalReminders(true);
  ScriptApp.newTrigger('weeklyApprovalReminders').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).create();
  props_().setProperty('WEEKLY_APPROVALS', 'on');
  SpreadsheetApp.getUi().alert('Approval reminders will go out every Monday around 8 AM, only if the Reg Zone exports in Drive are less than ' +
    APPROVALS.MAX_DATA_AGE_HOURS + ' hours old. You get a summary email each time (or a warning if the data is stale).');
}
function turnOffWeeklyApprovalReminders(quiet) {
  ScriptApp.getProjectTriggers().forEach(function (t) { if (t.getHandlerFunction() === 'weeklyApprovalReminders') ScriptApp.deleteTrigger(t); });
  if (quiet !== true) { props_().deleteProperty('WEEKLY_APPROVALS'); SpreadsheetApp.getUi().alert('Weekly approval reminders are off.'); }
}
