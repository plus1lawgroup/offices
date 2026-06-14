// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
//  PLUS ONE LAW GROUP â€” Google Apps Script Backend
//  Sheets: tenants Â· MeterData Â· Payments
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

const SHEETS = {
  TENANTS:     'tenants',
  ELECTRICITY: 'MeterData',
  PAYMENTS:    'Payments'
};

const DOCUMENT_TEMPLATES = {
  CONTRACT:  '1ZjHj8rmzSyJkcs6hpOOvCLev8c_fFOl3vn-cYSt7VWU',
  EXTENSION: '1_5InwAOi9cXo1sccip0jJsvcHn-CYJfFFURhh0Ht_hE'
};

const TENANT_DOCUMENTS_FOLDER = 'Plus One - Tenant Documents';
const MAX_TENANT_UPLOAD_BYTES = 15 * 1024 * 1024;

const DOCUMENT_TEMPLATE_CHECKS = {
  CONTRACT: {
    minimumParagraphs: 117,
    requiredText: [
      'ՊԱՅՄԱՆԱԳՐԻ ԱՌԱՐԿԱՆ',
      'ՊԱՅՄԱՆԱԳՐԻ ԺԱՄԿԵՏԸ',
      'ՊԱՅՄԱՆԱԳՐԻ ԳԻՆԸ ԵՎ ՎՃԱՐՄԱՆ ԿԱՐԳԸ',
      'ԿՈՂՄԵՐԻ ԻՐԱՎՈՒՆՔՆԵՐԸ ԵՎ ՊԱՐՏԱԿԱՆՈՒԹՅՈՒՆՆԵՐԸ',
      'ԿՈՂՄԵՐԻ ՊԱՏԱՍԽԱՆԱՏՎՈՒԹՅՈՒՆԸ',
      'ՎԱՐՁԱԿԱԼԱԾ ՕԲՅԵԿՏԻ ԲԱՐԵԼԱՎՈՒՄՆԵՐԸ',
      'ՀԱՏՈՒԿ ՊԱՅՄԱՆՆԵՐ',
      'ՊԱՅՄԱՆԱԳՐԻ ՎԱՂԱԺԱՄԿԵՏ ԼՈՒԾՄԱՆ ՀԻՄՔԵՐԸ',
      'ԱՆՀԱՂԹԱՀԱՐԵԼԻ ՈԻԺԻ ԱԶԴԵՑՈՒԹՅՈՒՆԸ',
      'ԵԶՐԱՓԱԿԻՉ ԴՐՈՒՅԹՆԵՐ',
      'ԿՈՂՄԵՐԻ ՏՎՅԱԼՆԵՐԸ և ԻՐԱՎԱԲԱՆԱԿԱՆ ՀԱՍՑԵՆԵՐԸ'
    ]
  },
  EXTENSION: {
    minimumParagraphs: 27,
    requiredText: [
      'ԼՐԱՑՈՒՑԻՉ ՀԱՄԱՁԱՅՆԱԳԻՐ',
      'Հոդված 1. Պայմանագրի ժամկետի երկարաձգում',
      'Հոդված 2. Վճարման հատուկ պայմաններ',
      'Հոդված 3. Եզրափակիչ դրույթներ',
      'ԿՈՂՄԵՐԻ ՏՎՅԱԼՆԵՐԸ ԵՎ ՍՏՈՐԱԳՐՈՒԹՅՈՒՆՆԵՐԸ'
    ]
  }
};


// â”€â”€ Electricity tariff rates (AMD per kWh) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
const AMD_PER_KWH_DAY   = 54;   // T1 Day tariff   â€” edit here
const AMD_PER_KWH_NIGHT = 45;   // T2 Night tariff â€” edit here

// Headers that will be auto-created if missing
const TENANT_HEADERS = [
  'Գրասենյակ',
  'Տեսակ',
  'Ֆիզ. անձի անուն','Ֆիզ. անձի անձնագիր','Ֆիզ. անձի հասցե','Ֆիզ. անձի տրված է','Ֆիզ. անձի վավեր է մինչև',
  'Իրավ. անձի անվանում','ՀՎՀՀ','Իրավ. հասցե','Գրանցման համար',
  'Տնօրենի անուն','Տնօրենի անձնագիր','Տնօրենի հասցե','Տնօրենի տրված է','Տնօրենի վավեր է մինչև',
  'Սկիզբ','Ավարտ',
  'Վճարման օր','Վարձ','Ինտերնետ',
  'Հեռախոս','Կապ','Ակտիվ','Սկզբնական ամիս',
  'Ստեղծվել է','Թարմացվել է','Նախորդ կոմունալը չի վճարում'
];

const PAYMENT_HEADERS = [
  'Month','Apt','Tenant','Rent Status','Electricity Status',
  'Rent Amount','Electricity Amount','Internet Amount','Total Due',
  'Rent Paid Amount','Electricity Paid Amount','Internet Paid Amount','Total Paid',
  'Rent Paid At','Electricity Paid At','Internet Paid At','Updated At'
];

// â”€â”€ Router â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
function doGet(e) {
  try {
    const action = String((e.parameter.action || '')).trim();
    if (action === 'getTenants')    return json(e, { ok: true, tenants: getTenants_() });
    if (action === 'getElectricity') return json(e, { ok: true, electricity: getElectricity_() });
    if (action === 'getAll')        return json(e, { ok: true, data: getPayments_() });
    if (action === 'getSnapshot')   return json(e, {
      ok: true,
      tenants: getTenants_(),
      electricity: getElectricity_(),
      data: getPayments_(),
      generatedAt: new Date().toISOString()
    });
    if (action === 'update')        return json(e, updatePayment_(e.parameter));
    if (action === 'saveMeter')     return json(e, saveMeter_(e.parameter));
    if (action === 'updateMeterRecord') return json(e, updateMeterRecord_(e.parameter));
    if (action === 'repairMeterHistory') return json(e, repairMeterHistory_(e.parameter));
    if (action === 'deletePayment') return json(e, deletePayment_(e.parameter));
    if (action === 'updateTenant')  return json(e, updateTenant_(e.parameter));
    if (action === 'getTenantFolder') return json(e, getTenantFolder_(e.parameter));
    if (action === 'generateContract') return json(e, idempotentDocumentGeneration_(
      action, e.parameter, () => generateContract_(e.parameter)
    ));
    if (action === 'generateExtension') return json(e, idempotentDocumentGeneration_(
      action, e.parameter, () => generateExtension_(e.parameter)
    ));
    if (action === 'testDocuments') return json(e, testDocumentTemplates_());
    if (action === 'test')          return json(e, runTest_());
    return json(e, { ok: false, error: 'Unknown action: ' + action });
  } catch (err) {
    return json(e, { ok: false, error: String(err && err.message ? err.message : err) });
  }
}

// â”€â”€ JSONP / JSON response â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
function doPost(e) {
  try {
    const action = String((e.parameter.action || '')).trim();
    if (action === 'uploadTenantFile') return json(e, uploadTenantFile_(e.parameter));
    return json(e, { ok: false, error: 'Unknown POST action: ' + action });
  } catch (err) {
    return json(e, { ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function json(e, obj) {
  const callback = e && e.parameter && e.parameter.callback;
  if (callback) {
    return ContentService
      .createTextOutput(callback + '(' + JSON.stringify(obj) + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// â”€â”€ Spreadsheet helpers â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }

function getSheet_(name, createIfMissing) {
  let sh = ss_().getSheetByName(name);
  if (!sh && createIfMissing) sh = ss_().insertSheet(name);
  if (!sh) throw new Error('Sheet not found: ' + name);
  return sh;
}

// Normalise header name for comparison. Keep Armenian letters, remove only
// spacing and punctuation so "Rent Amount" and "RentAmount" still match.
function norm_(value) {
  return String(value || '').toLowerCase().replace(/[\s._\-()\/\\:;,'"Â«Â»ÕÕžÕœÕ›]+/g, '');
}

function headerMap_(headers) {
  const map = {};
  headers.forEach((h, i) => { if (h !== '') map[norm_(h)] = i; });
  return map;
}

// Return first matching value from a list of possible header names
function first_(row, map, names, fallback) {
  for (let i = 0; i < names.length; i++) {
    const idx = map[norm_(names[i])];
    if (idx !== undefined) return row[idx];
  }
  return fallback;
}

// Return first matching value that is not blank. Useful while old and new
// tenant columns both exist in the worksheet.
function firstNonBlank_(row, map, names, fallback) {
  for (let i = 0; i < names.length; i++) {
    const idx = map[norm_(names[i])];
    if (idx === undefined) continue;
    const value = row[idx];
    if (value !== null && value !== undefined && String(value).trim() !== '') return value;
  }
  return fallback;
}

function officeKey_(value) {
  return String(value || '').trim().replace(/\s+/g, '').toUpperCase();
}

function num_(value) {
  if (value === null || value === undefined || value === '') return 0;
  return Number(String(value).replace(/[, Ö\s]/g, '')) || 0;
}

function bool_(value) {
  const v = String(value || '').toLowerCase().trim();
  return value === true || v === 'true' || v === 'yes' || v === 'y' || v === '1' || v === 'internet' || v === 'Õ¡ÕµÕ¸';
}

function cleanStatus_(value) {
  const v = String(value || '').trim();
  return v === 'null' ? '' : v;
}

function dateOrBlank_(value) {
  if (value === null || value === undefined || value === '') return '';
  const d = new Date(value);
  return isNaN(d.getTime()) ? value : d;
}

function getRows_(sheetName) {
  const sh = getSheet_(sheetName, false);
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return { headers: values[0] || [], rows: [] };
  return { headers: values[0], rows: values.slice(1) };
}

// â”€â”€ TENANTS â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// Accepts any column named: Office, Apt, Apartment, Room
// Accepts channel in: Ch, Channel
// Active column is OPTIONAL â€” if missing or empty, tenant is active
function tenantsSheet_() {
  const sh = getSheet_(SHEETS.TENANTS, true);
  if (sh.getLastRow() === 0 || sh.getLastColumn() === 0) {
    sh.getRange(1, 1, 1, TENANT_HEADERS.length).setValues([TENANT_HEADERS]);
    return sh;
  }
  const existing = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const existingNorm = existing.map(norm_);
  const missing = TENANT_HEADERS.filter(h => existingNorm.indexOf(norm_(h)) === -1);
  if (missing.length) {
    sh.getRange(1, existing.length + 1, 1, missing.length).setValues([missing]);
  }
  return sh;
}

function getTenants_() {
  const sh = tenantsSheet_();
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return [];

  const headers = values[0];
  const map = headerMap_(headers);

  return values.slice(1)
    .map(row => {
      // apt: prefer Office (your actual column), fall back to Apt/Apartment/Room
      const apt = officeKey_(firstNonBlank_(row, map,
        ['Գրասենյակ','Գրս','Office','Apt','Apartment','Room','Ô³Ö€Õ¡Õ½Õ¥Õ¶ÕµÕ¡Õ¯','Ô³Ö€Õ½'], ''));
      return { row, apt };
    })
    .filter(({ apt }) => apt !== '' && apt.toLowerCase() !== 'Õ¡Õ¦Õ¡Õ¿')
    .filter(({ row }) => {
      const active = first_(row, map, ['Ակտիվ','Active','Status','Ô±Õ¯Õ¿Õ«Õ¾'], '');
      // if Active column is empty/missing â†’ treat as active
      const v = String(active || '').toLowerCase().trim();
      return v !== 'no';
    })
    .map(({ row, apt }) => ({
      apt,
      name:     String(firstNonBlank_(row, map, ['Ֆիզ. անձի անուն','Իրավ. անձի անվանում','Name','Tenant','Tenant Name','Õ–Õ«Õ¦. Õ¡Õ¶Õ±Õ« Õ¡Õ¶Õ¸Ö‚Õ¶','Ô»Ö€Õ¡Õ¾. Õ¡Õ¶Õ±Õ« Õ¡Õ¶Õ¾Õ¡Õ¶Õ¸Ö‚Õ´'], '')).trim(),
      phone:    String(firstNonBlank_(row, map, ['Հեռախոս','Phone','Tel','Telephone','Õ€Õ¥Õ¼Õ¡Õ­Õ¸Õ½'], '')).trim(),
      ch:       String(firstNonBlank_(row, map, ['Կապ','Channel','Ch','Contact','Ô¿Õ¡Õº'], '')).trim() || 'wa',
      day:      num_(firstNonBlank_(row, map, ['Վճարման օր','Day','Due Day','Payment Day','ÕŽÕ³Õ¡Ö€Õ´Õ¡Õ¶ Ö…Ö€'], 0)),
      rent:     num_(firstNonBlank_(row, map, ['Վարձ','Rent','Rent Amount','ÕŽÕ¡Ö€Õ±'], 0)),
      internet: bool_(firstNonBlank_(row, map, ['Ինտերնետ','Internet','Net','Ô»Õ¶Õ¿Õ¥Ö€Õ¶Õ¥Õ¿'], false)),
      skipPreviousUtilities: bool_(firstNonBlank_(row, map, ['Նախորդ կոմունալը չի վճարում','Skip Previous Utilities','Õ†Õ¡Õ­Õ¸Ö€Õ¤ Õ¯Õ¸Õ´Õ¸Ö‚Õ¶Õ¡Õ¬Õ¨ Õ¹Õ« Õ¾Õ³Õ¡Ö€Õ¸Ö‚Õ´'], false)),
      startMonth: String(firstNonBlank_(row, map, ['Սկզբնական ամիս','Start Month','Move In Month','First Month','ÕÕ¯Õ¦Õ¢Õ¶Õ¡Õ¯Õ¡Õ¶ Õ¡Õ´Õ«Õ½'], '')).trim(),
      entityType: String(firstNonBlank_(row, map, ['Տեսակ','Entity Type','ÕÕ¥Õ½Õ¡Õ¯'], '')).trim(),
      physicalName: String(firstNonBlank_(row, map, ['Ֆիզ. անձի անուն','Physical Name','Õ–Õ«Õ¦. Õ¡Õ¶Õ±Õ« Õ¡Õ¶Õ¸Ö‚Õ¶'], '')).trim(),
      physicalPassport: String(firstNonBlank_(row, map, ['Ֆիզ. անձի անձնագիր','Physical Passport','Õ–Õ«Õ¦. Õ¡Õ¶Õ±Õ« Õ¡Õ¶Õ±Õ¶Õ¡Õ£Õ«Ö€'], '')).trim(),
      physicalAddress: String(firstNonBlank_(row, map, ['Ֆիզ. անձի հասցե','Physical Address','Õ–Õ«Õ¦. Õ¡Õ¶Õ±Õ« Õ°Õ¡Õ½ÖÕ¥'], '')).trim(),
      physicalIssuedFrom: String(firstNonBlank_(row, map, ['Ֆիզ. անձի տրված է','Physical Issued From','Õ–Õ«Õ¦. Õ¡Õ¶Õ±Õ« Õ¿Ö€Õ¾Õ¡Õ® Õ§'], '')).trim(),
      physicalValidUntil: firstNonBlank_(row, map, ['Ֆիզ. անձի վավեր է մինչև','Physical Valid Until','Õ–Õ«Õ¦. Õ¡Õ¶Õ±Õ« Õ¾Õ¡Õ¾Õ¥Ö€ Õ§ Õ´Õ«Õ¶Õ¹Ö‡'], ''),
      legalName: String(firstNonBlank_(row, map, ['Իրավ. անձի անվանում','Legal Name','Ô»Ö€Õ¡Õ¾. Õ¡Õ¶Õ±Õ« Õ¡Õ¶Õ¾Õ¡Õ¶Õ¸Ö‚Õ´'], '')).trim(),
      hvhh: String(firstNonBlank_(row, map, ['ՀՎՀՀ','HVHH','Tax ID','Õ€ÕŽÕ€Õ€'], '')).trim(),
      legalAddress: String(firstNonBlank_(row, map, ['Իրավ. հասցե','Legal Address','Ô»Ö€Õ¡Õ¾. Õ°Õ¡Õ½ÖÕ¥'], '')).trim(),
      registrationNumber: String(firstNonBlank_(row, map, ['Գրանցման համար','Registration Number','Ô³Ö€Õ¡Õ¶ÖÕ´Õ¡Õ¶ Õ°Õ¡Õ´Õ¡Ö€'], '')).trim(),
      ceoName: String(firstNonBlank_(row, map, ['Տնօրենի անուն','CEO Name','ÕÕ¶Ö…Ö€Õ¥Õ¶Õ« Õ¡Õ¶Õ¸Ö‚Õ¶'], '')).trim(),
      ceoPassport: String(firstNonBlank_(row, map, ['Տնօրենի անձնագիր','CEO Passport','ÕÕ¶Ö…Ö€Õ¥Õ¶Õ« Õ¡Õ¶Õ±Õ¶Õ¡Õ£Õ«Ö€'], '')).trim(),
      ceoAddress: String(firstNonBlank_(row, map, ['Տնօրենի հասցե','CEO Address','ÕÕ¶Ö…Ö€Õ¥Õ¶Õ« Õ°Õ¡Õ½ÖÕ¥'], '')).trim(),
      ceoIssuedFrom: String(firstNonBlank_(row, map, ['Տնօրենի տրված է','CEO Issued From','ÕÕ¶Ö…Ö€Õ¥Õ¶Õ« Õ¿Ö€Õ¾Õ¡Õ® Õ§'], '')).trim(),
      ceoValidUntil: firstNonBlank_(row, map, ['Տնօրենի վավեր է մինչև','CEO Valid Until','ÕÕ¶Ö…Ö€Õ¥Õ¶Õ« Õ¾Õ¡Õ¾Õ¥Ö€ Õ§ Õ´Õ«Õ¶Õ¹Ö‡'], ''),
      startDate: firstNonBlank_(row, map, ['Սկիզբ','Start Date','ÕÕ¯Õ«Õ¦Õ¢'], ''),
      endDate: firstNonBlank_(row, map, ['Ավարտ','End Date','Ô±Õ¾Õ¡Ö€Õ¿'], ''),
      paymentDate: firstNonBlank_(row, map, ['Վճարման օր','Payment Date','ÕŽÕ³Õ¡Ö€Õ´Õ¡Õ¶ Ö…Ö€'], ''),
      createdAt: firstNonBlank_(row, map, ['Ստեղծվել է','Created At','ÕÕ¿Õ¥Õ²Õ®Õ¾Õ¥Õ¬ Õ§'], '')
    }));
}


// â”€â”€ METER DATA SHEET â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// Exact column order stored in Google Sheet:
// Office | Serial | TotalPrevious | T1prev | T2prev | T1current | T2current |
// KwhDay | KwhNight | AMDDay | AMDNight | AMDTotal | Month
const METER_HEADERS = [
  'Office','Serial','TotalPrevious','T1prev','T2prev','T1current','T2current',
  'KwhDay','KwhNight','AMDDay','AMDNight','AMDTotal','Month',
  'PreviousMonth','PreviousAssumed','EnteredAt'
];

function meterSheet_() {
  const sh = getSheet_(SHEETS.ELECTRICITY, true);
  if (sh.getLastRow() === 0 || sh.getLastColumn() === 0) {
    sh.getRange(1, 1, 1, METER_HEADERS.length).setValues([METER_HEADERS]);
    const hr = sh.getRange(1, 1, 1, METER_HEADERS.length);
    hr.setBackground('#1a3a1a'); hr.setFontColor('#ffffff'); hr.setFontWeight('bold');
    sh.setFrozenRows(1);
    return sh;
  }
  // Add any missing headers without disturbing existing data
  const existing = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const existingNorm = existing.map(norm_);
  const missing = METER_HEADERS.filter(h => existingNorm.indexOf(norm_(h)) === -1);
  if (missing.length) {
    sh.getRange(1, existing.length + 1, 1, missing.length).setValues([missing]);
  }
  const values = sh.getDataRange().getValues();
  const headers = values[0] || [];
  const orderedHeaders = METER_HEADERS.slice();
  const orderedNorm = orderedHeaders.map(norm_);
  headers.forEach(h => {
    if (h !== '' && orderedNorm.indexOf(norm_(h)) === -1) orderedHeaders.push(h);
  });
  const sameOrder = orderedHeaders.length === headers.length &&
    orderedHeaders.every((h, i) => norm_(h) === norm_(headers[i]));
  if (!sameOrder) {
    const oldMap = headerMap_(headers);
    const reordered = values.map((row, rowIndex) => {
      if (rowIndex === 0) return orderedHeaders;
      return orderedHeaders.map(h => {
        const idx = oldMap[norm_(h)];
        return idx === undefined ? '' : row[idx];
      });
    });
    sh.clearContents();
    sh.getRange(1, 1, reordered.length, orderedHeaders.length).setValues(reordered);
  }
  return sh;
}

// â”€â”€ GET ELECTRICITY â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// Returns map keyed "Month_OFFICE" with full meter data incl. AMD amounts.
function getElectricity_() {
  const sh = meterSheet_();
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return {};
  const headers = values[0];
  const map = headerMap_(headers);
  const result = {};
  values.slice(1).forEach(row => {
    const apt = officeKey_(firstNonBlank_(row, map,
      ['Office','Ô³Ñ€Ð°ÑÐµÐ½rak','Apt','Apartment','Room'], ''));
    if (!apt) return;
    const month    = String(first_(row, map, ['Month','Ô±Õ´is'], '')).trim();
    if (!month) return;
    const amdDay   = num_(first_(row, map, ['AMDDay','Amd Day','AMD Day'], 0));
    const amdNight = num_(first_(row, map, ['AMDNight','Amd Night','AMD Night'], 0));
    const kwhDay   = num_(first_(row, map, ['KwhDay','Kwh Day'], 0));
    const kwhNight = num_(first_(row, map, ['KwhNight','Kwh Night'], 0));
    const t1prev   = num_(first_(row, map, ['T1prev','T1 prev'], 0));
    const t2prev   = num_(first_(row, map, ['T2prev','T2 prev'], 0));
    const totalPrevious = num_(first_(row, map, ['TotalPrevious','Total Previous'], t1prev + t2prev));
    const t1cur    = num_(first_(row, map, ['T1current','T1 current'], 0));
    const t2cur    = num_(first_(row, map, ['T2current','T2 current'], 0));
    const amdTotal = num_(first_(row, map, ['AMDTotal','AMD Total'], amdDay + amdNight));
    const serial   = String(first_(row, map, ['Serial'], '') || '').trim();
    const internet = bool_(first_(row, map, ['Internet','Net','Ô»Õ¶Õ¿ÐµÑ€Ð½ÐµÑ‚'], false));
    const previousMonth = String(first_(row, map, ['PreviousMonth','Previous Month'], '') || '').trim();
    const previousAssumed = bool_(first_(row, map, ['PreviousAssumed','Previous Assumed'], false));
    result[month + '_' + apt] = {
      apt, month, serial,
      totalPrevious, t1prev, t2prev, t1current: t1cur, t2current: t2cur,
      kwhDay, kwhNight, amdDay, amdNight, amdTotal, internet,
      previousMonth, previousAssumed
    };
  });
  return result;
}

// â”€â”€ SAVE METER READING â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
// Frontend sends only: apt, month, t1current, t2current, internet.
// Previous readings and serial are derived from the latest saved row before
// the selected month. This keeps meter entry to two inputs.
// Backend computes: KwhDay = T1cur-T1prev, KwhNight = T2cur-T2prev,
//                  AMDDay = KwhDay * AMD_PER_KWH_DAY (rounded),
//                  AMDNight = KwhNight * AMD_PER_KWH_NIGHT (rounded)
// Stores ONE row per office per month (upsert).
function saveMeter_(p) {
  const sh = meterSheet_();
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const map = headerMap_(headers);

  const month    = String(p.month || '').trim();
  const apt      = officeKey_(p.apt);
  if (!month || !apt) throw new Error('Missing month or apt');

  if (String(p.t1current || '').trim() === '' || String(p.t2current || '').trim() === '') {
    throw new Error('T1 and T2 current readings are required');
  }

  const t1cur = num_(p.t1current);
  const t2cur = num_(p.t2current);

  // Find the existing month row before deriving the baseline. When the first
  // stored month has no earlier row, preserve its imported serial and prior
  // readings instead of replacing them with blanks and zeroes.
  let rowNumber = 0;
  if (sh.getLastRow() >= 2) {
    const vals = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
    for (let i = 0; i < vals.length; i++) {
      const rm = String(first_(vals[i], map, ['Month'], '')).trim();
      const ra = officeKey_(firstNonBlank_(vals[i], map,
        ['Office','Apt','Apartment','Room'], ''));
      if (rm === month && ra === apt) { rowNumber = i + 2; break; }
    }
  }

  const numCols = sh.getLastColumn();
  const rowData = rowNumber
    ? sh.getRange(rowNumber, 1, 1, numCols).getValues()[0].slice()
    : new Array(numCols).fill('');
  const previous = findPreviousMeter_(sh, map, apt, month);
  const existingSerial = String(first_(rowData, map, ['Serial'], '') || '').trim();
  const serial = previous.serial || existingSerial || String(p.serial || '').trim();
  const t1prev = previous.found
    ? previous.t1current
    : num_(first_(rowData, map, ['T1prev','T1 prev'], 0));
  const t2prev = previous.found
    ? previous.t2current
    : num_(first_(rowData, map, ['T2prev','T2 prev'], 0));
  const previousMonth = previous.found
    ? previous.month
    : String(first_(rowData, map, ['PreviousMonth','Previous Month'], '') || '').trim();
  const previousAssumed = previous.found
    ? previous.assumed
    : bool_(first_(rowData, map, ['PreviousAssumed','Previous Assumed'], true));

  if (previous.found && (t1cur < t1prev || t2cur < t2prev)) {
    throw new Error(
      'Current reading cannot be below previous reading (' +
      t1prev + ' / ' + t2prev + ')'
    );
  }

  const kwhDay   = Math.max(0, t1cur - t1prev);
  const kwhNight = Math.max(0, t2cur - t2prev);
  const amdDay   = Math.round(kwhDay   * AMD_PER_KWH_DAY);
  const amdNight = Math.round(kwhNight * AMD_PER_KWH_NIGHT);
  const totalPrevious = t1prev + t2prev;
  const amdTotal = amdDay + amdNight;

  function setCol(name, val) {
    const idx = map[norm_(name)];
    if (idx !== undefined) rowData[idx] = val;
  }
  setCol('Office',    apt);
  setCol('Serial',    serial);
  setCol('TotalPrevious', totalPrevious);
  setCol('T1prev',    t1prev);
  setCol('T2prev',    t2prev);
  setCol('T1current', t1cur);
  setCol('T2current', t2cur);
  setCol('KwhDay',    kwhDay);
  setCol('KwhNight',  kwhNight);
  setCol('AMDDay',    amdDay);
  setCol('AMDNight',  amdNight);
  setCol('AMDTotal',  amdTotal);
  setCol('Month',     month);
  setCol('PreviousMonth', previousMonth);
  setCol('PreviousAssumed', previousAssumed ? 'yes' : 'no');
  setCol('EnteredAt', new Date());

  if (rowNumber) {
    sh.getRange(rowNumber, 1, 1, numCols).setValues([rowData]);
  } else {
    sh.getRange(sh.getLastRow() + 1, 1, 1, numCols).setValues([rowData]);
  }

  return { ok: true, apt, month, serial, t1prev, t2prev,
           t1current: t1cur, t2current: t2cur,
           totalPrevious, kwhDay, kwhNight, amdDay, amdNight, amdTotal,
           previousMonth, previousAssumed };
}

// Full utility-card edit. Unlike saveMeter_, this intentionally accepts
// corrected previous readings and serial numbers for the selected month.
function updateMeterRecord_(p) {
  const sh = meterSheet_();
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const map = headerMap_(headers);
  const month = String(p.month || '').trim();
  const apt = officeKey_(p.apt);
  if (!month || !apt) throw new Error('Missing month or apt');

  ['t1prev', 't2prev', 't1current', 't2current'].forEach(key => {
    if (String(p[key] === undefined ? '' : p[key]).trim() === '') {
      throw new Error('All previous and current readings are required');
    }
  });

  const serial = String(p.serial || '').trim();
  const t1prev = num_(p.t1prev);
  const t2prev = num_(p.t2prev);
  const t1cur = num_(p.t1current);
  const t2cur = num_(p.t2current);
  const internet = bool_(p.internet);
  if (t1cur < t1prev || t2cur < t2prev) {
    throw new Error('Current readings cannot be below previous readings');
  }

  let rowNumber = 0;
  if (sh.getLastRow() >= 2) {
    const values = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
    for (let i = 0; i < values.length; i++) {
      const rowMonth = String(first_(values[i], map, ['Month'], '')).trim();
      const rowApt = officeKey_(firstNonBlank_(values[i], map,
        ['Office','Apt','Apartment','Room'], ''));
      if (rowMonth === month && rowApt === apt) { rowNumber = i + 2; break; }
    }
  }

  const numCols = sh.getLastColumn();
  const rowData = rowNumber
    ? sh.getRange(rowNumber, 1, 1, numCols).getValues()[0].slice()
    : new Array(numCols).fill('');
  const kwhDay = t1cur - t1prev;
  const kwhNight = t2cur - t2prev;
  const amdDay = Math.round(kwhDay * AMD_PER_KWH_DAY);
  const amdNight = Math.round(kwhNight * AMD_PER_KWH_NIGHT);
  const amdTotal = amdDay + amdNight;

  function setCol(name, value) {
    const idx = map[norm_(name)];
    if (idx !== undefined) rowData[idx] = value;
  }
  setCol('Office', apt);
  setCol('Serial', serial);
  setCol('TotalPrevious', t1prev + t2prev);
  setCol('T1prev', t1prev);
  setCol('T2prev', t2prev);
  setCol('T1current', t1cur);
  setCol('T2current', t2cur);
  setCol('KwhDay', kwhDay);
  setCol('KwhNight', kwhNight);
  setCol('AMDDay', amdDay);
  setCol('AMDNight', amdNight);
  setCol('AMDTotal', amdTotal);
  setCol('Internet', internet ? 'yes' : 'no');
  setCol('Month', month);
  setCol('PreviousMonth', String(p.previousMonth || '').trim());
  setCol('PreviousAssumed', bool_(p.previousAssumed) ? 'yes' : 'no');
  setCol('EnteredAt', new Date());

  if (rowNumber) sh.getRange(rowNumber, 1, 1, numCols).setValues([rowData]);
  else sh.getRange(sh.getLastRow() + 1, 1, 1, numCols).setValues([rowData]);

  // Keep the tenant-level Internet flag synchronized with the utility card.
  const tenantSheet = tenantsSheet_();
  const tenantValues = tenantSheet.getDataRange().getValues();
  const tenantMap = headerMap_(tenantValues[0]);
  for (let i = 1; i < tenantValues.length; i++) {
    const tenantApt = officeKey_(firstNonBlank_(tenantValues[i], tenantMap,
      ['Գրասենյակ','Գրս','Office','Apt','Apartment','Room','Ô³Ö€Õ¡Õ½Õ¥Õ¶ÕµÕ¡Õ¯','Ô³Ö€Õ½'], ''));
    if (tenantApt === apt) {
      setTenantCell_(tenantSheet, i + 1, tenantMap,
        ['Ինտերնետ','Internet','Net','Ô»Õ¶Õ¿Õ¥Ö€Õ¶Õ¥Õ¿'], internet ? 'yes' : 'no');
      break;
    }
  }

  return {
    ok: true, apt, month, serial, t1prev, t2prev,
    t1current: t1cur, t2current: t2cur, totalPrevious: t1prev + t2prev,
    kwhDay, kwhNight, amdDay, amdNight, amdTotal, internet
  };
}

function meterMonthIndex_(month) {
  const names = [
    'Հունվար','Փետրվար','Մարտ','Ապրիլ','Մայիս','Հունիս',
    'Հուլիս','Օգոստոս','Սեպտեմբեր','Հոկտեմբեր','Նոյեմբեր','Դեկտեմբեր'
  ];
  const parts = String(month || '').trim().split('-');
  const monthIndex = names.indexOf(parts[0]);
  const year = Number(parts[1]);
  return monthIndex < 0 || !year ? null : year * 12 + monthIndex;
}

function findPreviousMeter_(sh, map, apt, targetMonth) {
  const targetIndex = meterMonthIndex_(targetMonth);
  const result = {
    found: false, month: '', serial: '', t1current: 0, t2current: 0, assumed: true
  };
  if (targetIndex === null || sh.getLastRow() < 2) return result;

  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  let bestIndex = -Infinity;
  rows.forEach(row => {
    const rowApt = officeKey_(firstNonBlank_(row, map,
      ['Office','Apt','Apartment','Room'], ''));
    if (rowApt !== apt) return;
    const rowMonth = String(first_(row, map, ['Month'], '')).trim();
    const rowIndex = meterMonthIndex_(rowMonth);
    if (rowIndex === null || rowIndex >= targetIndex || rowIndex <= bestIndex) return;
    bestIndex = rowIndex;
    result.found = true;
    result.month = rowMonth;
    result.serial = String(first_(row, map, ['Serial'], '') || '').trim();
    result.t1current = num_(first_(row, map, ['T1current','T1 current'], 0));
    result.t2current = num_(first_(row, map, ['T2current','T2 current'], 0));
  });

  result.assumed = !result.found || bestIndex !== targetIndex - 1;
  return result;
}

// â”€â”€ PAYMENTS â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
function repairMeterHistory_(p) {
  if (String(p.key || '') !== 'plusone-april-may-2026') {
    throw new Error('Invalid repair key');
  }

  // Office, Serial, March T1/T2, April T1/T2, May T1/T2.
  const source = [
    [201,1085286,1061,10,1214,16,1261,19],[301,1084580,4070,410,4111,411,4117,419],
    [302,1086975,2397,171,2401,171,2403,171],[303,1086331,1031,2,1053,2,1055,2],
    [304,1086325,3872,2872,3903,2954,3926,2981],[305,1086415,3702,343,3702,343,3702,343],
    [306,1086322,1303,27,1303,28,1304,28],[502,1086923,4226,351,4514,381,4614,406],
    [503,1087994,4543,739,4543,739,4550,740],[505,1085654,1138,727,1162,766,1175,770],
    [803,1086959,3331,765,3331,765,3340,765],[804,1086312,3419,1067,3419,1067,3443,1071],
    [901,1087350,724,8,741,8,745,8],[1001,1081001,1304,401,1342,414,1365,422],
    [1101,1084505,107,24,123,24,129,24],[1102,1084381,700,71,704,72,704,73],
    [1103,1085680,789,418,831,438,853,446],[1104,1084954,275,2,317,7,332,9],
    [1105,1086181,31,0,63,0,64,0],[1106,1085779,131,5,174,8,187,12],
    [1107,1086213,274,73,311,74,323,74],[1108,1086517,712,30,815,33,838,34],
    [1109,1085226,11,1,12,1,15,2],[1110,1086855,529,20,604,19,621,19],
    [1111,1086161,0,0,8,2,10,3],[1112,1085910,534,100,603,130,610,133],
    [1113,1087114,0,0,183,0,189,1],[1114,1085214,0,0,54,7,75,12],
    [1202,1086488,0,0,0,0,10,7],[1203,1085967,0,0,52,48,66,65],
    [1204,1086198,604,355,641,393,658,405],[1205,1085232,319,224,447,303,466,319],
    [1301,1085468,2139,168,2250,169,2318,175],[1302,1085666,2113,423,2279,478,2325,507],
    [1601,1086191,2458,754,2797,919,2893,988],[1602,1086372,371,31,385,33,388,33],
    [1603,1084133,3694,292,3772,296,3793,298]
  ];

  const sh = meterSheet_();
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const map = headerMap_(headers);
  const existing = sh.getLastRow() < 2 ? [] :
    sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
  const keep = existing.filter(row => {
    const month = String(first_(row, map, ['Month'], '')).trim();
    return month !== 'Ապրիլ-2026' && month !== 'Մայիս-2026';
  });

  function makeRow(item, month, previousMonth, prevT1, prevT2, curT1, curT2, assumed) {
    const row = new Array(headers.length).fill('');
    const set = (name, value) => {
      const index = map[norm_(name)];
      if (index !== undefined) row[index] = value;
    };
    const kwhDay = curT1 - prevT1;
    const kwhNight = curT2 - prevT2;
    const amdDay = Math.round(kwhDay * AMD_PER_KWH_DAY);
    const amdNight = Math.round(kwhNight * AMD_PER_KWH_NIGHT);
    set('Office', item[0]); set('Serial', item[1]);
    set('TotalPrevious', prevT1 + prevT2);
    set('T1prev', prevT1); set('T2prev', prevT2);
    set('T1current', curT1); set('T2current', curT2);
    set('KwhDay', kwhDay); set('KwhNight', kwhNight);
    set('AMDDay', amdDay); set('AMDNight', amdNight);
    set('AMDTotal', amdDay + amdNight); set('Month', month);
    set('PreviousMonth', previousMonth);
    set('PreviousAssumed', assumed ? 'yes' : 'no');
    set('EnteredAt', new Date());
    return row;
  }

  const repaired = [];
  source.forEach(item => {
    repaired.push(makeRow(item, 'Ապրիլ-2026', 'Մարտ-2026',
      item[2], item[3], item[4], item[5], true));
    repaired.push(makeRow(item, 'Մայիս-2026', 'Ապրիլ-2026',
      item[4], item[5], item[6], item[7], false));
  });

  if (existing.length) {
    sh.getRange(2, 1, existing.length, headers.length).clearContent();
  }
  const output = keep.concat(repaired);
  if (output.length) sh.getRange(2, 1, output.length, headers.length).setValues(output);
  return { ok: true, repaired: repaired.length, april: source.length, may: source.length };
}

function paymentsSheet_() {
  const sh = getSheet_(SHEETS.PAYMENTS, true);
  // Auto-add any missing headers
  if (sh.getLastRow() === 0 || sh.getLastColumn() === 0) {
    sh.getRange(1, 1, 1, PAYMENT_HEADERS.length).setValues([PAYMENT_HEADERS]);
    return sh;
  }
  const existing = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const existingNorm = existing.map(norm_);
  const missing = PAYMENT_HEADERS.filter(h => existingNorm.indexOf(norm_(h)) === -1);
  if (missing.length) {
    sh.getRange(1, existing.length + 1, 1, missing.length).setValues([missing]);
  }
  return sh;
}

function getPayments_() {
  const sh = paymentsSheet_();
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return {};

  const headers = values[0];
  const map = headerMap_(headers);
  const result = {};

  values.slice(1).forEach(row => {
    const month = String(first_(row, map, ['Month'], '')).trim();
    const apt   = officeKey_(firstNonBlank_(row, map, ['Apt','Ô³Ö€Õ¡Õ½Õ¥Õ¶ÕµÕ¡Õ¯','Office','Apartment','Room','Ô³Ö€Õ½'], ''));
    if (!month || !apt) return;

    result[month + '_' + apt] = {
      month, apt,
      tenant:                String(first_(row, map, ['Tenant'], '')).trim(),
      rentStatus:            cleanStatus_(first_(row, map, ['Rent Status','rentStatus'], '')),
      elecStatus:            cleanStatus_(first_(row, map, ['Electricity Status','Elec Status','elecStatus'], '')),
      rentAmount:            num_(first_(row, map, ['Rent Amount'], 0)),
      electricityAmount:     num_(first_(row, map, ['Electricity Amount'], 0)),
      internetAmount:        num_(first_(row, map, ['Internet Amount'], 0)),
      totalDue:              num_(first_(row, map, ['Total Due'], 0)),
      rentPaidAmount:        num_(first_(row, map, ['Rent Paid Amount'], 0)),
      electricityPaidAmount: num_(first_(row, map, ['Electricity Paid Amount'], 0)),
      internetPaidAmount:    num_(first_(row, map, ['Internet Paid Amount'], 0)),
      totalPaid:             num_(first_(row, map, ['Total Paid'], 0)),
      rentPaidAt:            first_(row, map, ['Rent Paid At'], ''),
      electricityPaidAt:     first_(row, map, ['Electricity Paid At'], ''),
      internetPaidAt:        first_(row, map, ['Internet Paid At'], '')
    };
  });

  return result;
}

// â”€â”€ UPDATE PAYMENT â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
function updatePayment_(p) {
  const sh = paymentsSheet_();
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const map = headerMap_(headers);

  const month  = String(p.month || '').trim();
  const apt    = officeKey_(p.apt);
  const type   = String(p.type  || '').trim();
  const status = cleanStatus_(p.status);
  if (!month || !apt) throw new Error('Missing month or apt');

  // Clean up any duplicate rows for this month+apt before writing
  dedupPaymentRows_(sh, map, month, apt);

  let rowNumber = findPaymentRow_(sh, map, month, apt);
  if (!rowNumber) {
    rowNumber = sh.getLastRow() + 1;
    setCell_(sh, rowNumber, map, 'Month', month);
    setCell_(sh, rowNumber, map, 'Apt',   apt);
  }

  setCell_(sh, rowNumber, map, 'Tenant', p.tenant || '');
  if (type === 'rent') setCell_(sh, rowNumber, map, 'Rent Status',        status);
  if (type === 'elec') setCell_(sh, rowNumber, map, 'Electricity Status', status);

  // Only overwrite the fields belonging to this type â€” never clobber the other side
  if (type === 'rent') {
    setCell_(sh, rowNumber, map, 'Rent Amount',      num_(p.rentAmount));
    setCell_(sh, rowNumber, map, 'Rent Paid Amount', num_(p.rentPaidAmount));
    // Recalculate Total Due / Total Paid from existing elec + new rent values
    const existingElecAmt  = getCellVal_(sh, rowNumber, map, 'Electricity Amount');
    const existingNetAmt   = getCellVal_(sh, rowNumber, map, 'Internet Amount');
    const existingElecPaid = getCellVal_(sh, rowNumber, map, 'Electricity Paid Amount');
    const existingNetPaid  = getCellVal_(sh, rowNumber, map, 'Internet Paid Amount');
    setCell_(sh, rowNumber, map, 'Total Due',  num_(p.rentAmount) + num_(existingElecAmt) + num_(existingNetAmt));
    setCell_(sh, rowNumber, map, 'Total Paid', num_(p.rentPaidAmount) + num_(existingElecPaid) + num_(existingNetPaid));
  } else if (type === 'elec') {
    setCell_(sh, rowNumber, map, 'Electricity Amount',      num_(p.electricityAmount));
    setCell_(sh, rowNumber, map, 'Internet Amount',         num_(p.internetAmount));
    setCell_(sh, rowNumber, map, 'Electricity Paid Amount', num_(p.electricityPaidAmount));
    setCell_(sh, rowNumber, map, 'Internet Paid Amount',    num_(p.internetPaidAmount));
    // Recalculate totals from existing rent + new elec values
    const existingRentAmt  = getCellVal_(sh, rowNumber, map, 'Rent Amount');
    const existingRentPaid = getCellVal_(sh, rowNumber, map, 'Rent Paid Amount');
    setCell_(sh, rowNumber, map, 'Total Due',  num_(existingRentAmt) + num_(p.electricityAmount) + num_(p.internetAmount));
    setCell_(sh, rowNumber, map, 'Total Paid', num_(existingRentPaid) + num_(p.electricityPaidAmount) + num_(p.internetPaidAmount));
  }

  // â”€â”€ Rent paid-at â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
  if (type === 'rent') {
    if (isRentPaidStatus_(status)) {
      // Use the timestamp the client recorded (optimistic UI moment); fall back to now
      const paidAt = dateOrBlank_(p.rentPaidAt) || new Date();
      setCell_(sh, rowNumber, map, 'Rent Paid At', paidAt);
    } else if (p.clearRentPaidAt === '1' || !isRentPaidStatus_(status)) {
      setCell_(sh, rowNumber, map, 'Rent Paid At', '');
    }
  }

  // â”€â”€ Electricity & Internet paid-at â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
  if (type === 'elec') {
    if (status === 'paid') {
      // Use the timestamp the client recorded; fall back to now
      const paidAt = dateOrBlank_(p.electricityPaidAt) || new Date();
      setCell_(sh, rowNumber, map, 'Electricity Paid At', paidAt);
      // Internet is always paid together with electricity
      if (num_(p.internetAmount) > 0) {
        setCell_(sh, rowNumber, map, 'Internet Paid At', paidAt);
      }
    } else if (p.clearElecPaidAt === '1' || !status || status === 'unpaid') {
      setCell_(sh, rowNumber, map, 'Electricity Paid At', '');
      setCell_(sh, rowNumber, map, 'Internet Paid At',    '');
    }
  }

  setCell_(sh, rowNumber, map, 'Updated At', new Date());
  return { ok: true };
}

// â”€â”€ DELETE PAYMENT â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
function deletePayment_(p) {
  const sh = paymentsSheet_();
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const map = headerMap_(headers);

  const month = String(p.month || '').trim();
  const apt   = officeKey_(p.apt);
  if (!month || !apt) throw new Error('Missing month or apt');

  const rowNumber = findPaymentRow_(sh, map, month, apt);
  if (!rowNumber) return { ok: true, deleted: false };

  sh.deleteRow(rowNumber);
  return { ok: true, deleted: true };
}

// â”€â”€ UPDATE TENANT â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
function updateTenant_(p) {
  const sh = tenantsSheet_();
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const map = headerMap_(headers);
  const apt = officeKey_(p.apt);
  if (!apt) throw new Error('Missing apt');

  let rowNumber = 0;
  for (let i = 1; i < values.length; i++) {
    const rowApt = officeKey_(firstNonBlank_(values[i], map,
      ['Գրասենյակ','Գրս','Office','Apt','Apartment','Room','Ô³Ö€Õ¡Õ½Õ¥Õ¶ÕµÕ¡Õ¯','Ô³Ö€Õ½'], ''));
    if (rowApt === apt) { rowNumber = i + 1; break; }
  }

  const isNew = !rowNumber;
  if (!rowNumber) rowNumber = sh.getLastRow() + 1;

  // Write to whichever apt column exists (Office or Apt)
  setTenantCell_(sh, rowNumber, map, ['Գրասենյակ','Գրս','Office','Apt','Apartment','Room','Ô³Ö€Õ¡Õ½Õ¥Õ¶ÕµÕ¡Õ¯','Ô³Ö€Õ½'], apt);
  setTenantCell_(sh, rowNumber, map, ['Հեռախոս','Phone','Tel','Telephone','Õ€Õ¥Õ¼Õ¡Õ­Õ¸Õ½'],      p.phone || '');
  setTenantCell_(sh, rowNumber, map, ['Կապ','Channel','Ch','Contact','Ô¿Õ¡Õº'],           p.ch    || 'wa');
  setTenantCell_(sh, rowNumber, map, ['Վճարման օր','Day','Due Day','Payment Day','ÕŽÕ³Õ¡Ö€Õ´Õ¡Õ¶ Ö…Ö€'],       num_(p.day));
  setTenantCell_(sh, rowNumber, map, ['Վարձ','Rent','Rent Amount','ÕŽÕ¡Ö€Õ±'],              num_(p.rent));
  setTenantCell_(sh, rowNumber, map, ['Ինտերնետ','Internet','Net','Ô»Õ¶Õ¿Õ¥Ö€Õ¶Õ¥Õ¿'],              bool_(p.internet) ? 'yes' : 'no');
  setTenantCell_(sh, rowNumber, map, ['Ակտիվ','Active','Status','Ô±Õ¯Õ¿Õ«Õ¾'],                         String(p.active || 'yes').toLowerCase() === 'no' ? 'no' : 'yes');
  setTenantCell_(sh, rowNumber, map, ['Նախորդ կոմունալը չի վճարում','Skip Previous Utilities','Õ†Õ¡Õ­Õ¸Ö€Õ¤ Õ¯Õ¸Õ´Õ¸Ö‚Õ¶Õ¡Õ¬Õ¨ Õ¹Õ« Õ¾Õ³Õ¡Ö€Õ¸Ö‚Õ´'], bool_(p.skipPreviousUtilities) ? 'yes' : 'no');
  if (isNew || p.startMonth) setTenantCell_(sh, rowNumber, map, ['Սկզբնական ամիս','Start Month','Move In Month','First Month','ÕÕ¯Õ¦Õ¢Õ¶Õ¡Õ¯Õ¡Õ¶ Õ¡Õ´Õ«Õ½'], p.startMonth || '');
  const entityType = p.entityType || 'physical';
  const displayName = p.name || p.physicalName || p.legalName || '';
  setTenantCell_(sh, rowNumber, map, ['Տեսակ','Entity Type','ÕÕ¥Õ½Õ¡Õ¯'], entityType);
  setTenantCell_(sh, rowNumber, map, ['Ֆիզ. անձի անուն','Physical Name','Õ–Õ«Õ¦. Õ¡Õ¶Õ±Õ« Õ¡Õ¶Õ¸Ö‚Õ¶'], entityType === 'physical' ? (p.physicalName || displayName) : (p.physicalName || ''));
  setTenantCell_(sh, rowNumber, map, ['Ֆիզ. անձի անձնագիր','Physical Passport','Õ–Õ«Õ¦. Õ¡Õ¶Õ±Õ« Õ¡Õ¶Õ±Õ¶Õ¡Õ£Õ«Ö€'], p.physicalPassport || '');
  setTenantCell_(sh, rowNumber, map, ['Ֆիզ. անձի հասցե','Physical Address','Õ–Õ«Õ¦. Õ¡Õ¶Õ±Õ« Õ°Õ¡Õ½ÖÕ¥'], p.physicalAddress || '');
  setTenantCell_(sh, rowNumber, map, ['Ֆիզ. անձի տրված է','Physical Issued From','Õ–Õ«Õ¦. Õ¡Õ¶Õ±Õ« Õ¿Ö€Õ¾Õ¡Õ® Õ§'], p.physicalIssuedFrom || '');
  setTenantCell_(sh, rowNumber, map, ['Ֆիզ. անձի վավեր է մինչև','Physical Valid Until','Õ–Õ«Õ¦. Õ¡Õ¶Õ±Õ« Õ¾Õ¡Õ¾Õ¥Ö€ Õ§ Õ´Õ«Õ¶Õ¹Ö‡'], dateOrBlank_(p.physicalValidUntil));
  setTenantCell_(sh, rowNumber, map, ['Իրավ. անձի անվանում','Legal Name','Ô»Ö€Õ¡Õ¾. Õ¡Õ¶Õ±Õ« Õ¡Õ¶Õ¾Õ¡Õ¶Õ¸Ö‚Õ´'], entityType === 'legal' ? (p.legalName || displayName) : (p.legalName || ''));
  setTenantCell_(sh, rowNumber, map, ['ՀՎՀՀ','HVHH','Tax ID','Õ€ÕŽÕ€Õ€'], p.hvhh || '');
  setTenantCell_(sh, rowNumber, map, ['Իրավ. հասցե','Legal Address','Ô»Ö€Õ¡Õ¾. Õ°Õ¡Õ½ÖÕ¥'], p.legalAddress || '');
  setTenantCell_(sh, rowNumber, map, ['Գրանցման համար','Registration Number','Ô³Ö€Õ¡Õ¶ÖÕ´Õ¡Õ¶ Õ°Õ¡Õ´Õ¡Ö€'], p.registrationNumber || '');
  setTenantCell_(sh, rowNumber, map, ['Տնօրենի անուն','CEO Name','ÕÕ¶Ö…Ö€Õ¥Õ¶Õ« Õ¡Õ¶Õ¸Ö‚Õ¶'], p.ceoName || '');
  setTenantCell_(sh, rowNumber, map, ['Տնօրենի անձնագիր','CEO Passport','ÕÕ¶Ö…Ö€Õ¥Õ¶Õ« Õ¡Õ¶Õ±Õ¶Õ¡Õ£Õ«Ö€'], p.ceoPassport || '');
  setTenantCell_(sh, rowNumber, map, ['Տնօրենի հասցե','CEO Address','ÕÕ¶Ö…Ö€Õ¥Õ¶Õ« Õ°Õ¡Õ½ÖÕ¥'], p.ceoAddress || '');
  setTenantCell_(sh, rowNumber, map, ['Տնօրենի տրված է','CEO Issued From','ÕÕ¶Ö…Ö€Õ¥Õ¶Õ« Õ¿Ö€Õ¾Õ¡Õ® Õ§'], p.ceoIssuedFrom || '');
  setTenantCell_(sh, rowNumber, map, ['Տնօրենի վավեր է մինչև','CEO Valid Until','ÕÕ¶Ö…Ö€Õ¥Õ¶Õ« Õ¾Õ¡Õ¾Õ¥Ö€ Õ§ Õ´Õ«Õ¶Õ¹Ö‡'], dateOrBlank_(p.ceoValidUntil));
  setTenantCell_(sh, rowNumber, map, ['Սկիզբ','Start Date','ÕÕ¯Õ«Õ¦Õ¢'], dateOrBlank_(p.startDate));
  setTenantCell_(sh, rowNumber, map, ['Ավարտ','End Date','Ô±Õ¾Õ¡Ö€Õ¿'], dateOrBlank_(p.endDate));
  setTenantCell_(sh, rowNumber, map, ['Վճարման օր','Payment Date','ÕŽÕ³Õ¡Ö€Õ´Õ¡Õ¶ Ö…Ö€'], p.paymentDate || p.day || '');
  if (isNew) setTenantCell_(sh, rowNumber, map, ['Ստեղծվել է','Created At','ÕÕ¿Õ¥Õ²Õ®Õ¾Õ¥Õ¬ Õ§'], new Date());
  setTenantCell_(sh, rowNumber, map, ['Թարմացվել է','Updated At','Ô¹Õ¡Ö€Õ´Õ¡ÖÕ¾Õ¥Õ¬ Õ§'], new Date());

  return { ok: true };
}

// â”€â”€ HELPERS â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
function isRentPaidStatus_(status) {
  return status === 'paid' || status === 'deposit' || status === 'prepaid';
}

function findPaymentRow_(sh, map, month, apt) {
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return 0;
  const values = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();
  for (let i = 0; i < values.length; i++) {
    const rowMonth = String(first_(values[i], map, ['Month'], '')).trim();
    const rowApt   = officeKey_(firstNonBlank_(values[i], map, ['Apt','Ô³Ö€Õ¡Õ½Õ¥Õ¶ÕµÕ¡Õ¯','Office','Apartment','Room','Ô³Ö€Õ½'], ''));
    if (rowMonth === month && rowApt === apt) return i + 2;
  }
  return 0;
}

function setCell_(sh, row, map, header, value) {
  const col = map[norm_(header)];
  if (col === undefined) throw new Error('Missing Payments column: ' + header);
  sh.getRange(row, col + 1).setValue(value);
}

// Read a single cell value from a row by header name (returns '' if missing)
function getCellVal_(sh, row, map, header) {
  const col = map[norm_(header)];
  if (col === undefined) return '';
  return sh.getRange(row, col + 1).getValue();
}

// Delete all duplicate rows for month+apt, keeping only the first one found
function dedupPaymentRows_(sh, map, month, apt) {
  const lastRow = sh.getLastRow();
  if (lastRow < 3) return;
  const values = sh.getRange(2, 1, lastRow - 1, sh.getLastColumn()).getValues();
  // Collect row numbers of duplicates (skip first match, delete the rest)
  let found = false;
  const toDelete = [];
  for (let i = values.length - 1; i >= 0; i--) {
    const rowMonth = String(first_(values[i], map, ['Month'], '')).trim();
    const rowApt   = officeKey_(firstNonBlank_(values[i], map, ['Apt','Ô³Ö€Õ¡Õ½Õ¥Õ¶ÕµÕ¡Õ¯','Office','Apartment','Room','Ô³Ö€Õ½'], ''));
    if (rowMonth === month && rowApt === apt) {
      if (found) toDelete.push(i + 2); // delete later duplicates (working bottom-up)
      else found = true;
    }
  }
  // Delete from bottom up so row numbers stay valid
  toDelete.forEach(r => sh.deleteRow(r));
}

function setTenantCell_(sh, row, map, possibleHeaders, value) {
  for (let i = 0; i < possibleHeaders.length; i++) {
    const col = map[norm_(possibleHeaders[i])];
    if (col !== undefined) {
      sh.getRange(row, col + 1).setValue(value);
      return;
    }
  }
  throw new Error('Missing tenants column: ' + possibleHeaders[0]);
}

// Contract templates are copied as-is; only {{VARIABLE}} markers are replaced.
function generateContract_(p) {
  const values = commonDocumentValues_(p);
  const termStart = p.termStart || p.documentDate;
  const termEnd = p.termEnd || addMonthsDateValue_(termStart, 3);
  values.WING = contractWing_(p.wing);
  values.UNIT_DESCRIPTION = contractUnit_(p.unitDescription);
  values.TERM_START = documentDate_(termStart);
  values.TERM_END = documentDate_(termEnd);
  values.MONTHLY_RENT = documentMoney_(p.monthlyRent);
  values.RENT_IN_WORDS = String(p.rentInWords || '').trim();
  values.PAYMENT_DAY = String(p.paymentDay || '').trim();
  values.FIRST_PAYMENT_DATE = documentDate_(p.firstPaymentDate || termStart);
  values.DEPOSIT_PAYMENT_DATE = documentDate_(p.depositPaymentDate || termStart);

  return createFromDocumentTemplate_(
    DOCUMENT_TEMPLATES.CONTRACT,
    'Պայմանագիր - ' + values.TENANT_NAME + ' - ' + values.DOCUMENT_DATE,
    values,
    DOCUMENT_TEMPLATE_CHECKS.CONTRACT,
    tenantFolder_(p)
  );
}

function idempotentDocumentGeneration_(action, p, createFn) {
  const requestId = String(p.requestId || '').trim();
  if (!requestId) return createFn();

  const cache = CacheService.getScriptCache();
  const cacheKey = 'document:' + action + ':' + requestId;
  const cached = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const afterLock = cache.get(cacheKey);
    if (afterLock) return JSON.parse(afterLock);

    const result = createFn();
    cache.put(cacheKey, JSON.stringify(result), 21600);
    return result;
  } finally {
    lock.releaseLock();
  }
}

function generateExtension_(p) {
  const values = commonDocumentValues_(p);
  const extensionStart = p.termStart || p.specialPeriodEnd || p.documentDate;
  const extensionEnd = p.termEnd || addMonthsDateValue_(extensionStart, 1);
  const specialPeriodEnd = p.specialPeriodEnd || extensionStart;
  const specialPeriodStart = p.specialPeriodStart || addMonthsDateValue_(specialPeriodEnd, -1);
  values.EXTENSION_NUMBER = String(p.extensionNumber || '1').trim();
  values.ORIGINAL_CONTRACT_DATE = documentDate_(p.originalContractDate || p.documentDate);
  values.WING = String(p.wing || '').trim();
  values.UNIT_DESCRIPTION = extensionUnit_(p.unitDescription);
  values.EXTENSION_PERIOD_TEXT = String(p.extensionPeriodText || '').trim();
  values.EXTENSION_START = documentDate_(extensionStart);
  values.EXTENSION_END = documentDate_(extensionEnd);
  values.SPECIAL_PERIOD_START = documentDate_(specialPeriodStart);
  values.SPECIAL_PERIOD_END = documentDate_(specialPeriodEnd);
  values.SPECIAL_RENT_AMOUNT = documentMoney_(p.specialRentAmount || p.monthlyRent);
  values.SPECIAL_RENT_IN_WORDS = String(p.specialRentInWords || p.rentInWords || '').trim();
  values.SPECIAL_PAYMENT_DEADLINE = documentDate_(p.specialPaymentDeadline || p.documentDate);

  return createFromDocumentTemplate_(
    DOCUMENT_TEMPLATES.EXTENSION,
    'Լրացուցիչ համաձայնագիր թիվ ' + values.EXTENSION_NUMBER + ' - ' +
      values.TENANT_NAME + ' - ' + values.DOCUMENT_DATE,
    values,
    DOCUMENT_TEMPLATE_CHECKS.EXTENSION,
    tenantFolder_(p)
  );
}

function commonDocumentValues_(p) {
  return {
    CITY: String(p.city || 'ք. Երևան').trim(),
    DOCUMENT_DATE: documentDate_(p.documentDate),
    TENANT_NAME: String(p.tenantName || '').trim(),
    TENANT_NAME_QUOTED: quotedTenantName_(p.tenantName, p.entityType),
    PROPERTY_ADDRESS: String(p.propertyAddress || '').trim(),
    FLOOR: String(p.floor || '').trim().replace(/\s*հարկ$/i, ''),
    PURPOSE: String(p.purpose || '').trim(),
    TENANT_ADDRESS: String(p.tenantAddress || p.legalAddress || '').trim(),
    TENANT_TAX_ID: String(p.hvhh || '').trim(),
    TENANT_BANK: String(p.tenantBank || '').trim(),
    TENANT_BANK_ACCOUNT: String(p.tenantBankAccount || '').trim(),
    TENANT_PHONE: documentPhone_(p.phone)
  };
}

function documentPhone_(value) {
  let digits = String(value || '').replace(/\D/g, '');
  if (digits.length === 9 && digits.charAt(0) === '0') digits = digits.slice(1);
  if (digits.length === 11 && digits.indexOf('374') === 0) digits = digits.slice(3);
  if (digits.length !== 8) return String(value || '').trim();
  return '+374 ' + digits.slice(0, 2) + ' ' + digits.slice(2);
}

function tenantFolderName_(p) {
  const apt = String(p.apt || '').trim();
  const tenantName = String(p.tenantName || '').trim();
  if (!apt) throw new Error('Tenant office is required');
  return ('Office ' + apt + (tenantName ? ' - ' + tenantName : ''))
    .replace(/[\\/:*?"<>|]/g, '-')
    .replace(/\s+/g, ' ')
    .trim();
}

function tenantFolderPropertyKey_(apt) {
  return 'TENANT_FOLDER_' + Utilities.base64EncodeWebSafe(String(apt || '').trim())
    .replace(/=+$/g, '');
}

function tenantDocumentsRoot_() {
  const props = PropertiesService.getScriptProperties();
  const propertyKey = 'TENANT_DOCUMENTS_ROOT_ID_V2';
  const savedId = props.getProperty(propertyKey);
  if (savedId) {
    try { return DriveApp.getFolderById(savedId); } catch (err) {}
  }

  const myDrive = DriveApp.getRootFolder();
  const matches = myDrive.getFoldersByName(TENANT_DOCUMENTS_FOLDER);
  const folder = matches.hasNext() ? matches.next() : myDrive.createFolder(TENANT_DOCUMENTS_FOLDER);
  props.setProperty(propertyKey, folder.getId());
  return folder;
}

function moveFolderIntoTenantRoot_(folder) {
  const root = tenantDocumentsRoot_();
  const parents = folder.getParents();
  while (parents.hasNext()) {
    if (parents.next().getId() === root.getId()) return folder;
  }
  folder.moveTo(root);
  return folder;
}

function tenantFolder_(p) {
  const apt = String(p.apt || '').trim();
  const desiredName = tenantFolderName_(p);
  const props = PropertiesService.getScriptProperties();
  const key = tenantFolderPropertyKey_(apt);
  const savedId = props.getProperty(key);
  if (savedId) {
    try {
      const saved = DriveApp.getFolderById(savedId);
      if (saved.getName() !== desiredName) saved.setName(desiredName);
      return moveFolderIntoTenantRoot_(saved);
    } catch (err) {
      props.deleteProperty(key);
    }
  }

  const lock = LockService.getUserLock();
  lock.waitLock(30000);
  try {
    const afterLockId = props.getProperty(key);
    if (afterLockId) {
      try { return moveFolderIntoTenantRoot_(DriveApp.getFolderById(afterLockId)); } catch (err) {}
    }

    const root = tenantDocumentsRoot_();
    const exact = root.getFoldersByName(desiredName);
    const folder = exact.hasNext() ? exact.next() : root.createFolder(desiredName);
    props.setProperty(key, folder.getId());
    return folder;
  } finally {
    lock.releaseLock();
  }
}

function getTenantFolder_(p) {
  const folder = tenantFolder_(p);
  return { ok: true, folderId: folder.getId(), folderUrl: folder.getUrl() };
}

function uploadTenantFile_(p) {
  const fileName = String(p.fileName || '').trim()
    .replace(/[\\/:*?"<>|]/g, '-');
  const mimeType = String(p.mimeType || 'application/octet-stream').trim();
  const base64 = String(p.fileData || '').replace(/^data:[^;]+;base64,/, '');
  if (!fileName) throw new Error('File name is required');
  if (!base64) throw new Error('File data is required');

  const bytes = Utilities.base64Decode(base64);
  if (bytes.length > MAX_TENANT_UPLOAD_BYTES) {
    throw new Error('File is too large. Maximum size is 15 MB.');
  }

  const folder = tenantFolder_(p);
  const file = folder.createFile(Utilities.newBlob(bytes, mimeType, fileName));
  return {
    ok: true,
    fileName: file.getName(),
    fileUrl: file.getUrl(),
    folderUrl: folder.getUrl()
  };
}

function createFromDocumentTemplate_(templateId, name, values, integrityCheck, outputFolder) {
  const template = DriveApp.getFileById(templateId);
  const folder = outputFolder || null;
  const copy = folder ? template.makeCopy(name, folder) : template.makeCopy(name);
  const document = DocumentApp.openById(copy.getId());
  const body = document.getBody();

  assertDocumentTemplateIntegrity_(body, integrityCheck);

  Object.keys(values).forEach(key => {
    body.replaceText(
      '\\{\\{' + key + '\\}\\}',
      safeDocumentReplacement_(values[key])
    );
  });

  document.saveAndClose();

  const completedDocument = DocumentApp.openById(copy.getId());
  const completedBody = completedDocument.getBody();
  assertDocumentTemplateIntegrity_(completedBody, integrityCheck);
  const remaining = completedBody.getText().match(/\{\{[A-Z0-9_]+\}\}/g);
  if (remaining && remaining.length) {
    copy.setTrashed(true);
    throw new Error('Unfilled template variables: ' + remaining.join(', '));
  }

  completedDocument.saveAndClose();
  Utilities.sleep(1000);
  const pdfBlob = copy.getBlob().getAs(MimeType.PDF).setName(name + '.pdf');
  const pdfFile = folder ? folder.createFile(pdfBlob) : DriveApp.createFile(pdfBlob);
  return {
    ok: true,
    docUrl: 'https://docs.google.com/document/d/' + copy.getId() + '/edit',
    pdfUrl: pdfFile.getUrl(),
    folderUrl: folder ? folder.getUrl() : ''
  };
}

function assertDocumentTemplateIntegrity_(body, check) {
  if (!check) return;
  const paragraphs = body.getParagraphs();
  if (paragraphs.length < check.minimumParagraphs) {
    throw new Error(
      'Document template is incomplete: expected at least ' +
      check.minimumParagraphs + ' paragraphs, found ' + paragraphs.length
    );
  }

  const text = body.getText();
  const missing = check.requiredText.filter(value => text.indexOf(value) === -1);
  if (missing.length) {
    throw new Error('Document template is missing required clauses: ' + missing.join(' | '));
  }
}

function testDocumentTemplates_() {
  const results = {};
  [
    ['contract', DOCUMENT_TEMPLATES.CONTRACT, DOCUMENT_TEMPLATE_CHECKS.CONTRACT],
    ['extension', DOCUMENT_TEMPLATES.EXTENSION, DOCUMENT_TEMPLATE_CHECKS.EXTENSION]
  ].forEach(item => {
    const document = DocumentApp.openById(item[1]);
    const body = document.getBody();
    assertDocumentTemplateIntegrity_(body, item[2]);
    results[item[0]] = {
      ok: true,
      paragraphCount: body.getParagraphs().length,
      templateId: item[1]
    };
  });
  return { ok: true, templates: results };
}

// Visible in the Apps Script function dropdown. Run once to authorize Drive/Docs.
function authorizeDocuments() {
  return testDocumentTemplates_();
}

function safeDocumentReplacement_(value) {
  return String(value == null ? '' : value).replace(/\\/g, '\\\\').replace(/\$/g, '\\$');
}

function documentDate_(value) {
  const text = String(value || '').trim();
  if (!text) return '';
  const iso = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return iso[3] + '.' + iso[2] + '.' + iso[1];
  return text.replace(/թ\.?$/, '');
}

function addMonthsDateValue_(value, months) {
  const text = String(value || '').trim();
  let match = text.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!match) {
    const armenianDate = text.match(/^(\d{2})\.(\d{2})\.(\d{4})/);
    if (!armenianDate) return '';
    match = [armenianDate[0], armenianDate[3], armenianDate[2], armenianDate[1]];
  }

  const date = new Date(Date.UTC(Number(match[1]), Number(match[2]) - 1 + months, Number(match[3])));
  return Utilities.formatDate(date, 'UTC', 'yyyy-MM-dd');
}

function documentMoney_(value) {
  const number = Math.round(num_(value));
  return number ? number.toLocaleString('en-US') : '';
}

function contractWing_(value) {
  const wing = String(value || '').trim();
  if (!wing) return '';
  return /ից$/.test(wing) ? wing : wing + 'ից';
}

function contractUnit_(value) {
  return String(value || '').trim().replace(/տարածք(?:ը|ի)?$/, 'տարածքը');
}

function extensionUnit_(value) {
  return String(value || '').trim().replace(/տարածք(?:ը|ի)?$/, 'տարածքի');
}

function quotedTenantName_(value, entityType) {
  const name = String(value || '').trim();
  if (!name) return '';
  if (String(entityType || '').toLowerCase() !== 'legal') return name;

  const match = name.match(/^(.*?)(\s+(?:ԱՁ|ՍՊԸ|ՓԲԸ|ԲԲԸ|ՀԿ|Հիմնադրամ))$/);
  return match ? '«' + match[1].trim() + '»' + match[2] : '«' + name + '»';
}

// â”€â”€ TEST (open ?action=test in browser to diagnose) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
function runTest_() {
  const result = { ok: true, sheets: {} };
  try {
    const t = getTenants_();
    result.sheets.tenants = { count: t.length, sample: t.slice(0, 2) };
  } catch(e) { result.sheets.tenants = { error: e.message }; }
  try {
    const elec = getElectricity_();
    const keys = Object.keys(elec);
    result.sheets.meterData = { count: keys.length, sample: keys.slice(0, 2).map(k => elec[k]) };
  } catch(e) { result.sheets.meterData = { error: e.message }; }
  try {
    const pay = getPayments_();
    const keys = Object.keys(pay);
    result.sheets.payments = { count: keys.length, sample: keys.slice(0, 2).map(k => pay[k]) };
  } catch(e) { result.sheets.payments = { error: e.message }; }
  return result;
}
