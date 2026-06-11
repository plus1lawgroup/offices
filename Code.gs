// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•
//  PLUS ONE LAW GROUP â€” Google Apps Script Backend
//  Sheets: tenants Â· MeterData Â· Payments
// â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•â•

const SHEETS = {
  TENANTS:     'tenants',
  ELECTRICITY: 'MeterData',
  PAYMENTS:    'Payments'
};


// â”€â”€ Electricity tariff rates (AMD per kWh) â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
const AMD_PER_KWH_DAY   = 42.39;   // T1 Day tariff   â€” edit here
const AMD_PER_KWH_NIGHT = 21.19;   // T2 Night tariff â€” edit here

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
    if (action === 'update')        return json(e, updatePayment_(e.parameter));
    if (action === 'saveMeter')     return json(e, saveMeter_(e.parameter));
    if (action === 'deletePayment') return json(e, deletePayment_(e.parameter));
    if (action === 'updateTenant')  return json(e, updateTenant_(e.parameter));
    if (action === 'test')          return json(e, runTest_());
    return json(e, { ok: false, error: 'Unknown action: ' + action });
  } catch (err) {
    return json(e, { ok: false, error: String(err && err.message ? err.message : err) });
  }
}

// â”€â”€ JSONP / JSON response â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€â”€
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
  return String(value || '').toLowerCase().replace(/[\s._\-()\/\\:;,'"Â«Â»ÕÕžÕœÕ›]+/g, '');
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
  return Number(String(value).replace(/[, Ö\s]/g, '')) || 0;
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
      startMonth: String(firstNonBlank_(row, map, ['Սկզբնական ամիս','Start Month','Move In Month','First Month','ÕÕ¯Õ¦Õ¢Õ¶Õ¡Õ¯Õ¡Õ¶ Õ¡Õ´Õ«Õ½'], '')).trim(),
      entityType: String(firstNonBlank_(row, map, ['Տեսակ','Entity Type','ÕÕ¥Õ½Õ¡Õ¯'], '')).trim(),
      physicalName: String(firstNonBlank_(row, map, ['Ֆիզ. անձի անուն','Physical Name','Õ–Õ«Õ¦. Õ¡Õ¶Õ±Õ« Õ¡Õ¶Õ¸Ö‚Õ¶'], '')).trim(),
      physicalPassport: String(firstNonBlank_(row, map, ['Ֆիզ. անձի անձնագիր','Physical Passport','Õ–Õ«Õ¦. Õ¡Õ¶Õ±Õ« Õ¡Õ¶Õ±Õ¶Õ¡Õ£Õ«Ö€'], '')).trim(),
      physicalAddress: String(firstNonBlank_(row, map, ['Ֆիզ. անձի հասցե','Physical Address','Õ–Õ«Õ¦. Õ¡Õ¶Õ±Õ« Õ°Õ¡Õ½ÖÕ¥'], '')).trim(),
      physicalIssuedFrom: String(firstNonBlank_(row, map, ['Ֆիզ. անձի տրված է','Physical Issued From','Õ–Õ«Õ¦. Õ¡Õ¶Õ±Õ« Õ¿Ö€Õ¾Õ¡Õ® Õ§'], '')).trim(),
      physicalValidUntil: firstNonBlank_(row, map, ['Ֆիզ. անձի վավեր է մինչև','Physical Valid Until','Õ–Õ«Õ¦. Õ¡Õ¶Õ±Õ« Õ¾Õ¡Õ¾Õ¥Ö€ Õ§ Õ´Õ«Õ¶Õ¹Ö‡'], ''),
      legalName: String(firstNonBlank_(row, map, ['Իրավ. անձի անվանում','Legal Name','Ô»Ö€Õ¡Õ¾. Õ¡Õ¶Õ±Õ« Õ¡Õ¶Õ¾Õ¡Õ¶Õ¸Ö‚Õ´'], '')).trim(),
      hvhh: String(firstNonBlank_(row, map, ['ՀՎՀՀ','HVHH','Tax ID','Õ€ÕŽÕ€Õ€'], '')).trim(),
      legalAddress: String(firstNonBlank_(row, map, ['Իրավ. հասցե','Legal Address','Ô»Ö€Õ¡Õ¾. Õ°Õ¡Õ½ÖÕ¥'], '')).trim(),
      registrationNumber: String(firstNonBlank_(row, map, ['Գրանցման համար','Registration Number','Ô³Ö€Õ¡Õ¶ÖÕ´Õ¡Õ¶ Õ°Õ¡Õ´Õ¡Ö€'], '')).trim(),
      ceoName: String(firstNonBlank_(row, map, ['Տնօրենի անուն','CEO Name','ÕÕ¶Ö…Ö€Õ¥Õ¶Õ« Õ¡Õ¶Õ¸Ö‚Õ¶'], '')).trim(),
      ceoPassport: String(firstNonBlank_(row, map, ['Տնօրենի անձնագիր','CEO Passport','ÕÕ¶Ö…Ö€Õ¥Õ¶Õ« Õ¡Õ¶Õ±Õ¶Õ¡Õ£Õ«Ö€'], '')).trim(),
      ceoAddress: String(firstNonBlank_(row, map, ['Տնօրենի հասցե','CEO Address','ÕÕ¶Ö…Ö€Õ¥Õ¶Õ« Õ°Õ¡Õ½ÖÕ¥'], '')).trim(),
      ceoIssuedFrom: String(firstNonBlank_(row, map, ['Տնօրենի տրված է','CEO Issued From','ÕÕ¶Ö…Ö€Õ¥Õ¶Õ« Õ¿Ö€Õ¾Õ¡Õ® Õ§'], '')).trim(),
      ceoValidUntil: firstNonBlank_(row, map, ['Տնօրենի վավեր է մինչև','CEO Valid Until','ÕÕ¶Ö…Ö€Õ¥Õ¶Õ« Õ¾Õ¡Õ¾Õ¥Ö€ Õ§ Õ´Õ«Õ¶Õ¹Ö‡'], ''),
      startDate: firstNonBlank_(row, map, ['Սկիզբ','Start Date','ÕÕ¯Õ«Õ¦Õ¢'], ''),
      endDate: firstNonBlank_(row, map, ['Ավարտ','End Date','Ô±Õ¾Õ¡Ö€Õ¿'], ''),
      paymentDate: firstNonBlank_(row, map, ['Վճարման օր','Payment Date','ÕŽÕ³Õ¡Ö€Õ´Õ¡Õ¶ Ö…Ö€'], ''),
      createdAt: firstNonBlank_(row, map, ['Ստեղծվել է','Created At','ÕÕ¿Õ¥Õ²Õ®Õ¾Õ¥Õ¬ Õ§'], '')
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
      ['Office','Ô³Ñ€Ð°ÑÐµÐ½rak','Apt','Apartment','Room'], ''));
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

  const t1cur    = num_(p.t1current);
  const t2cur    = num_(p.t2current);
  const previous = findPreviousMeter_(sh, map, apt, month);
  const serial   = previous.serial;
  const t1prev   = previous.t1current;
  const t2prev   = previous.t2current;

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

  // Find existing row for this month + office (upsert)
  let rowNumber = 0;
  if (sh.getLastRow() >= 2) {
    const vals = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues();
    for (let i = 0; i < vals.length; i++) {
      const rm = String(first_(vals[i], map, ['Month','Ô±Õ´is'], '')).trim();
      const ra = officeKey_(firstNonBlank_(vals[i], map,
        ['Office','Ô³Ñ€Ð°ÑÐµÐ½rak','Apt','Apartment','Room'], ''));
      if (rm === month && ra === apt) { rowNumber = i + 2; break; }
    }
  }

  // Build row array in METER_HEADERS column order
  const numCols = sh.getLastColumn();
  const rowData = rowNumber
    ? sh.getRange(rowNumber, 1, 1, numCols).getValues()[0].slice()
    : new Array(numCols).fill('');

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
  setCol('PreviousMonth', previous.month);
  setCol('PreviousAssumed', previous.assumed ? 'yes' : 'no');
  setCol('EnteredAt', new Date());

  if (rowNumber) {
    sh.getRange(rowNumber, 1, 1, numCols).setValues([rowData]);
  } else {
    sh.getRange(sh.getLastRow() + 1, 1, 1, numCols).setValues([rowData]);
  }

  return { ok: true, apt, month, serial, t1prev, t2prev,
           t1current: t1cur, t2current: t2cur,
           totalPrevious, kwhDay, kwhNight, amdDay, amdNight, amdTotal,
           previousMonth: previous.month, previousAssumed: previous.assumed };
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
  if (isNew || p.startMonth) setTenantCell_(sh, rowNumber, map, ['Սկզբնական ամիս','Start Month','Move In Month','First Month','ÕÕ¯Õ¦Õ¢Õ¶Õ¡Õ¯Õ¡Õ¶ Õ¡Õ´Õ«Õ½'], p.startMonth || '');
  const entityType = p.entityType || 'physical';
  const displayName = p.name || p.physicalName || p.legalName || '';
  setTenantCell_(sh, rowNumber, map, ['Տեսակ','Entity Type','ÕÕ¥Õ½Õ¡Õ¯'], entityType);
  setTenantCell_(sh, rowNumber, map, ['Ֆիզ. անձի անուն','Physical Name','Õ–Õ«Õ¦. Õ¡Õ¶Õ±Õ« Õ¡Õ¶Õ¸Ö‚Õ¶'], entityType === 'physical' ? (p.physicalName || displayName) : (p.physicalName || ''));
  setTenantCell_(sh, rowNumber, map, ['Ֆիզ. անձի անձնագիր','Physical Passport','Õ–Õ«Õ¦. Õ¡Õ¶Õ±Õ« Õ¡Õ¶Õ±Õ¶Õ¡Õ£Õ«Ö€'], p.physicalPassport || '');
  setTenantCell_(sh, rowNumber, map, ['Ֆիզ. անձի հասցե','Physical Address','Õ–Õ«Õ¦. Õ¡Õ¶Õ±Õ« Õ°Õ¡Õ½ÖÕ¥'], p.physicalAddress || '');
  setTenantCell_(sh, rowNumber, map, ['Ֆիզ. անձի տրված է','Physical Issued From','Õ–Õ«Õ¦. Õ¡Õ¶Õ±Õ« Õ¿Ö€Õ¾Õ¡Õ® Õ§'], p.physicalIssuedFrom || '');
  setTenantCell_(sh, rowNumber, map, ['Ֆիզ. անձի վավեր է մինչև','Physical Valid Until','Õ–Õ«Õ¦. Õ¡Õ¶Õ±Õ« Õ¾Õ¡Õ¾Õ¥Ö€ Õ§ Õ´Õ«Õ¶Õ¹Ö‡'], dateOrBlank_(p.physicalValidUntil));
  setTenantCell_(sh, rowNumber, map, ['Իրավ. անձի անվանում','Legal Name','Ô»Ö€Õ¡Õ¾. Õ¡Õ¶Õ±Õ« Õ¡Õ¶Õ¾Õ¡Õ¶Õ¸Ö‚Õ´'], entityType === 'legal' ? (p.legalName || displayName) : (p.legalName || ''));
  setTenantCell_(sh, rowNumber, map, ['ՀՎՀՀ','HVHH','Tax ID','Õ€ÕŽÕ€Õ€'], p.hvhh || '');
  setTenantCell_(sh, rowNumber, map, ['Իրավ. հասցե','Legal Address','Ô»Ö€Õ¡Õ¾. Õ°Õ¡Õ½ÖÕ¥'], p.legalAddress || '');
  setTenantCell_(sh, rowNumber, map, ['Գրանցման համար','Registration Number','Ô³Ö€Õ¡Õ¶ÖÕ´Õ¡Õ¶ Õ°Õ¡Õ´Õ¡Ö€'], p.registrationNumber || '');
  setTenantCell_(sh, rowNumber, map, ['Տնօրենի անուն','CEO Name','ÕÕ¶Ö…Ö€Õ¥Õ¶Õ« Õ¡Õ¶Õ¸Ö‚Õ¶'], p.ceoName || '');
  setTenantCell_(sh, rowNumber, map, ['Տնօրենի անձնագիր','CEO Passport','ÕÕ¶Ö…Ö€Õ¥Õ¶Õ« Õ¡Õ¶Õ±Õ¶Õ¡Õ£Õ«Ö€'], p.ceoPassport || '');
  setTenantCell_(sh, rowNumber, map, ['Տնօրենի հասցե','CEO Address','ÕÕ¶Ö…Ö€Õ¥Õ¶Õ« Õ°Õ¡Õ½ÖÕ¥'], p.ceoAddress || '');
  setTenantCell_(sh, rowNumber, map, ['Տնօրենի տրված է','CEO Issued From','ÕÕ¶Ö…Ö€Õ¥Õ¶Õ« Õ¿Ö€Õ¾Õ¡Õ® Õ§'], p.ceoIssuedFrom || '');
  setTenantCell_(sh, rowNumber, map, ['Տնօրենի վավեր է մինչև','CEO Valid Until','ÕÕ¶Ö…Ö€Õ¥Õ¶Õ« Õ¾Õ¡Õ¾Õ¥Ö€ Õ§ Õ´Õ«Õ¶Õ¹Ö‡'], dateOrBlank_(p.ceoValidUntil));
  setTenantCell_(sh, rowNumber, map, ['Սկիզբ','Start Date','ÕÕ¯Õ«Õ¦Õ¢'], dateOrBlank_(p.startDate));
  setTenantCell_(sh, rowNumber, map, ['Ավարտ','End Date','Ô±Õ¾Õ¡Ö€Õ¿'], dateOrBlank_(p.endDate));
  setTenantCell_(sh, rowNumber, map, ['Վճարման օր','Payment Date','ÕŽÕ³Õ¡Ö€Õ´Õ¡Õ¶ Ö…Ö€'], p.paymentDate || p.day || '');
  if (isNew) setTenantCell_(sh, rowNumber, map, ['Ստեղծվել է','Created At','ÕÕ¿Õ¥Õ²Õ®Õ¾Õ¥Õ¬ Õ§'], new Date());
  setTenantCell_(sh, rowNumber, map, ['Թարմացվել է','Updated At','Ô¹Õ¡Ö€Õ´Õ¡ÖÕ¾Õ¥Õ¬ Õ§'], new Date());

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
