/**
 * Vehicle shipping details — submission endpoint
 * ------------------------------------------------
 * Accepts one submission containing one or more vehicles and writes one row
 * per vehicle. Deploy as a web app: Deploy > New deployment > Web app,
 * Execute as Me, Who has access Anyone, then copy the /exec URL into
 * index.html (the ENDPOINT constant).
 *
 * Optional check: run testAppend() once from the editor; a sample row should
 * appear in the "Submissions" tab.
 *
 * The tab and header row are created automatically if they don't exist.
 */

const SHEET_NAME = 'Submissions';

const HEADERS = [
  'Timestamp',
  'Customer ID',
  'Company',
  'VIN',
  'Brand',
  'Model',
  'Engine Type',
  'Engine Capacity',
  'Year',
  'Status',
  'Notes',
  'Date received'
];

// The form only ever writes columns A-I. J-L are yours to fill in by hand,
// and the script never touches their cells.
const DATA_COLUMNS = 9;
const VIN_COLUMN = 4;
const CAPACITY_COLUMN = 8;
const MAX_VEHICLES_PER_SUBMISSION = 50;

function doPost(e) {
  const lock = LockService.getScriptLock();
  let acquired = false;

  try {
    if (!e || !e.postData || !e.postData.contents) {
      return respondJson({ ok: false, error: 'Missing or empty request body.' });
    }

    let payload;
    try {
      payload = JSON.parse(e.postData.contents);
    } catch (parseError) {
      return respondJson({ ok: false, error: 'Malformed JSON payload.' });
    }

    if (!payload || typeof payload !== 'object' || Array.isArray(payload)) {
      return respondJson({ ok: false, error: 'Payload must be a JSON object.' });
    }

    const customerId = sanitizeString(payload.customerId, 100);
    const company = sanitizeString(payload.company, 150);

    if (!customerId) {
      return respondJson({ ok: false, error: 'Customer ID is required.' });
    }
    if (!company) {
      return respondJson({ ok: false, error: 'Company / account name is required.' });
    }

    // New shape is { vehicles: [...] }. A single vehicle posted flat is still
    // accepted so an older cached copy of the page doesn't lose data.
    const submitted = Array.isArray(payload.vehicles) ? payload.vehicles : [payload];

    if (submitted.length === 0) {
      return respondJson({ ok: false, error: 'Add at least one vehicle.' });
    }
    if (submitted.length > MAX_VEHICLES_PER_SUBMISSION) {
      return respondJson({
        ok: false,
        error: 'Too many vehicles in one submission (limit ' + MAX_VEHICLES_PER_SUBMISSION + ').'
      });
    }

    const vehicles = [];
    for (let i = 0; i < submitted.length; i++) {
      const checked = validateVehicle(submitted[i]);
      if (!checked.ok) {
        return respondJson({ ok: false, error: 'Vehicle ' + (i + 1) + ': ' + checked.error });
      }
      vehicles.push(checked.vehicle);
    }

    if (!lock.tryLock(30000)) {
      return respondJson({ ok: false, error: 'Server busy. Please try again.' });
    }
    acquired = true;

    const sheet = getSheet();
    const timestamp = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      'yyyy-MM-dd HH:mm:ss'
    );

    const rows = vehicles.map(function (vehicle) {
      return [
        timestamp,
        customerId,
        company,
        vehicle.vin,
        vehicle.brand,
        vehicle.model,
        vehicle.engineType,
        vehicle.engineCapacity,
        vehicle.year
      ];
    });

    // One batch write rather than a row-at-a-time append. Only A-I: the
    // working columns J-L are left untouched.
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rows.length, DATA_COLUMNS).setValues(rows);

    // Keep VIN and capacity as text so Sheets never reformats them.
    sheet.getRange(startRow, VIN_COLUMN, rows.length, 1).setNumberFormat('@');
    sheet.getRange(startRow, CAPACITY_COLUMN, rows.length, 1).setNumberFormat('@');

    return respondJson({ ok: true, added: vehicles.length });
  } catch (err) {
    return respondJson({ ok: false, error: (err && err.message) || 'Internal server error.' });
  } finally {
    if (acquired) lock.releaseLock();
  }
}

function doGet() {
  return respondJson({ ok: true, message: 'Vehicle intake endpoint is live. Send a POST request.' });
}

/**
 * Validates one vehicle object. Returns { ok: true, vehicle } or
 * { ok: false, error } so the caller can name the offending vehicle.
 */
function validateVehicle(raw) {
  const source = (raw && typeof raw === 'object') ? raw : {};

  const vinInput = String(source.vin == null ? '' : source.vin).trim();
  if (!vinInput) return { ok: false, error: 'VIN is required.' };
  const vin = sanitizeVin(vinInput);
  if (vin === null) {
    return { ok: false, error: 'VIN should be 11 to 17 letters and numbers, without I, O or Q.' };
  }

  const brand = sanitizeString(source.brand, 100);
  if (!brand) return { ok: false, error: 'Brand is required.' };

  const model = sanitizeString(source.model, 100);
  if (!model) return { ok: false, error: 'Model is required.' };

  const engineType = sanitizeString(source.engineType, 100);
  if (!engineType) return { ok: false, error: 'Engine type is required.' };

  const engineCapacity = sanitizeString(source.engineCapacity, 100);
  if (!engineCapacity) return { ok: false, error: 'Engine capacity is required.' };

  const year = validateYear(source.year);
  if (year === null) {
    return { ok: false, error: 'Year of car make should be a year between 1900 and ' + maxYear() + '.' };
  }

  return {
    ok: true,
    vehicle: { vin: vin, brand: brand, model: model, engineType: engineType, engineCapacity: engineCapacity, year: year }
  };
}

function maxYear() {
  return new Date().getFullYear() + 1;
}

/**
 * Trims strings, bounds length, and neutralises spreadsheet formula injection.
 */
function sanitizeString(value, maxLength) {
  if (value === null || value === undefined) return '';
  let text = String(value).trim();
  if (maxLength && text.length > maxLength) {
    text = text.substring(0, maxLength);
  }
  if (/^[=+\-@]/.test(text)) {
    text = "'" + text;
  }
  return text;
}

/**
 * Returns the cleaned VIN, '' when missing, or null when invalid.
 */
function sanitizeVin(value) {
  const text = String(value == null ? '' : value).trim().toUpperCase();
  if (text === '') return '';
  if (!/^[A-HJ-NPR-Z0-9]{11,17}$/.test(text)) return null;
  return text;
}

/**
 * Returns the year as a number, or null when missing or out of range.
 */
function validateYear(value) {
  const text = String(value == null ? '' : value).trim();
  if (!/^\d{4}$/.test(text)) return null;
  const year = parseInt(text, 10);
  if (year < 1900 || year > maxYear()) return null;
  return year;
}

function getSheet() {
  const book = SpreadsheetApp.getActiveSpreadsheet();
  if (!book) {
    throw new Error('This script must be attached to a spreadsheet.');
  }

  let sheet = book.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = book.insertSheet(SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length)
      .setFontWeight('bold')
      .setBackground('#f1f3f4');
    sheet.getRange(1, DATA_COLUMNS + 1, 1, HEADERS.length - DATA_COLUMNS)
      .setFontStyle('italic');
    sheet.setFrozenRows(1);
  } else {
    ensureWorkingHeaders(sheet);
  }

  return sheet;
}

/**
 * Adds the manual working headers (Status, Notes, Date received) to an
 * existing tab. Only fills cells that are empty, so anything you have typed
 * in that row is left alone.
 */
function ensureWorkingHeaders(sheet) {
  const working = HEADERS.slice(DATA_COLUMNS);
  const existing = sheet.getRange(1, DATA_COLUMNS + 1, 1, working.length).getValues()[0];

  const missing = [];
  for (let i = 0; i < working.length; i++) {
    if (String(existing[i]).trim() === '') missing.push(i + 1);
  }
  if (missing.length === 0) return;

  missing.forEach(function (columnIndex) {
    const label = working[columnIndex - 1];
    sheet.getRange(1, DATA_COLUMNS + columnIndex)
      .setValue(label)
      .setFontWeight('bold')
      .setBackground('#f1f3f4')
      .setFontStyle('italic');
  });
}

function respondJson(body) {
  return ContentService
    .createTextOutput(JSON.stringify(body))
    .setMimeType(ContentService.MimeType.JSON);
}

/** Writes one sample row so you can confirm the sheet is wired up. */
function testAppend() {
  const sheet = getSheet();
  const row = [
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
    'test-customer-id',
    'Test Company',
    'TEST0000000000001',
    'Test',
    'Sample',
    'Petrol',
    '1998 cc',
    2020
  ];
  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, 1, DATA_COLUMNS).setValues([row]);
  sheet.getRange(startRow, VIN_COLUMN, 1, 1).setNumberFormat('@');
  sheet.getRange(startRow, CAPACITY_COLUMN, 1, 1).setNumberFormat('@');
}
