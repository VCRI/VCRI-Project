/**
 * VC&RI Namakkal — Dairy Project Report : Google Sheet logger
 * -----------------------------------------------------------
 * Writes EVERY report generation as ONE ROW, with one value per column
 * (no JSON). The header row is created automatically on the first write.
 *
 * Setup: Sheet ▸ Extensions ▸ Apps Script ▸ paste this ▸ Deploy as Web app
 *        (Execute as: Me · Who has access: Anyone). See SHEET-SETUP.md.
 */

var SHEET_ID  = '';          // blank = the spreadsheet this script is bound to
var LOG_SHEET = 'Reports';

/* Column layout — [ Header shown in the sheet , source , key ]
   source:  'now' = timestamp · 'gen' = generation count ·
            'data' = from the summary payload · 'det' = the value the user entered */
var COLS = [
  ['Timestamp',                      'now',  ''],
  ['Date',                           'data', 'date'],
  ['Time',                           'data', 'time'],
  ['Generation #',                   'gen',  ''],
  ['Name of Farmer / Unit',          'det',  'farmerName'],
  ['Mobile',                         'det',  'mobile'],
  ['E-mail',                         'det',  'email'],
  ['Aadhar',                         'det',  'aadhar'],
  ['Address of Farmer',              'det',  'farmerAddress'],
  ['Nature of Farming',              'det',  'natureFarming'],
  ['Farm Location',                  'det',  'farmLocation'],
  ['Financing Bank',                 'det',  'bankName'],
  ['Branch',                         'det',  'branchName'],

  ['Animal Type',                    'det',  'C7'],
  ['Breed',                          'det',  'C8'],
  ['Unit Size - Batch 1',            'det',  'D9'],
  ['Unit Size - Batch 2',            'det',  'F9'],
  ['Total Animals',                  'det',  'C9'],
  ['Cost per Animal (Rs)',           'det',  'C10'],
  ['Milk Yield (L/day)',             'det',  'C11'],
  ['Insurance Premium (%)',          'det',  'C12'],
  ['Transport (Rs/animal)',          'det',  'C13'],

  ['Floor Space - Animal Shed (sqft)','det', 'C16'],
  ['Cost - Animal Shed (Rs/sqft)',   'det',  'C17'],
  ['Floor Space - Calves (sqft)',    'det',  'C18'],
  ['Cost - Calves Shed (Rs/sqft)',   'det',  'C19'],
  ['Floor Space - Store Room (sqft)','det',  'C20'],
  ['Cost - Store Room (Rs/sqft)',    'det',  'C21'],

  ['Feeder & Waterer (Rs/animal)',   'det',  'C24'],
  ['Milking Can (Rs)',               'det',  'C25'],
  ['Milking Machine (Rs)',           'det',  'C26'],
  ['Chaff Cutter (Rs)',              'det',  'C27'],
  ['Brush Cutter (Rs)',              'det',  'C28'],
  ['Feed Making Machinery (Rs)',     'det',  'C29'],
  ['Milk Processing Machinery (Rs)', 'det',  'C30'],
  ['Motor for Irrigation (Rs)',      'det',  'C31'],
  ['Sprinkler & Fogger (Rs)',        'det',  'C32'],
  ['Computers (Rs)',                 'det',  'C33'],
  ['CCTV (Rs)',                      'det',  'C34'],

  ['Floor Space - Labour Quarters',  'det',  'C37'],
  ['Cost - Labour Quarters',         'det',  'C38'],
  ['Overhead Tank (Rs)',             'det',  'C39'],
  ['Fencing (Rs)',                   'det',  'C40'],
  ['Land Levelling (Rs)',            'det',  'C41'],
  ['Borewell (Rs)',                  'det',  'C42'],
  ['Floor Space - Feed Mill',        'det',  'C43'],
  ['Cost - Feed Mill',               'det',  'C44'],
  ['Floor Space - Milk Proc. Unit',  'det',  'C45'],
  ['Cost - Milk Proc. Unit',         'det',  'C46'],
  ['Bio Gas Unit (Rs)',              'det',  'C47'],
  ['Vermicompost Unit (Rs)',         'det',  'C48'],
  ['Silage Making Unit (Rs)',        'det',  'C49'],

  ['Margin Money (%)',               'det',  'C52'],
  ['Rate of Interest (%)',           'det',  'C53'],
  ['Repayment Period (yrs)',         'det',  'C54'],
  ['Subsidy (%)',                    'det',  'C55'],

  ['Two-Wheelers (nos)',             'det',  'C58'],
  ['Cost per Two-Wheeler (Rs)',      'det',  'C59'],
  ['Four-Wheelers (nos)',            'det',  'C60'],
  ['Cost per Four-Wheeler (Rs)',     'det',  'C61'],

  ['Milking Period (days)',          'det',  'C64'],
  ['Dry Period (days)',              'det',  'C65'],
  ['Concentrate - Milking (kg/day)', 'det',  'C66'],
  ['Concentrate - Dry (kg/day)',     'det',  'C67'],
  ['Cost of Concentrate (Rs/kg)',    'det',  'C68'],
  ['Dry Fodder (kg/day)',            'det',  'C69'],
  ['Cost of Dry Fodder (Rs/kg)',     'det',  'C70'],
  ['Green Fodder Area (acres)',      'det',  'C71'],
  ['Green Fodder Yr1 (Rs/acre)',     'det',  'C72'],
  ['Green Fodder Yr2+ (Rs/acre)',    'det',  'C73'],
  ['Veterinary Aid (Rs/animal/yr)',  'det',  'C74'],
  ['Electricity (Rs/month)',         'det',  'C75'],
  ['Fuel (Rs/month)',                'det',  'C76'],
  ['Depreciation - Buildings (%)',   'det',  'C77'],
  ['Depreciation - Equipment (%)',   'det',  'C78'],
  ['Miscellaneous (Rs/animal/yr)',   'det',  'C79'],
  ['Feed Making Unit (Rs/yr)',       'det',  'C80'],
  ['Milk Processing Unit (Rs/yr)',   'det',  'C81'],
  ['Vermicompost Unit (Rs/yr)',      'det',  'C82'],
  ['Silage Making Unit (Rs/yr)',     'det',  'C83'],

  ['No. of Farm Workers',            'det',  'C86'],
  ['Wages - Farm Worker (Rs/mo)',    'det',  'C87'],
  ['No. of Farm Managers',           'det',  'C88'],
  ['Salary - Farm Manager (Rs/mo)',  'det',  'C89'],
  ['No. of Drivers',                 'det',  'C90'],
  ['Salary - Driver (Rs/mo)',        'det',  'C91'],
  ['No. of Office Staff',            'det',  'C92'],
  ['Salary - Office Staff (Rs/mo)',  'det',  'C93'],

  ['Sale Price of Milk (Rs/L)',      'det',  'C96'],
  ['Sale Price Male Calf (Rs)',      'det',  'C97'],
  ['Value Female Calf (Rs)',         'det',  'C98'],
  ['Sale Price Manure (Rs)',         'det',  'C99'],
  ['Income Feed Making Unit (Rs/yr)','det',  'C100'],
  ['Income Milk Proc. Unit (Rs/yr)', 'det',  'C101'],

  ['Total Project Cost (Rs)',        'data', 'projectCost'],
  ['Bank Loan (Rs)',                 'data', 'loan'],
  ['NPV (Rs)',                       'data', 'npv'],
  ['BCR',                            'data', 'bcr'],
  ['IRR (%)',                        'data', 'irr'],
  ['Avg DSCR',                       'data', 'dscr']
];

function doPost(e) { return handle_(e); }
function doGet(e)  { return handle_(e); }   // the website uses a GET beacon (?payload=...)

function handle_(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);

    // --- parse incoming payload (POST body or GET ?payload=...) ---
    var data = {};
    if (e && e.postData && e.postData.contents) data = JSON.parse(e.postData.contents);
    else if (e && e.parameter && e.parameter.payload) data = JSON.parse(e.parameter.payload);
    else if (e && e.parameter) data = e.parameter;

    // the entered values come bundled in data.details (the form's full data object)
    var det = {};
    try { det = JSON.parse(data.details || '{}'); } catch (err) { det = {}; }

    var ss  = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
    var tz  = ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone();
    var now = new Date();
    var sh  = ss.getSheetByName(LOG_SHEET) || ss.insertSheet(LOG_SHEET);

    // write the header row once
    if (sh.getLastRow() === 0) {
      sh.appendRow(COLS.map(function (c) { return c[0]; }));
      sh.setFrozenRows(1);
    }

    // generation count for this mobile number
    var mobile = (det.mobile || data.mobile || '').toString().trim();
    var genNo = 1, mobileCol = -1;
    for (var i = 0; i < COLS.length; i++) if (COLS[i][0] === 'Mobile') { mobileCol = i + 1; break; }
    if (mobile && mobileCol > 0 && sh.getLastRow() > 1) {
      var col = sh.getRange(2, mobileCol, sh.getLastRow() - 1, 1).getValues();
      for (var j = 0; j < col.length; j++) if ((col[j][0] || '').toString().trim() === mobile) genNo++;
    }

    // build the row — one value per column
    var row = COLS.map(function (c) {
      var src = c[1], key = c[2], v;
      if (src === 'now') return Utilities.formatDate(now, tz, 'dd-MM-yyyy HH:mm:ss');
      if (src === 'gen') return genNo;
      v = (src === 'data') ? data[key] : det[key];
      return (v === undefined || v === null) ? '' : v;
    });
    sh.appendRow(row);

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, generation: genNo }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}
