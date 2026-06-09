/**
 * VC&RI Namakkal — Livestock Project Report : Google Sheet logger  (Stage 4)
 * --------------------------------------------------------------------------
 * Writes EVERY report generation as ONE ROW, one value per column.
 *
 * Tabs created automatically on first write:
 *   • Master   — one row per generation across ALL species (common fields).
 *   • Cow      — full Cow / Dairy parameters per generation.
 *   • Goat     — full Goat parameters per generation.
 *   • Chicken  — full Native Chicken (Hen) parameters per generation.
 *
 * Every generation appends ONE row to Master AND ONE row to its species tab.
 *
 * Setup: Sheet ▸ Extensions ▸ Apps Script ▸ paste this ▸ Deploy as Web app
 *        (Execute as: Me · Who has access: Anyone). See SHEET-SETUP.md.
 */

var SHEET_ID   = '';          // blank = the spreadsheet this script is bound to
var MASTER_TAB = 'Master';

/* Column layout — [ Header shown in the sheet , source , key ]
   source:  'now'   = timestamp
            'gen'   = generation count for this mobile
            'sp'    = species (friendly name)
            'breed' = species-aware breed
            'unit'  = species-aware unit-size summary
            'data'  = from the summary payload (date/time/metrics)
            'det'   = the value the user entered (from details JSON)
            'json'  = full details object as JSON                              */

/* ── MASTER tab : common across every species ───────────────────────────── */
var MASTER_COLS = [
  ['Timestamp',                'now',  ''],
  ['Date',                     'data', 'date'],
  ['Time',                     'data', 'time'],
  ['Generation #',             'gen',  ''],
  ['Species',                  'sp',   ''],
  ['Name of Farmer / Unit',    'det',  'farmerName'],
  ['Mobile',                   'det',  'mobile'],
  ['E-mail',                   'det',  'email'],
  ['Aadhar',                   'det',  'aadhar'],
  ['Address of Farmer',        'det',  'farmerAddress'],
  ['Nature of Farming',        'det',  'natureFarming'],
  ['Farm Location',            'det',  'farmLocation'],
  ['Financing Bank',           'det',  'bankName'],
  ['Branch',                   'det',  'branchName'],
  ['Breed',                    'breed',''],
  ['Unit Size',                'unit', ''],
  ['Total Project Cost (Rs)',  'data', 'projectCost'],
  ['Bank Loan (Rs)',           'data', 'loan'],
  ['NPV (Rs)',                 'data', 'npv'],
  ['BCR',                      'data', 'bcr'],
  ['IRR (%)',                  'data', 'irr'],
  ['Avg DSCR',                 'data', 'dscr'],
  ['Details (JSON)',           'json', '']
];

/* ── shared head & tail used by every species tab ───────────────────────── */
var DETAIL_HEAD = [
  ['Timestamp',                'now',  ''],
  ['Date',                     'data', 'date'],
  ['Time',                     'data', 'time'],
  ['Generation #',             'gen',  ''],
  ['Name of Farmer / Unit',    'det',  'farmerName'],
  ['Mobile',                   'det',  'mobile'],
  ['E-mail',                   'det',  'email'],
  ['Aadhar',                   'det',  'aadhar'],
  ['Address of Farmer',        'det',  'farmerAddress'],
  ['Nature of Farming',        'det',  'natureFarming'],
  ['Farm Location',            'det',  'farmLocation'],
  ['Financing Bank',           'det',  'bankName'],
  ['Branch',                   'det',  'branchName']
];

var SUMMARY = [
  ['Total Project Cost (Rs)',  'data', 'projectCost'],
  ['Bank Loan (Rs)',           'data', 'loan'],
  ['NPV (Rs)',                 'data', 'npv'],
  ['BCR',                      'data', 'bcr'],
  ['IRR (%)',                  'data', 'irr'],
  ['Avg DSCR',                 'data', 'dscr']
];

/* ── COW / DAIRY parameters ─────────────────────────────────────────────── */
var COW_DET = [
  ['Animal Type',                     'det', 'C7'],
  ['Breed',                           'det', 'C8'],
  ['Unit Size - Batch 1',             'det', 'D9'],
  ['Unit Size - Batch 2',             'det', 'F9'],
  ['Total Animals',                   'det', 'C9'],
  ['Cost per Animal (Rs)',            'det', 'C10'],
  ['Milk Yield (L/day)',              'det', 'C11'],
  ['Insurance Premium (%)',           'det', 'C12'],
  ['Transport (Rs/animal)',           'det', 'C13'],
  ['Floor Space - Animal Shed (sqft)','det', 'C16'],
  ['Cost - Animal Shed (Rs/sqft)',    'det', 'C17'],
  ['Floor Space - Calves (sqft)',     'det', 'C18'],
  ['Cost - Calves Shed (Rs/sqft)',    'det', 'C19'],
  ['Floor Space - Store Room (sqft)', 'det', 'C20'],
  ['Cost - Store Room (Rs/sqft)',     'det', 'C21'],
  ['Feeder & Waterer (Rs/animal)',    'det', 'C24'],
  ['Milking Can (Rs)',                'det', 'C25'],
  ['Milking Machine (Rs)',            'det', 'C26'],
  ['Chaff Cutter (Rs)',               'det', 'C27'],
  ['Brush Cutter (Rs)',               'det', 'C28'],
  ['Feed Making Machinery (Rs)',      'det', 'C29'],
  ['Milk Processing Machinery (Rs)',  'det', 'C30'],
  ['Motor for Irrigation (Rs)',       'det', 'C31'],
  ['Sprinkler & Fogger (Rs)',         'det', 'C32'],
  ['Computers (Rs)',                  'det', 'C33'],
  ['CCTV (Rs)',                       'det', 'C34'],
  ['Floor Space - Labour Quarters',   'det', 'C37'],
  ['Cost - Labour Quarters',          'det', 'C38'],
  ['Overhead Tank (Rs)',              'det', 'C39'],
  ['Fencing (Rs)',                    'det', 'C40'],
  ['Land Levelling (Rs)',             'det', 'C41'],
  ['Borewell (Rs)',                   'det', 'C42'],
  ['Floor Space - Feed Mill',         'det', 'C43'],
  ['Cost - Feed Mill',                'det', 'C44'],
  ['Floor Space - Milk Proc. Unit',   'det', 'C45'],
  ['Cost - Milk Proc. Unit',          'det', 'C46'],
  ['Bio Gas Unit (Rs)',               'det', 'C47'],
  ['Vermicompost Unit (Rs)',          'det', 'C48'],
  ['Silage Making Unit (Rs)',         'det', 'C49'],
  ['Margin Money (%)',                'det', 'C52'],
  ['Rate of Interest (%)',            'det', 'C53'],
  ['Repayment Period (yrs)',          'det', 'C54'],
  ['Subsidy (%)',                     'det', 'C55'],
  ['Two-Wheelers (nos)',              'det', 'C58'],
  ['Cost per Two-Wheeler (Rs)',       'det', 'C59'],
  ['Four-Wheelers (nos)',             'det', 'C60'],
  ['Cost per Four-Wheeler (Rs)',      'det', 'C61'],
  ['Milking Period (days)',           'det', 'C64'],
  ['Dry Period (days)',               'det', 'C65'],
  ['Concentrate - Milking (kg/day)',  'det', 'C66'],
  ['Concentrate - Dry (kg/day)',      'det', 'C67'],
  ['Cost of Concentrate (Rs/kg)',     'det', 'C68'],
  ['Dry Fodder (kg/day)',             'det', 'C69'],
  ['Cost of Dry Fodder (Rs/kg)',      'det', 'C70'],
  ['Green Fodder Area (acres)',       'det', 'C71'],
  ['Green Fodder Yr1 (Rs/acre)',      'det', 'C72'],
  ['Green Fodder Yr2+ (Rs/acre)',     'det', 'C73'],
  ['Veterinary Aid (Rs/animal/yr)',   'det', 'C74'],
  ['Electricity (Rs/month)',          'det', 'C75'],
  ['Fuel (Rs/month)',                 'det', 'C76'],
  ['Depreciation - Buildings (%)',    'det', 'C77'],
  ['Depreciation - Equipment (%)',    'det', 'C78'],
  ['Miscellaneous (Rs/animal/yr)',    'det', 'C79'],
  ['Feed Making Unit (Rs/yr)',        'det', 'C80'],
  ['Milk Processing Unit (Rs/yr)',    'det', 'C81'],
  ['Vermicompost Unit (Rs/yr)',       'det', 'C82'],
  ['Silage Making Unit (Rs/yr)',      'det', 'C83'],
  ['No. of Farm Workers',             'det', 'C86'],
  ['Wages - Farm Worker (Rs/mo)',     'det', 'C87'],
  ['No. of Farm Managers',            'det', 'C88'],
  ['Salary - Farm Manager (Rs/mo)',   'det', 'C89'],
  ['No. of Drivers',                  'det', 'C90'],
  ['Salary - Driver (Rs/mo)',         'det', 'C91'],
  ['No. of Office Staff',             'det', 'C92'],
  ['Salary - Office Staff (Rs/mo)',   'det', 'C93'],
  ['Sale Price of Milk (Rs/L)',       'det', 'C96'],
  ['Sale Price Male Calf (Rs)',       'det', 'C97'],
  ['Value Female Calf (Rs)',          'det', 'C98'],
  ['Sale Price Manure (Rs)',          'det', 'C99'],
  ['Income Feed Making Unit (Rs/yr)', 'det', 'C100'],
  ['Income Milk Proc. Unit (Rs/yr)',  'det', 'C101']
];

/* ── GOAT parameters ────────────────────────────────────────────────────── */
var GOAT_DET = [
  ['Breed',                            'det', 'gBreed'],
  ['System of Rearing',                'det', 'gSystem'],
  ['Unit Size - Bucks',                'det', 'gBucks'],
  ['Unit Size - Does',                 'det', 'gDoes'],
  ['Cost per Buck (Rs)',               'det', 'gCostMale'],
  ['Cost per Doe (Rs)',                'det', 'gCostFemale'],
  ['Insurance Premium (%)',            'det', 'gInsurance'],
  ['Transport (Rs/animal)',            'det', 'gTransport'],
  ['Floor Space - Buck (sqft)',        'det', 'gSpaceBuck'],
  ['Floor Space - Doe (sqft)',         'det', 'gSpaceDoe'],
  ['Floor Space - Kid (sqft)',         'det', 'gSpaceKid'],
  ['Construction Cost - Shed (Rs/sqft)','det','gShedRate'],
  ['Floor Space - Store Room (sqft)',  'det', 'gStoreArea'],
  ['Cost - Store Room (Rs/sqft)',      'det', 'gStoreRate'],
  ['Equipment - Adult (Rs/animal)',    'det', 'gEquipAdult'],
  ['Equipment - Kid (Rs/kid)',         'det', 'gEquipKid'],
  ['No. of Chaff Cutters',             'det', 'gChaffNos'],
  ['Cost of Chaff Cutter (Rs)',        'det', 'gChaffCost'],
  ['Brush Cutter (Rs)',                'det', 'gBrush'],
  ['Mini Feed Making Machinery (Rs)',  'det', 'gFeedMachine'],
  ['Motor for Irrigation (Rs)',        'det', 'gMotor'],
  ['Sprinkler & Fogger (Rs)',          'det', 'gSprinkler'],
  ['Computers (Rs)',                   'det', 'gComputers'],
  ['CCTV (Rs)',                        'det', 'gCCTV'],
  ['Floor Space - Labour Quarters (sqft)','det','gLqArea'],
  ['Cost - Labour Quarters (Rs/sqft)', 'det', 'gLqRate'],
  ['Overhead / Water Tank (Rs)',       'det', 'gTank'],
  ['Fencing (Rs)',                     'det', 'gFencing'],
  ['Land Levelling (Rs)',              'det', 'gLandLevel'],
  ['Borewell (Rs)',                    'det', 'gBorewell'],
  ['Margin Money (%)',                 'det', 'gMargin'],
  ['Rate of Interest (%)',             'det', 'gInterest'],
  ['Repayment Period (yrs)',           'det', 'gRepay'],
  ['Subsidy (%)',                      'det', 'gSubsidy'],
  ['Two-Wheelers (nos)',               'det', 'gTwoW'],
  ['Cost per Two-Wheeler (Rs)',        'det', 'gTwoWCost'],
  ['Four-Wheelers (nos)',              'det', 'gFourW'],
  ['Cost per Four-Wheeler (Rs)',       'det', 'gFourWCost'],
  ['Concentrate - Buck (kg/day)',      'det', 'gConcBuck'],
  ['Concentrate - Doe (kg/day)',       'det', 'gConcDoe'],
  ['Concentrate - Kid (kg/day)',       'det', 'gConcKid'],
  ['Cost of Concentrate (Rs/kg)',      'det', 'gConcCost'],
  ['Dry Fodder (Rs/yr)',               'det', 'gDryFodder'],
  ['Green Fodder Area (acres)',        'det', 'gGreenArea'],
  ['Green Fodder Yr1 (Rs/acre)',       'det', 'gGreenY1'],
  ['Green Fodder Yr2+ (Rs/acre)',      'det', 'gGreenY2'],
  ['Veterinary Aid - Adult (Rs/yr)',   'det', 'gVetAdult'],
  ['Veterinary Aid - Kid (Rs/yr)',     'det', 'gVetKid'],
  ['Electricity (Rs/month)',           'det', 'gElectricity'],
  ['Fuel (Rs/month)',                  'det', 'gFuel'],
  ['Miscellaneous (Rs/adult/yr)',      'det', 'gMisc'],
  ['Feed Making Unit (Rs/yr)',         'det', 'gFeedUnit'],
  ['Depreciation - Buildings (%)',     'det', 'gDepBuilding'],
  ['Depreciation - Equipment (%)',     'det', 'gDepEquip'],
  ['Kidding Interval (months)',        'det', 'gKiddingInterval'],
  ['Number of Kiddings (/doe/yr)',     'det', 'gKiddings'],
  ['Average Litter Size',              'det', 'gLitter'],
  ['Conception Rate (%)',              'det', 'gConception'],
  ['Kidding Percentage (%)',           'det', 'gKiddingPct'],
  ['Mortality of Kids (%)',            'det', 'gKidMortality'],
  ['Saleable Age of Kids (months)',    'det', 'gSaleAge'],
  ['Replacement Stock - Male',         'det', 'gReplMale'],
  ['Replacement Stock - Female',       'det', 'gReplFemale'],
  ['Culling Percentage (%)',           'det', 'gCullingPct'],
  ['Adult Males Culled',               'det', 'gCulledMale'],
  ['Adult Females Culled',             'det', 'gCulledFemale'],
  ['No. of Farm Workers',              'det', 'gWorkers'],
  ['Wages - Farm Worker (Rs/mo)',      'det', 'gWage'],
  ['No. of Farm Managers',             'det', 'gManagers'],
  ['Salary - Farm Manager (Rs/mo)',    'det', 'gMgrSal'],
  ['No. of Drivers',                   'det', 'gDrivers'],
  ['Salary - Driver (Rs/mo)',          'det', 'gDrvSal'],
  ['No. of Office Staff',              'det', 'gOffice'],
  ['Salary - Office Staff (Rs/mo)',    'det', 'gOffSal'],
  ['Sale Price - Male Kids (Rs)',      'det', 'gPriceMale'],
  ['Sale Price - Female Kids (Rs)',    'det', 'gPriceFemale'],
  ['Sale Price - Culled Animal (Rs)',  'det', 'gCulledPrice'],
  ['Income from Manure (Rs/animal/yr)','det', 'gManure']
];

/* ── HEN / NATIVE CHICKEN parameters ────────────────────────────────────── */
var HEN_DET = [
  ['Breed',                            'det', 'hBreed'],
  ['Chicks per Batch',                 'det', 'hBirds'],
  ['Batches at a Time in Shed',        'det', 'hBatchesAtTime'],
  ['Batches Introduced - Year 1',      'det', 'hBatchesY1'],
  ['Batches Introduced - Year 2+',     'det', 'hBatchesY2'],
  ['Rearing Period (weeks)',           'det', 'hRearWeeks'],
  ['Construction Period (months)',     'det', 'hConstrMonths'],
  ['Cost of Chicks (Rs/chick)',        'det', 'hChickCost'],
  ['Mortality (%)',                    'det', 'hMortality'],
  ['Floor Space (sqft/bird)',          'det', 'hFloorSpace'],
  ['Construction Cost (Rs/sqft)',      'det', 'hConstrCost'],
  ['Equipment - Feeder & Waterer (Rs/bird)','det','hEquipBird'],
  ['Margin Money (%)',                 'det', 'hMargin'],
  ['Rate of Interest (%)',             'det', 'hInterest'],
  ['Repayment Period (yrs)',           'det', 'hRepay'],
  ['Subsidy (%)',                      'det', 'hSubsidy'],
  ['Concentrate Feed (kg/bird)',       'det', 'hFeedPerBird'],
  ['Concentrate Feed Cost (Rs/kg)',    'det', 'hFeedCost'],
  ['Deep Litter Material (Rs/chick)',  'det', 'hDeepLitter'],
  ['Veterinary Expenses (Rs/bird)',    'det', 'hVetBird'],
  ['Electricity (Rs/month)',           'det', 'hElectricity'],
  ['Miscellaneous (Rs/bird)',          'det', 'hMisc'],
  ['Depreciation - Building (%)',      'det', 'hDepBuilding'],
  ['Depreciation - Equipment (%)',     'det', 'hDepEquip'],
  ['No. of Labour Required',           'det', 'hWorkers'],
  ['Labour Wages (Rs/month)',          'det', 'hWage'],
  ['Sale Price of Bird (Rs/kg)',       'det', 'hSalePrice'],
  ['Average Weight at Marketing (kg)', 'det', 'hWeight'],
  ['Income from Manure (Rs/bird)',     'det', 'hManure']
];

/* species code (from the website) → { tab name , full column layout } */
var SPECIES_TABS = {
  Cow:  { tab: 'Cow',     cols: DETAIL_HEAD.concat(COW_DET,  SUMMARY) },
  Goat: { tab: 'Goat',    cols: DETAIL_HEAD.concat(GOAT_DET, SUMMARY) },
  Hen:  { tab: 'Chicken', cols: DETAIL_HEAD.concat(HEN_DET,  SUMMARY) }
};

/* ── entry points ───────────────────────────────────────────────────────── */
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

    // species: prefer the explicit field, fall back to the one stored in details
    var species = (data.species || det.species || 'Cow').toString();

    var ss  = SHEET_ID ? SpreadsheetApp.openById(SHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
    var tz  = ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone();
    var now = new Date();

    // ensure the Master tab + header exist (also the per-mobile counter source)
    var master = ss.getSheetByName(MASTER_TAB) || ss.insertSheet(MASTER_TAB);
    ensureHeader_(master, MASTER_COLS);

    var mobile = (det.mobile || data.mobile || '').toString().trim();
    var genNo  = genCount_(master, mobile);

    var ctx = { now: now, tz: tz, genNo: genNo, species: species, data: data, det: det };

    // 1) Master row (common fields, all species together)
    master.appendRow(buildRow_(MASTER_COLS, ctx));

    // 2) species-specific tab row (full parameters)
    var spec = SPECIES_TABS[species] || SPECIES_TABS.Cow;
    var sh = ss.getSheetByName(spec.tab) || ss.insertSheet(spec.tab);
    ensureHeader_(sh, spec.cols);
    sh.appendRow(buildRow_(spec.cols, ctx));

    return ContentService
      .createTextOutput(JSON.stringify({ ok: true, species: speciesName_(species), tab: spec.tab, generation: genNo }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ ok: false, error: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  } finally {
    try { lock.releaseLock(); } catch (e2) {}
  }
}

/* ── helpers ─────────────────────────────────────────────────────────────── */
function ensureHeader_(sheet, cols) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(cols.map(function (c) { return c[0]; }));
    sheet.setFrozenRows(1);
  }
}

function buildRow_(cols, ctx) {
  return cols.map(function (c) {
    var src = c[1], key = c[2], v;
    if (src === 'now')   return Utilities.formatDate(ctx.now, ctx.tz, 'dd-MM-yyyy HH:mm:ss');
    if (src === 'gen')   return ctx.genNo;
    if (src === 'sp')    return speciesName_(ctx.species);
    if (src === 'breed') return breed_(ctx.species, ctx.det);
    if (src === 'unit')  return unitText_(ctx.species, ctx.det);
    if (src === 'json')  return JSON.stringify(ctx.det);
    v = (src === 'data') ? ctx.data[key] : ctx.det[key];
    return (v === undefined || v === null) ? '' : v;
  });
}

// running count of reports for this mobile (counted from the Master tab)
function genCount_(master, mobile) {
  if (!mobile) return 1;
  var mc = -1;
  for (var i = 0; i < MASTER_COLS.length; i++) {
    if (MASTER_COLS[i][2] === 'mobile') { mc = i + 1; break; }
  }
  var n = 1;
  if (mc > 0 && master.getLastRow() > 1) {
    var col = master.getRange(2, mc, master.getLastRow() - 1, 1).getValues();
    for (var j = 0; j < col.length; j++) {
      if ((col[j][0] || '').toString().trim() === mobile) n++;
    }
  }
  return n;
}

function speciesName_(species) {
  if (species === 'Hen')  return 'Chicken';
  if (species === 'Goat') return 'Goat';
  return 'Cow';
}

function breed_(species, det) {
  if (species === 'Goat') return det.gBreed || '';
  if (species === 'Hen')  return det.hBreed || '';
  return det.C8 || '';
}

function unitText_(species, det) {
  var n;
  if (species === 'Goat') {
    n = (+det.gBucks || 0) + (+det.gDoes || 0);
    return n + ' (Bucks ' + (det.gBucks || 0) + ' / Does ' + (det.gDoes || 0) + ')';
  }
  if (species === 'Hen') {
    return (det.hBirds || 0) + ' birds/batch';
  }
  n = det.C9 || ((+det.D9 || 0) + (+det.F9 || 0));
  return n + ' ' + (det.C7 || 'animals');
}
