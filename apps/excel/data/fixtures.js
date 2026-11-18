/* =============================================================================
   Excel Trainer — data/fixtures.js
   The practice workbooks.

   Everything in here is GENERATED AND SYNTHETIC. No value from any real
   workbook appears in this trainer. What is copied from the real files is their
   STRUCTURE, because that is what has to be learned:

     - column positions           MIMS SAVE FILE really does key on column M,
                                  with a second key in V and a K->L chain
     - key formats                a scanner barcode, a product code, a site code
     - sheet names                Samplesin, scanMODWB, FILTERWBTODAY,
                                  'MIMS SAVE FILE', FINALEMAIL WB, Calendars
     - the five layers            landing, mirror, filter, workings, output
     - and the dirt               trailing spaces, CHAR(160), numbers stored as
                                  text, zero padding, mixed date formats, blank
                                  rows, export padding zeros

   The dirt is not decoration. Catalog §12 is explicit: "the sandbox paste-blocks
   should include trailing spaces, CHAR(160), numbers stored as text,
   zero-padding, and inconsistent date formats. The real exports do."

   Generation is deterministic from a seed, so an exercise grades the same every
   time and a hint can refer to "row 7" and still be right tomorrow.
   ============================================================================= */
(function (XT) {
  'use strict';
  XT.data = XT.data || {};
  var D = XT.D;

  /* -------------------------------------------------------- deterministic RNG */

  // mulberry32: tiny, fast, and repeatable across browsers.
  function rng(seed) {
    var a = seed >>> 0;
    return function () {
      a = (a + 0x6D2B79F5) >>> 0;
      var t = Math.imul(a ^ (a >>> 15), 1 | a);
      t = (t + Math.imul(t ^ (t >>> 7), 61 | t)) ^ t;
      return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    };
  }
  function pick(r, list) { return list[Math.floor(r() * list.length)]; }
  function int(r, lo, hi) { return lo + Math.floor(r() * (hi - lo + 1)); }

  /* ------------------------------------------------------------------- dirt -- */

  var NBSP = ' ';

  var dirt = {
    // A number that is really text. The single most common real-world trap.
    asText: function (n) { return { v: String(n) }; },
    // Trailing space: invisible, and it breaks every exact-match lookup.
    trail: function (s) { return { v: s + ' ' }; },
    lead: function (s) { return { v: ' ' + s }; },
    // A non-breaking space, as pasted out of a web page or a PLU export.
    nbsp: function (s) { return { v: String(s) + NBSP }; },
    // Zero-padded code that must stay text or it loses its zeros.
    padded: function (n, width) {
      var s = String(n);
      while (s.length < width) s = '0' + s;
      return { v: s };
    }
  };
  XT.data.dirt = dirt;

  /* --------------------------------------------------------------- the world */

  // Invented products, sites and tests. Deliberately generic.
  var PRODUCTS = [
    ['Breast Fillet 300g', 'FILLET', 300], ['Breast Fillet 600g', 'FILLET', 600],
    ['Thigh Fillet 400g', 'THIGH', 400], ['Drumstick 500g', 'DRUM', 500],
    ['Whole Bird 1.4kg', 'WHOLE', 1400], ['Whole Bird 1.8kg', 'WHOLE', 1800],
    ['Diced Breast 350g', 'DICED', 350], ['Mini Fillet 300g', 'MINI', 300],
    ['Wing 700g', 'WING', 700], ['Leg Quarter 900g', 'LEGQ', 900],
    ['Butterfly 450g', 'BFLY', 450], ['Escalope 250g', 'ESCAL', 250]
  ];
  var RETAILERS = ['Northgate', 'Westvale', 'Harbour', 'Greenfield', 'Lowry'];
  var LINES = ['L1', 'L2', 'L3'];
  var SHIFTS = ['AM', 'PM', 'NIGHT'];
  var SITES = ['SCU', 'FLX', 'WIL', 'BNH'];
  var TESTS = ['TVC', 'ENT', 'ECOLI', 'SALM', 'CAMPY', 'LISTERIA'];
  var CATEGORIES = ['Finished Product', 'Raw Material', 'Environmental Swab', 'Water'];
  var OPERATORS = ['op.a', 'op.b', 'op.c', 'op.d'];

  /* Barcodes look like a 13-digit EAN. A PURE function of the row index, so
     every sheet that mentions row i produces the same barcode -- which is what
     makes the joins in the exercises actually join. */
  function barcode(i) { return '50' + String(100000000000 + i * 7919).slice(1, 12); }

  function productCode(i) { return 'P' + String(1000 + i); }

  /* =========================================================== sheet builders */

  /* MIMS SAVE FILE -- 37 columns, A..AK.
     From the census: the key is column M (113,022 of its 116,892 references are
     as a lookup key), with a second key in V used for the two-key join, and a
     K -> L chain where L is looked up by K rather than by M.

     Column meanings beyond that are INFERRED from how the formulas use them --
     see docs/DATA-DICTIONARY.md, which marks every column inferred or unknown. */
  function mimsSaveFile(seed, n) {
    var r = rng(seed), rows = [];
    rows.push(['PRODUCT', 'MATERIAL', 'RISE', 'UNIT', 'DESCRIPTN', 'SIZE', 'RETAILER',
      'SHIFT', 'LINE', 'OTHER_ID', 'TEST_GROUP', 'TEST_NAME', 'BARCODE', 'N1', 'N2', 'N3',
      'N4', 'SPEC_REF', 'SITE', 'DUE_DAYS', 'SUB_TYPE', 'SHIFT_KEY', 'W1', 'X1',
      'LAB', 'F1', 'F2', 'F3', 'F4', 'F5', 'F6', 'F7', 'F8', 'F9', 'F10', 'F11', 'F12']);
    for (var i = 0; i < n; i++) {
      var p = PRODUCTS[i % PRODUCTS.length];
      var shift = SHIFTS[i % SHIFTS.length];
      var bc = barcode(i);
      var test = TESTS[i % TESTS.length];
      var row = new Array(37).fill(null);
      row[0] = p[0];                                   // A PRODUCT
      row[1] = p[1];                                   // B MATERIAL
      row[2] = int(r, 1, 4);                           // C RISE
      row[3] = 'g';                                    // D UNIT
      row[4] = p[0] + ' ' + pick(r, RETAILERS);        // E DESCRIPTN
      row[5] = p[2];                                   // F SIZE
      row[6] = RETAILERS[i % RETAILERS.length];        // G RETAILER
      row[7] = shift;                                  // H SHIFT
      row[8] = LINES[i % LINES.length];                // I LINE
      row[9] = productCode(i) + '-' + shift;           // J OTHER_ID (two-key result)
      row[10] = 'G' + (1 + (i % 4));                   // K TEST_GROUP
      row[11] = test;                                  // L TEST_NAME (found via K)
      // M the KEY. Deliberately mixed: most are text, a few slipped in as numbers,
      // which is exactly the fault that makes a lookup fail on three rows only.
      row[12] = (i % 11 === 5) ? Number(bc) : { v: bc };
      row[17] = 'SPEC-' + (200 + (i % 9));             // R SPEC_REF
      row[18] = SITES[i % SITES.length];               // S SITE
      row[19] = int(r, 2, 10);                         // T DUE_DAYS
      row[20] = CATEGORIES[i % CATEGORIES.length];     // U SUB_TYPE
      row[21] = shift;                                 // V SHIFT_KEY (the 2nd key)
      row[24] = 'Lab ' + pick(r, ['North', 'South']);  // Y LAB
      // Z..AK: test flags, padded with zeros by the export -- which is why the
      // real formulas wrap them in IF(...=0,"") to blank them out.
      for (var k = 25; k < 37; k++) row[k] = (r() < 0.25) ? 1 : 0;
      rows.push(row);
    }
    return rows;
  }

  /* Samplesin -- what the barcode scanner writes. 12 columns, A..L.
     Column I is a full timestamp (date AND time) because the shift cut-off
     logic depends on the time part; column L is the device. */
  function samplesIn(seed, n, day) {
    var r = rng(seed), rows = [];
    rows.push(['Change', 'Barcode', 'Name', 'Description', 'Quantity', 'Category',
      'Notes', 'Location', 'Date', 'Changed By', 'Changed On', 'Changed Using']);
    for (var i = 0; i < n; i++) {
      var src = i % 14;                   // which MIMS row this scan refers to
      var bc = barcode(src);
      var p = PRODUCTS[src % PRODUCTS.length];
      var hh = int(r, 4, 20), mm = int(r, 0, 59);
      var stamp = day + D.hmsToSerial(hh, mm, int(r, 0, 59));
      var row = [
        'Add',
        // The scanner writes the barcode as TEXT, always. MIMS has a few as
        // numbers. That mismatch is the lesson in level 4.
        { v: bc },
        p[0],
        // A description with a trailing space on some rows: invisible, and it
        // breaks an exact-match lookup every time.
        (i % 7 === 3) ? dirt.trail(p[0]) : p[0],
        // Quantity: mostly real numbers. Some arrive as text, and one row has a
        // non-breaking space stuck on the end, which no amount of TRIM will shift.
        (i % 9 === 4) ? dirt.asText(int(r, 1, 6))
          : (i % 13 === 7 ? dirt.nbsp(int(r, 1, 6)) : int(r, 1, 6)),
        CATEGORIES[src % CATEGORIES.length],
        (i % 5 === 0) ? 'recheck' : null,
        'Line ' + LINES[src % LINES.length],
        { v: stamp, fmt: 'dd/mm/yyyy hh:mm:ss' },
        OPERATORS[i % OPERATORS.length],
        { v: day, fmt: 'dd/mm/yyyy' },
        pick(r, ['Scanner 1', 'Scanner 2', 'Handheld'])
      ];
      rows.push(row);
      // A blank row partway down, as the export sometimes leaves.
      if (i === 8) rows.push(new Array(12).fill(null));
    }
    return rows;
  }

  /* Product Lookup -- a small reference table with the KEY IN COLUMN A.

     This exists so that VLOOKUP can be taught before XLOOKUP. MIMS SAVE FILE
     keys on column M and the useful columns are to its LEFT, which VLOOKUP
     simply cannot do -- and that is the honest reason the real workbooks moved
     to XLOOKUP. Teaching VLOOKUP on a table it can actually handle, then
     meeting the wall, is the right order.

     One row has its barcode stored as a NUMBER instead of text. Everything
     else is text. That single row produces an #N/A that looks impossible until
     you check it with ISNUMBER -- which is the whole of levels 4 and 5. */
  function productLookup(n) {
    var rows = [['BARCODE', 'PRODUCT', 'MATERIAL', 'TEST', 'DUE_DAYS', 'RETAILER']];
    for (var i = 0; i < n; i++) {
      var p = PRODUCTS[i % PRODUCTS.length];
      rows.push([
        (i % 11 === 5) ? Number(barcode(i)) : { v: barcode(i) },
        p[0], p[1], TESTS[i % TESTS.length], 2 + (i % 7), RETAILERS[i % RETAILERS.length]
      ]);
    }
    return rows;
  }

  /* Calendars -- 11 columns A..K. The census says:
       A  composite key, built as F&B, used by VLOOKUP 102,579 times
       F  week label         G  week start      H  week end (and itself a key)
       I,J,K  the things looked up: period, label, year
     This is the sheet that exists because retailer weeks are not ISO weeks. */
  function calendars(startDate, weeks) {
    var rows = [['KEY', 'YEAR', 'PERIOD', 'QTR', 'MONTH', 'WEEK', 'START', 'END',
      'PERIOD_LABEL', 'WEEK_LABEL', 'LOOKUP']];
    // Walk back to the Saturday on or before startDate, then forward week by week.
    var d = startDate - (D.weekday(startDate, 1) - 1);
    for (var i = 0; i < weeks; i++) {
      var start = d + i * 7;
      var end = start + 6;
      var wk = i + 1;
      var period = Math.floor(i / 4) + 1;
      var label = 'WK' + (wk < 10 ? '0' + wk : wk);
      rows.push([
        { f: '=F' + (i + 2) + '&B' + (i + 2) },        // A composite key
        2026,                                           // B YEAR
        period,                                          // C PERIOD
        Math.floor(i / 13) + 1,                          // D QTR
        { f: '=MONTH(G' + (i + 2) + '+2)' },            // E MONTH (the +2 nudge)
        label,                                           // F WEEK
        { v: start, fmt: 'dd/mm/yyyy' },                 // G START
        { v: end, fmt: 'dd/mm/yyyy' },                   // H END
        'P' + period,                                    // I PERIOD_LABEL
        label + ' 2026',                                 // J WEEK_LABEL
        'P' + period + '/' + label                       // K LOOKUP
      ]);
    }
    return rows;
  }

  /* picklists -- one column per dropdown, exposed as defined names. 26 of the
     152 workbooks have exactly this sheet. */
  function picklists() {
    var cols = [
      ['Failure_type', ['Foreign body', 'Label error', 'Short weight', 'Pack damage',
        'Date code', 'Seal failure', 'Temperature', 'Taint']],
      ['Regulatory_Microbiology', ['Salmonella', 'Campylobacter', 'Listeria']],
      ['Raw_Material_Microbiology', ['TVC', 'Enterobacteriaceae', 'E.coli']],
      ['Finished_Product_Microbiology', ['TVC', 'ENT', 'ECOLI', 'SALM', 'LISTERIA']],
      ['Environmental_Swabbing', ['Floor', 'Drain', 'Wall', 'Conveyor', 'Hand contact',
        'Blade', 'Chute']],
      ['Water', ['Mains in', 'Chiller', 'Washdown', 'Ice']],
      ['Shift', ['AM', 'PM', 'NIGHT']],
      ['Answer', ['Yes', 'No', 'N/A']]
    ];
    var height = 1;
    cols.forEach(function (c) { height = Math.max(height, c[1].length + 1); });
    var rows = [];
    for (var r = 0; r < height; r++) {
      var row = [];
      cols.forEach(function (c) { row.push(r === 0 ? c[0] : (c[1][r - 1] || null)); });
      rows.push(row);
    }
    return { rows: rows, cols: cols };
  }

  /* ========================================================= S1 -- Sample Hub */

  /* The gentlest scenario, and the one the catalog recommends starting with:
     one key, many columns. Scanner output lands, is mirrored, filtered to
     today, enriched 37 columns wide from MIMS SAVE FILE, then presented. */
  function buildS1(opts) {
    opts = opts || {};
    var day = opts.today || D.ymdToSerial(2026, 11, 13);
    var seed = opts.seed || 20261113;
    var pl = picklists();

    var def = {
      sheets: [
        { name: 'Home', layer: 5, note: 'The front panel the operator sees.',
          rows: [
            ['Sample Hub — daily submission'],
            [],
            ['Today', { f: '=TODAY()', fmt: 'dddd dd/mm/yyyy' }],
            ['Shift cut-off', { v: D.hmsToSerial(6, 0, 0), fmt: 'hh:mm:ss' }],
            ['Scans landed', { f: '=COUNTA(Samplesin!B2:B60)' }],
            ['Matched to MIMS', null],
            ['Unmatched', null],
            ['Default due days', 5],
            [],
            ['Tests by retailer'],
            [null, 'Northgate', 'Westvale', 'Harbour'],
            ['TVC', null, null, null],
            ['ENT', null, null, null],
            ['ECOLI', null, null, null],
            [],
            ['Quantity by category'],
            ['Finished Product', null],
            ['Raw Material', null],
            ['Environmental Swab', null],
            ['Water', null],
            [],
            ['Scans today', null],
            ['Scratch', null]
          ],
          colWidth: { 0: 160, 1: 130 } },

        { name: 'Samplesin', layer: 1,
          note: 'LANDING. What the scanner wrote. Never edit this sheet.',
          rows: samplesIn(seed, 22, day) },

        { name: 'MIMS SAVE FILE', layer: 1,
          note: 'REFERENCE. The product master. Key is column M — so the ' +
            'useful columns are to its LEFT, and VLOOKUP cannot reach them.',
          rows: mimsSaveFile(seed, 14) },

        { name: 'Product Lookup', layer: 1,
          note: 'REFERENCE. A tidy table with the key in column A. ' +
            'One barcode in here is a number instead of text.',
          rows: productLookup(14) },

        { name: 'scanMODWB', layer: 2,
          note: 'MIRROR. A clean 1:1 copy of the landing sheet. You build this.',
          rows: [['Change', 'Barcode', 'Name', 'Description', 'Quantity', 'Category',
            'Notes', 'Location', 'Date', 'Changed By', 'Changed On', 'Changed Using']] },

        { name: 'FILTERWBTODAY', layer: 3,
          note: 'FILTER. Only today’s scans, with the shift and due date worked out.',
          rows: [['Row', 'Barcode', 'Name', 'Qty', 'Stamp', 'Shift', 'Due']],
          colWidth: { 1: 130, 4: 150 } },

        { name: 'MIMS DAILY WB OLD', layer: 4,
          note: 'WORKINGS. The record rebuilt beside each barcode.',
          rows: [['KEY', 'PRODUCT (VLOOKUP)', 'DUE DAYS', 'BROKEN ON PURPOSE',
            'PRODUCT', 'MATERIAL', 'RISE', 'UNIT', 'DESCRIPTN', 'SIZE', 'RETAILER',
            'SHIFT', 'OTHER_ID', 'IN LOOKUP?', 'REGION', 'SIGNATURE', 'VERDICT']],
          colWidth: { 0: 130, 1: 150, 8: 170, 15: 150, 16: 200 } },

        { name: 'FINALEMAIL WB', layer: 5,
          note: 'OUTPUT. What gets pasted into the email to the lab.',
          rows: [['Barcode', 'Seq key', 'Scan time', 'Product', 'Test', 'Note']],
          colWidth: { 0: 130, 1: 130, 2: 150 } },

        { name: 'REF', layer: 1,
          note: 'REFERENCE. Retailer to region — the table that replaces a ' +
            'nested IF ladder.',
          rows: [['RETAILER', 'REGION'],
            ['Northgate', 'North'], ['Westvale', 'West'], ['Harbour', 'Coast'],
            ['Greenfield', 'Inland'], ['Lowry', 'Inland']] },

        { name: 'Calendars', layer: 1, hidden: false,
          note: 'REFERENCE. Retailer weeks. They are not ISO weeks.',
          rows: calendars(D.ymdToSerial(2026, 1, 3), 56) },

        { name: 'picklists', layer: 1, hidden: true,
          note: 'REFERENCE. One column per dropdown.',
          rows: pl.rows }
      ],
      names: {},
      tables: []
    };

    pl.cols.forEach(function (c, i) {
      var col = XT.R.numToCol(i);
      def.names[c[0]] = 'picklists!$' + col + '$2:$' + col + '$' + (c[1].length + 1);
    });
    return def;
  }

  /* ================================================ S4 -- vision log parsing */

  /* One text column, one line per pack, and it has to become typed columns.
     Both the old way (count characters) and the new way (count fields) are
     possible on the same data, which is the point. */
  function buildS4(opts) {
    opts = opts || {};
    var seed = opts.seed || 4477;
    var r = rng(seed);
    var rows = [['TS', 'LINE', 'MESSAGE']];
    var jobs = [7781, 7782, 7783, 7784, 7785];
    var day = opts.today || D.ymdToSerial(2026, 11, 13);

    jobs.forEach(function (job, ji) {
      var prod = PRODUCTS[ji % PRODUCTS.length];
      var code = productCode(ji);
      // A JOB header line, then a run of detail lines, then a verdict line.
      rows.push([{ v: day + D.hmsToSerial(6 + ji, ji * 7, 0), fmt: 'hh:mm:ss' },
        'PACKLINE ' + (1 + ji % 2),
        'JOB ' + job + ' PACKLINE ' + (1 + ji % 2) + ' SLEEVE CHECK ' + code]);
      var packs = int(r, 3, 6);
      for (var k = 0; k < packs; k++) {
        var checked = int(r, 40, 120);
        var failed = (r() < 0.3) ? int(r, 1, 4) : 0;
        rows.push([{ v: day + D.hmsToSerial(6 + ji, ji * 7 + k + 1, 0), fmt: 'hh:mm:ss' },
          'PACKLINE ' + (1 + ji % 2),
          'PACK ' + code + ' CHECKED ' + checked + ' PACKS FAILED ' + failed + ' PACKS']);
      }
      var verdict = (r() < 0.35)
        ? 'SLEEVE VERIFY FLASHLABEL; Invalid: Mismatch'
        : ((r() < 0.5) ? 'SLEEVE VERIFY RESULT; Invalid: Mismatch' : 'SLEEVE VERIFY RESULT; Valid: OK');
      rows.push([{ v: day + D.hmsToSerial(6 + ji, ji * 7 + packs + 1, 0), fmt: 'hh:mm:ss' },
        'PACKLINE ' + (1 + ji % 2), verdict]);
    });

    return {
      sheets: [
        { name: 'Home', layer: 5, note: 'The verification summary.',
          rows: [['Line verification — sleeve check'], [],
            ['Lines in log', { f: '=COUNTA(LVALL!C2:C60)' }],
            ['Mismatches', null], ['Missing sleeves', null]],
          colWidth: { 0: 170, 1: 120 } },
        { name: 'LVALL', layer: 1,
          note: 'LANDING. The vision system log. One text line per event.',
          rows: rows, colWidth: { 2: 460 } },
        { name: 'Sheet4', layer: 2,
          note: 'WORKINGS. Turn the log into columns. Old way and new way.',
          rows: [['Carried code', 'Line', 'Verdict', 'Job line', 'Job code',
            'Detail line', 'Field 4', 'Checked', 'Field 6', 'Failed']],
          colWidth: { 0: 90, 3: 300, 5: 300 } },
        { name: 'Home Config', layer: 1, hidden: true,
          note: 'REFERENCE. Products and their packs-per-case.',
          rows: [['CODE', 'PRODUCT', 'PACKS']].concat(PRODUCTS.slice(0, 5).map(function (p, i) {
            return [productCode(i), p[0], int(r, 6, 24)];
          })) }
      ]
    };
  }

  /* ============================================= S3 -- the micro week grid */

  function buildS3(opts) {
    opts = opts || {};
    var seed = opts.seed || 909;
    var r = rng(seed);
    var day = opts.today || D.ymdToSerial(2026, 11, 13);

    // ECOLI MICRO PASTE: the census says A,B,C,E,F,G,H,K,M then a long gap to
    // AB,AC,AD,AH and the CF/CG/CK/DD/DH cluster. We reproduce the shape with
    // the columns the exercises use, and leave the gap honest.
    var rows = [new Array(30).fill(null)];
    rows[0][0] = 'SAMPLE_REF'; rows[0][1] = 'SITE'; rows[0][2] = 'SAMPLE_DATE';
    rows[0][4] = 'PRODUCT'; rows[0][5] = 'MATERIAL'; rows[0][6] = 'LINE';
    rows[0][7] = 'SHIFT'; rows[0][10] = 'TEST'; rows[0][12] = 'OPERATOR';
    rows[0][27] = 'LOD'; rows[0][28] = 'RESULT_RAW'; rows[0][29] = 'RESULT';

    for (var i = 0; i < 48; i++) {
      var row = new Array(30).fill(null);
      var when = day - int(r, 0, 55);
      row[0] = 'S' + (40000 + i);
      row[1] = SITES[i % SITES.length];
      row[2] = { v: when, fmt: 'dd/mm/yyyy' };
      row[4] = PRODUCTS[i % PRODUCTS.length][0];
      row[5] = PRODUCTS[i % PRODUCTS.length][1];
      row[6] = LINES[i % LINES.length];
      row[7] = SHIFTS[i % SHIFTS.length];
      row[10] = 'ECOLI';
      row[12] = OPERATORS[i % OPERATORS.length];
      // Below the limit of detection the lab reports "<10", as TEXT. The real
      // formula floors it at 10 with IF(AB<=10,10,AH) -- pattern in S3.
      var below = r() < 0.3;
      row[27] = below ? 10 : int(r, 11, 20);
      row[28] = below ? { v: '<10' } : int(r, 11, 4000);
      row[29] = below ? 10 : row[28];
      rows.push(row);
    }

    return {
      sheets: [
        { name: 'Current month', layer: 5, note: 'The trend the vet reads.',
          rows: [['E.coli trending'], [],
            ['Week ending', null, 'Samples', 'Over limit', '% over'], []],
          colWidth: { 0: 130 } },
        { name: 'ECOLI MICRO PASTE', layer: 1,
          note: 'LANDING. The lab export. Note the "<10" results are TEXT.',
          rows: rows },
        { name: 'Ecoli Final Data', layer: 4,
          note: 'WORKINGS. Reorder, floor the result, attach the week.',
          rows: [['SAMPLE_REF', null, 'SAMPLE_DATE', 'PRODUCT', 'MATERIAL', 'LINE',
            'SHIFT', 'TEST', null, null, 'OPERATOR', null, null, null, null, null,
            null, null, null, null, 'RESULT', 'WEEK', 'WK TOTAL', 'WK MEAN', 'OVER?',
            'WK RATE', 'SITE', 'WK KEY']] },
        { name: 'Calendars', layer: 1,
          note: 'REFERENCE. Week start and end pairs, for the boundary lookup.',
          rows: calendars(D.ymdToSerial(2026, 1, 3), 56) }
      ]
    };
  }

  /* ========================================== S2 -- complaints, fallback keys */

  function buildS2(opts) {
    opts = opts || {};
    var seed = opts.seed || 1212;
    var r = rng(seed);
    var day = opts.today || D.ymdToSerial(2026, 11, 13);

    // Volumes Paste: three key columns A, B, C -- three different code systems,
    // which is precisely why the complaint join needs a fallback cascade.
    // Dates run ACROSS row 13, which is why SUMIFS works on rows here.
    var vol = [];
    for (var i = 0; i < 20; i++) vol.push(new Array(14).fill(null));
    vol[0][0] = 'RETAILER_CODE'; vol[0][1] = 'EAN'; vol[0][2] = 'INTERNAL_CODE';
    vol[0][3] = 'PRODUCT'; vol[0][4] = 'SITE'; vol[0][5] = 'RETAILER';
    vol[0][6] = 'MATCH_KEY'; vol[0][7] = 'PACKS'; vol[0][8] = 'CASES';
    for (var p = 0; p < 10; p++) {
      var prod = PRODUCTS[p % PRODUCTS.length];
      vol[p + 1][0] = 'R' + (500 + p);
      vol[p + 1][1] = (p % 4 === 1) ? dirt.padded(9000 + p, 8) : { v: '59' + (10000000 + p * 37) };
      vol[p + 1][2] = productCode(p);
      vol[p + 1][3] = prod[0];
      vol[p + 1][4] = SITES[p % SITES.length];
      vol[p + 1][5] = RETAILERS[p % RETAILERS.length];
      vol[p + 1][6] = 'MK' + (700 + p);
      vol[p + 1][7] = int(r, 20000, 90000);
      vol[p + 1][8] = int(r, 500, 4000);
    }
    // Row 13 (index 12) is the date header row; row 17 (index 16) a volume row.
    vol[12][0] = 'WEEK ENDING';
    for (var w = 0; w < 10; w++) {
      vol[12][w + 1] = { v: day - 7 * (9 - w), fmt: 'dd/mm/yyyy' };
      vol[16][w + 1] = int(r, 100000, 400000);
    }
    vol[16][0] = 'TOTAL PACKS';

    // Complaints Paste: 35 columns with the gaps the census found. The ones the
    // exercises use are A, I, J, L, N, O, P, Q, R, X (the date) and AA.
    var comp = [new Array(40).fill(null)];
    comp[0][0] = 'REF'; comp[0][8] = 'RETAILER'; comp[0][9] = 'SITE';
    comp[0][11] = 'RETAILER_CODE'; comp[0][13] = 'EAN'; comp[0][14] = 'INTERNAL_CODE';
    comp[0][15] = 'DESCRIPTION'; comp[0][16] = 'CATEGORY'; comp[0][17] = 'FAILURE';
    comp[0][23] = 'COMPLAINT_DATE'; comp[0][26] = 'CONFIRMED';
    for (var k = 0; k < 24; k++) {
      var row = new Array(40).fill(null);
      var src = k % 12;
      row[0] = 'C' + (90000 + k);
      row[8] = RETAILERS[src % RETAILERS.length];
      row[9] = SITES[src % SITES.length];
      // Each complaint carries ONE of the three code systems, not all three.
      var which = k % 3;
      if (which === 0 && src < 10) row[11] = 'R' + (500 + src);
      else if (which === 1 && src < 10) row[13] = { v: '59' + (10000000 + src * 37) };
      else if (src < 10) row[14] = productCode(src);
      else row[14] = productCode(99);                 // a code that is in nothing
      row[15] = PRODUCTS[src % PRODUCTS.length][0];
      row[16] = CATEGORIES[src % CATEGORIES.length];
      row[17] = pick(r, ['Foreign body', 'Label error', 'Short weight', 'Pack damage']);
      row[23] = { v: day - int(r, 0, 60), fmt: 'dd/mm/yyyy' };
      row[26] = (r() < 0.6) ? 'Yes' : 'No';
      comp.push(row);
    }

    return {
      sheets: [
        { name: 'Site KPI Reporting', layer: 5, note: 'The weekly KPI page.',
          rows: [['Complaints per million packs'], [],
            ['Site', 'Complaints', 'Packs', 'CPM']],
          colWidth: { 0: 120 } },
        { name: 'Complaints Paste', layer: 1,
          note: 'LANDING. The complaints export. Three rival code systems.',
          rows: comp },
        { name: 'Volumes Paste', layer: 1,
          note: 'LANDING. Volumes. Keys in A, B and C; dates across row 13.',
          rows: vol },
        { name: 'Complaints Paste Review', layer: 2,
          note: 'MIRROR + the fallback key cascade. You build this.',
          rows: [['Ref', 'Seq', 'Count', 'First', 'Category', 'Period',
            'REF', 'RETAILER', 'SITE', 'R_CODE', 'has R?', 'EAN', 'has EAN?',
            'INT', 'has INT?', 'MATCH_KEY', 'VERDICT']] },
        { name: 'Data Extract', layer: 4,
          note: 'WORKINGS. Period windows.',
          rows: [['Period extract']] }
      ]
    };
  }

  /* ========================================== S5 -- the many-source assembly */

  function buildS5(opts) {
    opts = opts || {};
    var seed = opts.seed || 1717;
    var r = rng(seed);
    var day = opts.today || D.ymdToSerial(2026, 11, 13);

    // Four separate pasted exports, each starting at a different row because
    // each report has its own header block. That is faithful and it matters.
    function lbi() {
      var rows = [[], [], [], [], ['LBI EXPORT'], ['BATCH', 'KILL_DATE', 'BIRDS',
        null, null, null, null, null, null, null, null, null, null, null, 'LINE',
        null, null, null, 'CRATE', 'FARM', 'WEIGHT']];
      for (var i = 0; i < 18; i++) {
        var row = new Array(20).fill(null);
        row[0] = 'B' + (3000 + i);
        row[1] = { v: day - int(r, 0, 5), fmt: 'dd/mm/yyyy' };
        row[2] = int(r, 2000, 6000);
        row[14] = LINES[i % 2];
        row[18] = 'C' + (100 + i);
        row[19] = int(r, 1500, 2600);
        rows.push(row);
      }
      return rows;
    }
    function despatch() {
      var rows = [[], [], [], [], [],
        ['PALLET', 'BATCH', 'PRODUCT', 'CASES', 'USE_BY', 'DEPOT', 'TEMP']];
      for (var i = 0; i < 16; i++) {
        rows.push(['P' + (700 + i), 'B' + (3000 + (i % 18)),
          PRODUCTS[i % PRODUCTS.length][0], int(r, 10, 90),
          { v: day + int(r, 2, 9), fmt: 'dd/mm/yyyy' },
          pick(r, RETAILERS), (r() < 0.9) ? int(r, 1, 4) : 0]);
      }
      return rows;
    }
    function transactions() {
      var rows = [[], [], [], [], [], [],
        ['TRANS_REF', null, null, null, null, null, 'USE_BY', 'PACKED', 'STAMP', 'CASES']];
      for (var i = 0; i < 20; i++) {
        var row = new Array(10).fill(null);
        row[0] = 'T' + (8000 + i);
        row[6] = { v: day + int(r, 2, 9), fmt: 'dd/mm/yyyy' };
        row[7] = { v: day - int(r, 0, 3), fmt: 'dd/mm/yyyy' };
        row[8] = { v: day + D.hmsToSerial(int(r, 5, 18), int(r, 0, 59), 0),
          fmt: 'dd/mm/yyyy hh:mm' };
        row[9] = int(r, 5, 40);
        rows.push(row);
      }
      return rows;
    }

    return {
      sheets: [
        { name: 'Label Info', layer: 5, note: 'OUTPUT. What prints on the label.',
          rows: [['Sample label'], [], ['Batch', null], ['Use by', null],
            ['Kill date', null], ['Line', null], ['Test', null]],
          colWidth: { 0: 110, 1: 150 } },
        { name: 'LBI Paste', layer: 1, note: 'LANDING 1. Batch and kill data. Starts at row 6.',
          rows: lbi() },
        { name: 'Despatch Paste', layer: 1, note: 'LANDING 2. Pallets out. Starts at row 6.',
          rows: despatch() },
        { name: 'Transaction Listings Paste', layer: 1,
          note: 'LANDING 3. Packing transactions. Starts at row 7.',
          rows: transactions() },
        { name: 'Despatch Hidden', layer: 2, hidden: true,
          note: 'MIRROR of Despatch Paste, with the export’s zeros blanked.',
          rows: [[]] },
        { name: 'LBI Hidden', layer: 4, hidden: true,
          note: 'WORKINGS. Joined and derived values.',
          rows: [['BATCH', 'BIRDS', 'LINE', 'KILL_DATE', 'OFFSET', 'KILL_TIME']] },
        { name: 'Batch Data Information', layer: 4,
          note: 'WORKINGS. The batch-level summary the label needs.',
          rows: [['Batch summary'], [], ['Earliest use-by', null],
            ['Latest use-by', null], ['Single use-by?', null], ['Batches', null]],
          colWidth: { 0: 150, 1: 140 } },
        { name: 'Error Coding', layer: 5,
          note: 'OUTPUT. The workbook’s own verdict on its data.',
          rows: [['Checks'], [], ['Use-by consistent', null], ['All batches matched', null],
            ['Verdict', null]],
          colWidth: { 0: 170, 1: 220 } }
      ]
    };
  }

  /* ========================================================= the scenario map */

  XT.data.scenarios = {
    S1: {
      id: 'S1', title: 'Sample Hub — the enrichment fan-out',
      source: 'Barcode Sample Submission v15.xlsm — 19 sheets, 85,163 formulas',
      job: 'A lab operator scans sample barcodes. The workbook has to produce a ' +
        'submission file for the external lab and a summary for the email.',
      why: 'The gentlest place to start: one key, many columns. It teaches $ anchoring, ' +
        'IFERROR, fill-right, and the payoff of a well-built lookup table — then ' +
        'escalates to two-key joins and the shift cut-off.',
      patterns: ['P1', 'P2', 'P3', 'P5', 'P9', 'P10'],
      layers: ['Samplesin', 'scanMODWB', 'FILTERWBTODAY', 'MIMS DAILY WB OLD', 'FINALEMAIL WB'],
      build: buildS1
    },
    S4: {
      id: 'S4', title: 'Line verification — parsing a machine log',
      source: '26th October LV Verification V20.xlsx — 4 sheets, 1,000,076 formulas',
      job: 'A vision system emits one text line per pack. Turn it into a count of ' +
        'label mismatches and missing sleeves per product.',
      why: 'The best demonstration in the whole catalog: the old way (count characters ' +
        'with RIGHT and LEFT) and the new way (count fields with TEXTAFTER) solving the ' +
        'same problem in the same file.',
      patterns: ['P3', 'P6', 'P7', 'P13'],
      layers: ['LVALL', 'Sheet4', 'Home'],
      build: buildS4
    },
    S3: {
      id: 'S3', title: 'Micro trending — the week grid',
      source: 'All Lab helpers…xlsm and Live E.coli Trending — 667,402 formulas',
      job: 'Micro results per site per week, percentage over the action limit, trended.',
      why: 'The pure "reorganise, do not calculate" scenario, and the only place real ' +
        'statistics appear. Teaches the retailer week grid and the limit-of-detection floor.',
      patterns: ['P1', 'P8', 'P12', 'P16', 'P20'],
      layers: ['ECOLI MICRO PASTE', 'Ecoli Final Data', 'Current month'],
      build: buildS3
    },
    S2: {
      id: 'S2', title: 'Complaints analysis — the fallback-key join',
      source: 'Complaints Analysis Weekly Database2.xlsm — 515,992 formulas',
      job: 'Turn a complaints export and a volumes export into complaints per million ' +
        'packs, by site and period.',
      why: 'The hardest reasoning in the estate: three rival code systems, a fallback ' +
        'cascade, and a failure that speaks to a human rather than showing #N/A.',
      patterns: ['P4', 'P11', 'P12', 'P14'],
      layers: ['Complaints Paste', 'Complaints Paste Review', 'Data Extract', 'Site KPI Reporting'],
      build: buildS2
    },
    S5: {
      id: 'S5', title: 'Labels toolkit — the many-source assembly',
      source: 'Campy WB and Necks labels toolkit v17.xlsm — 48 sheets, 441,660 formulas',
      job: 'From four pasted exports, assemble printable sample labels with the right ' +
        'batch, use-by and kill-time — and report the workbook’s own data errors.',
      why: 'The capstone. Four sources, a hidden workings layer, line-specific rules, ' +
        'and a self-audit sheet.',
      patterns: ['P2', 'P5', 'P15', 'P17', 'P19'],
      layers: ['LBI Paste', 'Despatch Hidden', 'LBI Hidden', 'Batch Data Information',
        'Label Info', 'Error Coding'],
      build: buildS5
    }
  };

  XT.data.scenarioOrder = ['S1', 'S4', 'S3', 'S2', 'S5'];

  /* Build a scenario into a live workbook. */
  XT.data.DEFAULT_TODAY = D.ymdToSerial(2026, 11, 13);   // a Friday

  XT.data.makeWorkbook = function (scenarioId, opts) {
    var sc = XT.data.scenarios[scenarioId];
    if (!sc) return null;
    opts = opts || {};
    /* Freeze TODAY() to the fixture's day. Without this, every exercise that
       filters on "today" would grade differently tomorrow, and a saved answer
       would rot overnight. Sandbox mode can unfreeze it. */
    D.frozenToday = opts.today || XT.data.DEFAULT_TODAY;
    var def = sc.build(opts);
    var wb = XT.newWorkbook(sc.title);
    def.sheets.forEach(function (sd) {
      wb.addSheet(sd.name, { hidden: sd.hidden, note: sd.note, layer: sd.layer });
    });
    wb.load(def);
    wb.recalc();
    return wb;
  };

  /* A TSV block for the paste exercises: the same generator, rendered as the
     text a real export would put on the clipboard. */
  XT.data.pasteBlock = function (which, opts) {
    opts = opts || {};
    var day = opts.today || D.ymdToSerial(2026, 11, 13);
    var rows;
    if (which === 'samplesin') rows = samplesIn(opts.seed || 777, opts.n || 12, day);
    else if (which === 'mims') rows = mimsSaveFile(opts.seed || 20261113, opts.n || 14);
    else return '';
    return rows.map(function (row) {
      return row.map(function (v) {
        if (v === null || v === undefined) return '';
        if (typeof v === 'object') {
          if (v.f !== undefined) return v.f;
          if (typeof v.v === 'number' && v.fmt) return XT.F.format(v.v, v.fmt);
          return String(v.v);
        }
        if (typeof v === 'number' && which === 'samplesin') return String(v);
        return String(v);
      }).join('\t');
    }).join('\n');
  };

})(window.XT = window.XT || {});
