/* =============================================================================
   Excel Trainer — data/toolkits.js
   The four toolkits that matter day to day:

     T1  Shelf life toolkit        Shelf life toolkit V12 — 14 sheets, 3,085 formulas
     T2  Neck & Caeca toolkit      (that is scenario S5, already in fixtures.js)
     T3  Nutritional toolkit       2sfg UKP Nutritional Toolkit — 803 formulas
     T4  Micro schedule            2026 Schedule V22 — 23 sheets, 70,000+ formulas

   Same rule as fixtures.js: every VALUE here is invented. What is copied is the
   STRUCTURE — sheet names, which column the key sits in, the paired result
   columns, the "." placeholder, the Y/N risk ladder, the week grid — because
   that structure is the thing that has to be learned.

   Data is hardcoded rather than generated so that a hint can say "row 7" and a
   drill's expected answer never moves.
   ============================================================================= */
(function (XT) {
  'use strict';
  XT.data = XT.data || {};
  var D = XT.D;

  function dt(y, m, d) { return { v: D.ymdToSerial(y, m, d), fmt: 'dd/mm/yyyy' }; }
  function txt(v) { return { v: String(v) }; }          // a number stored as text
  function trail(s) { return { v: s + ' ' }; }          // trailing space

  /* ======================================================= T1 — SHELF LIFE == */

  /* The real pipeline:
       MICRO PASTE  (lab export, one flat block)
         -> SET MOD        which sample id belongs to which test point
         -> SAMPLE ID NO   XLOOKUP, because the key is to the RIGHT of the answer
         -> MICRO MOD      VLOOKUP the result, then strip the "<" to get a number
         -> Test Profile 1 the sheet the technologist signs                        */
  function buildT1() {
    var POINTS = ['Day 0', 'Day 3', 'Day 6', 'Day 9', 'Day 12', 'Day 15', 'Day 18', 'Day 21'];
    var IDS = ['SL2611-001', 'SL2611-002', 'SL2611-003', 'SL2611-004',
               'SL2611-005', 'SL2611-006', 'SL2611-007', 'SL2611-008'];

    /* MICRO PASTE. Columns D and E are the real paired-result trick: when E
       holds "." the usable number is in D; when E holds something, E wins.
       Several results are "<10" or ">1000000", which are TEXT, not numbers. */
    var paste = [
      ['SAMPLE ID', 'PRODUCT', 'RECEIVED', 'TVC cfu/g', 'TVC ALT',
       'YEAST & MOULD', 'ENTEROS', 'PSEUDOMONAS', 'LAB STATUS'],
      ['SL2611-001', 'Aromatic Butter Chicken BDF', dt(2026, 11, 2), '<10', '.', '<100', '<10', '<10', 'Final'],
      ['SL2611-002', 'Aromatic Butter Chicken BDF', dt(2026, 11, 5), 120, '.', '<100', '<10', 40, 'Final'],
      ['SL2611-003', 'Aromatic Butter Chicken BDF', dt(2026, 11, 8), 1800, '.', 150, '<10', 900, 'Final'],
      ['SL2611-004', 'Aromatic Butter Chicken BDF', dt(2026, 11, 11), txt(24000), '.', 320, 20, 11000, 'Final'],
      ['SL2611-005', 'Aromatic Butter Chicken BDF', dt(2026, 11, 14), '.', 190000, 2400, 90, 85000, 'Final'],
      ['SL2611-006', 'Aromatic Butter Chicken BDF', dt(2026, 11, 17), '>1000000', '.', 15000, 300, 420000, 'Final'],
      [null, null, null, null, null, null, null, null, null],
      ['SL2611-007', 'Aromatic Butter Chicken BDF', dt(2026, 11, 20), '.', 2700000, 46000, 1200, 980000, 'Interim'],
      ['SL2611-008', 'Aromatic Butter Chicken BDF', dt(2026, 11, 23), '>1000000', '.', 88000, 3400, 1500000, 'Interim']
    ];

    /* SET MOD. The key is TEST POINT in column H; the answer, SAMPLE ID, is in
       column A — to its LEFT. VLOOKUP cannot look left. This single fact is why
       the real workbook uses XLOOKUP here and nowhere else. */
    var setmod = [['SAMPLE ID', 'SET', 'TEMP', 'PACK', 'SUITE', 'ANALYST', 'NOTE', 'TEST POINT']];
    POINTS.forEach(function (p, i) {
      setmod.push([IDS[i], 'S5', '3C', 'MAP', 'Full micro', 'KD', '', p]);
    });

    var sampleId = [['ROW', 'TEST POINT', 'SAMPLE ID (used)', 'SAMPLE ID (XLOOKUP)']];
    POINTS.forEach(function (p, i) { sampleId.push([i + 1, p, null, null]); });

    var micromod = [['DAY', 'TEST POINT', 'SAMPLE ID', 'TVC RAW', 'TVC NUMBER',
                     'Y&M RAW', 'Y&M NUMBER', 'PRODUCT', 'SAMPLE NO', 'LABEL']];
    POINTS.forEach(function (p, i) {
      micromod.push([i * 3, p, null, null, null, null, null, null, null, null]);
    });

    var profile = [
      ['Shelf life toolkit — Test Profile 1'],
      [],
      ['Product', null, 'Aromatic Butter Chicken BDF'],
      ['Shelf life applied for (days)', null, 21],
      ['Test type', null, 'Routine'],
      ['Multiplier', null, null],
      ['TVC action limit (cfu/g)', null, null],
      ['Y&M action limit (cfu/g)', null, null],
      [],
      [],
      ['DAY', 'TEST POINT', 'PRODUCT', 'TVC', 'Y&M', 'OVER LIMIT?',
       'REP 1', 'REP 2', 'REP 3', 'AVERAGE', 'TOTAL'],
      [0, 'Day 0', null, null, null, null, 3, 3, 3, null, null],
      [3, 'Day 3', null, null, null, null, 3, 3, 3, null, null],
      [6, 'Day 6', null, null, null, null, 3, 3, 3, null, null],
      [9, 'Day 9', null, null, null, null, 3, 3, 3, null, null],
      [12, 'Day 12', null, null, null, null, 3, 3, 3, null, null],
      [15, 'Day 15', null, null, null, null, 3, 3, 3, null, null],
      [18, 'Day 18', null, null, null, null, 3, 2, 3, null, null],
      [21, 'Day 21', null, null, null, null, 3, 3, 3, null, null],
      [],
      ['Samples analysed', null, null],
      ['Days over limit', null, null],
      ['Verdict', null, null]
    ];

    /* The key in column A is the product and the organism glued together with
       "|". One product has two limits, so the product on its own cannot be a
       key — this is the estate's standard answer to that, and the reason &
       turns up inside so many lookups. */
    var cop = [
      ['KEY', 'PRODUCT', 'ORGANISM', 'ACTION LIMIT', 'MAX LIMIT', 'UNIT'],
      ['Aromatic Butter Chicken BDF|TVC', 'Aromatic Butter Chicken BDF', 'TVC', 100000, 1000000, 'cfu/g'],
      ['Aromatic Butter Chicken BDF|Yeast & Mould', 'Aromatic Butter Chicken BDF', 'Yeast & Mould', 10000, 100000, 'cfu/g'],
      ['Bulk Drum Fillet No Gas|TVC', 'Bulk Drum Fillet No Gas', 'TVC', 1000000, 10000000, 'cfu/g'],
      ['Bulk Drum Fillet No Gas|Yeast & Mould', 'Bulk Drum Fillet No Gas', 'Yeast & Mould', 100000, 1000000, 'cfu/g'],
      ['Cooked Diced Chicken|TVC', 'Cooked Diced Chicken', 'TVC', 10000, 100000, 'cfu/g'],
      ['Cooked Diced Chicken|Yeast & Mould', 'Cooked Diced Chicken', 'Yeast & Mould', 1000, 10000, 'cfu/g']
    ];

    return {
      sheets: [
        { name: 'MICRO PASTE', layer: 1,
          note: 'LANDING. The lab export, pasted whole. Column E holds "." when ' +
            'the usable number is in D — that pair is the real file’s own trick. ' +
            'Row 8 is blank, "<10" and ">1000000" are text, and 24000 is a number ' +
            'stored as text.',
          rows: paste, colWidth: { 0: 110, 1: 215, 2: 100 } },

        { name: 'SET MOD', layer: 1, hidden: true,
          note: 'REFERENCE. Test point is the KEY, in column H. Sample id, the ' +
            'answer, is in column A — to its LEFT. VLOOKUP cannot do this.',
          rows: setmod, colWidth: { 0: 110, 7: 100 } },

        { name: 'SAMPLE ID NO', layer: 2,
          note: 'MIRROR. Which sample id belongs to each test point.',
          rows: sampleId, colWidth: { 1: 100, 2: 140, 3: 150 } },

        { name: 'COP', layer: 1,
          note: 'REFERENCE. Code of practice limits. One product has two limits, ' +
            'so the key in column A is product and organism glued with "|".',
          rows: cop, colWidth: { 0: 260, 1: 215, 2: 120 } },

        { name: 'MICRO MOD', layer: 4,
          note: 'WORKINGS. The hidden sheet that turns "<10" and ">1000000" into ' +
            'numbers you can average and plot.',
          rows: micromod, colWidth: { 1: 100, 2: 110, 3: 110, 4: 110, 7: 215, 9: 260 } },

        { name: 'Test Profile 1', layer: 5,
          note: 'OUTPUT. The sheet the technologist signs and sends.',
          rows: profile, colWidth: { 0: 180, 1: 100, 2: 215 } }
      ]
    };
  }

  /* ====================================================== T3 — NUTRITIONAL == */

  /* The real pipeline:
       PASTE      one flat key -> value block out of the lab system
         -> background   VLOOKUP each analyte, strip "<", work out energy
         -> toolkit      three reps averaged, compared with the declaration,
                         and judged against the UK tolerance bands             */
  function buildT3() {
    var paste = [
      ['KEY', 'VALUE', 'UNIT', 'LOD', 'METHOD'],
      ['S5-2026-0412', 'Aldi Chicken Breast Fillets 650g', '', '', ''],
      ['S5-2026-0413', 'Aldi Chicken Thigh Fillets 500g', '', '', ''],
      ['FAT', 2.4, 'g/100g', 0.1, 'Soxhlet'],
      ['SATURATES', 0.7, 'g/100g', 0.1, 'GC-FID'],
      ['PROTEIN', 23.8, 'g/100g', 0.5, 'Kjeldahl'],
      ['CARBOHYDRATE', '<0.5', 'g/100g', 0.5, 'Calculation'],
      ['SUGARS', '<0.5', 'g/100g', 0.5, 'HPLC'],
      ['FIBRE', '<0.5', 'g/100g', 0.5, 'AOAC 991.43'],
      ['SALT', 0.21, 'g/100g', 0.01, 'Calculation'],
      ['SODIUM', 84, 'mg/100g', 1, 'ICP-OES'],
      ['MOISTURE', 73.1, 'g/100g', 0.1, 'Oven'],
      ['ASH', txt(1.1), 'g/100g', 0.1, 'Furnace']
    ];

    var labels = [
      ['RETAILER', 'PRODUCT', 'NUTRIENT', 'VALUE', 'REF VALUE', 'STATE'],
      ['Aldi', 'Chicken Breast Fillets 650g', 'Energy kJ', 480, 8400, 'Declared'],
      ['Aldi', 'Chicken Breast Fillets 650g', 'Fat', 2.0, 70, 'Declared'],
      ['Aldi', 'Chicken Breast Fillets 650g', 'Saturates', 0.6, 20, 'Declared'],
      ['Aldi', 'Chicken Breast Fillets 650g', 'Carbohydrate', 0.5, 260, 'Declared'],
      ['Aldi', 'Chicken Breast Fillets 650g', 'Sugars', 0.5, 90, 'Declared'],
      ['Aldi', 'Chicken Breast Fillets 650g', 'Fibre', 0.5, 0, 'Declared'],
      ['Aldi', 'Chicken Breast Fillets 650g', 'Protein', 22.0, 50, 'Declared'],
      ['Aldi', 'Chicken Breast Fillets 650g', 'Salt', 0.20, 6, 'Declared']
    ];

    var ANALYTES = ['FAT', 'SATURATES', 'PROTEIN', 'CARBOHYDRATE', 'SUGARS',
                    'FIBRE', 'SALT', 'SODIUM', 'MOISTURE', 'ASH'];
    var bg = [['ANALYTE', 'RAW FROM PASTE', 'NUMBER', 'UNIT']];
    ANALYTES.forEach(function (a) { bg.push([a, null, null, null]); });
    bg.push([]);
    bg.push(['Energy kJ (calculated)', null]);
    bg.push(['Energy kcal (calculated)', null]);

    /* The declaration check. Three reps, their average, the declared figure,
       the difference, and the UK tolerance band for that nutrient. */
    var tk = [
      ['UKP Nutritional Toolkit — declaration check'],
      [],
      ['Sample ref', null, 'S5-2026-0412'],
      ['Product', null, null],
      ['Declared basis (g)', null, 100],
      ['Reps per nutrient', null, 3],
      [],
      [],
      ['NUTRIENT', 'DECLARED', 'REP 1', 'REP 2', 'REP 3', 'AVERAGE', 'LAB VALUE',
       'DIFFERENCE', '% DIFF', 'TOLERANCE +/-%', 'UNDER?', 'OVER?', 'VERDICT'],
      ['Fat', 2.0, 2.3, 2.5, 2.4, null, null, null, null, 20, null, null, null],
      ['Saturates', 0.6, 0.9, 0.8, 0.9, null, null, null, null, 20, null, null, null],
      ['Protein', 22.0, 23.9, 23.6, 23.9, null, null, null, null, 20, null, null, null],
      ['Carbohydrate', 0.5, 0.4, 0.5, 0.4, null, null, null, null, 20, null, null, null],
      ['Sugars', 0.5, 0.4, 0.5, 0.5, null, null, null, null, 20, null, null, null],
      ['Fibre', 0.5, 0.5, 0.4, 0.5, null, null, null, null, 20, null, null, null],
      ['Salt', 0.20, 0.21, 0.22, 0.20, null, null, null, null, 20, null, null, null],
      ['Sodium', 80, 84, 85, 83, null, null, null, null, 20, null, null, null],
      ['Moisture', 73.0, 73.1, 73.4, 72.8, null, null, null, null, 20, null, null, null],
      ['Ash', 1.0, 1.1, 1.1, 1.0, null, null, null, null, 20, null, null, null],
      [],
      ['Nutrients outside tolerance', null],
      ['Declaration verdict', null]
    ];

    return {
      sheets: [
        { name: 'PASTE', layer: 1, hidden: true,
          note: 'LANDING. One flat key to value block. The key is column A and ' +
            'the answer is column B, which is why every lookup in the real file ' +
            'asks for column 2. "<0.5" is TEXT, and ASH is a number stored as text.',
          rows: paste, colWidth: { 0: 140, 1: 230, 2: 90 } },

        { name: 'LABELS INFOS', layer: 1,
          note: 'REFERENCE. The declaration as printed, and the reference intakes.',
          rows: labels, colWidth: { 1: 230, 2: 120 } },

        { name: 'background', layer: 4, hidden: true,
          note: 'WORKINGS. Where "<0.5" becomes 0.5 and the energy is calculated ' +
            'from protein, fat, carbohydrate and fibre.',
          rows: bg, colWidth: { 0: 180, 1: 140, 2: 110 } },

        { name: '2sfg UKP Nutritional Toolkit', layer: 5,
          note: 'OUTPUT. The signed declaration check.',
          rows: tk, colWidth: { 0: 150, 9: 110, 12: 150 } }
      ]
    };
  }

  /* =================================================== T4 — MICRO SCHEDULE == */

  /* The real pipeline:
       Swab Results Database  what came back from the lab
         -> Risk Assessment   Y/N answers become a score, a band, a frequency
         -> QTR 1             a 13-week grid, one row per point
         -> Point List        the rollup, and the date chain  =prev+7            */
  function buildT4() {
    var PTS = [
      [1, 'Evisceration', 'Vent cutter blade', 'Y', 'Y', 'Y', 'Y'],
      [2, 'Evisceration', 'Shackle line guard', 'N', 'Y', 'N', 'Y'],
      [3, 'Chilling', 'Air chiller inlet rail', 'N', 'Y', 'N', 'N'],
      [4, 'Chilling', 'Drip tray under rail 3', 'N', 'N', 'Y', 'Y'],
      [5, 'Cutting', 'Breast fillet conveyor', 'Y', 'Y', 'N', 'N'],
      [6, 'Cutting', 'Thigh deboner nozzle', 'Y', 'Y', 'Y', 'N'],
      [7, 'Cutting', 'Weigh hopper 2', 'Y', 'Y', 'N', 'Y'],
      [8, 'Packing', 'Tray sealer film track', 'N', 'Y', 'N', 'N'],
      [9, 'Packing', 'Hand wash basin tap', 'N', 'N', 'N', 'Y'],
      [10, 'Packing', 'Pack-off table edge', 'Y', 'Y', 'Y', 'Y'],
      [11, 'Despatch', 'Pallet wrapper base', 'N', 'N', 'N', 'N'],
      [12, 'Despatch', 'Loading bay floor drain', 'N', 'N', 'Y', 'Y'],
      [13, 'Hygiene', 'Cleaning hose reel 4', 'N', 'N', 'Y', 'Y'],
      [14, 'Hygiene', 'Boot wash brush', 'N', 'N', 'Y', 'N'],
      [15, 'Intake', 'Crate washer exit', 'N', 'Y', 'Y', 'Y']
    ];

    var home = [
      ['Micro swab schedule 2026'],
      [],
      ['Week 1 Monday', null, dt(2026, 1, 5)],
      ['Weeks per quarter', null, 13],
      ['Points on schedule', null, null],
      ['High risk points', null, null],
      ['Target % complete', null, 95],
      [],
      ['BAND', 'WEEKS BETWEEN SWABS'],
      ['High', 1],
      ['Medium', 2],
      ['Low', 4],
      ['No Risk', 13],
      [],
      ['QUARTER 1 SUMMARY'],
      ['Swabs taken', null, null],
      ['Swabs over limit', null, null],
      ['Cutting fails in the lab database', null, null],
      ['Limit total for Cutting swabs', null, null],
      ['% clear', null, null],
      ['First swab date', null, null],
      ['Last swab date', null, null]
    ];

    var ra = [
      ['Risk Assessment — swab points'],
      [],
      ['POINT', 'AREA', 'DESCRIPTION', 'FOOD CONTACT', 'OPEN PRODUCT',
       'PAST FAILURE', 'HARD TO CLEAN', 'SC CONTACT', 'SC OPEN', 'SC FAILURE',
       'SC CLEAN', 'TOTAL', 'BAND', 'WEEKS BETWEEN']
    ];
    PTS.forEach(function (p) {
      ra.push([p[0], p[1], p[2], p[3], p[4], p[5], p[6], null, null, null, null, null, null, null]);
    });

    var WEEKS = 13;
    var q1 = [
      ['QTR 1 — week grid'],
      [],
      [null, null, null, null, null, null, null, null, null, null, 'WEEK COMMENCING'],
      [],
      ['POINT', 'AREA', 'DESCRIPTION', 'BAND', null, null, null, null, null, null,
       'W1', 'W2', 'W3', 'W4', 'W5', 'W6', 'W7', 'W8', 'W9', 'W10', 'W11', 'W12', 'W13']
    ];
    /* Seeded results: 1 = swabbed and clear, 2 = swabbed and over limit,
       blank = not due. Deliberately irregular, as a real quarter is. */
    var SEED = [
      [1, 1, 1, 2, 1, 1, 1, 1, 2, 1, 1, 1, 1],
      [1, '', 1, '', 1, '', 1, '', 1, '', 1, '', 1],
      [1, '', '', '', 1, '', '', '', 2, '', '', '', 1],
      [1, '', 1, '', 1, '', 2, '', 1, '', 1, '', 1],
      [1, 1, 1, 1, 1, 2, 1, 1, 1, 1, 1, 1, 1],
      [1, 1, 2, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1],
      [1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1],
      [1, '', 1, '', 1, '', 1, '', 1, '', 1, '', 1],
      ['', '', '', '', 1, '', '', '', '', '', '', '', 1],
      [1, 1, 1, 1, 2, 1, 1, 1, 1, 2, 1, 1, 1],
      ['', '', '', '', '', '', 1, '', '', '', '', '', ''],
      [1, '', '', '', 1, '', '', '', 1, '', '', '', 1],
      [1, '', '', '', 1, '', '', '', 1, '', '', '', 1],
      ['', '', '', '', 1, '', '', '', '', '', '', '', 1],
      [1, '', 1, '', 1, '', 1, '', 2, '', 1, '', 1]
    ];
    PTS.forEach(function (p, i) {
      var row = [p[0], p[1], p[2], null, null, null, null, null, null, null];
      for (var w = 0; w < WEEKS; w++) {
        var v = SEED[i][w];
        row.push(v === '' ? null : v);
      }
      q1.push(row);
    });
    q1.push([]);
    q1.push(['Swabs taken']);
    q1.push(['Swabs over limit']);
    q1.push(['% clear']);

    var pl = [
      ['Point List — quarter rollup'],
      [],
      ['POINT', 'DESCRIPTION', 'BAND', 'WEIGHTED TOTAL', 'QTR 1 OVER LIMIT',
       '% CLEAR', 'FIRST FAIL WEEK']
    ];
    PTS.forEach(function (p) { pl.push([p[0], p[2], null, null, null, null, null]); });

    var db = [
      ['DATE', 'POINT', 'AREA', 'ORGANISM', 'RESULT', 'LIMIT', 'VERDICT'],
      [dt(2026, 1, 5), 1, 'Evisceration', 'TVC', 1200, 10000, 'Pass'],
      [dt(2026, 1, 5), 5, 'Cutting', 'TVC', txt(24000), 10000, 'Fail'],
      [dt(2026, 1, 5), 10, 'Packing', 'Enteros', '<10', 100, 'Pass'],
      [dt(2026, 1, 12), 1, 'Evisceration', 'TVC', 8400, 10000, 'Pass'],
      [dt(2026, 1, 12), 6, 'Cutting', 'Enteros', 140, 100, 'Fail'],
      [dt(2026, 1, 19), 3, 'Chilling', 'TVC', 600, 10000, 'Pass'],
      [dt(2026, 1, 19), trail('10'), 'Packing', 'TVC', 3100, 10000, 'Pass'],
      [dt(2026, 1, 26), 1, 'Evisceration', 'TVC', 22000, 10000, 'Fail'],
      [dt(2026, 1, 26), 7, 'Cutting', 'TVC', 900, 10000, 'Pass'],
      [dt(2026, 2, 2), 2, 'Evisceration', 'Enteros', '<10', 100, 'Pass']
    ];

    return {
      sheets: [
        { name: 'Home', layer: 5,
          note: 'The front panel. One date in C3 drives all 52 week headers.',
          rows: home, colWidth: { 0: 180, 1: 180, 2: 120 } },

        { name: 'Swab Results Database', layer: 1,
          note: 'LANDING. What came back from the lab. Point 10 arrives with a ' +
            'trailing space in row 8 and one result is text.',
          rows: db, colWidth: { 0: 110, 2: 130, 3: 110 } },

        { name: 'Risk Assessment', layer: 2,
          note: 'The scoring sheet. Four Y/N answers become a score, a total, a ' +
            'band, and the swab frequency that band earns.',
          rows: ra, colWidth: { 2: 200, 12: 100 } },

        { name: 'QTR 1', layer: 4,
          note: 'WORKINGS. 13 weeks across. 1 = swabbed and clear, 2 = over limit, ' +
            'blank = not due.',
          rows: q1, colWidth: { 2: 200 } },

        { name: 'Point List', layer: 5,
          note: 'OUTPUT. The rollup the manager reads.',
          rows: pl, colWidth: { 1: 200 } }
      ]
    };
  }

  /* ========================================================== registration == */

  XT.data.scenarios.T1 = {
    id: 'T1', title: 'Shelf life toolkit',
    source: 'Shelf life toolkit V12 — 14 sheets, 3,085 formulas',
    job: 'A shelf life study runs to 21 days. Eight sample points come back from ' +
      'the lab as one flat block. Turn them into a signed test profile with a ' +
      'verdict on each day.',
    why: 'The only place in your estate where XLOOKUP earns its keep: the key is ' +
      'to the right of the answer. Then it is all about "<10" — a result that ' +
      'looks like a number and is text.',
    patterns: ['P1', 'P3', 'P5', 'P6', 'P10'],
    layers: ['MICRO PASTE', 'SET MOD', 'SAMPLE ID NO', 'MICRO MOD', 'Test Profile 1'],
    build: buildT1
  };

  XT.data.scenarios.T3 = {
    id: 'T3', title: 'Nutritional toolkit',
    source: '2sfg UKP Nutritional Toolkit — 4 sheets, 803 formulas',
    job: 'Three replicate results per nutrient. Average them, compare with what ' +
      'the pack declares, and say whether the declaration still stands.',
    why: 'One flat key-to-value paste and one answer column. The whole file is ' +
      'VLOOKUP column 2, AVERAGE, and a tolerance test — done ten times.',
    patterns: ['P1', 'P3', 'P5', 'P12'],
    layers: ['PASTE', 'background', '2sfg UKP Nutritional Toolkit'],
    build: buildT3
  };

  XT.data.scenarios.T4 = {
    id: 'T4', title: 'Micro schedule',
    source: '2026 Schedule V22 — 23 sheets, 70,000+ formulas',
    job: 'Fifteen swab points, each scored for risk, each earning a swab ' +
      'frequency. Lay the year out in weeks and report what actually got done.',
    why: 'The counting scenario. SUMIF and COUNTIF appear 21,000 times in this ' +
      'one workbook, and the Y/N risk ladder is the clearest nested IF in the estate.',
    patterns: ['P1', 'P2', 'P8', 'P12', 'P16'],
    layers: ['Swab Results Database', 'Risk Assessment', 'QTR 1', 'Point List', 'Home'],
    build: buildT4
  };

  /* The neck and caeca toolkit is scenario S5 in fixtures.js. Give it the name
     it has on the shelf, and put the four toolkits first. */
  XT.data.scenarios.S5.title = 'Neck & Caeca labels toolkit';
  XT.data.scenarioOrder = ['T1', 'S5', 'T3', 'T4', 'S1', 'S4', 'S3', 'S2'];

})(window.XT = window.XT || {});
