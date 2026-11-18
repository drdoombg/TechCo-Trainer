/* =============================================================================
   Excel Trainer — trainer/levels.js
   The curriculum: catalog §11, levels 0-18, in dependency order.
   Each exercise is data. `setup` stands in for earlier layers so any exercise
   can be attempted on its own; `solution` is the model answer the grader reads
   expected values off, so the generated fixtures and the answers never drift.
   ============================================================================= */
(function (XT) {
  'use strict';
  XT.data = XT.data || {};

  var S1 = 'S1', S3 = 'S3', S4 = 'S4';

  // Reusable setup blocks.
  var MIRROR = { 'scanMODWB!A2:L24': '=IF(Samplesin!A2="","",Samplesin!A2)' };
  var KEYCOL = { 'MIMS DAILY WB OLD!A2:A24': '=IF(Samplesin!B2="","",Samplesin!B2)' };
  var FILTER = {
    'FILTERWBTODAY!B2:B24': '=IF(Samplesin!B2="","",Samplesin!B2)',
    'FILTERWBTODAY!C2:C24': '=IF(Samplesin!C2="","",Samplesin!C2)',
    'FILTERWBTODAY!E2:E24': '=IF(Samplesin!I2="","",Samplesin!I2)'
  };
  var FANOUT = {};
  for (var k in KEYCOL) FANOUT[k] = KEYCOL[k];
  FANOUT['MIMS DAILY WB OLD!E2:K24'] =
    '=IFERROR(_xlfn.XLOOKUP($A2,\'MIMS SAVE FILE\'!$M:$M,\'MIMS SAVE FILE\'!A:A),"")';

  function merge() {
    var out = {};
    for (var i = 0; i < arguments.length; i++) {
      var o = arguments[i];
      for (var key in o) out[key] = o[key];
    }
    return out;
  }

  XT.data.levels = [
    /* ------------------------------------------------------------------ 0 --- */
    { id: 0, title: 'The grid, and what a cell really holds',
      concepts: ['typing values', 'text vs number vs date', 'alignment as a tell'],
      intro: 'Before any formula: Excel decides what you meant the moment you press ' +
        'Enter, and it does not always agree with you. Numbers and dates go to the ' +
        'RIGHT of the cell, text goes LEFT. That alignment is the quickest way to spot ' +
        'a "number" that is secretly text \u2014 which is behind a large share of the ' +
        'broken lookups in your real files.',
      exercises: [
        { id: 'L0a', scenario: S1, sheet: 'Home', level: 0, layer: 5,
          title: 'Type two things and watch where they land',
          task: 'In Home, type 007 into E3 and 13/11/2026 into E4. Then look at which ' +
            'side of the cell each one sits on.',
          targets: ['E3', 'E4'],
          solution: { 'Home!E3': '007', 'Home!E4': '13/11/2026' },
          shapeRules: { mustBeFormula: false, mustFill: false },
          hints: ['Just type them in. No = needed.',
            '007 keeps its zeros, so Excel stores it as TEXT and left-aligns it.',
            '13/11/2026 becomes the NUMBER 46339, shown as a date, right-aligned.'],
          realWorldNote: 'Every product code with a leading zero in your files is text, ' +
            'and will never match a number, however identical they look.' },
        { id: 'L0b', scenario: S1, sheet: 'Home', level: 0, layer: 5,
          title: 'A date is a number',
          task: 'In E6, add 7 to the date in E4. You should get a date a week later.',
          setup: { 'Home!E4': '13/11/2026' },
          targets: ['E6'], solution: { 'Home!E6': '=E4+7' },
          shapeRules: { mustNotHardcode: true, mustFill: false },
          hints: ['Start with = and point at E4.', 'Just add 7.',
            'Dates are plain numbers, so arithmetic on them works.'],
          patterns: ['P19'] }
      ] },

    /* ------------------------------------------------------------------ 1 --- */
    { id: 1, title: 'Pointing at another sheet', patterns: ['P1'],
      concepts: ['cross-sheet references', 'quoted sheet names', 'fill down'],
      intro: 'The commonest formula in your whole estate: 867,462 cells that just say ' +
        '=OtherSheet!A2. It mirrors the pasted export onto a working sheet so that ' +
        're-pasting never breaks anything downstream. Boring, and load-bearing.',
      exercises: [
        { id: 'L1a', scenario: S1, sheet: 'scanMODWB', level: 1, layer: 2,
          title: 'Mirror one column',
          task: 'On scanMODWB, make A2 read column A of Samplesin, then fill it down ' +
            'to A24.',
          targets: ['A2:A24'], solution: { 'A2': '=Samplesin!A2' },
          shapeRules: { mustReference: ['Samplesin'], mustNotHardcode: true,
            requirePattern: 'P1' },
          hints: ['Type =, click the Samplesin tab, click A2, press Enter.',
            'Then drag the little square at the bottom-right of the cell down.',
            'Or select A2:A24 and press Ctrl+D.'],
          realWorldNote: 'Row 11 of Samplesin is blank \u2014 the export left a gap. ' +
            'Watch what your mirror shows there. That 0 is the reason level 2 exists.' },
        { id: 'L1b', scenario: S1, sheet: 'scanMODWB', level: 1, layer: 2,
          title: 'Mirror the whole block',
          task: 'Fill the mirror across all twelve columns and down to row 24.',
          targets: ['A2:L24'], solution: { 'A2': '=Samplesin!A2' },
          shapeRules: { mustReference: ['Samplesin'], mustNotHardcode: true },
          hints: ['Get A2 right first.', 'Fill right to L2, then select A2:L24 and Ctrl+D.',
            'No $ anywhere: every reference should move with the cell.'] }
      ] },

    /* ------------------------------------------------------------------ 2 --- */
    { id: 2, title: 'Showing nothing instead of noise', patterns: ['P2', 'P3'],
      concepts: ['IF', 'EMPTY versus ""', 'why a mirror shows zeros'],
      intro: 'A formula that reads an empty cell shows 0. Across a half-empty export ' +
        'that fills the sheet with zeros that print on labels and look like data. ' +
        '900,000 cells in your files exist to suppress them.',
      exercises: [
        { id: 'L2a', scenario: S1, sheet: 'scanMODWB', level: 2, layer: 2,
          title: 'Guard the mirror',
          task: 'Rebuild A2:A24 so a blank source row shows nothing instead of 0.',
          targets: ['A2:A24'],
          solution: { 'A2': '=IF(Samplesin!A2="","",Samplesin!A2)' },
          shapeRules: { mustUse: ['IF'], requirePattern: 'P3', mustNotHardcode: true },
          hints: ['=IF( the source is empty , "" , the source )',
            'Test it with ="" \u2014 that is true for a truly empty cell too.',
            'The whole formula is =IF(Samplesin!A2="","",Samplesin!A2)'],
          realWorldNote: 'The cell now holds empty TEXT, not nothing. ISBLANK says ' +
            'FALSE and COUNTA counts it. Select A11 and read the Inspector panel.' },
        { id: 'L2b', scenario: S1, sheet: 'scanMODWB', level: 2, layer: 2,
          title: 'The zero-suppressing version, and its trap',
          task: 'Mirror the Quantity column into E2:E24 using the >0 test that your ' +
            'real files use 472,402 times.',
          targets: ['E2:E24'],
          solution: { 'E2': '=IF(Samplesin!E2>0,Samplesin!E2,"")' },
          shapeRules: { mustUse: ['IF'], requirePattern: 'P2' },
          hints: ['=IF(source>0, source, "")',
            'Now look at rows 6, 9 and 16.',
            'Those quantities are TEXT, and in Excel any text is GREATER than any ' +
            'number \u2014 so the guard lets them straight through.'],
          defectLink: 'D5',
          realWorldNote: 'This is defect D5. The honest test is ="" , not >0.' }
      ] },

    /* ------------------------------------------------------------------ 3 --- */
    { id: 3, title: '$ \u2014 the anchor', patterns: ['P1'],
      concepts: ['relative, absolute and mixed references', 'F4'],
      intro: 'A bare A1 moves when you fill. $A$1 never moves. A$1 keeps its row, ' +
        '$A1 keeps its column. Press F4 while editing to cycle through all four.',
      exercises: [
        { id: 'L3a', scenario: S1, sheet: 'FILTERWBTODAY', level: 3, layer: 3,
          title: 'One setting, used by every row',
          task: 'In G2:G24, work out the due date: the scan date plus the Default due ' +
            'days from Home B8. Blank rows stay blank.',
          setup: FILTER, targets: ['G2:G24'],
          solution: { 'G2': '=IF(E2="","",INT(E2)+Home!$B$8)' },
          shapeRules: { mustUse: ['IF'], anchorPattern: '^\\$B\\$8$',
            anchorWhy: 'Home!B8 has to be locked as $B$8, or it slides down the Home ' +
              'sheet as you fill and starts reading blank cells.' },
          hints: ['INT(E2) strips the time off the timestamp.',
            'Add Home!B8 \u2014 then put the cursor on it and press F4.',
            '=IF(E2="","",INT(E2)+Home!$B$8)'] },
        { id: 'L3b', scenario: S1, sheet: 'MIMS DAILY WB OLD', level: 3, layer: 4,
          title: 'Lock a column, not a cell',
          task: 'Put the barcode in A2:A24, anchored so that it will survive being ' +
            'filled sideways later.',
          targets: ['A2:A24'],
          solution: { 'A2': '=IF(Samplesin!$B2="","",Samplesin!$B2)' },
          shapeRules: { mustUse: ['IF'], anchorPattern: '^\\$B\\d+$',
            anchorWhy: 'The column must be locked ($B2) and the row must not be, so ' +
              'the key follows the row down but stays on column B when you fill right.' },
          hints: ['Same guard as level 2.', 'Use $B2, not B2 and not $B$2.',
            'F4 three times gets you from B2 to $B2.'] }
      ] },

    /* ------------------------------------------------------------------ 4 --- */
    { id: 4, title: 'The join', patterns: ['P5'],
      concepts: ['VLOOKUP', 'whole-column blocks', 'exact match', 'col_index_num'],
      intro: '689,179 cells. The house style is a whole-column block and an explicit ' +
        'FALSE. The number counts from the FIRST COLUMN OF THE BLOCK, not from column A.',
      exercises: [
        { id: 'L4a', scenario: S1, sheet: 'MIMS DAILY WB OLD', level: 4, layer: 4,
          title: 'Look up the product',
          task: 'In B2:B24, find each barcode in Product Lookup and bring back the ' +
            'product name.',
          setup: KEYCOL, targets: ['B2:B24'],
          solution: { 'B2': '=VLOOKUP($A2,\'Product Lookup\'!$A:$F,2,FALSE)' },
          shapeRules: { mustUse: ['VLOOKUP'], mustNotHardcode: true,
            anchorPattern: '^\\$A:\\$F$' },
          hints: ['=VLOOKUP( what to find , where to look , which column , FALSE )',
            'The block is \'Product Lookup\'!$A:$F \u2014 whole columns, locked.',
            'The product name is the 2nd column of that block.'],
          realWorldNote: 'Three rows come back #N/A. Two of them are a real fault worth ' +
            'chasing \u2014 level 5 and the Inspector will show you which.' },
        { id: 'L4b', scenario: S1, sheet: 'MIMS DAILY WB OLD', level: 4, layer: 4,
          title: 'Count the columns properly',
          task: 'In C2:C24, bring back DUE_DAYS from the same table.',
          setup: KEYCOL, targets: ['C2:C24'],
          solution: { 'C2': '=VLOOKUP($A2,\'Product Lookup\'!$A:$F,5,FALSE)' },
          shapeRules: { mustUse: ['VLOOKUP'], mustNotHardcode: true },
          hints: ['Same formula, different column number.',
            'Count from A: A=1, B=2, C=3, D=4, E=5.', 'DUE_DAYS is column E, so 5.'] },
        { id: 'L4c', scenario: S1, sheet: 'MIMS DAILY WB OLD', level: 4, layer: 4,
          title: 'Find out why one row fails',
          task: 'Row 7 comes back #N/A even though the barcode is clearly there. ' +
            'Put =ISNUMBER on the two cells in N2 and N3 to prove why.',
          setup: KEYCOL, targets: ['N2', 'N3'],
          solution: { 'MIMS DAILY WB OLD!N2': '=ISNUMBER(\'Product Lookup\'!A7)',
            'MIMS DAILY WB OLD!N3': '=ISNUMBER(Samplesin!B7)' },
          shapeRules: { mustUse: ['ISNUMBER'], mustFill: false },
          hints: ['ISNUMBER tells you what a cell really holds.',
            'Check Product Lookup!A7 and Samplesin!B7.',
            'One is a number, one is text. They look identical and will never match.'],
          realWorldNote: 'This single mismatch is the most common cause of "my VLOOKUP ' +
            'is broken" anywhere.' }
      ] },

    /* ------------------------------------------------------------------ 5 --- */
    { id: 5, title: 'Guarding the join \u2014 and what guards hide',
      patterns: ['P5'], concepts: ['IFERROR', 'IFNA', '#REF! vs #N/A'],
      intro: 'IFERROR appears 750,917 times, IFNA 45. IFERROR catches EVERYTHING, ' +
        'including a formula that is structurally broken \u2014 which is how breakage ' +
        'stays hidden for years. IFNA catches only "not found".',
      exercises: [
        { id: 'L5a', scenario: S1, sheet: 'MIMS DAILY WB OLD', level: 5, layer: 4,
          title: 'Turn "not found" into a blank',
          task: 'Wrap the product lookup in B2:B24 so that a miss shows nothing.',
          setup: KEYCOL, targets: ['B2:B24'],
          solution: { 'B2': '=IFERROR(VLOOKUP($A2,\'Product Lookup\'!$A:$F,2,FALSE),"")' },
          shapeRules: { mustUse: ['IFERROR'], requirePattern: 'P5' },
          hints: ['=IFERROR( the lookup you already have , "" )',
            'Keep the inner formula exactly as it was.',
            'This is pattern P5, the most-used join in your files.'],
          realWorldNote: 'Now select B7 and read the Inspector. It will tell you what ' +
            'the guard swallowed \u2014 Excel would not.' },
        { id: 'L5b', scenario: S1, sheet: 'MIMS DAILY WB OLD', level: 5, layer: 4,
          title: 'Let a broken formula shout',
          task: 'D2:D24 asks for column 9 of a six-column block, which is wrong. ' +
            'Guard it with IFNA instead of IFERROR, so the #REF! still shows.',
          setup: KEYCOL, targets: ['D2:D24'],
          solution: { 'D2': '=IFNA(VLOOKUP($A2,\'Product Lookup\'!$A:$F,9,FALSE),"")' },
          shapeRules: { mustUse: ['IFNA'], mustNotUse: ['IFERROR'],
            mustNotUseWhy: 'IFERROR would hide the #REF! and the column mistake would ' +
              'never be found. That is the point of this exercise.' },
          hints: ['Same shape, IFNA instead of IFERROR.',
            'Yes, the cells will show #REF!. That is correct \u2014 the formula IS wrong.',
            'IFNA says "hide not-found, but let real breakage through".'],
          defectLink: 'D2' }
      ] },

    /* ------------------------------------------------------------------ 6 --- */
    { id: 6, title: 'The modern join', patterns: ['P10'],
      concepts: ['XLOOKUP', 'looking leftwards', 'fill right'],
      intro: 'MIMS SAVE FILE keys on column M, and everything useful is to its LEFT. ' +
        'VLOOKUP cannot reach backwards. That is the real reason your files moved to ' +
        'XLOOKUP \u2014 and then wrote it out 37 times by hand.',
      exercises: [
        { id: 'L6a', scenario: S1, sheet: 'MIMS DAILY WB OLD', level: 6, layer: 4,
          title: 'Look to the left',
          task: 'In E2:E24, get PRODUCT (column A of MIMS SAVE FILE) by matching the ' +
            'barcode against column M.',
          setup: KEYCOL, targets: ['E2:E24'],
          solution: { 'E2': '=IFERROR(_xlfn.XLOOKUP($A2,\'MIMS SAVE FILE\'!$M:$M,' +
            '\'MIMS SAVE FILE\'!A:A),"")' },
          shapeRules: { mustUse: ['XLOOKUP'], anchorPattern: '^\\$M:\\$M$',
            anchorWhy: 'The key column $M:$M must be locked; the return column must ' +
              'NOT be, so it can walk right in the next exercise.' },
          hints: ['=XLOOKUP( key , where to look , what to bring back )',
            'Lock the lookup column as $M:$M, leave the return as A:A.',
            'Wrap it in IFERROR(...,"") as usual.'] },
        { id: 'L6b', scenario: S1, sheet: 'MIMS DAILY WB OLD', level: 6, layer: 4,
          title: 'One formula, seven columns',
          task: 'Fill that XLOOKUP right across E2:K24 so it rebuilds seven columns of ' +
            'the record.',
          setup: KEYCOL, targets: ['E2:K24'],
          solution: { 'E2': '=IFERROR(_xlfn.XLOOKUP($A2,\'MIMS SAVE FILE\'!$M:$M,' +
            '\'MIMS SAVE FILE\'!A:A),"")' },
          shapeRules: { mustUse: ['XLOOKUP'], anchorPattern: '^\\$M:\\$M$' },
          hints: ['Get E2 right, then fill right to K2, then fill the block down.',
            'If the key moves sideways you locked the wrong thing.',
            '$A2 and $M:$M stay put; A:A becomes B:B, C:C, and so on.'],
          realWorldNote: 'In the real file this is 37 separate formulas, typed out. ' +
            'Anchored properly it is one.' }
      ] },

    /* ------------------------------------------------------------------ 7 --- */
    { id: 7, title: 'Two keys at once', patterns: ['P9'],
      concepts: ['& concatenation', 'array concatenation', 'helper columns'],
      intro: 'One column is often not enough to identify a row: barcode AND shift, ' +
        'week AND site. XLOOKUP can glue the keys together inline. VLOOKUP needs a ' +
        'physical helper column \u2014 and the old files are full of them.',
      exercises: [
        { id: 'L7a', scenario: S1, sheet: 'MIMS DAILY WB OLD', level: 7, layer: 4,
          title: 'Get the shift first',
          task: 'In L2:L24, bring back SHIFT (column H of MIMS SAVE FILE).',
          setup: KEYCOL, targets: ['L2:L24'],
          solution: { 'L2': '=IFERROR(_xlfn.XLOOKUP($A2,\'MIMS SAVE FILE\'!$M:$M,' +
            '\'MIMS SAVE FILE\'!$H:$H),"")' },
          shapeRules: { mustUse: ['XLOOKUP'] },
          hints: ['Same as level 6, return column H.', 'Lock both columns this time.',
            'You will need this shift in the next exercise.'] },
        { id: 'L7b', scenario: S1, sheet: 'MIMS DAILY WB OLD', level: 7, layer: 4,
          title: 'Join on barcode AND shift',
          task: 'In M2:M24, get OTHER_ID (column J) by matching barcode AND shift ' +
            'together: column M AND column V of MIMS SAVE FILE.',
          setup: merge(KEYCOL, { 'MIMS DAILY WB OLD!L2:L24':
            '=IFERROR(_xlfn.XLOOKUP($A2,\'MIMS SAVE FILE\'!$M:$M,\'MIMS SAVE FILE\'!$H:$H),"")' }),
          targets: ['M2:M24'],
          solution: { 'M2': '=IFERROR(_xlfn.XLOOKUP($A2&L2,\'MIMS SAVE FILE\'!$M:$M&' +
            '\'MIMS SAVE FILE\'!$V:$V,\'MIMS SAVE FILE\'!$J:$J),"")' },
          shapeRules: { mustUse: ['XLOOKUP'], requirePattern: 'P9' },
          hints: ['Glue the two keys with &: $A2&L2',
            'Glue the two lookup columns the same way: $M:$M&$V:$V',
            'XLOOKUP can take a computed array; VLOOKUP cannot.'],
          realWorldNote: 'Notice row 7 now works, where the single-key lookup failed. ' +
            '& turns the number into text, so the types finally agree. A real quirk.' }
      ] },

    /* ------------------------------------------------------------------ 8 --- */
    { id: 8, title: 'Counting what matters', patterns: ['P11', 'P13'],
      concepts: ['COUNTIF as an existence test', 'COUNTIFS', 'mixed anchoring'],
      intro: 'COUNTIFS is in 39 of your 152 workbooks \u2014 more than any other ' +
        'aggregator. Every KPI number is one of these.',
      exercises: [
        { id: 'L8a', scenario: S1, sheet: 'MIMS DAILY WB OLD', level: 8, layer: 4,
          title: 'Does this key exist at all?',
          task: 'In N2:N24, flag 1 when the barcode exists in Product Lookup, blank ' +
            'otherwise.',
          setup: KEYCOL, targets: ['N2:N24'],
          solution: { 'N2': '=IF($A2="","",IF(COUNTIF(\'Product Lookup\'!$A:$A,$A2)>=1,1,""))' },
          shapeRules: { mustUse: ['COUNTIF'], requirePattern: 'P11' },
          hints: ['COUNTIF counts matches; more than zero means it exists.',
            'Guard the blank rows first, as always.',
            'This is how the complaints workbook picks between three rival code systems.'] },
        { id: 'L8b', scenario: S1, sheet: 'Home', level: 8, layer: 5,
          title: 'A crosstab from one formula',
          task: 'Fill B12:D14 with the count of MIMS rows for each test (column L) and ' +
            'retailer (column G). One formula, filled both right and down.',
          targets: ['B12:D14'],
          solution: { 'Home!B12': '=COUNTIFS(\'MIMS SAVE FILE\'!$L:$L,$A12,' +
            '\'MIMS SAVE FILE\'!$G:$G,B$11)' },
          shapeRules: { mustUse: ['COUNTIFS'], requirePattern: 'P13' },
          hints: ['Two criteria pairs: test column then retailer column.',
            'The test name is in column A of this sheet; the retailer is in row 11.',
            '$A12 locks the column, B$11 locks the row. That is what lets ONE formula ' +
            'fill in both directions.'],
          realWorldNote: 'Get the $ wrong here and it works in the first cell only. ' +
            'This is the single highest-value lesson for KPI work.' }
      ] },

    /* ------------------------------------------------------------------ 9 --- */
    { id: 9, title: 'Totalling a window', patterns: ['P12', 'P20'],
      concepts: ['SUMIFS', 'criteria assembled with &', 'rates'],
      intro: 'The thing that catches everyone: a criterion is a STRING with an ' +
        'operator inside it, and it is usually built on the fly \u2014 ">=" & B3.',
      exercises: [
        { id: 'L9a', scenario: S1, sheet: 'Home', level: 9, layer: 5,
          title: 'Quantity by category',
          task: 'In B17:B20, total the mirrored Quantity column for each category.',
          setup: MIRROR, targets: ['B17:B20'],
          solution: { 'Home!B17': '=SUMIFS(scanMODWB!$E$2:$E$24,scanMODWB!$F$2:$F$24,$A17)' },
          shapeRules: { mustUse: ['SUMIFS|SUMIF'] },
          hints: ['SUMIFS( what to add , where to test , what to test for )',
            'Lock both ranges fully so they do not slide as you fill down.',
            'The category to match is in column A of this sheet.'],
          realWorldNote: 'The totals are lower than you would expect: the text ' +
            'quantities in rows 6, 9 and 16 are silently skipped. Text in a range is ' +
            'never added.' },
        { id: 'L9b', scenario: S1, sheet: 'Home', level: 9, layer: 5,
          title: 'A date window',
          task: 'In B22, count how many scans happened on the Today date in B3. ' +
            'The timestamps carry a time, so you need two bounds.',
          targets: ['B22'],
          solution: { 'Home!B22': '=COUNTIFS(Samplesin!$I$2:$I$24,">="&$B$3,' +
            'Samplesin!$I$2:$I$24,"<"&$B$3+1)' },
          shapeRules: { mustUse: ['COUNTIFS'], mustFill: false },
          hints: ['Two criteria on the same column: at least B3, and less than B3+1.',
            'Build each criterion with & : ">="&$B$3',
            'A timestamp is the date plus a fraction, so "<B3+1" means "before midnight".'] }
      ] },

    /* ----------------------------------------------------------------- 10 --- */
    { id: 10, title: 'Dates, weeks and shifts', patterns: ['P8', 'P19'],
      concepts: ['WEEKDAY', 'week-ending dates', 'TIME/HOUR/MINUTE'],
      intro: 'WEEKDAY appears 177,467 times, almost all inside one idiom that rounds a ' +
        'date to its week-ending Saturday. WEEKNUM appears ONCE in 152 workbooks, ' +
        'because retailer weeks are not ISO weeks \u2014 so the calendar has to be a table.',
      exercises: [
        { id: 'L10a', scenario: S1, sheet: 'FILTERWBTODAY', level: 10, layer: 3,
          title: 'Which side of the shift cut-off?',
          task: 'In F2:F24, say NIGHT if the scan time is before the cut-off in ' +
            'Home B4, otherwise DAY.',
          setup: FILTER, targets: ['F2:F24'],
          solution: { 'F2': '=IF(E2="","",IF(TIME(HOUR(E2),MINUTE(E2),SECOND(E2))<' +
            'Home!$B$4,"NIGHT","DAY"))' },
          shapeRules: { mustUse: ['TIME'] },
          hints: ['You cannot compare a full timestamp to a time of day.',
            'Rebuild just the time: TIME(HOUR(E2),MINUTE(E2),SECOND(E2))',
            'Then compare it to Home!$B$4, locked.'] },
        { id: 'L10b', scenario: S3, sheet: 'Ecoli Final Data', level: 10, layer: 4,
          title: 'The week-ending date',
          task: 'In V2:V49, find each sample\u2019s week label by rounding its date to ' +
            'the week-ending Saturday and looking that up in Calendars.',
          setup: { 'Ecoli Final Data!C2:C49': '=\'ECOLI MICRO PASTE\'!C2' },
          targets: ['V2:V49'],
          solution: { 'V2': '=IF(C2="","",VLOOKUP((C2-WEEKDAY(C2,1))+7,' +
            'Calendars!$H:$J,3,FALSE))' },
          shapeRules: { mustUse: ['WEEKDAY'], requirePattern: 'P8' },
          hints: ['(date - WEEKDAY(date,1)) + 7 gives the Saturday that ends the week.',
            'Calendars column H holds those Saturdays.',
            'The week label is the 3rd column of the block H:J.'],
          realWorldNote: 'Reach for WEEKNUM here and you will get the wrong week.' }
      ] },

    /* ----------------------------------------------------------------- 11 --- */
    { id: 11, title: 'Cutting up text', patterns: ['P6', 'P7'],
      concepts: ['LEFT/RIGHT', 'TEXTBEFORE/TEXTAFTER', 'VALUE', 'CHAR(160)'],
      intro: 'Your Technical folder is 204,153 RIGHTs and 200,242 LEFTs, because a ' +
        'vision system emits one line of text per pack. The same workbook also ' +
        'contains the modern answer. Seeing both is the best lesson in the catalog.',
      exercises: [
        { id: 'L11a', scenario: S4, sheet: 'Sheet4', level: 11, layer: 2,
          title: 'Classify by counting characters',
          task: 'In C2:C40, mark a log line Missing Sleeve or Mismatch using RIGHT, ' +
            'exactly as the real file does.',
          targets: ['C2:C40'],
          solution: { 'C2': '=IF(RIGHT(LVALL!C2,29)="FLASHLABEL; Invalid: Mismatch",' +
            '"Missing Sleeve",IF(RIGHT(LVALL!C2,8)="Mismatch","Mismatch",""))' },
          shapeRules: { mustUse: ['RIGHT'], requirePattern: 'P6' },
          hints: ['RIGHT(text, 29) takes the last 29 characters.',
            'Test for the long message first, then the short one.',
            'Anything else gets "".'],
          defectLink: 'D6',
          realWorldNote: 'The 29 and the 8 were counted by hand. The day the vendor ' +
            'rewords a message, 400,000 cells stop classifying and nobody notices.' },
        { id: 'L11b', scenario: S4, sheet: 'Sheet4', level: 11, layer: 2,
          title: 'Count fields instead',
          task: 'In H2:H40, pull the CHECKED count out of the detail lines using ' +
            'TEXTAFTER and TEXTBEFORE, and turn it into a real number.',
          targets: ['H2:H40'],
          solution: { 'H2': '=IFERROR(VALUE(_xlfn.TEXTBEFORE(' +
            '_xlfn.TEXTAFTER(LVALL!C2," ",3)," ",1)),"")' },
          shapeRules: { mustUse: ['TEXTAFTER'], requirePattern: 'P7' },
          hints: ['TEXTAFTER(line," ",3) gives everything after the 3rd space.',
            'TEXTBEFORE(that," ",1) takes the first field off that.',
            'Slicing text gives TEXT, so wrap the lot in VALUE() \u2014 and IFERROR for ' +
            'the lines that have no 3rd space.'] },
        { id: 'L11c', scenario: S1, sheet: 'scanMODWB', level: 11, layer: 2,
          title: 'The invisible character',
          task: 'Samplesin E9 looks like a number and will not add up. In O2, turn it ' +
            'into a real number.',
          targets: ['O2'],
          solution: { 'scanMODWB!O2': '=VALUE(SUBSTITUTE(Samplesin!E9,CHAR(160),""))' },
          shapeRules: { mustUse: ['SUBSTITUTE'], mustFill: false },
          hints: ['Try =VALUE(Samplesin!E9) first and read the error message.',
            'There is a non-breaking space on the end. TRIM will not touch it.',
            'SUBSTITUTE(x,CHAR(160),"") then VALUE.'],
          defectLink: 'D7' }
      ] },

    /* ----------------------------------------------------------------- 12 --- */
    { id: 12, title: 'Banding, and when to stop nesting', patterns: ['P18'],
      concepts: ['nested IF', 'refactoring a ladder into a table'],
      intro: '88,814 cells map a category to a value through nested IFs, up to six ' +
        'deep. It works. It is also a lookup table written as code.',
      exercises: [
        { id: 'L12a', scenario: S1, sheet: 'MIMS DAILY WB OLD', level: 12, layer: 4,
          title: 'The IF ladder',
          task: 'In O2:O24, turn the retailer in column K into a region: Northgate is ' +
            'North, Westvale is West, Harbour is Coast, anything else is Inland.',
          setup: FANOUT, targets: ['O2:O24'],
          solution: { 'O2': '=IF(K2="","",IF(K2="Northgate","North",IF(K2="Westvale",' +
            '"West",IF(K2="Harbour","Coast","Inland"))))' },
          shapeRules: { mustUse: ['IF'] },
          hints: ['Guard the blanks, then one IF per retailer.',
            'The last branch is the catch-all, so it needs no test.',
            'Watch the closing brackets \u2014 one per IF.'] },
        { id: 'L12b', scenario: S1, sheet: 'MIMS DAILY WB OLD', level: 12, layer: 4,
          title: 'The same thing as a table',
          task: 'Replace the ladder in O2:O24 with a lookup against the REF sheet. ' +
            'Same answers, no IF at all.',
          setup: FANOUT, targets: ['O2:O24'],
          solution: { 'O2': '=IFNA(_xlfn.XLOOKUP(K2,REF!$A:$A,REF!$B:$B),"")' },
          shapeRules: { mustUse: ['XLOOKUP|VLOOKUP'], mustNotUse: ['IF'],
            mustNotUseWhy: 'The whole point is that the table replaces the ladder. ' +
              'Use IFNA for the blank rows instead of an IF.' },
          hints: ['REF has retailer in A and region in B.',
            'A blank retailer will not be found, so guard with IFNA(...,"").',
            'Adding a sixth retailer is now a row in a table, not a formula edit.'],
          defectLink: 'D4' }
      ] },

    /* ----------------------------------------------------------------- 13 --- */
    { id: 13, title: 'Making the sheet check itself', patterns: ['P17'],
      concepts: ['signature columns', 'verdict strings'],
      intro: 'Your files do this everywhere: build a signature, compare it to the ' +
        'expected one, and print a verdict for the person who has to fix it. ' +
        '"PLEASE REVIEW AND AMEND", not #N/A.',
      exercises: [
        { id: 'L13a', scenario: S1, sheet: 'MIMS DAILY WB OLD', level: 13, layer: 5,
          title: 'Do the two lookups agree?',
          task: 'Column B came from Product Lookup, column E from MIMS SAVE FILE. ' +
            'In Q2:Q24, say DATA MATCH when they agree and PLEASE REVIEW AND AMEND ' +
            'when they do not.',
          setup: merge(FANOUT, { 'MIMS DAILY WB OLD!B2:B24':
            '=IFERROR(VLOOKUP($A2,\'Product Lookup\'!$A:$F,2,FALSE),"")' }),
          targets: ['Q2:Q24'],
          solution: { 'Q2': '=IF($A2="","",IF(B2=E2,"DATA MATCH","PLEASE REVIEW AND AMEND"))' },
          shapeRules: { mustUse: ['IF'], requirePattern: 'P17' },
          hints: ['Guard the blank rows, then compare B2 with E2.',
            'Two sources that should agree, and a sentence for when they do not.',
            'Row 7 will disagree \u2014 that is the numeric-barcode row from level 4.'],
          realWorldNote: 'Treat adding a self-check column as part of the job, not an ' +
            'extra. It is the cheapest quality control there is.' }
      ] },

    /* -------------------------------------------------------------- 14, 15 -- */
    { id: 14, title: 'Colour as output', v2: true,
      concepts: ['conditional formatting'],
      intro: '292,327 conditional-format rules in your files, and colour scales are ' +
        'barely used \u2014 it is nearly all cellIs/between with TEXT criteria like ' +
        '"<500", plus formula rules such as $J16="Removed". A result sheet that is not ' +
        'red/amber/green has not done its job. Not built yet: this is the v2 list.',
      exercises: [] },
    { id: 15, title: 'Dropdowns and picklists', v2: true,
      concepts: ['data validation', 'defined names'],
      intro: '159 list rules across your workbooks, nearly all bound to a picklists ' +
        'sheet through a defined name (Shift, Failure_type, Water). The engine already ' +
        'resolves those names \u2014 try =COUNTA(Shift) in the sandbox. The dropdown UI ' +
        'itself is on the v2 list.',
      exercises: [] },

    /* ----------------------------------------------------------------- 16 --- */
    { id: 16, title: 'INDEX, MATCH, and the nth match', patterns: ['P14', 'P15', 'P16'],
      concepts: ['expanding ranges', 'array formulas', 'ROW as a generator'],
      intro: 'VLOOKUP only ever finds the FIRST match. To reach the second and third ' +
        'sample of the same product, your files manufacture a unique key \u2014 or use ' +
        'the INDEX/SMALL/IF array formula, which is the ceiling of the estate\u2019s ' +
        'complexity.',
      exercises: [
        { id: 'L16a', scenario: S1, sheet: 'FINALEMAIL WB', level: 16, layer: 5,
          title: 'Number the repeats',
          task: 'In B2:B24, turn each barcode into a unique key by numbering its ' +
            'repeats: barcode-1, barcode-2, and so on.',
          setup: { 'FINALEMAIL WB!A2:A24': '=IF(Samplesin!B2="","",Samplesin!B2)' },
          targets: ['B2:B24'],
          solution: { 'B2': '=IF($A2="","",$A2&"-"&COUNTIF($A$2:$A2,$A2))' },
          shapeRules: { mustUse: ['COUNTIF'], requirePattern: 'P14' },
          hints: ['COUNTIF over the range that GROWS as you fill: $A$2:$A2',
            'Work out why those particular $ signs make it grow one way only.',
            'Then glue it on with & .'] },
        { id: 'L16b', scenario: S1, sheet: 'FINALEMAIL WB', level: 16, layer: 5,
          title: 'Get the nth match',
          task: 'In C2:C24, bring back the scan TIME for each occurrence \u2014 the 1st ' +
            'scan of a barcode, then the 2nd, then the 3rd. Enter it with ' +
            'Ctrl+Shift+Enter.',
          setup: { 'FINALEMAIL WB!A2:A24': '=IF(Samplesin!B2="","",Samplesin!B2)' },
          targets: ['C2:C24'],
          solution: { 'C2': { f: '=IFERROR(INDEX(Samplesin!$I$2:$I$24,SMALL(IF(' +
            'Samplesin!$B$2:$B$24=$A2,ROW(Samplesin!$I$2:$I$24)-1),' +
            'COUNTIF($A$2:$A2,$A2))),"")', cse: true } },
          shapeRules: { mustUse: ['SMALL'], mustBeArray: true, requirePattern: 'P15' },
          hints: ['IF(range=key, ROW(range)-1) gives the positions of every match.',
            'SMALL(..., k) picks the kth of them, and COUNTIF on the expanding range ' +
            'supplies k.',
            'Press Ctrl+Shift+Enter, not Enter \u2014 without it the ranges collapse to ' +
            'one row and the answer is quietly wrong.'] },
        { id: 'L16c', scenario: S3, sheet: 'Ecoli Final Data', level: 16, layer: 4,
          title: 'Which week does this date fall inside?',
          task: 'Calendars stores start and end dates, not a key. In AB2:AB49, find the ' +
            'week whose range contains each sample date. Ctrl+Shift+Enter.',
          setup: { 'Ecoli Final Data!C2:C49': '=\'ECOLI MICRO PASTE\'!C2' },
          targets: ['AB2:AB49'],
          solution: { 'AB2': { f: '=INDEX(Calendars!$F$2:$F$57,MATCH(1,' +
            '(C2>=Calendars!$G$2:$G$57)*(C2<=Calendars!$H$2:$H$57),0))', cse: true } },
          shapeRules: { mustUse: ['MATCH'], mustBeArray: true, requirePattern: 'P16' },
          hints: ['(date>=starts) gives an array of TRUE/FALSE; so does (date<=ends).',
            'Multiply them: TRUE*TRUE is 1, so MATCH(1,...,0) finds the straddling row.',
            'Then INDEX the week labels by that position.'] }
      ] },

    /* ----------------------------------------------------------------- 17 --- */
    { id: 17, title: 'Building the five layers', patterns: ['P1', 'P3', 'P5'],
      concepts: ['workbook architecture', 'one job per sheet'],
      intro: 'This is the real skill. Landing, mirror, filter, workings, output \u2014 ' +
        'each sheet doing one job and pointing only at the one before it. It is what ' +
        'makes a workbook survivable, and it is graded structurally here, not cell by cell.',
      exercises: [
        { id: 'L17a', scenario: S1, sheet: 'scanMODWB', level: 17, layer: 0,
          mode: 'layers',
          title: 'The whole chain, properly layered',
          task: 'Fill in all five layers so each one reads only from the layer before ' +
            'it: Samplesin \u2192 scanMODWB \u2192 FILTERWBTODAY \u2192 MIMS DAILY WB ' +
            'OLD \u2192 FINALEMAIL WB. Reference sheets (MIMS SAVE FILE, Product Lookup, ' +
            'Calendars, REF) may be used from anywhere.',
          layerOrder: ['Samplesin', 'scanMODWB', 'FILTERWBTODAY', 'MIMS DAILY WB OLD',
            'FINALEMAIL WB'],
          solution: merge(MIRROR, {
            'FILTERWBTODAY!B2:B24': '=IF(scanMODWB!B2="","",scanMODWB!B2)',
            'FILTERWBTODAY!E2:E24': '=IF(scanMODWB!I2="","",scanMODWB!I2)',
            'MIMS DAILY WB OLD!A2:A24': '=IF(FILTERWBTODAY!$B2="","",FILTERWBTODAY!$B2)',
            'MIMS DAILY WB OLD!E2:E24': '=IFERROR(_xlfn.XLOOKUP($A2,' +
              '\'MIMS SAVE FILE\'!$M:$M,\'MIMS SAVE FILE\'!A:A),"")',
            'FINALEMAIL WB!A2:A24': '=IF(\'MIMS DAILY WB OLD\'!A2="","",' +
              '\'MIMS DAILY WB OLD\'!A2)',
            'FINALEMAIL WB!D2:D24': '=IF(\'MIMS DAILY WB OLD\'!E2="","",' +
              '\'MIMS DAILY WB OLD\'!E2)'
          }),
          hints: ['Start at the top and work down one sheet at a time.',
            'If a formula on FINALEMAIL WB reads Samplesin directly, it has skipped ' +
            'three layers \u2014 that is the thing being graded.',
            'Press "Show me" if you want the model answer put in for you.'] }
      ] },

    /* ----------------------------------------------------------------- 18 --- */
    { id: 18, title: 'Debugging someone else\u2019s workbook',
      concepts: ['the error inspector', 'tracing', 'reading a formula critically'],
      intro: 'Five real defects from the catalog, seeded into one sheet. Use the ' +
        'Inspector, the Stepper and Trace precedents. Ctrl+` shows all the formulas ' +
        'at once.',
      exercises: [
        { id: 'L18a', scenario: S1, sheet: 'MIMS DAILY WB OLD', level: 18, layer: 4,
          mode: 'defects',
          title: 'Find and fix five faults',
          task: 'Each cell below is wrong in a way that is in your real files. Fix all ' +
            'five.',
          setup: merge(KEYCOL, {
            'MIMS DAILY WB OLD!D2:D24':
              '=IFERROR(VLOOKUP($A2,\'Product Lookup\'!$A:$F,9,FALSE),"")',
            'scanMODWB!A2:A24': '=IFERROR(IF(Samplesin!A2="","",Samplesin!A2),"")',
            'scanMODWB!E2:E24': '=IF(Samplesin!E2>0,Samplesin!E2,"")',
            'MIMS DAILY WB OLD!O2:O24':
              '=IF(K2="Northgate","North",IF(K2="Westvale","West",))',
            'Home!E10': '=IF(B3=0/1/1900,"epoch","not the epoch")'
          }),
          defects: [
            { id: 'D1', cell: 'MIMS DAILY WB OLD!D2', test: 'no-hidden-error',
              title: 'an IFERROR hiding a #REF!',
              hint: 'Select D2 and read the Inspector. The column index is past the end ' +
                'of the block. Fix the index, or swap IFERROR for IFNA so it shows.',
              fixedNote: 'Nothing is being swallowed any more.' },
            { id: 'D12', cell: 'scanMODWB!A2', test: 'formula-lacks', text: 'IFERROR',
              title: 'a redundant wrapper, copied 329,778 times',
              hint: 'A mirror of a cell cannot produce an error. Take the IFERROR off ' +
                'and leave the blank guard.',
              fixedNote: 'One less thing to read.' },
            { id: 'D5', cell: 'scanMODWB!E2', test: 'formula-lacks', text: '>0',
              title: 'a numeric guard on text data',
              hint: 'Text is greater than any number in Excel, so >0 lets the text ' +
                'quantities straight through. Test for ="" instead.',
              fixedNote: 'The guard now means what it looks like it means.' },
            { id: 'D4', cell: 'MIMS DAILY WB OLD!O4', test: 'not-false',
              title: 'an empty final argument',
              hint: 'The ladder ends with ,)) \u2014 an argument written as nothing, ' +
                'which returns 0 and prints on the report. Give it a real catch-all ' +
                'such as "Inland".',
              fixedNote: 'Unmatched retailers now say something sensible.' },
            { id: 'D3', cell: 'Home!E10', test: 'formula-lacks', text: '0/1/1900',
              title: '0/1/1900 is division, not a date',
              hint: 'Excel reads 0/1/1900 as 0 divided by 1 divided by 1900, which is 0. ' +
                'Use DATE(1900,1,1) if you really mean the epoch.',
              fixedNote: 'It is a date comparison now.' }
          ],
          hints: ['Press Ctrl+` to see every formula at once.',
            'Work through the checklist below one at a time; each says which cell.',
            'The Inspector panel explains any error in plain words, including the ones ' +
            'a guard has hidden.'] }
      ] }
  ];

  XT.data.exercises = [];
  XT.data.levels.forEach(function (L) {
    (L.exercises || []).forEach(function (ex) {
      ex.level = L.id;
      ex.levelTitle = L.title;
      if (!ex.patterns) ex.patterns = L.patterns || [];
      XT.data.exercises.push(ex);
    });
  });

  XT.data.exercise = function (id) {
    for (var i = 0; i < XT.data.exercises.length; i++) {
      if (XT.data.exercises[i].id === id) return XT.data.exercises[i];
    }
    return null;
  };

})(window.XT = window.XT || {});
