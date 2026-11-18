/* =============================================================================
   Excel Trainer — trainer/formulas.js
   LEARN BY FORMULA.

   The course in levels.js teaches a pipeline: landing, mirror, filter, workings,
   output. This file teaches one function at a time, on the same four worksheets,
   in the order the four toolkits actually use them:

     IF       297,488    VLOOKUP   34,289    IFERROR  33,455    SUMIF  13,104
     COUNTIF    8,421    SUM        3,886    MIN/MAX   1,339    MATCH     822
     COUNT        753    COUNTIFS     681    INDEX       652    AND/OR    129
     ISERROR      121    TEXTAFTER     69    VALUE        56    XLOOKUP    42
     RIGHT         34    AVERAGE       26    ROUND         6

   Those counts are measured: they are how many cells in the shelf life, neck
   and caeca, nutritional and micro schedule workbooks call each function. The
   order of this track is the order of that list, so the time you spend lines
   up with the formulas you will actually meet.

   Each lesson is: what it does, what it does IN YOUR FILES, the trap, and one
   or two graded drills on a real toolkit sheet. Every drill is a normal
   exercise object, so the same grader applies: values AND shape, and a formula
   that works in the first cell but breaks when filled still fails.
   ============================================================================= */
(function (XT) {
  'use strict';
  XT.data = XT.data || {};

  var drills = [];

  /* A compact way to write a drill. `s` is the sheet, `t` the targets,
     `sol` the model answer, `set` the setup that stands in for earlier work. */
  function drill(o) {
    o.track = 'formula';
    o.level = null;
    drills.push(o);
    return o;
  }

  /* ------------------------------------------------------------- setup blocks */

  /* T4 Risk Assessment: the four scores and the total. Several lessons need the
     sheet in this state without making you rebuild it. */
  var RA_SCORES = {
    'Risk Assessment!H4:H18': '=IF(D4="Y",2,0)',
    'Risk Assessment!I4:I18': '=IF(E4="Y",2,0)',
    'Risk Assessment!J4:J18': '=IF(F4="Y",2,0)',
    'Risk Assessment!K4:K18': '=IF(G4="Y",1,0)',
    'Risk Assessment!L4:L18': '=SUM(H4:K4)'
  };
  var RA_BAND = merge(RA_SCORES, {
    'Risk Assessment!M4:M18':
      '=IF(L4>=6,"High",IF(L4>=4,"Medium",IF(L4>=2,"Low","No Risk")))'
  });

  /* T1: the sample id for each test point, and the resolved raw results. */
  var T1_IDS = {
    'MICRO MOD!C2:C9': '=XLOOKUP($B2,\'SET MOD\'!$H$2:$H$9,\'SET MOD\'!$A$2:$A$9)'
  };
  var T1_RAW = merge(T1_IDS, {
    'MICRO MOD!D2:D9':
      '=IFERROR(IF(VLOOKUP($C2,\'MICRO PASTE\'!$A$2:$I$10,5,FALSE)=".",' +
      'VLOOKUP($C2,\'MICRO PASTE\'!$A$2:$I$10,4,FALSE),' +
      'VLOOKUP($C2,\'MICRO PASTE\'!$A$2:$I$10,5,FALSE)),"")',
    'MICRO MOD!F2:F9': '=IFERROR(VLOOKUP($C2,\'MICRO PASTE\'!$A$2:$I$10,6,FALSE),"")'
  });
  var T1_NUM = merge(T1_RAW, {
    'MICRO MOD!E2:E9': '=IFERROR(VALUE(IFERROR(TEXTAFTER($D2,"<"),$D2)),' +
      'VALUE(IFERROR(TEXTAFTER($D2,">"),$D2)))',
    'MICRO MOD!G2:G9': '=IFERROR(VALUE(IFERROR(TEXTAFTER($F2,"<"),$F2)),' +
      'VALUE(IFERROR(TEXTAFTER($F2,">"),$F2)))'
  });
  var T1_LIMITS = {
    'Test Profile 1!C7': '=VLOOKUP($C$3&"|TVC",COP!$A$2:$F$7,4,FALSE)',
    'Test Profile 1!C8': '=VLOOKUP($C$3&"|Yeast & Mould",COP!$A$2:$F$7,4,FALSE)'
  };
  var T1_PROFILE = merge(T1_NUM, T1_LIMITS, {
    'Test Profile 1!D12:D19': '=\'MICRO MOD\'!$E2',
    'Test Profile 1!E12:E19': '=\'MICRO MOD\'!$G2'
  });

  /* T3: the hidden workings sheet, and the averages on the front sheet. */
  var T3_BG = { 'background!B2:B11': '=IFERROR(VLOOKUP($A2,PASTE!$A:$E,2,FALSE),"")' };
  var T3_NUM = merge(T3_BG, {
    'background!C2:C11': '=VALUE(IFERROR(TEXTAFTER($B2,"<"),$B2))'
  });
  var T3_AVG = { '2sfg UKP Nutritional Toolkit!F10:F19': '=AVERAGE($C10:$E10)' };
  var T3_DIFF = merge(T3_AVG, {
    '2sfg UKP Nutritional Toolkit!H10:H19': '=IF(ISERROR($F10-$B10),"-",$F10-$B10)',
    '2sfg UKP Nutritional Toolkit!I10:I19': '=IFERROR(ROUND(($F10-$B10)/$B10*100,1),"")'
  });

  /* T4 QTR 1: the week columns counted up. */
  var Q1_COUNTS = {
    'QTR 1!K22:W22': '=COUNT(K$6:K$20)',
    'QTR 1!K23:W23': '=COUNTIF(K$6:K$20,2)'
  };

  function merge() {
    var out = {};
    for (var i = 0; i < arguments.length; i++) {
      var o = arguments[i];
      for (var k in o) out[k] = o[k];
    }
    return out;
  }

  /* ================================================================ LESSONS == */

  XT.data.formulaTrack = [

    /* ------------------------------------------------------------------ IF -- */
    { fn: 'IF', title: 'IF — the one that does everything',
      cells: 297488, scenario: 'T4',
      what: 'IF(test, what to do if true, what to do if false). It is three ' +
        'arguments and that is all there is to it.',
      estate: 'Six out of every seven formulas in your four toolkits are an IF or ' +
        'contain one. Two jobs dominate: turning a Y/N answer into a number, and ' +
        'stopping a formula showing 0 when its source row is empty.',
      trap: 'Excel does not stop you writing IF(A1="Y",2,0) when A1 holds "Y " with ' +
        'a trailing space. The test is simply false and you get 0, silently.',
      drills: [
        drill({ id: 'F-IF-1', fn: 'IF', scenario: 'T4', sheet: 'Risk Assessment',
          title: 'Turn a Y into a score',
          task: 'Score the FOOD CONTACT answer in H4: a "Y" is worth 2 points, ' +
            'anything else is worth 0. Fill it down to H18.',
          targets: ['H4:H18'],
          solution: { 'H4:H18': '=IF(D4="Y",2,0)' },
          shapeRules: { mustUse: ['IF'] },
          hints: ['=IF( the test , value if true , value if false )',
            'The test is D4="Y" — note the quotes around Y.',
            'The whole thing is =IF(D4="Y",2,0)'],
          realWorldNote: 'This exact shape appears 822 times in the real Risk ' +
            'Assessment sheet, once per question per point.' }),

        drill({ id: 'F-IF-2', fn: 'IF', scenario: 'T4', sheet: 'QTR 1',
          title: 'The guard: show nothing, not zero',
          task: 'In D6, bring the risk band across from the Risk Assessment sheet ' +
            '— but if that row has no total yet, show nothing at all rather than 0. ' +
            'Fill down to D20.',
          setup: RA_BAND,
          targets: ['D6:D20'],
          solution: { 'D6:D20': '=IF(\'Risk Assessment\'!L4="","",\'Risk Assessment\'!M4)' },
          shapeRules: { mustUse: ['IF'], mustReference: ['Risk Assessment'],
            mustNotHardcode: true },
          hints: ['Point at the other sheet by typing =IF( then clicking its tab.',
            'Test the TOTAL in column L, but return the BAND in column M.',
            '=IF(\'Risk Assessment\'!L4="","",\'Risk Assessment\'!M4)'],
          realWorldNote: 'The real QTR 1 sheet holds 2,740 copies of this one ' +
            'formula. It is the single most repeated thing in the schedule workbook.' })
      ] },

    /* ----------------------------------------------------------- nested IF -- */
    { fn: 'IF nested', title: 'IF inside IF — the ladder',
      cells: 41501, scenario: 'T4',
      what: 'The false branch of an IF can be another IF. Read it as "otherwise, ' +
        'if…". Stop at the first test that is true.',
      estate: 'Every band, grade and verdict in the estate is one of these. The ' +
        'order matters: the first true test wins, so the tests have to run from ' +
        'the most extreme down.',
      trap: 'A ladder that ends with ,)) returns 0 for everything that fell ' +
        'through, and 0 prints. Always finish with a real catch-all.',
      drills: [
        drill({ id: 'F-NIF-1', fn: 'IF nested', scenario: 'T4', sheet: 'Risk Assessment',
          title: 'Score to band',
          task: 'In M4 turn the total in L4 into a band: 6 or more is "High", ' +
            '4 or 5 is "Medium", 2 or 3 is "Low", below that is "No Risk". ' +
            'Fill down to M18.',
          setup: RA_SCORES,
          targets: ['M4:M18'],
          solution: { 'M4:M18':
            '=IF(L4>=6,"High",IF(L4>=4,"Medium",IF(L4>=2,"Low","No Risk")))' },
          shapeRules: { mustUse: ['IF'] },
          hints: ['Start from the top band and work down.',
            'Each "otherwise" is another IF in the third slot.',
            '=IF(L4>=6,"High",IF(L4>=4,"Medium",IF(L4>=2,"Low","No Risk")))'],
          realWorldNote: 'The real sheet runs this ladder six deep. Six is about ' +
            'where it stops being readable — past that, use a lookup table.' }),

        drill({ id: 'F-NIF-2', fn: 'IF nested', scenario: 'T1', sheet: 'Test Profile 1',
          title: 'One setting changes the whole study',
          task: 'C6 is the multiplier. If the Test type in C5 is "Challenge test" ' +
            'it is 2, otherwise it is 1.',
          targets: ['C6'],
          solution: { 'C6': '=IF($C$5="Challenge test",2,1)' },
          shapeRules: { mustUse: ['IF'], mustFill: false },
          hints: ['Only one cell. Point at C5.',
            'Lock it with $ — F4 cycles the dollars — because 61 other cells will ' +
            'read this one.',
            '=IF($C$5="Challenge test",2,1)'],
          realWorldNote: 'In the real file this is O37, and 61 cells depend on it. ' +
            'One typo in C5 and the whole profile doubles.' }),

        drill({ id: 'F-NIF-3', fn: 'IF nested', scenario: 'T1', sheet: 'MICRO MOD',
          title: 'The paired result columns',
          task: 'The lab writes the usable result in column D of MICRO PASTE, ' +
            'unless column E has something — then E wins, and D holds ".". ' +
            'In D2, fetch whichever one is real. Fill down to D9.',
          setup: T1_IDS,
          targets: ['D2:D9'],
          solution: { 'D2:D9':
            '=IFERROR(IF(VLOOKUP($C2,\'MICRO PASTE\'!$A$2:$I$10,5,FALSE)=".",' +
            'VLOOKUP($C2,\'MICRO PASTE\'!$A$2:$I$10,4,FALSE),' +
            'VLOOKUP($C2,\'MICRO PASTE\'!$A$2:$I$10,5,FALSE)),"")' },
          shapeRules: { mustUse: ['IF'], mustReference: ['MICRO PASTE'] },
          hints: ['Look up column 5 first and test it against ".".',
            'If it IS ".", you want column 4. If it is not, you want column 5.',
            'So the same VLOOKUP appears three times: once to test, once for each ' +
            'branch. Wrap the lot in IFERROR so a missing sample gives "".'],
          realWorldNote: 'This is formula G10 in the real MICRO MOD, copied 42 ' +
            'times. Yes, it looks up the same value twice. That is what the file does.' })
      ] },

    /* ------------------------------------------------------------- VLOOKUP -- */
    { fn: 'VLOOKUP', title: 'VLOOKUP — fetch a row by its key',
      cells: 34289, scenario: 'T3',
      what: 'VLOOKUP(what to find, where to look, which column to bring back, FALSE). ' +
        'It searches the FIRST column of the block and counts columns from there.',
      estate: '34,289 cells across the four toolkits. Almost always with FALSE, ' +
        'almost always wrapped in IFERROR, and almost always asking for column 2 — ' +
        'because the exports are key-then-answer.',
      trap: 'Three traps, in order of how often they bite. FALSE missing: you get ' +
        'the nearest match, which is a wrong answer that looks right. Key column ' +
        'not first: #N/A forever. Key is text and the lookup value is a number, ' +
        'or vice versa: #N/A, however identical they look on screen.',
      drills: [
        drill({ id: 'F-VL-1', fn: 'VLOOKUP', scenario: 'T3',
          sheet: '2sfg UKP Nutritional Toolkit',
          title: 'Name the product from its sample ref',
          task: 'C4 should show the product name for the sample ref in C3, ' +
            'looked up in the PASTE sheet.',
          targets: ['C4'],
          solution: { 'C4': '=VLOOKUP($C$3,PASTE!$A:$E,2,FALSE)' },
          shapeRules: { mustUse: ['VLOOKUP'], mustFill: false, mustNotHardcode: true,
            mustReference: ['PASTE'] },
          hints: ['PASTE is hidden. Use the sheet tabs — hidden sheets still show ' +
            'in this trainer, greyed out.',
            'The key is column A and the answer is column B, so the column number is 2.',
            '=VLOOKUP($C$3,PASTE!$A:$E,2,FALSE)'],
          realWorldNote: 'The real formula is =IFERROR(VLOOKUP(C18,PASTE!$A:$DN,2,' +
            'FALSE),"") and it appears 54 times. DN is column 118 — the paste is that wide.' }),

        drill({ id: 'F-VL-2', fn: 'VLOOKUP', scenario: 'T4', sheet: 'Risk Assessment',
          title: 'Replace a ladder with a table',
          task: 'N4 should say how many weeks apart this point gets swabbed. The ' +
            'band-to-frequency table is on Home, in A10:B13. Fill down to N18.',
          setup: RA_BAND,
          targets: ['N4:N18'],
          solution: { 'N4:N18': '=VLOOKUP($M4,Home!$A$10:$B$13,2,FALSE)' },
          shapeRules: { mustUse: ['VLOOKUP'], mustReference: ['Home'],
            mustNotUse: ['IF'],
            mustNotUseWhy: 'A four-rung IF ladder would work. The point of this ' +
              'drill is that a table does the same job and can be edited by someone ' +
              'who does not write formulas.' },
          hints: ['The thing you are looking up is the band in M4.',
            'Lock the table with $ so filling down does not slide it off the bottom: ' +
            'Home!$A$10:$B$13.',
            '=VLOOKUP($M4,Home!$A$10:$B$13,2,FALSE)'],
          realWorldNote: 'When the frequency policy changes, this version needs one ' +
            'number changed on Home. The IF-ladder version needs 15 formulas edited.' })
      ] },

    /* ------------------------------------------------------------- IFERROR -- */
    { fn: 'IFERROR', title: 'IFERROR — and what it hides',
      cells: 33455, scenario: 'T1',
      what: 'IFERROR(try this, use this instead if it errors). Any error at all: ' +
        '#N/A, #VALUE!, #REF!, #DIV/0!.',
      estate: '33,455 cells. Almost always IFERROR(VLOOKUP(…),"") — a missing ' +
        'match should leave the cell looking empty rather than printing #N/A on a ' +
        'label. The second argument does not have to be "": it can be a whole ' +
        'second attempt.',
      trap: 'It hides #REF! and #VALUE! just as happily as #N/A. A column of ' +
        'blanks can mean "nothing matched" or "this formula has been broken since ' +
        'March". Use the Inspector panel: it shows you the error an IFERROR ate.',
      drills: [
        drill({ id: 'F-IER-1', fn: 'IFERROR', scenario: 'T1', sheet: 'MICRO MOD',
          title: 'A missing sample should look empty',
          task: 'In F2, fetch the YEAST & MOULD result (column 6 of MICRO PASTE) ' +
            'for this row’s sample id. If there is no such sample, show nothing. ' +
            'Fill down to F9.',
          setup: T1_IDS,
          targets: ['F2:F9'],
          solution: { 'F2:F9': '=IFERROR(VLOOKUP($C2,\'MICRO PASTE\'!$A$2:$I$10,6,FALSE),"")' },
          shapeRules: { mustUse: ['IFERROR'], mustUse2: null,
            mustReference: ['MICRO PASTE'] },
          hints: ['Write the VLOOKUP first and get it working.',
            'Then wrap it: =IFERROR( your VLOOKUP , "" )',
            '=IFERROR(VLOOKUP($C2,\'MICRO PASTE\'!$A$2:$I$10,6,FALSE),"")'],
          realWorldNote: 'The sheet name has a space in it, so Excel quotes it: ' +
            '\'MICRO PASTE\'!A2. That is why so many real formulas are full of ' +
            'apostrophes.' }),

        drill({ id: 'F-IER-2', fn: 'IFERROR', scenario: 'T1', sheet: 'Test Profile 1',
          title: 'A second attempt, not a blank',
          task: 'D12 should show the TVC figure: the clean number from MICRO MOD ' +
            'column E if there is one, and failing that the raw text from column D. ' +
            'Fill down to D19.',
          setup: T1_NUM,
          targets: ['D12:D19'],
          solution: { 'D12:D19': '=IFERROR(\'MICRO MOD\'!$E2,\'MICRO MOD\'!$D2)' },
          shapeRules: { mustUse: ['IFERROR'], mustReference: ['MICRO MOD'] },
          hints: ['Both arguments point at MICRO MOD — different columns.',
            'Row 12 here lines up with row 2 there, so the offset holds as you fill.',
            '=IFERROR(\'MICRO MOD\'!$E2,\'MICRO MOD\'!$D2)'],
          realWorldNote: 'The real one is =IFERROR(\'MICRO MOD\'!P10,\'MICRO MOD\'!G10), ' +
            'used 114 times. It means "give me the number, or if we could not make a ' +
            'number, give me whatever the lab actually wrote".' })
      ] },

    /* --------------------------------------------------------------- SUMIF -- */
    { fn: 'SUMIF', title: 'SUMIF — add up the ones that match',
      cells: 13104, scenario: 'T4',
      what: 'SUMIF(range to test, the condition, [range to add up]). Leave the ' +
        'third one out and it adds up the range it tested.',
      estate: '13,104 cells, nearly all of them in the schedule workbook, and ' +
        'nearly all of them with a comparison in quotes: SUMIF(range,"<3").',
      trap: 'The criteria go in quotes as a little piece of text: "<3", ">=100", ' +
        '"Cutting". To compare against a cell you have to glue it: ">="&$C$7. ' +
        'Forget the & and you are searching for the literal text ">=C7".',
      drills: [
        drill({ id: 'F-SIF-1', fn: 'SUMIF', scenario: 'T4', sheet: 'Point List',
          title: 'A weighted total across the quarter',
          task: 'In D4, add up this point’s thirteen weeks on QTR 1 (K to W), ' +
            'counting only the entries below 3. Fill down to D18.',
          targets: ['D4:D18'],
          solution: { 'D4:D18': '=SUMIF(\'QTR 1\'!$K6:$W6,"<3")' },
          shapeRules: { mustUse: ['SUMIF'], mustReference: ['QTR 1'] },
          hints: ['Row 4 here is row 6 over there. Get the first one right and the ' +
            'fill takes care of the rest.',
            'Lock the columns with $K and $W, but leave the row numbers free.',
            '=SUMIF(\'QTR 1\'!$K6:$W6,"<3")'],
          realWorldNote: 'A clear swab is 1 and an over-limit swab is 2, so this ' +
            'total is swabs taken plus failures. The real Point List sheet does ' +
            'exactly this, 7,761 times.' }),

        drill({ id: 'F-SIF-2', fn: 'SUMIF', scenario: 'T4', sheet: 'Home',
          title: 'Test one column, add up another',
          task: 'C19 should total the LIMIT column of the Swab Results Database ' +
            'for the Cutting area only.',
          targets: ['C19'],
          solution: { 'C19': '=SUMIF(\'Swab Results Database\'!$C$2:$C$11,"Cutting",' +
            '\'Swab Results Database\'!$F$2:$F$11)' },
          shapeRules: { mustUse: ['SUMIF'], mustFill: false,
            mustReference: ['Swab Results Database'] },
          hints: ['Three arguments this time: test column C, condition "Cutting", ' +
            'add up column F.',
            'The two ranges must be the same height.',
            '=SUMIF(\'Swab Results Database\'!$C$2:$C$11,"Cutting",' +
            '\'Swab Results Database\'!$F$2:$F$11)'] })
      ] },

    /* ------------------------------------------------------------- COUNTIF -- */
    { fn: 'COUNTIF', title: 'COUNTIF — how many match',
      cells: 8421, scenario: 'T4',
      what: 'COUNTIF(range, condition). Same criteria rules as SUMIF, and it ' +
        'counts instead of adding.',
      estate: '8,421 cells. Two jobs: counting failures in a grid, and the ' +
        'existence check — COUNTIF(key column, my key) = 0 means "do not bother ' +
        'looking this one up".',
      trap: 'COUNTIF ignores a cell holding "" but COUNTA counts it, so a column ' +
        'of guarded formulas gives two different answers depending on which you ask.',
      drills: [
        drill({ id: 'F-CIF-1', fn: 'COUNTIF', scenario: 'T4', sheet: 'QTR 1',
          title: 'Failures per week',
          task: 'Row 23 is "Swabs over limit". In K23, count the 2s in that week’s ' +
            'column (K6:K20). Fill right to W23.',
          targets: ['K23:W23'],
          solution: { 'K23:W23': '=COUNTIF(K$6:K$20,2)' },
          shapeRules: { mustUse: ['COUNTIF'] },
          hints: ['A 2 means swabbed and over limit.',
            'You are filling RIGHT, so lock the ROW numbers and let the column ' +
            'letter move: K$6:K$20.',
            '=COUNTIF(K$6:K$20,2)'],
          realWorldNote: 'Lock the wrong half and the formula still works in K and ' +
            'is wrong in every column after it. Ctrl+` shows all thirteen at once.' }),

        drill({ id: 'F-CIF-2', fn: 'COUNTIF', scenario: 'T4', sheet: 'Home',
          title: 'How many high risk points',
          task: 'C6 should count the points whose band is "High".',
          setup: RA_BAND,
          targets: ['C6'],
          solution: { 'C6': '=COUNTIF(\'Risk Assessment\'!$M$4:$M$18,"High")' },
          shapeRules: { mustUse: ['COUNTIF'], mustFill: false,
            mustReference: ['Risk Assessment'] },
          hints: ['The bands are in column M of the Risk Assessment sheet.',
            'The condition is just the word, in quotes.',
            '=COUNTIF(\'Risk Assessment\'!$M$4:$M$18,"High")'] })
      ] },

    /* ----------------------------------------------------------------- SUM -- */
    { fn: 'SUM', title: 'SUM — and the blank-instead-of-zero wrapper',
      cells: 3886, scenario: 'T4',
      what: 'SUM(range) adds the numbers and ignores text and blanks instead of ' +
        'erroring on them.',
      estate: '3,886 cells. Interesting mostly for what is wrapped round it: ' +
        'IF(SUM(…)=0,"",SUM(…)), so an untouched row prints as blank, not 0.',
      trap: 'SUM ignores numbers stored as text. A column that adds to 34,000 ' +
        'when the real total is 58,000 is usually this, not a broken formula.',
      drills: [
        drill({ id: 'F-SUM-1', fn: 'SUM', scenario: 'T4', sheet: 'Risk Assessment',
          title: 'Total the four scores',
          task: 'L4 is the risk total: add up the four scores in H4:K4. ' +
            'Fill down to L18.',
          setup: RA_SCORES,
          targets: ['L4:L18'],
          solution: { 'L4:L18': '=SUM(H4:K4)' },
          shapeRules: { mustUse: ['SUM'] },
          hints: ['=SUM( then drag across H4 to K4.',
            'No dollars needed: every reference should move down with the fill.',
            '=SUM(H4:K4)'] }),

        drill({ id: 'F-SUM-2', fn: 'SUM', scenario: 'T1', sheet: 'Test Profile 1',
          title: 'Blank, not zero',
          task: 'K12 is the total number of plates for that day: add G12:I12. ' +
            'But if that comes to 0, show nothing instead. Fill down to K19.',
          targets: ['K12:K19'],
          solution: { 'K12:K19': '=IF(SUM($G12:$I12)=0,"",SUM($G12:$I12))' },
          shapeRules: { mustUse: ['SUM', 'IF'] },
          hints: ['You will write SUM twice: once in the test, once as the answer.',
            'Test for =0, and return "" when it is.',
            '=IF(SUM($G12:$I12)=0,"",SUM($G12:$I12))'],
          realWorldNote: 'The real formula is =IF(SUM(G118:J118)=0,"",SUM(G118:J118)). ' +
            'Writing SUM twice is the cost of not printing a zero. Nobody has found ' +
            'a shorter way that keeps a truly blank look.' })
      ] },

    /* ----------------------------------------------------------- MIN / MAX -- */
    { fn: 'MIN', title: 'MIN and MAX — first and last, because dates are numbers',
      cells: 1339, scenario: 'T4',
      what: 'MIN and MAX over a range. On a column of dates they give you the ' +
        'earliest and the latest, because a date IS a number.',
      estate: '1,339 cells, and the usual job is exactly that: MIN of a pasted ' +
        'date column is "when does this batch start", MAX is "when does it end".',
      trap: 'MIN of an empty range is 0, and 0 formats as 00/01/1900. If you see ' +
        'that date on a report, nothing matched.',
      drills: [
        drill({ id: 'F-MM-1', fn: 'MIN', scenario: 'T4', sheet: 'Home',
          title: 'First and last swab date',
          task: 'C21 is the earliest date in the Swab Results Database and C22 is ' +
            'the latest. The dates are in column A, rows 2 to 11.',
          targets: ['C21', 'C22'],
          solution: { 'C21': '=MIN(\'Swab Results Database\'!$A$2:$A$11)',
            'C22': '=MAX(\'Swab Results Database\'!$A$2:$A$11)' },
          shapeRules: { mustUse: ['MIN'], mustFill: false,
            mustReference: ['Swab Results Database'] },
          hints: ['Two different cells, two different formulas.',
            'MIN for the earliest, MAX for the latest.',
            '=MIN(\'Swab Results Database\'!$A$2:$A$11) and the same with MAX.'],
          realWorldNote: 'In the neck and caeca toolkit this is how the batch ' +
            'window is found: MIN(\'Transaction Listings Paste\'!I:I) for the start ' +
            'and MAX of the same column for the end.' })
      ] },

    /* --------------------------------------------------------------- MATCH -- */
    { fn: 'MATCH', title: 'MATCH — not the value, the position',
      cells: 822, scenario: 'T4',
      what: 'MATCH(what to find, where to look, 0). It returns WHICH ONE it is — ' +
        '1 for the first, 2 for the second — not the value.',
      estate: '822 cells. On its own it answers "which week did this first go ' +
        'wrong"; paired with INDEX it does what VLOOKUP cannot.',
      trap: 'The third argument must be 0 for an exact match. Leave it out and ' +
        'MATCH assumes your list is sorted ascending and returns nonsense if it ' +
        'is not.',
      drills: [
        drill({ id: 'F-MAT-1', fn: 'MATCH', scenario: 'T4', sheet: 'Point List',
          title: 'Which week did it first fail',
          task: 'G4 should give the week number of this point’s first ' +
            'over-limit result across QTR 1 columns K to W, or nothing at all if ' +
            'it never failed. Fill down to G18.',
          targets: ['G4:G18'],
          solution: { 'G4:G18': '=IFERROR(MATCH(2,\'QTR 1\'!$K6:$W6,0),"")' },
          shapeRules: { mustUse: ['MATCH'], mustReference: ['QTR 1'] },
          hints: ['You are looking for the number 2 — an over-limit swab.',
            'MATCH gives #N/A when it finds nothing, so wrap it in IFERROR.',
            '=IFERROR(MATCH(2,\'QTR 1\'!$K6:$W6,0),"")'],
          realWorldNote: 'The neck and caeca toolkit uses MATCH the same way to ' +
            'find the row where a flag is set, then reads that row with VLOOKUP.' })
      ] },

    /* --------------------------------------------------------------- INDEX -- */
    { fn: 'INDEX', title: 'INDEX + MATCH — looking left',
      cells: 652, scenario: 'T3',
      what: 'INDEX(a column, which row). Put MATCH in the second slot and you ' +
        'have a lookup that does not care which side of the key the answer sits on.',
      estate: '652 cells. It is what you reach for when the key is not the first ' +
        'column of the block, which VLOOKUP simply cannot handle.',
      trap: 'The INDEX range and the MATCH range must be the same height and ' +
        'start on the same row, or you get the right position in the wrong list.',
      drills: [
        drill({ id: 'F-IDX-1', fn: 'INDEX', scenario: 'T3',
          sheet: '2sfg UKP Nutritional Toolkit',
          title: 'Pull the lab figure out of the hidden sheet',
          task: 'G10 should bring back the number the background sheet worked out ' +
            'for this nutrient — matched by name, not by row. Fill down to G19.',
          setup: T3_NUM,
          targets: ['G10:G19'],
          solution: { 'G10:G19': '=IFERROR(INDEX(background!$C$2:$C$11,' +
            'MATCH($A10,background!$A$2:$A$11,0)),"")' },
          shapeRules: { mustUse: ['INDEX', 'MATCH'], mustReference: ['background'] },
          hints: ['INDEX over the NUMBER column, MATCH over the ANALYTE column.',
            'The names on the background sheet are in capitals. Excel does not ' +
            'care — matching is case-blind.',
            '=IFERROR(INDEX(background!$C$2:$C$11,MATCH($A10,background!$A$2:$A$11,0)),"")'],
          realWorldNote: 'Matching by name rather than by row is what lets someone ' +
            'insert a nutrient into the middle of the list without silently shifting ' +
            'every figure below it.' })
      ] },

    /* ------------------------------------------------------------- COUNTIFS -- */
    { fn: 'COUNTIFS', title: 'COUNTIFS — two conditions at once',
      cells: 681, scenario: 'T4',
      what: 'COUNTIFS(range1, condition1, range2, condition2, …). Every condition ' +
        'has to hold on the same row. SUMIFS is the same idea, with the range to ' +
        'add up first instead of last.',
      estate: '681 cells. Note the argument order flips between SUMIF and SUMIFS: ' +
        'SUMIF ends with the range to add, SUMIFS starts with it.',
      trap: 'All the ranges must be the same height. Mix A2:A11 with C2:C20 and ' +
        'you get #VALUE!, not a warning.',
      drills: [
        drill({ id: 'F-CIFS-1', fn: 'COUNTIFS', scenario: 'T4', sheet: 'Home',
          title: 'Fails in one area only',
          task: 'C18 should count the rows of the Swab Results Database that are ' +
            'both in the Cutting area and a Fail.',
          targets: ['C18'],
          solution: { 'C18': '=COUNTIFS(\'Swab Results Database\'!$C$2:$C$11,"Cutting",' +
            '\'Swab Results Database\'!$G$2:$G$11,"Fail")' },
          shapeRules: { mustUse: ['COUNTIFS'], mustFill: false,
            mustReference: ['Swab Results Database'] },
          hints: ['Area is column C, verdict is column G.',
            'Range, condition, range, condition.',
            '=COUNTIFS(\'Swab Results Database\'!$C$2:$C$11,"Cutting",' +
            '\'Swab Results Database\'!$G$2:$G$11,"Fail")'] })
      ] },

    /* ------------------------------------------------------------- AND / OR -- */
    { fn: 'OR', title: 'AND and OR — more than one test',
      cells: 129, scenario: 'T1',
      what: 'AND(test, test, …) is true only if all of them are. OR(…) is true if ' +
        'any of them is. They go inside the first slot of an IF.',
      estate: 'Only 129 cells, which is the surprise: the estate prefers nesting ' +
        'IFs. AND and OR are usually the clearer choice when the tests are ' +
        'genuinely independent.',
      trap: 'AND and OR return one TRUE or FALSE for the whole thing. They cannot ' +
        'give you a different answer per row — that is what IF is for.',
      drills: [
        drill({ id: 'F-AO-1', fn: 'OR', scenario: 'T1', sheet: 'Test Profile 1',
          title: 'Over on either count',
          task: 'F12 should say "OVER" if the TVC in D12 is above the limit in C7 ' +
            'OR the Y&M in E12 is above the limit in C8, and "OK" otherwise. ' +
            'Fill down to F19.',
          setup: merge(T1_NUM, T1_LIMITS, {
            'Test Profile 1!D12:D19': '=\'MICRO MOD\'!$E2',
            'Test Profile 1!E12:E19': '=\'MICRO MOD\'!$G2'
          }),
          targets: ['F12:F19'],
          solution: { 'F12:F19': '=IF(OR($D12>$C$7,$E12>$C$8),"OVER","OK")' },
          shapeRules: { mustUse: ['OR', 'IF'] },
          hints: ['The two limits are single cells, so lock them: $C$7 and $C$8.',
            'The two results move down with the fill, so $D12 and $E12.',
            '=IF(OR($D12>$C$7,$E12>$C$8),"OVER","OK")'],
          realWorldNote: 'Lock the limits with F4 as you type. Forget, and row 19 ' +
            'will be comparing against C14 and C15, which are empty — so everything ' +
            'reads OVER.' })
      ] },

    /* ------------------------------------------------------------- ISERROR -- */
    { fn: 'ISERROR', title: 'ISERROR — the long way round',
      cells: 121, scenario: 'T3',
      what: 'ISERROR(x) is TRUE when x is an error. The old pattern is ' +
        'IF(ISERROR(x),"-",x) — which computes x twice.',
      estate: '121 cells in the nutritional toolkit alone. It predates IFERROR, ' +
        'and it is worth being able to read because half the estate is written ' +
        'this way.',
      trap: 'It evaluates the same expression twice, so it is slower and twice as ' +
        'easy to get wrong when you edit it — change one copy and not the other ' +
        'and the cell lies to you.',
      drills: [
        drill({ id: 'F-ISE-1', fn: 'ISERROR', scenario: 'T3',
          sheet: '2sfg UKP Nutritional Toolkit',
          title: 'The old guard, as written',
          task: 'H10 is the difference between the average in F10 and the declared ' +
            'figure in B10 — but using the IF(ISERROR(…)) form, with "-" when it ' +
            'cannot be worked out. Fill down to H19.',
          setup: T3_AVG,
          targets: ['H10:H19'],
          solution: { 'H10:H19': '=IF(ISERROR($F10-$B10),"-",$F10-$B10)' },
          shapeRules: { mustUse: ['ISERROR', 'IF'] },
          hints: ['The subtraction appears twice: once inside ISERROR, once as the ' +
            'answer.',
            'The fallback is "-", not "".',
            '=IF(ISERROR($F10-$B10),"-",$F10-$B10)'],
          realWorldNote: 'The real file writes =IF(ISERROR(Q21-L21),"-",(Q21-L21)). ' +
            'IFERROR($F10-$B10,"-") is the same thing in half the characters — but ' +
            'you need to recognise the long form to maintain what is already there.' })
      ] },

    /* ----------------------------------------------------- TEXTAFTER / VALUE -- */
    { fn: 'TEXTAFTER', title: 'TEXTAFTER and VALUE — making "<10" into a number',
      cells: 125, scenario: 'T3',
      what: 'TEXTAFTER(text, delimiter) gives you everything after the delimiter. ' +
        'VALUE(text) turns a digit string into a real number.',
      estate: 'This pair is the whole answer to micro results. The lab writes ' +
        '"<10" and ">1000000" because that is honest, and both are text, so ' +
        'nothing can be averaged or plotted until they are stripped.',
      trap: 'TEXTAFTER errors when the delimiter is not there, which is why the ' +
        'real formula is VALUE(IFERROR(TEXTAFTER(x,"<"),x)) — "strip the < if ' +
        'there is one, otherwise use it as it stands".',
      drills: [
        drill({ id: 'F-TXT-1', fn: 'TEXTAFTER', scenario: 'T3', sheet: 'background',
          title: 'Strip the less-than',
          task: 'C2 should be the number behind the raw result in B2. Results like ' +
            '"<0.5" must come through as 0.5. Fill down to C11.',
          setup: T3_BG,
          targets: ['C2:C11'],
          solution: { 'C2:C11': '=VALUE(IFERROR(TEXTAFTER($B2,"<"),$B2))' },
          shapeRules: { mustUse: ['TEXTAFTER', 'VALUE'] },
          hints: ['Innermost first: TEXTAFTER($B2,"<") gives "0.5" for "<0.5" and ' +
            'an error for a plain 2.4.',
            'IFERROR turns that error back into the original value.',
            '=VALUE(IFERROR(TEXTAFTER($B2,"<"),$B2))'],
          realWorldNote: 'Row 11 is ASH, which the lab sent as the TEXT "1.1". ' +
            'VALUE fixes that too, which is the other half of why it is there.' }),

        drill({ id: 'F-TXT-2', fn: 'TEXTAFTER', scenario: 'T1', sheet: 'MICRO MOD',
          title: 'Both directions',
          task: 'E2 should be the number behind the raw TVC in D2 — which may be ' +
            '"<10", or ">1000000", or already a number. Fill down to E9.',
          setup: T1_RAW,
          targets: ['E2:E9'],
          solution: { 'E2:E9': '=IFERROR(VALUE(IFERROR(TEXTAFTER($D2,"<"),$D2)),' +
            'VALUE(IFERROR(TEXTAFTER($D2,">"),$D2)))' },
          shapeRules: { mustUse: ['TEXTAFTER', 'VALUE', 'IFERROR'] },
          hints: ['Build the "<" version first and watch what ">1000000" does to it.',
            'Then wrap the whole thing in another IFERROR whose fallback is the ' +
            'same formula with ">".',
            '=IFERROR(VALUE(IFERROR(TEXTAFTER($D2,"<"),$D2)),' +
            'VALUE(IFERROR(TEXTAFTER($D2,">"),$D2)))'],
          realWorldNote: 'This is formula P10 in the real MICRO MOD, 168 copies of ' +
            'it. Open the Evaluate panel on your answer and step through it: four ' +
            'functions deep is exactly where that panel starts to earn its keep.' })
      ] },

    /* ------------------------------------------------------------- XLOOKUP -- */
    { fn: 'XLOOKUP', title: 'XLOOKUP — when the key is on the wrong side',
      cells: 42, scenario: 'T1',
      what: 'XLOOKUP(what to find, the column to search, the column to return, ' +
        '[what to say if not found]). Two separate ranges, so they can be in any ' +
        'order.',
      estate: '42 cells in the whole shelf life toolkit, all on one sheet, all for ' +
        'one reason: on SET MOD the test point is in column H and the sample id is ' +
        'in column A. VLOOKUP cannot look left. That is the entire justification.',
      trap: 'The search range and the return range must be the same height, and ' +
        'older Excel will not open it at all. Elsewhere in the estate, VLOOKUP is ' +
        'the safer choice.',
      drills: [
        drill({ id: 'F-XL-1', fn: 'XLOOKUP', scenario: 'T1', sheet: 'SAMPLE ID NO',
          title: 'Look left',
          task: 'D2 should give the sample id for this test point. On SET MOD the ' +
            'test points are in H2:H9 and the sample ids are in A2:A9. ' +
            'Fill down to D9.',
          targets: ['D2:D9'],
          solution: { 'D2:D9':
            '=XLOOKUP($B2,\'SET MOD\'!$H$2:$H$9,\'SET MOD\'!$A$2:$A$9)' },
          shapeRules: { mustUse: ['XLOOKUP'], mustReference: ['SET MOD'],
            mustNotUse: ['VLOOKUP'],
            mustNotUseWhy: 'VLOOKUP searches the first column of its block and ' +
              'counts right from there. The answer here is to the LEFT of the key, ' +
              'so there is no column number that reaches it.' },
          hints: ['Three ranges in all: the value, the column to search, the ' +
            'column to return.',
            'Lock all of them with $ so the fill does not drag them down.',
            '=XLOOKUP($B2,\'SET MOD\'!$H$2:$H$9,\'SET MOD\'!$A$2:$A$9)'],
          realWorldNote: 'The real formula is XLOOKUP(B2,\'SET MOD\'!$H$134:$H$147,' +
            '\'SET MOD\'!$A$134:$A$147). Row 134 — the table lives a long way down a ' +
            'hidden sheet, which is normal and is why the Trace panel exists.' }),

        drill({ id: 'F-XL-2', fn: 'XLOOKUP', scenario: 'T1', sheet: 'SAMPLE ID NO',
          title: 'Say what happened instead of #N/A',
          task: 'C2 should do the same lookup, but say "not scheduled" when the ' +
            'test point is not in the table — using XLOOKUP’s own fourth ' +
            'argument, not IFERROR. Fill down to C9.',
          targets: ['C2:C9'],
          solution: { 'C2:C9': '=XLOOKUP($B2,\'SET MOD\'!$H$2:$H$9,' +
            '\'SET MOD\'!$A$2:$A$9,"not scheduled")' },
          shapeRules: { mustUse: ['XLOOKUP'], mustNotUse: ['IFERROR'],
            mustNotUseWhy: 'IFERROR would catch a #REF! or a #VALUE! as well, and ' +
              'report it as "not scheduled". XLOOKUP’s fourth argument fires ' +
              'only on a genuine no-match, so a real fault still shows itself.' },
          hints: ['Same formula, one more argument on the end.',
            'The text goes in quotes.',
            '=XLOOKUP($B2,\'SET MOD\'!$H$2:$H$9,\'SET MOD\'!$A$2:$A$9,"not scheduled")'],
          realWorldNote: 'This is the one thing XLOOKUP does that no amount of ' +
            'IFERROR can: distinguish "not in the list" from "something is broken".' })
      ] },

    /* --------------------------------------------------------------- & glue -- */
    { fn: '&', title: '& — glue, and the composite key',
      cells: 2000, scenario: 'T1',
      what: 'a & b joins two values into one piece of text. CONCAT does the same ' +
        'thing with brackets.',
      estate: 'Two jobs. Building label text — "2037, Scunthorpe, [KILL LINE " & ' +
        'line & "]" — and building a key out of two columns when neither one is ' +
        'unique on its own.',
      trap: 'Anything you glue becomes TEXT. Glue a number and then compare it ' +
        'with a number and you get FALSE. In your files this is the single ' +
        'commonest cause of a lookup that cannot find a row that is plainly there.',
      drills: [
        drill({ id: 'F-AMP-1', fn: '&', scenario: 'T1', sheet: 'Test Profile 1',
          title: 'A key made of two things',
          task: 'The COP sheet keys on product AND organism, glued with "|". ' +
            'C7 should fetch the TVC action limit for the product in C3, and C8 ' +
            'the Yeast & Mould one.',
          targets: ['C7', 'C8'],
          solution: { 'C7': '=VLOOKUP($C$3&"|TVC",COP!$A$2:$F$7,4,FALSE)',
            'C8': '=VLOOKUP($C$3&"|Yeast & Mould",COP!$A$2:$F$7,4,FALSE)' },
          shapeRules: { mustUse: ['VLOOKUP'], mustFill: false,
            mustReference: ['COP'] },
          hints: ['Look at COP column A to see exactly how the key is spelled.',
            'Build the key inside the VLOOKUP: $C$3&"|TVC".',
            '=VLOOKUP($C$3&"|TVC",COP!$A$2:$F$7,4,FALSE) — and column 4 is the ' +
            'action limit.'],
          realWorldNote: 'One product, two limits, so the product alone cannot be a ' +
            'key. Gluing is how the estate solves this everywhere, and it is why so ' +
            'many reference sheets have a first column nobody reads.' }),

        drill({ id: 'F-AMP-2', fn: '&', scenario: 'T1', sheet: 'MICRO MOD',
          title: 'Build the label line',
          task: 'J2 should read like "Aromatic Butter Chicken BDF Day 0 (001)" — ' +
            'the product from H2, the test point from B2, and the sample number ' +
            'from I2 in brackets. Fill down to J9.',
          setup: merge(T1_NUM, {
            'MICRO MOD!H2:H9': '=IFERROR(VLOOKUP($C2,\'MICRO PASTE\'!$A$2:$I$10,2,FALSE),"")',
            'MICRO MOD!I2:I9': '=RIGHT($C2,3)'
          }),
          targets: ['J2:J9'],
          solution: { 'J2:J9': '=$H2&" "&$B2&" ("&$I2&")"' },
          shapeRules: { mustFill: true, mustNotHardcode: true },
          hints: ['Spaces and brackets are text, so they go in quotes.',
            'Four & signs in all.',
            '=$H2&" "&$B2&" ("&$I2&")"'],
          realWorldNote: 'This is how every label in the neck and caeca toolkit is ' +
            'assembled, and why a trailing space in a source column shows up as a ' +
            'double space on the printed label.' })
      ] },

    /* ------------------------------------------------- RIGHT / LEFT / MID --- */
    { fn: 'RIGHT', title: 'RIGHT, LEFT and MID — counting characters',
      cells: 34, scenario: 'T1',
      what: 'RIGHT(text, n) takes the last n characters, LEFT(text, n) the first ' +
        'n, MID(text, start, n) n characters from a position.',
      estate: '34 cells, mostly RIGHT(code, 8) to drop a prefix off a batch code. ' +
        'The newer TEXTAFTER counts FIELDS instead of characters, which does not ' +
        'break when the prefix changes length.',
      trap: 'Counting characters assumes every code is the same shape forever. ' +
        'One six-character site code among the five-character ones and every row ' +
        'is quietly wrong.',
      drills: [
        drill({ id: 'F-RLM-1', fn: 'RIGHT', scenario: 'T1', sheet: 'MICRO MOD',
          title: 'The last three characters',
          task: 'I2 should be just the sample number — the last three characters ' +
            'of the sample id in C2. Fill down to I9.',
          setup: T1_IDS,
          targets: ['I2:I9'],
          solution: { 'I2:I9': '=RIGHT($C2,3)' },
          shapeRules: { mustUse: ['RIGHT'], mustNotHardcode: true },
          hints: ['=RIGHT( the cell , how many characters )',
            'Three characters.',
            '=RIGHT($C2,3)'],
          realWorldNote: 'The result is TEXT, even though it looks like 001. ' +
            'Compare it with the number 1 and you get FALSE. Check the alignment: ' +
            'it sits left.' })
      ] },

    /* ------------------------------------------------------------- AVERAGE -- */
    { fn: 'AVERAGE', title: 'AVERAGE — and what it quietly skips',
      cells: 26, scenario: 'T3',
      what: 'AVERAGE(range). The mean of the numbers in there.',
      estate: '26 cells, almost all of them averaging replicate results — three ' +
        'plates, or three sub-samples of the same batch.',
      trap: 'AVERAGE skips blanks and text but NOT zeros. Three results of 10, 20 ' +
        'and blank average 15; three results of 10, 20 and 0 average 10. Which one ' +
        'you want depends on whether the blank means "not done" or "none found".',
      drills: [
        drill({ id: 'F-AVG-1', fn: 'AVERAGE', scenario: 'T3',
          sheet: '2sfg UKP Nutritional Toolkit',
          title: 'Average the three reps',
          task: 'F10 is the average of the three replicate results in C10:E10. ' +
            'Fill down to F19.',
          targets: ['F10:F19'],
          solution: { 'F10:F19': '=AVERAGE($C10:$E10)' },
          shapeRules: { mustUse: ['AVERAGE'], mustNotHardcode: true },
          hints: ['=AVERAGE( then drag across the three.',
            'Lock the columns, leave the row free: $C10:$E10.',
            '=AVERAGE($C10:$E10)'] }),

        drill({ id: 'F-AVG-2', fn: 'AVERAGE', scenario: 'T1', sheet: 'Test Profile 1',
          title: 'Average the plate counts',
          task: 'J12 is the average of the three plate counts in G12:I12. ' +
            'Fill down to J19.',
          targets: ['J12:J19'],
          solution: { 'J12:J19': '=AVERAGE($G12:$I12)' },
          shapeRules: { mustUse: ['AVERAGE'] },
          hints: ['Same shape as before.',
            'Row 18 has a 2 among the 3s — watch the average move.',
            '=AVERAGE($G12:$I12)'] })
      ] },

    /* ----------------------------------------------------- COUNT / COUNTA --- */
    { fn: 'COUNT', title: 'COUNT and COUNTA — two different questions',
      cells: 753, scenario: 'T4',
      what: 'COUNT counts NUMBERS. COUNTA counts anything that is not truly empty ' +
        '— including a cell holding "".',
      estate: '753 cells. COUNT is "how many swabs were actually taken"; COUNTA is ' +
        '"how far down does this list go".',
      trap: 'A cell holding "" from an IF guard is empty to COUNT and to the eye, ' +
        'but COUNTA counts it. That is the difference between "15 points on the ' +
        'schedule" and "600 points on the schedule".',
      drills: [
        drill({ id: 'F-CNT-1', fn: 'COUNT', scenario: 'T4', sheet: 'QTR 1',
          title: 'Swabs taken per week',
          task: 'Row 22 is "Swabs taken". In K22, count the entries in that ' +
            'week’s column, K6:K20. Fill right to W22.',
          targets: ['K22:W22'],
          solution: { 'K22:W22': '=COUNT(K$6:K$20)' },
          shapeRules: { mustUse: ['COUNT'] },
          hints: ['Both 1s and 2s are swabs that were taken.',
            'Filling RIGHT, so lock the rows: K$6:K$20.',
            '=COUNT(K$6:K$20)'] }),

        drill({ id: 'F-CNT-2', fn: 'COUNT', scenario: 'T4', sheet: 'Home',
          title: 'How long is the list',
          task: 'C5 should say how many swab points are on the Risk Assessment ' +
            'sheet. The point numbers are in A4:A18.',
          targets: ['C5'],
          solution: { 'C5': '=COUNTA(\'Risk Assessment\'!$A$4:$A$18)' },
          shapeRules: { mustUse: ['COUNTA|COUNT'], mustFill: false,
            mustReference: ['Risk Assessment'] },
          hints: ['Either would give 15 here, because every point number is a ' +
            'number.',
            'COUNTA is the honest choice: it still works if someone enters a point ' +
            'as "12a".',
            '=COUNTA(\'Risk Assessment\'!$A$4:$A$18)'] })
      ] },

    /* ----------------------------------------------------------- ROUND / % -- */
    { fn: 'ROUND', title: 'ROUND, and dividing without #DIV/0!',
      cells: 6, scenario: 'T3',
      what: 'ROUND(number, decimal places). And the division guard: wrap any ' +
        'division in IFERROR, because a zero denominator is a matter of time.',
      estate: 'Only 6 cells round anything — but every percentage in the estate is ' +
        'wrapped: IFERROR((a/b)*100,""). A blank is a readable "no data"; ' +
        '#DIV/0! on a printed report is not.',
      trap: 'Rounding for display and rounding the stored value are different ' +
        'things. A number format changes what you see; ROUND changes what the next ' +
        'formula gets.',
      drills: [
        drill({ id: 'F-RND-1', fn: 'ROUND', scenario: 'T3',
          sheet: '2sfg UKP Nutritional Toolkit',
          title: 'Percentage difference, to one decimal',
          task: 'I10 is how far the lab average in F10 sits from the declared ' +
            'figure in B10, as a percentage of the declaration, to one decimal ' +
            'place — and blank if it cannot be worked out. Fill down to I19.',
          setup: T3_AVG,
          targets: ['I10:I19'],
          solution: { 'I10:I19': '=IFERROR(ROUND(($F10-$B10)/$B10*100,1),"")' },
          shapeRules: { mustUse: ['ROUND', 'IFERROR'] },
          hints: ['(average - declared) / declared * 100.',
            'ROUND(…,1) for one decimal place.',
            '=IFERROR(ROUND(($F10-$B10)/$B10*100,1),"")'],
          realWorldNote: 'Fat comes out at 20.0% against a tolerance of 20 — right ' +
            'on the line. That is the number a declaration argument turns on, which ' +
            'is why the rounding is decided deliberately and not left to the ' +
            'number format.' }),

        drill({ id: 'F-RND-2', fn: 'ROUND', scenario: 'T4', sheet: 'QTR 1',
          title: 'Percent clear, guarded',
          task: 'Row 24 is "% clear". In K24, work out what share of that week’s ' +
            'swabs came back clear: taken in row 22, over limit in row 23. Weeks ' +
            'with no swabs should be blank, not an error. Fill right to W24.',
          setup: Q1_COUNTS,
          targets: ['K24:W24'],
          solution: { 'K24:W24': '=IFERROR((K$22-K$23)/K$22*100,"")' },
          shapeRules: { mustUse: ['IFERROR'] },
          hints: ['Clear swabs are taken minus over-limit.',
            'Week 11 has no swabs at all — that is the division by zero IFERROR is ' +
            'there for.',
            '=IFERROR((K$22-K$23)/K$22*100,"")'],
          realWorldNote: 'The real formula is =IFERROR((K58/K57)*100,""). Exactly ' +
            'this, 156 times.' })
      ] },

    /* ----------------------------------------------------------- the dates -- */
    { fn: 'TODAY', title: 'Dates — a date is a number, so you can add 7 to it',
      cells: 561, scenario: 'T4',
      what: 'A date is the count of days since 1900. 13/11/2026 is the number ' +
        '46339. So next week is just +7, and TODAY() is today’s number.',
      estate: '561 cells in the schedule workbook are literally =previous+7. One ' +
        'date on Home, and every week header in the year follows from it.',
      trap: 'TODAY() changes under you. Anything graded or signed should read a ' +
        'date from a cell, not from TODAY(), or last month’s report will not ' +
        'reprint the same.',
      drills: [
        drill({ id: 'F-DAT-1', fn: 'TODAY', scenario: 'T4', sheet: 'QTR 1',
          title: 'The first week comes from Home',
          task: 'K3 is the week 1 date. Take it from Home C3 rather than typing it.',
          targets: ['K3'],
          solution: { 'K3': '=Home!$C$3' },
          shapeRules: { mustFill: false, mustNotHardcode: true,
            mustReference: ['Home'] },
          hints: ['Type =, click the Home tab, click C3, press Enter.',
            'Lock it with F4 — this cell is the origin of the whole year.',
            '=Home!$C$3'] }),

        drill({ id: 'F-DAT-2', fn: 'TODAY', scenario: 'T4', sheet: 'QTR 1',
          title: 'Then just add seven',
          task: 'L3 should be a week after K3, and so on across to W3.',
          setup: { 'QTR 1!K3': '=Home!$C$3' },
          targets: ['L3:W3'],
          solution: { 'L3:W3': '=K3+7' },
          shapeRules: { mustNotHardcode: true },
          hints: ['Point at the cell to its left and add 7.',
            'No dollars at all: each one should look at its own neighbour.',
            '=K3+7'],
          realWorldNote: 'The real sheet writes it =+R20+7. That leading + is a ' +
            'Lotus habit from the 1980s that Excel still accepts; it does nothing, ' +
            'and it is all over your files.' })
      ] },

    /* ------------------------------------------------- putting it together -- */
    { fn: 'the lot', title: 'Putting it together — the output row',
      cells: 0, scenario: 'T1',
      what: 'Nothing new. These are the cells someone signs, and each one is two ' +
        'or three of the functions above stacked up.',
      estate: 'This is what every formula in the four toolkits is for: one line ' +
        'that a person reads and acts on.',
      trap: 'Build it in layers and check each one in the Evaluate panel. A single ' +
        'cell containing six functions typed in one go is where an afternoon goes.',
      drills: [
        drill({ id: 'F-END-1', fn: 'the lot', scenario: 'T1', sheet: 'Test Profile 1',
          title: 'Sign off the shelf life study',
          task: 'Three summary cells. C21: how many days were analysed. C22: how ' +
            'many of them came back over limit. C23: the verdict — if nothing was ' +
            'over limit say "Shelf life supported", otherwise say which day it ' +
            'first failed, as "Fails at day 15".',
          setup: merge(T1_PROFILE, {
            'Test Profile 1!F12:F19': '=IF(OR($D12>$C$7,$E12>$C$8),"OVER","OK")'
          }),
          targets: ['C21', 'C22', 'C23'],
          solution: {
            'C21': '=COUNT($D$12:$D$19)',
            'C22': '=COUNTIF($F$12:$F$19,"OVER")',
            'C23': '=IF($C$22=0,"Shelf life supported","Fails at day "&' +
              'INDEX($A$12:$A$19,MATCH("OVER",$F$12:$F$19,0)))'
          },
          shapeRules: { mustUse: ['COUNT', 'COUNTIF', 'MATCH'], mustFill: false },
          hints: ['C21 is a COUNT of the TVC column. C22 is a COUNTIF of the ' +
            'verdict column for "OVER".',
            'For C23: MATCH("OVER",…,0) gives the position of the first failure, ' +
            'and INDEX turns that position into the day number in column A.',
            'Then glue it: "Fails at day "&INDEX(…).'],
          realWorldNote: 'Every piece of this is something you built earlier in ' +
            'this track. That is what the real files are: five or six of these ' +
            'functions, stacked, repeated.' }),

        drill({ id: 'F-END-2', fn: 'the lot', scenario: 'T3',
          sheet: '2sfg UKP Nutritional Toolkit',
          title: 'Judge the declaration',
          task: 'K10 flags a result more than the tolerance BELOW the declaration ' +
            '(1 or 0), L10 flags one more than the tolerance ABOVE it, and M10 ' +
            'turns the pair into words: "Within tolerance" or "OUTSIDE TOLERANCE". ' +
            'Fill all three down to row 19.',
          setup: T3_DIFF,
          targets: ['K10:K19', 'L10:L19', 'M10:M19'],
          solution: {
            'K10:K19': '=IF($I10<-$J10,1,0)',
            'L10:L19': '=IF($I10>$J10,1,0)',
            'M10:M19': '=IF($K10+$L10=0,"Within tolerance","OUTSIDE TOLERANCE")'
          },
          shapeRules: { mustUse: ['IF'], mustFill: false },
          hints: ['The % difference is in I and the tolerance is in J.',
            'Below tolerance means I is less than MINUS J: $I10<-$J10.',
            'M is then just a test on the two flags adding to 0.'],
          realWorldNote: 'The real file writes =IF(O21<-(AE21),1,0) and ' +
            '=IF(O21>(AE21),1,0), 130 cells of them. Two 1/0 flags added together ' +
            'is the estate’s standard way of asking "did anything go wrong on ' +
            'this row".' })
      ] }
  ];

  /* ------------------------------------------------------------ registration */

  XT.data.drills = drills;

  /* Give every drill the fields the lesson panel and the grader expect. */
  XT.data.formulaTrack.forEach(function (L, i) {
    L.index = i;
    (L.drills || []).forEach(function (d) {
      d.levelTitle = 'Learn by formula';
      d.lessonTitle = L.title;
      d.lessonIndex = i;
      if (!d.patterns) d.patterns = [];
    });
  });

  /* XT.data.exercise(id) must find drills too, so the whole UI — load, grade,
     Show me, next/previous — works on them without a second code path. */
  var courseLookup = XT.data.exercise;
  XT.data.exercise = function (id) {
    var found = courseLookup(id);
    if (found) return found;
    for (var i = 0; i < drills.length; i++) if (drills[i].id === id) return drills[i];
    return null;
  };

  XT.data.drillsFor = function (fn) {
    return drills.filter(function (d) { return d.fn === fn; });
  };

})(window.XT = window.XT || {});
