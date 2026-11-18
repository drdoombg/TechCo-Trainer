/* =============================================================================
   Excel Trainer — data/patterns.js
   The twenty patterns from catalog §5, as data.

   `cells` is the measured count from the scan of 152 workbooks. The progress
   screen uses it to weight coverage, so "you have mastered 71% of this estate's
   formula mass" is a real number rather than a count of exercises done.

   `shapes` are exact normalised forms (see engine/shape.js). `test` entries are
   regexes against the normalised form, used when a formula is the same move
   with different parts. `example` is real formula text from the source files --
   formula text only; no real data values appear anywhere in this trainer.
   ============================================================================= */
(function (XT) {
  'use strict';
  XT.data = XT.data || {};

  XT.data.patterns = [
    {
      id: 'P1', title: 'Bare pass-through', level: 1, layer: 2,
      cells: 867462, files: 53, rank: 1,
      oneLine: 'Mirror a column from the landing sheet onto a working sheet.',
      example: '=SamplesinWB!A2',
      shapes: ['SH!R'],
      test: ['^SH!R$'],
      why: 'Keeps the pasted export pristine. Everything downstream points at the ' +
        'mirror, so re-pasting never breaks a formula. The estate’s version of ' +
        '"do not edit your source data".',
      teaches: ['cross-sheet references', 'quoted sheet names', 'fill down'],
      watch: 'It is the most common formula in the whole estate. Boring, and load-bearing.'
    },
    {
      id: 'P2', title: 'Zero-suppressing mirror', level: 2, layer: 2,
      cells: 472402, files: 2, rank: 2,
      oneLine: 'Mirror a cell, but show nothing instead of a zero.',
      example: '=IF(\'Despatch Paste\'!A6>0,\'Despatch Paste\'!A6,"")',
      shapes: ['IF(SH!R>N,SH!R,"S")'],
      test: ['^IF\\(SH!R>N,SH!R,"S"\\)$'],
      why: 'Exports pad empty cells with zeros, and zeros print on labels where they ' +
        'look like data.',
      teaches: ['IF', 'the "" result', 'guarding a mirror'],
      watch: 'The >0 test is numeric but the data is often text, and in Excel text is ' +
        'GREATER than any number. So this guard does not do what it looks like it does.',
      defect: 'D5'
    },
    {
      id: 'P3', title: 'Blank-guarded mirror', level: 2, layer: 2,
      cells: 428034, files: 6, rank: 3,
      oneLine: 'Only pull a value if the anchor cell has something in it.',
      example: '=IF(E1="","",LVALL!C2)',
      shapes: ['IF(R="S","S",SH!R)', 'IF(SH!R="S","S",SH!R)'],
      test: ['^IF\\((SH!)?R="S","S",SH!R\\)$'],
      why: 'Stops a column filled down a hundred thousand rows from showing junk below ' +
        'the data.',
      teaches: ['IF', 'EMPTY versus ""', 'anchor cells'],
      watch: 'The cell now holds TEXT of zero length, not nothing. ISBLANK says FALSE, ' +
        'COUNTA counts it, and SUM ignores it. This is the single most important ' +
        'subtlety in the whole dialect.'
    },
    {
      id: 'P4', title: 'Belt-and-braces mirror', level: 5, layer: 2,
      cells: 329778, files: 1, rank: 4,
      oneLine: 'A guarded mirror with an IFERROR wrapped round it as well.',
      example: '=IFERROR(IF(\'Complaints Paste\'!A1="","",\'Complaints Paste\'!A1),"")',
      shapes: ['IFERROR(IF(SH!R="S","S",SH!R),"S")'],
      test: ['^IFERROR\\(IF\\(SH!R="S","S",SH!R\\),"S"\\)$'],
      why: 'It does not have a reason. A mirror of a cell cannot produce an error.',
      teaches: ['reading a formula critically', 'when a wrapper earns its place'],
      watch: 'Copied 329,778 times without anyone asking what the IFERROR was for. ' +
        'Good practice at deciding what to delete.',
      defect: 'D12'
    },
    {
      id: 'P5', title: 'The guarded join', level: 4, layer: 4,
      cells: 237873, files: 9, rank: 5,
      oneLine: 'Enrich a row from a reference table; blank if the key is not found.',
      example: '=IFERROR(VLOOKUP(C18,PASTE!$A:$DN,2,FALSE),"")',
      shapes: ['IFERROR(VLOOKUP(R,SH!C:C,N,FALSE),"S")'],
      test: ['^IFERROR\\(VLOOKUP\\(R,SH!C:C,N,(FALSE|N)\\),"S"\\)$',
        '^VLOOKUP\\(R,SH!C:C,N,(FALSE|N)\\)$'],
      why: 'Every pipeline needs it: product code to description, barcode to product, ' +
        'farm to site.',
      teaches: ['VLOOKUP', 'whole-column ranges', 'exact match', 'IFERROR', '$ anchoring'],
      watch: 'col_index_num counts from the FIRST COLUMN OF THE BLOCK, not from column A. ' +
        'And VLOOKUP will not match a number against text that looks like the same number.'
    },
    {
      id: 'P6', title: 'Log-line slicing by character count', level: 11, layer: 2,
      cells: 400000, files: 2, rank: 6,
      oneLine: 'Classify a text log line by taking a fixed number of characters off an end.',
      example: '=IF(RIGHT(LVALL!C1,29)="FLASHLABEL; Invalid: Mismatch","Missing Sleeve",' +
        'IF(RIGHT(LVALL!C1,8)="Mismatch","Mismatch",""))',
      shapes: ['IF(RIGHT(SH!R,N)="S","S",IF(RIGHT(SH!R,N)="S","S","S"))',
        'IF(LEFT(SH!R,N)="S",LEFT(SH!R,N),"S")'],
      test: ['(RIGHT|LEFT)\\(SH!R,N\\)="S"'],
      why: 'The vision system emits a flat log, not a table, so somebody has to turn one ' +
        'text column into typed columns.',
      teaches: ['LEFT', 'RIGHT', 'text comparison', 'nested IF'],
      watch: 'The 29 and the 8 were counted by hand. The day the vendor rewords a ' +
        'message, every one of these 400,000 cells quietly stops classifying.',
      defect: 'D6'
    },
    {
      id: 'P7', title: 'Field extraction by delimiter', level: 11, layer: 2,
      cells: 17911, files: 17, rank: 7,
      oneLine: 'Take the nth delimited field out of a line, then convert it to a number.',
      example: '=IFERROR(VALUE(_xlfn.TEXTBEFORE(G1," ",1)),"")',
      shapes: ['IFERROR(_xlfn.TEXTAFTER(R,"S",N),"S")',
        'IFERROR(VALUE(_xlfn.TEXTBEFORE(R,"S",N)),"S")'],
      test: ['TEXT(AFTER|BEFORE)\\('],
      why: 'Counting fields instead of characters survives a change of wording.',
      teaches: ['TEXTBEFORE', 'TEXTAFTER', 'negative instance numbers', 'VALUE'],
      watch: 'A negative instance counts from the end: TEXTAFTER(M2," ",-1) is ' +
        '"everything after the last space". And slicing text gives TEXT, even when it ' +
        'looks like a number, so VALUE() is not optional.',
      pairWith: 'P6'
    },
    {
      id: 'P8', title: 'Week-ending date lookup', level: 10, layer: 4,
      cells: 177447, files: 13, rank: 8,
      oneLine: 'Round a date to its week-ending Saturday, then look up the retailer period.',
      example: '=IF(A4="","-",VLOOKUP(($B4-WEEKDAY($B4,1))+7,Calendars!$H:$J,3,FALSE))',
      shapes: ['IF(R="S","S",VLOOKUP((R-WEEKDAY(R,N))+N,SH!C:C,N,FALSE))'],
      test: ['WEEKDAY\\(R,N\\)\\)\\+N'],
      why: 'The business runs on retailer calendars, not ISO weeks.',
      teaches: ['date serials', 'WEEKDAY', 'date arithmetic', 'lookup into a date column'],
      watch: 'WEEKNUM is used ONCE in 152 workbooks. Retailer weeks do not match ISO ' +
        'weeks, so the calendar has to be a table.'
    },
    {
      id: 'P9', title: 'Two keys at once', level: 7, layer: 4,
      cells: 110961, files: 9, rank: 9,
      oneLine: 'Join on two columns by gluing them into one key with &.',
      example: '=IFERROR(_xlfn.XLOOKUP(FILTERWBTODAY!$B2&FILTERWBTODAY!G2,' +
        '\'MIMS SAVE FILE\'!$M:$M&\'MIMS SAVE FILE\'!V:V,\'MIMS SAVE FILE\'!J:J),"")',
      shapes: ['VLOOKUP(R&R,SH!R:R,N,N)',
        'IFERROR(_xlfn.XLOOKUP(SH!R&SH!R,SH!C:C&SH!C:C,SH!C:C),"S")'],
      test: ['XLOOKUP\\((SH!)?R&(SH!)?R,SH!C:C&SH!C:C', 'VLOOKUP\\(R&R,',
        '&R,SH!C:C&SH!C:C'],
      why: 'Barcode plus shift, week plus site: one column is not enough to identify a row.',
      teaches: ['& concatenation', 'array concatenation', 'helper columns', 'XLOOKUP'],
      watch: 'VLOOKUP needs a PHYSICAL helper column holding =F2&B2. XLOOKUP can build ' +
        'the key array inline. Learn both, because the old files are full of the first.'
    },
    {
      id: 'P10', title: 'XLOOKUP column fan-out', level: 6, layer: 4,
      cells: 99383, files: 9, rank: 10,
      oneLine: 'Rebuild a whole record beside a key: same key, one XLOOKUP per output column.',
      example: '=IFERROR(_xlfn.XLOOKUP(FILTERWBTODAY!$B2,\'MIMS SAVE FILE\'!$M:$M,' +
        '\'MIMS SAVE FILE\'!A:A),"")',
      shapes: ['IFERROR(_xlfn.XLOOKUP(SH!R,SH!C:C,SH!C:C),"S")',
        'IF(IFERROR(_xlfn.XLOOKUP(SH!R,SH!C:C,SH!C:C),"S")=N,"S")'],
      test: ['^IFERROR\\(_?x?l?f?n?\\.?XLOOKUP\\(SH!R,SH!C:C,SH!C:C\\),"S"\\)$',
        'XLOOKUP\\(SH!R,SH!C:C,SH!C:C\\)'],
      why: 'The scanner gives a barcode; the submission file needs the full 37-column ' +
        'product record in a fixed order.',
      teaches: ['XLOOKUP', '$ on the key', 'fill right', 'IFERROR'],
      watch: 'This is the right answer, written out 37 times by hand. Anchor the key ' +
        'column with $ and one formula fills across the lot.'
    },
    {
      id: 'P11', title: 'Fallback keys', level: 8, layer: 4,
      cells: 40000, files: 1, rank: 11,
      oneLine: 'Test three candidate keys for existence, use the first that exists, ' +
        'otherwise write a message to a human.',
      example: '=IFERROR(IF(L2=1,VLOOKUP(K2,\'Volumes Paste\'!$A:$G,7,FALSE),' +
        'IF(N2=1,VLOOKUP(M2,\'Volumes Paste\'!$B:$G,6,FALSE),' +
        'IF(P2=1,VLOOKUP(O2,\'Volumes Paste\'!$C:$G,5,FALSE),' +
        '"Check"&" "&"-"&" "&O2))),"Check"&" "&"-"&" "&O2)',
      shapes: ['IF(R="S","S",IF(COUNTIF(SH!C:C,SH!R)>=N,N,"S"))'],
      test: ['COUNTIF\\(SH!C:C,(SH!)?R\\)>=N', 'IF\\(R=N,VLOOKUP\\(R,SH!C:C,N,FALSE\\),IF\\(R=N,VLOOKUP'],
      why: 'A complaint record identifies a product by one of three different code ' +
        'systems depending on the retailer.',
      teaches: ['COUNTIF as an existence test', 'cascading IF', 'human-readable failure'],
      watch: 'The fallback is not an error value, it is the text "Check - 40231". The ' +
        'estate’s whole error philosophy in one formula: speak to the person who has ' +
        'to fix it.',
      defect: 'D3'
    },
    {
      id: 'P12', title: 'Date-window total', level: 9, layer: 5,
      cells: 17943, files: 3, rank: 12,
      oneLine: 'Total a week or period by giving SUMIFS two date bounds built with &.',
      example: '=SUMIFS(\'Volumes Paste\'!17:17,\'Volumes Paste\'!$13:$13,">=" &K$14,' +
        '\'Volumes Paste\'!$13:$13,"<="&K$15)',
      shapes: ['SUMIFS(SH!N:N,SH!$13:$13,"S"&R,SH!$13:$13,"S"&R)'],
      test: ['SUMIFS\\(.*"S"&R.*"S"&R'],
      why: 'Every report is "this week", "last period", "year to date".',
      teaches: ['SUMIFS', 'criteria as assembled strings', 'whole-row ranges', 'mixed anchoring'],
      watch: 'The criterion is a STRING: ">=" & K14 makes the text ">=46234". This is ' +
        'where most people come unstuck. Also note the ranges are ROWS, because the ' +
        'pasted volumes file has its dates across the top.'
    },
    {
      id: 'P13', title: 'Multi-criteria dashboard cell', level: 8, layer: 5,
      cells: 17325, files: 39, rank: 13,
      oneLine: 'One dashboard number = count of paste rows matching test code AND site ' +
        'AND week AND a flag.',
      example: '=COUNTIFS(\'Micro paste\'!$M:$M,$D$86,\'Micro paste\'!$H:$H,$C$86,' +
        '\'Micro paste\'!$B:$B,\'Reporting Template\'!$Q$23,\'Micro paste\'!$AA:$AA,"Yes")',
      shapes: ['COUNTIFS(SH!C:C,SH!R,SH!C:C,R,SH!C:C,SH!R,SH!C:C,"S")'],
      test: ['^COUNTIFS\\(SH!C:C,'],
      why: 'The entire KPI reporting layer is built from these.',
      teaches: ['COUNTIFS', 'four to six criteria pairs', 'mixed anchoring for fill ' +
        'both ways', 'cross-sheet criteria'],
      watch: 'B$133 against $A$132: row-locked versus fully-locked, so ONE formula fills ' +
        'both right and down. Get the $ wrong and it works in the first cell only. ' +
        'Highest-value single lesson for the KPI role.'
    },
    {
      id: 'P14', title: 'Manufacture a unique key', level: 7, layer: 4,
      cells: 10582, files: 2, rank: 14,
      oneLine: 'Number repeats 1, 2, 3 with COUNTIF on an expanding range, so the second ' +
        'and third match can be retrieved.',
      example: '=C5&":"&COUNTIF($C5:C$16,C5)',
      shapes: ['R&"S"&COUNTIF(R:R,R)', 'R&COUNTIF(R:R,R)'],
      test: ['R&"S"&COUNTIF\\(R:R,R\\)', 'R&COUNTIF\\(R:R,R\\)', 'COUNTIF\\(R:R,R\\)$'],
      why: 'VLOOKUP only ever finds the first match, so a unique key has to be invented ' +
        'to reach the 2nd and 3rd sample of the same product.',
      teaches: ['expanding ranges', 'the $ that makes a range grow one way', 'COUNTIF', '&'],
      watch: '$C5:C$16 grows as it fills down. Work out why that particular pair of $ ' +
        'signs does that and you have understood anchoring.'
    },
    {
      id: 'P15', title: 'Get the nth match', level: 16, layer: 4,
      cells: 2000, files: 1, rank: 15, cse: true,
      oneLine: 'Retrieve ALL matching rows, not just the first, with INDEX + SMALL + IF.',
      example: '=IFERROR(INDEX(FILTERCUT9AM!Q:Q,SMALL(IF((FILTERCUT9AM!B:B=FILTERCUT9AM!B2)*' +
        '(FILTERCUT9AM!F:F=FILTERCUT9AM!F2),ROW(FILTERCUT9AM!Q:Q)),' +
        'COUNTIFS(FILTERCUT9AM!$B$2:B2,FILTERCUT9AM!B2,FILTERCUT9AM!$F$2:F2,' +
        'FILTERCUT9AM!F2))),"")',
      shapes: ['IFERROR(INDEX(SH!C:C,SMALL(IF((SH!C:C=SH!R)*(SH!C:C=SH!R),ROW(SH!C:C)),' +
        'COUNTIFS(SH!R:R,SH!R,SH!R:R,SH!R))),"S")'],
      test: ['SMALL\\(IF\\('],
      why: 'This is the pre-FILTER idiom, and it is the ceiling of the estate’s ' +
        'complexity.',
      teaches: ['Ctrl+Shift+Enter arrays', 'boolean x boolean as AND', 'ROW(range) as a ' +
        'position generator', 'SMALL to pick the kth', 'COUNTIFS on an expanding range as k'],
      watch: 'Four ideas stacked in one cell. Build it from the inside out and it is ' +
        'readable; read it from the outside in and it is not.'
    },
    {
      id: 'P16', title: 'Boundary-match lookup', level: 16, layer: 4,
      cells: 9100, files: 5, rank: 16, cse: true,
      oneLine: 'Find which calendar week a date falls inside, where the calendar has ' +
        'start and end columns.',
      example: '=INDEX(Calendars!F:F, MATCH(1, (C2>=Calendars!G:G)*(C2<=Calendars!H:H), 0))',
      shapes: ['INDEX(SH!C:C,MATCH(N,(R>=SH!C:C)*(R<=SH!C:C),N))'],
      test: ['MATCH\\(N,\\(R>=SH![RC]'],
      why: 'The calendar is stored as date pairs, not as a single key.',
      teaches: ['the array-as-AND idiom', 'MATCH on a computed array', 'INDEX'],
      watch: 'The honest modern alternative is XLOOKUP(C2, Calendars!G:G, Calendars!F:F, ' +
        ', -1) with an approximate match. Worth comparing the two.'
    },
    {
      id: 'P17', title: 'Make the sheet check itself', level: 13, layer: 5,
      cells: 3000, files: 4, rank: 17,
      oneLine: 'Build a signature, compare it to the expected one, print a verdict for ' +
        'the operator.',
      example: '=IF(R2=1,"DATA MATCH","PLEASE REVIEW AND AMEND")',
      shapes: ['IF(R=R,N,N)', 'IF(R=N,"S","S")'],
      test: ['IF\\(R=R,"S","S"\\)', 'IF\\(R=N,"S","S"\\)', 'IF\\(R=R,N,N\\)'],
      why: 'The workbook checks itself and speaks to the person who has to fix it.',
      teaches: ['signature columns with &', 'comparison', 'verdict strings'],
      watch: 'This is the estate’s quality culture written as formulas. Treat adding ' +
        'a self-check column as part of the job, not an extra.'
    },
    {
      id: 'P18', title: 'Score to band (nested IF ladder)', level: 12, layer: 4,
      cells: 88814, files: 4, rank: 18,
      oneLine: 'Map a category to a value through nested IFs, up to six deep.',
      example: '=IF(S3="Aldi",U3,IF(S3="CO-OP",U3,IF(S3="Tesco",V3,IF(S3="M&S","2015/16",' +
        'IF(S3="Sainsbury",V3,IF(S3="Asda",V3,))))))',
      shapes: ['IF(SH!R=N,N,IF(SH!R=N,N,"S"))', 'IF(R="S",N,IF(R="S",N,N))'],
      test: ['^IF\\((SH!)?R=("S"|N),.*IF\\((SH!)?R=("S"|N),'],
      why: 'It works, and it is the first thing anyone reaches for.',
      teaches: ['nested IF', 'refactoring a ladder into a lookup table'],
      watch: 'Look at the end of the real example: ,V3,)))) -- an EMPTY final argument. ' +
        'That returns 0, not blank, and it prints on the report. A two-column table ' +
        'plus XLOOKUP replaces the whole thing.',
      defect: 'D4'
    },
    {
      id: 'P19', title: 'Spine and sequence generators', level: 10, layer: 1,
      cells: 37974, files: 17, rank: 19,
      oneLine: 'Manufacture the row or column spine a report hangs off.',
      example: '=+K13+1',
      alsoExample: '=IF(A2="","",IF((A2+1)<=$L$1,(A2+1),""))   (a self-limiting counter)',
      shapes: ['+R+N', 'ROW()', '+R', 'R+N',
        'IF(R="S","S",IF((R+N)<=R,(R+N),"S"))'],
      test: ['^\\+?R\\+N$', '^ROW\\(\\)$', '\\(R\\+N\\)<=R'],
      why: 'Dates, batch numbers, rolling windows: something has to count.',
      teaches: ['ROW()', 'self-limiting counters', 'COUNTA as a data-height probe'],
      watch: 'The estate has no SEQUENCE. =+K13+1 filled across is how a date spine gets ' +
        'built, and =IF(ROW()<COUNTA($B:$B)-49,"",…) is how a last-50-rows window gets built.'
    },
    {
      id: 'P20', title: 'Rate and percentage finishers', level: 9, layer: 5,
      cells: 5012, files: 12, rank: 20,
      oneLine: 'Flag each row 1/0, then average the flags to get a rate.',
      example: '=IF(U2>500,1,0)   then   =SUMIFS($Y:$Y,$V:$V,$V2,$AA:$AA,$AA2)/' +
        'COUNTIFS($V:$V,$V2,$AA:$AA,$AA2)',
      shapes: ['IFERROR((R/R)*N,"S")', 'IF(R>N,N,N)', 'SUMIFS(C:C,C:C,R,C:C,R)/COUNTIFS(C:C,R,C:C,R)'],
      test: ['^IFERROR\\(\\(R/R\\)\\*N,"S"\\)$', '^IF\\(R>N,N,N\\)$', 'SUMIFS\\(.*\\)/COUNTIFS\\('],
      why: 'The only real arithmetic in the estate: % compliance, % over limit, mean result.',
      teaches: ['IFERROR for divide-by-zero', 'the flag-then-average idiom', 'AVERAGEIFS'],
      watch: 'Flagging first, then averaging the flags, is easier to audit than one ' +
        'clever division -- and you can see the flag column.'
    }
  ];

  // Total measured formula mass, for the weighted progress figure.
  XT.data.patternMass = XT.data.patterns.reduce(function (t, p) { return t + p.cells; }, 0);

  XT.data.pattern = function (id) {
    var all = XT.data.patterns;
    for (var i = 0; i < all.length; i++) if (all[i].id === id) return all[i];
    return null;
  };

})(window.XT = window.XT || {});
