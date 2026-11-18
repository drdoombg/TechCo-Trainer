/* =============================================================================
   Excel Trainer — js/tests.js
   The conformance suite. Written BEFORE the lessons, because every exercise in
   the trainer is only as trustworthy as the engine underneath it, and the
   semantics that matter most here are the ones that are easy to get subtly
   wrong: EMPTY versus "", text-versus-number comparison, and what each error
   does when it meets an operator.

   Open the file tests.html in a browser. Green is a pass, red is a failure,
   amber is a case listed in docs/OPEN_QUESTIONS.md -- something where the
   expected answer needs confirming against real Excel. An amber case is never
   counted as a pass.
   ============================================================================= */
(function (XT) {
  'use strict';

  var V = XT.V, D = XT.D, F = XT.F, P = XT.P, S = XT.S;
  var tests = [];
  XT.testCases = tests;

  /* ------------------------------------------------------------- framework -- */

  var group = '';
  function G(name) { group = name; }

  // t(name, fn) -- fn throws on failure, via the assertions below.
  function t(name, fn, opts) {
    tests.push({ group: group, name: name, fn: fn, open: !!(opts && opts.open),
      note: opts && opts.note, ref: opts && opts.ref });
  }

  function show(v) {
    if (v === V.EMPTY) return 'EMPTY';
    if (V.isErr(v)) return v.code;
    if (V.isArr(v)) return 'array[' + v.flat().map(show).join(',') + ']';
    if (V.isRange(v)) return 'range';
    if (typeof v === 'string') return JSON.stringify(v);
    if (typeof v === 'number') return V.numToText(v);
    return String(v);
  }

  function fail(msg) { throw new Error(msg); }

  /* eq is STRICT about the thing that matters: EMPTY, "" and 0 are three
     different answers and the suite will not let them blur. */
  function eq(actual, expected, label) {
    var ok;
    if (expected === 'EMPTY') ok = (actual === V.EMPTY);
    else if (typeof expected === 'string' && /^#(N\/A|REF!|VALUE!|DIV\/0!|NAME\?|NUM!|NULL!)$/.test(expected)) {
      ok = V.isErr(actual) && actual.code === expected;
    } else if (typeof expected === 'number') {
      ok = (typeof actual === 'number') &&
        (Math.abs(actual - expected) < 1e-9 || actual === expected);
    } else if (typeof expected === 'string') {
      ok = (typeof actual === 'string') && actual === expected;
    } else if (typeof expected === 'boolean') {
      ok = (typeof actual === 'boolean') && actual === expected;
    } else ok = (actual === expected);
    if (!ok) fail((label ? label + ': ' : '') + 'expected ' + show(expected) +
      ' but got ' + show(actual));
  }

  function truthy(x, label) { if (!x) fail((label || 'assertion') + ' was not true'); }

  /* ------------------------------------------------------- scratch workbook -- */

  /* A tiny workbook to run formulas in. sheets is {Name: rows[][]} where a
     string starting with "=" is a formula and anything else is a literal. */
  function scratch(sheets, extra) {
    var wb = XT.newWorkbook('test');
    var def = { sheets: [] };
    sheets = sheets || { Sheet1: [[null]] };
    Object.keys(sheets).forEach(function (name) {
      def.sheets.push({ name: name, rows: sheets[name] });
    });
    if (!def.sheets.length) def.sheets.push({ name: 'Sheet1', rows: [] });
    if (extra && extra.names) def.names = extra.names;
    if (extra && extra.tables) def.tables = extra.tables;
    wb.load(def);
    var api = {
      wb: wb,
      // Evaluate a formula as if typed into a given cell (default Z1 on the
      // first sheet), and return its value.
      f: function (formula, at, sheetName) {
        var sn = sheetName || def.sheets[0].name;
        var a = XT.R.rc(at || 'Z1');
        wb.setInput(sn, a.r, a.c, formula.charAt(0) === '=' ? formula : '=' + formula);
        wb.recalc();
        return wb.cellValue(sn, a.r, a.c);
      },
      // Same, but entered as a legacy Ctrl+Shift+Enter array formula.
      cse: function (formula, at, sheetName) {
        var sn = sheetName || def.sheets[0].name;
        var a = XT.R.rc(at || 'Z1');
        wb.setInput(sn, a.r, a.c, formula.charAt(0) === '=' ? formula : '=' + formula, { cse: true });
        wb.recalc();
        return wb.cellValue(sn, a.r, a.c);
      },
      val: function (at, sheetName) {
        var a = XT.R.rc(at);
        return wb.cellValue(sheetName || def.sheets[0].name, a.r, a.c);
      },
      set: function (at, input, sheetName) {
        var a = XT.R.rc(at);
        wb.setInput(sheetName || def.sheets[0].name, a.r, a.c, input);
        wb.recalc();
        return api;
      },
      cell: function (at, sheetName) {
        var a = XT.R.rc(at);
        return wb.describe(sheetName || def.sheets[0].name, a.r, a.c);
      }
    };
    return api;
  }
  XT.scratch = scratch;

  // A serial we can name, so date tests do not depend on today.
  var NOV15 = D.ymdToSerial(2026, 11, 15);      // a Sunday
  var NOV11 = D.ymdToSerial(2026, 11, 11);      // a Wednesday

  /* =========================================================================
     §7.4  EMPTY versus "" versus a literal space
     The catalog calls this the most important semantic in the estate: 1.2M+
     formulas end in ,"") and every one of them makes a cell that looks blank
     and is not.
     ========================================================================= */

  G('§7.4 EMPTY vs "" vs space');

  /* One column, five different kinds of "nearly nothing":
       A1  truly empty -- nobody has ever typed into it
       A2  holds the formula =""   -> TEXT of zero length
       A3  holds a single literal space
       A4  holds the number 0
       A5  holds the text "abc"                                        */
  function threeStates() {
    return scratch({ T: [[null], ['=""'], [' '], [0], ['abc']] });
  }

  var table74 = [
    // formula,                      empty A1,  ""  A2,    space A3
    ['ISBLANK(%)', true, false, false],
    ['%=""', true, true, false],
    ['COUNTA(%)', 0, 1, 1],
    ['COUNT(%)', 0, 0, 0],
    ['COUNTBLANK(%)', 1, 1, 0],
    ['%>0', false, true, true],
    ['SUM(%)', 0, 0, 0],
    ['%&"!"', '!', '!', ' !'],
    ['%=0', true, false, false],
    ['LEN(%)', 0, 0, 1],
    ['ISTEXT(%)', false, true, true],
    ['ISNUMBER(%)', false, false, false],
    ['%=" "', false, false, true],
    ['IF(%="","blank","not")', 'blank', 'blank', 'not']
  ];

  ['A1', 'A2', 'A3'].forEach(function (addr, col) {
    table74.forEach(function (row) {
      var formula = row[0].replace(/%/g, addr);
      var expect = row[1 + col];
      var stateName = ['truly empty', '="" (empty text)', 'a literal space'][col];
      t('=' + formula + '   where ' + addr + ' is ' + stateName, function () {
        var s = threeStates();
        eq(s.f('=' + formula, 'Z1', 'T'), expect);
      });
    });
  });

  t('a cell holding ="" is text, not EMPTY', function () {
    var s = threeStates();
    eq(s.val('A2', 'T'), '');
    eq(s.val('A1', 'T'), 'EMPTY');
    // typeName is compared as a plain label, not through eq's error/EMPTY sentinels.
    truthy(V.typeName(s.val('A2', 'T')) === 'text ""', 'A2 reports as empty text');
    truthy(V.typeName(s.val('A1', 'T')) === 'EMPTY', 'A1 reports as EMPTY');
    truthy(V.typeName(s.val('A3', 'T')) === 'text', 'A3 reports as text');
  });

  /* The display rule that P2 and P3 exist to work around. */
  t('a FORMULA reading an empty cell shows 0, though its value is EMPTY', function () {
    var s = scratch({ Src: [[null]], Dst: [[null]] });
    s.set('A1', '=Src!A1', 'Dst');
    eq(s.val('A1', 'Dst'), 'EMPTY', 'the value is still EMPTY');
    truthy(s.wb.display('Dst', 0, 0) === '0', 'but it displays as 0');
    truthy(s.wb.display('Src', 0, 0) === '', 'while the empty cell itself displays nothing');
  }, { ref: 'P2' });

  t('which is why the guarded mirror is needed at all', function () {
    var s = scratch({ Src: [[null]], Dst: [[null]] });
    s.set('A1', '=IF(Src!A1="","",Src!A1)', 'Dst');
    truthy(s.wb.display('Dst', 0, 0) === '', 'the guard turns the 0 back into nothing');
    eq(s.val('A1', 'Dst'), '');
  }, { ref: 'P3' });

  t('COUNTA and COUNTBLANK disagree about "" -- and both are right', function () {
    var s = scratch({ T: [['=""'], ['=""'], [null]] });
    eq(s.f('=COUNTA(A1:A3)', 'Z1', 'T'), 2, 'COUNTA sees the two "" cells');
    eq(s.f('=COUNTBLANK(A1:A3)', 'Z2', 'T'), 3, 'COUNTBLANK counts all three as blank');
  });

  /* =========================================================================
     §7.5  Coercion
     ========================================================================= */

  G('§7.5 coercion');

  t('"10"=10 is FALSE -- the = operator does not coerce types', function () {
    eq(scratch().f('="10"=10'), false);
  });

  t('"10">5 is TRUE -- any text is greater than any number', function () {
    eq(scratch().f('="10">5'), true);
  });

  t('a cell of text "5" does not equal the number 5', function () {
    var s = scratch({ T: [['5']] });     // typed as text via the loader? check both ways
    s.set('A1', "'5", 'T');              // leading apostrophe forces text
    eq(s.cell('A1', 'T').type, 'text');
    eq(s.f('=A1=5', 'Z1', 'T'), false);
  });

  t('but COUNTIF matches it -- criteria are looser than =', function () {
    var s = scratch({ T: [[null]] });
    s.set('A1', "'5", 'T');
    eq(s.f('=COUNTIF(A1,5)', 'Z1', 'T'), 1);
  });

  t('VALUE trims ordinary spaces', function () {
    eq(scratch().f('=VALUE("  12.5 ")'), 12.5);
  });

  t('VALUE("abc") is #VALUE!', function () {
    eq(scratch().f('=VALUE("abc")'), '#VALUE!');
  });

  t('VALUE cannot strip CHAR(160) -- defect D7', function () {
    eq(scratch().f('=VALUE("5"&CHAR(160))'), '#VALUE!');
  }, { ref: 'D7' });

  t('TRIM does not remove CHAR(160) either -- defect D7', function () {
    eq(scratch().f('=LEN(TRIM("5"&CHAR(160)))'), 2);
  }, { ref: 'D7' });

  t('SUBSTITUTE(x,CHAR(160),"") then VALUE does work -- the fix', function () {
    eq(scratch().f('=VALUE(SUBSTITUTE("5"&CHAR(160),CHAR(160),""))'), 5);
  }, { ref: 'D7' });

  t('1/3&"" gives 15 significant figures', function () {
    eq(scratch().f('=1/3&""'), '0.333333333333333');
  });

  t('(0.1+0.2)&"" gives "0.3", not the binary artefact', function () {
    eq(scratch().f('=(0.1+0.2)&""'), '0.3');
  });

  t('TRUE*1 is 1', function () { eq(scratch().f('=TRUE*1'), 1); });

  t('AND(...)*1 turns a boolean into 1/0 -- the estate idiom', function () {
    eq(scratch().f('=AND(2>1,3>2)*1'), 1);
    eq(scratch().f('=AND(2>1,1>3)*1'), 0);
  });

  t('">="&46234 builds the criterion text ">=46234"', function () {
    eq(scratch().f('=">="&46234'), '>=46234');
  });

  t('text in a RANGE is ignored by SUM; a direct text argument is coerced', function () {
    var s = scratch({ T: [[null]] });
    s.set('A1', "'2", 'T');
    eq(s.f('=SUM(A1)', 'Z1', 'T'), 0, 'text inside a range');
    eq(s.f('=SUM("2")', 'Z2', 'T'), 2, 'direct text argument');
  });

  t('COUNT treats a range and a direct argument differently too', function () {
    var s = scratch({ T: [[null]] });
    s.set('A1', "'2", 'T');
    eq(s.f('=COUNT(A1)', 'Z1', 'T'), 0);
    eq(s.f('=COUNT("2")', 'Z2', 'T'), 1);
  });

  t('a number and text compare by type, not by appearance', function () {
    eq(scratch().f('=1&""="1"'), true);
    eq(scratch().f('=1="1"'), false);
  });

  /* =========================================================================
     §7.3  Errors
     ========================================================================= */

  G('§7.3 errors');

  var ERRS = ['#N/A', '#REF!', '#VALUE!', '#DIV/0!', '#NAME?', '#NUM!', '#NULL!'];

  ERRS.forEach(function (e) {
    t(e + ' propagates through +', function () {
      eq(scratch().f('=' + e + '+1'), e);
    });
    t(e + ' propagates through &', function () {
      eq(scratch().f('=' + e + '&"x"'), e);
    });
    t(e + ' propagates through a comparison', function () {
      eq(scratch().f('=' + e + '=1'), e);
    });
    t('IFERROR catches ' + e, function () {
      eq(scratch().f('=IFERROR(' + e + ',"caught")'), 'caught');
    });
    t('ISERROR(' + e + ') is TRUE', function () {
      eq(scratch().f('=ISERROR(' + e + ')'), true);
    });
  });

  t('IFNA catches #N/A', function () {
    eq(scratch().f('=IFNA(#N/A,"caught")'), 'caught');
  });

  t('IFNA does NOT catch #DIV/0! -- this is the D2 lesson', function () {
    eq(scratch().f('=IFNA(1/0,"caught")'), '#DIV/0!');
  }, { ref: 'D2' });

  t('IFNA does NOT catch #REF! -- a broken formula stays visible', function () {
    eq(scratch().f('=IFNA(#REF!,"caught")'), '#REF!');
  }, { ref: 'D2' });

  t('ISNA tells #N/A apart from the rest', function () {
    eq(scratch().f('=ISNA(#N/A)'), true);
    eq(scratch().f('=ISNA(#REF!)'), false);
  });

  t('an error inside a range reaches SUM', function () {
    var s = scratch({ T: [[1], ['=#REF!'], [3]] });
    eq(s.f('=SUM(A1:A3)', 'Z1', 'T'), '#REF!');
  });

  t('IFERROR records what it swallowed, so the inspector can show it', function () {
    var s = scratch({ T: [[1]] });
    s.set('B1', '=IFERROR(VLOOKUP("nope",A1:A1,1,FALSE),"")', 'T');
    var d = s.cell('B1', 'T');
    eq(d.value, '');
    truthy(d.hidden && d.hidden.length === 1, 'one hidden error recorded');
    truthy(d.hidden[0].code === '#N/A', 'and it was the #N/A from the failed lookup');
    truthy(/VLOOKUP/.test(d.hidden[0].from), 'the inspector can name the guilty argument');
  }, { ref: 'D1' });

  t('a circular reference gives 0 and is reported, not a hang', function () {
    var s = scratch({ T: [[null]] });
    s.set('A1', '=A1+1', 'T');
    eq(s.val('A1', 'T'), 0);
    truthy(s.wb.circular.length >= 1, 'the circular cell is named');
  });

  t('a divide by zero names the numerator in its explanation', function () {
    var v = scratch().f('=7/0');
    eq(v, '#DIV/0!');
    truthy(/7/.test(v.detail || ''), 'the detail mentions the 7');
  });

  /* =========================================================================
     §7.6  Dates
     ========================================================================= */

  G('§7.6 dates');

  t('serial 1 is 1900-01-01', function () {
    eq(scratch().f('=DATE(1900,1,1)'), 1);
  });

  t('2000-01-01 is serial 36526', function () {
    eq(scratch().f('=DATE(2000,1,1)'), 36526);
  });

  t('WEEKDAY(36526,1) is 7 -- 2000-01-01 was a Saturday', function () {
    eq(scratch().f('=WEEKDAY(36526,1)'), 7);
  });

  t('WEEKDAY type 2 makes Monday 1', function () {
    eq(scratch().f('=WEEKDAY(DATE(2026,11,16),2)'), 1, '2026-11-16 is a Monday');
  });

  t('(date - WEEKDAY(date,1)) + 7 lands on the week-ending Saturday -- P8', function () {
    var s = scratch();
    var v = s.f('=(' + NOV11 + '-WEEKDAY(' + NOV11 + ',1))+7');
    truthy(typeof v === 'number', 'it is a number');
    eq(D.weekday(v, 1), 7, 'the result is a Saturday');
    truthy(v >= NOV11 && v - NOV11 < 7, 'and it is within the same week');
  }, { ref: 'P8' });

  t('TIME(12,0,0) is 0.5', function () { eq(scratch().f('=TIME(12,0,0)'), 0.5); });

  t('TIME/HOUR/MINUTE/SECOND round-trip', function () {
    var s = scratch();
    eq(s.f('=HOUR(TIME(13,45,30))'), 13);
    eq(s.f('=MINUTE(TIME(13,45,30))', 'Z2'), 45);
    eq(s.f('=SECOND(TIME(13,45,30))', 'Z3'), 30);
  });

  t('rebuilding a time from its parts strips the date -- the shift cut-off idiom', function () {
    var s = scratch({ T: [[{ v: NOV15 + D.hmsToSerial(5, 30, 0), fmt: 'dd/mm/yyyy hh:mm' }]] });
    var v = s.f('=TIME(HOUR(A1),MINUTE(A1),SECOND(A1))', 'Z1', 'T');
    eq(v, D.hmsToSerial(5, 30, 0));
    truthy(v < 1, 'it is a fraction of a day, not a timestamp');
  });

  t('0/1/1900 evaluates to 0 -- it is arithmetic, not a date (defect D3)', function () {
    eq(scratch().f('=0/1/1900'), 0);
  }, { ref: 'D3' });

  t('IF(x=0/1/1900,…) can therefore only ever test for zero', function () {
    var s = scratch({ T: [[0], [NOV15]] });
    eq(s.f('=IF(A1=0/1/1900,"matched","no")', 'Z1', 'T'), 'matched');
    eq(s.f('=IF(A2=0/1/1900,"matched","no")', 'Z2', 'T'), 'no');
  }, { ref: 'D3' });

  t('DATE(1900,1,1) is what was meant, and it is not 0', function () {
    eq(scratch().f('=DATE(1900,1,1)=0/1/1900'), false);
  }, { ref: 'D3' });

  t('TEXT(date,"MMM-YY") gives the period label', function () {
    eq(scratch().f('=TEXT(' + NOV15 + ',"MMM-YY")'), 'Nov-26');
  });

  t('TEXT(date,"dd/mm/yyyy") is UK order', function () {
    eq(scratch().f('=TEXT(' + NOV15 + ',"dd/mm/yyyy")'), '15/11/2026');
  });

  t('TEXT(date,"m/d/yyyy") is what a US export arrives as', function () {
    eq(scratch().f('=TEXT(' + NOV15 + ',"m/d/yyyy")'), '11/15/2026');
  });

  t('TEXT(date,"ddd") and "dddd" give day names', function () {
    eq(scratch().f('=TEXT(' + NOV15 + ',"ddd")'), 'Sun');
    eq(scratch().f('=TEXT(' + NOV15 + ',"dddd")', 'Z2'), 'Sunday');
  });

  t('the m in hh:mm is minutes, the m in dd/mm is months', function () {
    var s = scratch();
    var ts = D.ymdToSerial(2026, 3, 4) + D.hmsToSerial(9, 7, 0);
    eq(s.f('=TEXT(' + ts + ',"dd/mm/yyyy hh:mm")'), '04/03/2026 09:07');
  });

  t('INT strips the time off a timestamp', function () {
    eq(scratch().f('=INT(' + (NOV15 + 0.75) + ')'), NOV15);
  });

  t('TODAY() is frozen for the trainer, so exercises grade the same tomorrow', function () {
    D.frozenToday = NOV15;
    try { eq(scratch().f('=TODAY()'), NOV15); }
    finally { D.frozenToday = null; }
  });

  /* =========================================================================
     §7.1 / §7.2  Parser
     ========================================================================= */

  G('§7.1-7.2 parser');

  t('a leading unary + parses: =+K13+1', function () {
    var s = scratch({ T: [[null]] });
    s.set('K13', 5, 'T');
    eq(s.f('=+K13+1', 'Z1', 'T'), 6);
  });

  t('=+Sheet!A1 parses too', function () {
    var s = scratch({ Home: [[7]], T: [[null]] });
    eq(s.f('=+Home!A1', 'Z1', 'T'), 7);
  });

  t('unary plus leaves text alone', function () {
    eq(scratch().f('=+"abc"'), 'abc');
  }, { open: true, note: 'Excel treats unary + as a no-op; confirm =+"abc" returns abc.' });

  t('unary minus binds tighter than ^ , so -2^2 is 4', function () {
    eq(scratch().f('=-2^2'), 4);
  });

  t('^ is left-associative here, so 2^3^2 is 64', function () {
    eq(scratch().f('=2^3^2'), 64);
  }, { open: true, note: 'Confirm whether Excel gives 64 (left) or 512 (right) for =2^3^2.' });

  t('precedence: 2+3*4 is 14', function () { eq(scratch().f('=2+3*4'), 14); });
  t('& binds looser than + : 1+2&"x" is "3x"', function () { eq(scratch().f('=1+2&"x"'), '3x'); });
  t('comparison binds loosest: 1+1=2 is TRUE', function () { eq(scratch().f('=1+1=2'), true); });
  t('percent is postfix: 50% is 0.5', function () { eq(scratch().f('=50%'), 0.5); });
  t('a doubled quote is one quote inside text', function () {
    eq(scratch().f('="say ""hi"""'), 'say "hi"');
  });

  t('_xlfn.XLOOKUP is the same function as XLOOKUP', function () {
    var s = scratch({ T: [['a', 1], ['b', 2]] });
    eq(s.f('=_xlfn.XLOOKUP("b",A1:A2,B1:B2)', 'Z1', 'T'), 2);
    eq(s.f('=XLOOKUP("b",A1:A2,B1:B2)', 'Z2', 'T'), 2);
  });

  t('a quoted sheet name with spaces works', function () {
    var s = scratch({ 'MIMS SAVE FILE': [[42]], T: [[null]] });
    eq(s.f("='MIMS SAVE FILE'!A1", 'Z1', 'T'), 42);
  });

  t('a quoted sheet name containing & works', function () {
    var s = scratch({ 'M&S Vision': [['ok']], T: [[null]] });
    eq(s.f("='M&S Vision'!A1", 'Z1', 'T'), 'ok');
  });

  t('a whole-row range parses and clips: 1:1', function () {
    var s = scratch({ T: [[1, 2, 3]] });
    eq(s.f('=SUM(T!1:1)', 'Z5', 'T'), 6);
  });

  t('a whole-column range parses and clips: $A:$C', function () {
    var s = scratch({ T: [[1, 2, 3], [4, 5, 6]] });
    eq(s.f('=SUM(T!$A:$C)', 'Z5', 'T'), 21);
  });

  t('a whole-column range is clipped to the used range, not a million rows', function () {
    var s = scratch({ T: [[1], [2], [3]] });
    var before = s.wb.evalCount;
    eq(s.f('=SUM(A:A)', 'Z1', 'T'), 6);
    truthy(s.wb.evalCount - before < 100, 'it did not walk the whole column');
  }, { ref: 'D13' });

  t('#REF! written inside a formula evaluates to #REF! and does not crash', function () {
    eq(scratch().f('=IFERROR(VLOOKUP("x",#REF!,8,0),"NA")'), 'NA');
    eq(scratch().f('=SUM(#REF!)', 'Z2'), '#REF!');
    eq(scratch().f('=1+#REF!', 'Z3'), '#REF!');
  }, { ref: 'D1' });

  t('a sheet-qualified #REF! parses: =+\'QTR 2 Old\'!#REF!', function () {
    var s = scratch({ 'QTR 2 Old': [[1]], T: [[null]] });
    eq(s.f("=+'QTR 2 Old'!#REF!", 'Z1', 'T'), '#REF!');
  });

  t('an external reference parses and reports that the link is unavailable', function () {
    var v = scratch().f('=[1]CODES!$G$4');
    eq(v, '#REF!');
    truthy(/not open/.test(v.detail || ''), 'the message says the workbook is not open');
  }, { ref: 'D10' });

  t('an external reference inside VLOOKUP still parses', function () {
    eq(scratch().f("=IFERROR(VLOOKUP(\"x\",[1]CODES!$G$4:$I$4975,3,0),\"nope\")"), 'nope');
  });

  t('a structured This-Row reference resolves', function () {
    var s = scratch({
      T: [['Barcode', 'SUPPLIER'], ['abc', 5], ['def', 9]]
    }, { tables: [{ name: 'Table4', sheet: 'T', ref: 'A1:B3', headers: ['Barcode', 'SUPPLIER'] }] });
    s.set('D2', '=Table4[[#This Row],[SUPPLIER]]', 'T');
    eq(s.val('D2', 'T'), 5);
    s.set('D3', '=Table4[[#This Row],[SUPPLIER]]*2', 'T');
    eq(s.val('D3', 'T'), 18);
  });

  t('a structured column reference resolves to the data range', function () {
    var s = scratch({
      T: [['Barcode', 'Qty'], ['abc', 5], ['def', 9]]
    }, { tables: [{ name: 'Table1', sheet: 'T', ref: 'A1:B3', headers: ['Barcode', 'Qty'] }] });
    eq(s.f('=SUM(Table1[Qty])', 'Z1', 'T'), 14);
  });

  t('a defined name resolves to its range', function () {
    var s = scratch({ picklists: [['Days'], ['Mon'], ['Tue'], ['Wed']], T: [[null]] },
      { names: { Shift: 'picklists!$A$2:$A$4' } });
    eq(s.f('=COUNTA(Shift)', 'Z1', 'T'), 3);
  });

  t('a defined name pointing at #REF! says so -- defect D8', function () {
    var s = scratch({ T: [[null]] }, { names: { boom: '#REF!' } });
    var v = s.f('=boom', 'Z1', 'T');
    eq(v, '#REF!');
    truthy(/boom/.test(v.detail || ''), 'the message names the broken name');
  }, { ref: 'D8' });

  t('an unknown name gives #NAME? with a useful message', function () {
    var v = scratch().f('=NoSuchThing');
    eq(v, '#NAME?');
  });

  t('an argument written as nothing: =TEXT(C1,) returns the plain text', function () {
    var s = scratch({ T: [[null]] });
    s.set('C1', 42, 'T');
    eq(s.f('=TEXT(C1,)', 'Z1', 'T'), '42');
  });

  t('=IF(FALSE,1,) returns 0 -- an empty third argument, not an absent one', function () {
    eq(scratch().f('=IF(1=2,1,)'), 0);
  }, { open: true, ref: 'D4',
    note: 'Confirm: =IF(FALSE,1,) gives 0 while =IF(FALSE,1) gives FALSE. The catalog ' +
      'entry for defect D4 says FALSE for the trailing-comma form; this suite says 0.' });

  t('=IF(FALSE,1) with no third argument at all returns FALSE', function () {
    eq(scratch().f('=IF(1=2,1)'), false);
  }, { open: true, ref: 'D4' });

  t('implicit intersection: a column range used as one value takes this row', function () {
    var s = scratch({ T: [[10], [20], [30]] });
    var a = XT.R.rc('C2');
    s.wb.setInput('T', a.r, a.c, '=A1:A3');
    s.wb.recalc();
    eq(s.val('C2', 'T'), 20, 'row 2 of the range');
  });

  t('implicit intersection with no overlap is #VALUE!', function () {
    var s = scratch({ T: [[10], [20], [30]] });
    eq(s.f('=A1:A3', 'Z9', 'T'), '#VALUE!');
  });

  t('a 1x1 range used as a value is just that value', function () {
    var s = scratch({ T: [[7]] });
    eq(s.f('=SUMIF(A1,">0")', 'Z1', 'T'), 7, 'a single cell is a legal range');
    eq(s.f('=A1:A1', 'Z2', 'T'), 7);
  });

  t('every formula shape in the parser test list round-trips through unparse', function () {
    var samples = [
      '=IFERROR(VLOOKUP(C18,PASTE!$A:$DN,2,FALSE),"")',
      '=+K13+1',
      '=SUMIFS(\'Volumes Paste\'!17:17,\'Volumes Paste\'!$13:$13,">=" &K$14)',
      '=IF(RIGHT(LVALL!C1,29)="x","Missing Sleeve","")',
      '=C5&":"&COUNTIF($C5:C$16,C5)',
      '=_xlfn.TEXTAFTER(M2," ",-1)',
      '=50%*2'
    ];
    samples.forEach(function (src) {
      var a = P.parse(src);
      truthy(a.ok, 'parses: ' + src);
      var again = P.parse('=' + P.unparse(a.ast));
      truthy(again.ok, 're-parses: ' + P.unparse(a.ast));
      eq(S.normalise('=' + P.unparse(a.ast)), S.normalise(src), 'same shape after round-trip');
    });
  });

  /* =========================================================================
     Fill translation -- the main interaction (§7.9)
     ========================================================================= */

  G('fill translation');

  var fillCases = [
    ['=A1', 1, 0, '=A2', 'relative row moves down'],
    ['=A1', 0, 1, '=B1', 'relative column moves right'],
    ['=$A$1', 5, 5, '=$A$1', 'fully absolute never moves'],
    ['=A$1', 3, 0, '=A$1', 'row-locked does not move down'],
    ['=A$1', 0, 2, '=C$1', 'row-locked still moves right'],
    ['=$A1', 0, 3, '=$A1', 'column-locked does not move right'],
    ['=$A1', 2, 0, '=$A3', 'column-locked still moves down'],
    ['=PASTE!$A:$DN', 500, 0, '=PASTE!$A:$DN', 'an absolute whole-column range never moves'],
    ['=PASTE!A:D', 0, 1, '=PASTE!B:E', 'a relative whole-column range DOES move right'],
    ['=SUM(A1:A5)', 1, 0, '=SUM(A2:A6)', 'a range moves as a block'],
    ['=COUNTIF($C5:C$16,C5)', 1, 0, '=COUNTIF($C6:C$16,C6)', 'the P14 expanding range grows'],
    ['=\'Micro paste\'!$M:$M', 1, 1, '=\'Micro paste\'!$M:$M', 'quoted sheet, absolute column'],
    ['=COUNTIFS(P!$M:$M,$D$86,P!$H:$H,C$1)', 0, 1,
      '=COUNTIFS(P!$M:$M,$D$86,P!$H:$H,D$1)', 'P13 mixed anchoring fills right correctly']
  ];

  fillCases.forEach(function (c) {
    t(c[0] + '  filled ' + c[1] + ' down, ' + c[2] + ' right  ->  ' + c[3] + '   (' + c[4] + ')',
      function () { eq(P.translateText(c[0], c[1], c[2]), c[3]); });
  });

  t('filling above row 1 produces #REF!, as in Excel', function () {
    eq(P.translateText('=A1', -1, 0), '=#REF!');
  });

  t('a fill-down of the guarded join keeps the lookup block still', function () {
    var src = '=IFERROR(VLOOKUP(A2,PASTE!$A:$D,2,FALSE),"")';
    eq(P.translateText(src, 1, 0), '=IFERROR(VLOOKUP(A3,PASTE!$A:$D,2,FALSE),"")');
  });

  t('the same join WITHOUT $ rots as it fills -- the lesson behind D13', function () {
    var src = '=IFERROR(VLOOKUP(A2,PASTE!A2:D500,2,FALSE),"")';
    eq(P.translateText(src, 1, 0), '=IFERROR(VLOOKUP(A3,PASTE!A3:D501,2,FALSE),"")');
  });

  /* =========================================================================
     Criteria engine
     ========================================================================= */

  G('criteria engine');

  function critSheet() {
    return scratch({ T: [[5, 'Apple'], [10, 'Banana'], [15, 'apple'], [0, ''], [null, null]] });
  }

  var critCases = [
    ['=COUNTIF(A1:A5,">5")', 2],
    ['=COUNTIF(A1:A5,">=5")', 3],
    ['=COUNTIF(A1:A5,"<>")', 4, 'non-blank'],
    ['=COUNTIF(A1:A5,"")', 1, 'only the truly empty cell is blank in column A'],
    ['=COUNTIF(B1:B5,"apple")', 2, 'text criteria are case-insensitive'],
    ['=COUNTIF(B1:B5,"A*")', 2, 'the * wildcard'],
    ['=COUNTIF(B1:B5,"?pple")', 2, 'the ? wildcard'],
    ['=COUNTIF(B1:B5,"<>Apple")', 3, 'negated text'],
    ['=COUNTIF(A1:A5,10)', 1],
    ['=SUMIF(A1:A5,">5")', 25],
    ['=SUMIF(A1:A5,">5",A1:A5)', 25, 'explicit sum_range'],
    ['=SUMIFS(A1:A5,A1:A5,">=5",A1:A5,"<=10")', 15],
    ['=COUNTIFS(A1:A5,">0",B1:B5,"A*")', 2],
    ['=AVERAGEIFS(A1:A5,A1:A5,">0")', 10]
  ];
  critCases.forEach(function (c) {
    t(c[0] + (c[2] ? '   (' + c[2] + ')' : ''), function () {
      eq(critSheet().f(c[0], 'Z1', 'T'), c[1]);
    });
  });

  t('a wildcard can be escaped with ~', function () {
    var s = scratch({ T: [['a*b'], ['axb']] });
    eq(s.f('=COUNTIF(A1:A2,"a~*b")', 'Z1', 'T'), 1);
  });

  t('COUNTIFS with mismatched range shapes is #VALUE!', function () {
    var s = scratch({ T: [[1, 2], [3, 4]] });
    eq(s.f('=COUNTIFS(A1:A2,">0",B1:B1,">0")', 'Z1', 'T'), '#VALUE!');
  });

  t('a date window built with & works as a criterion -- P12', function () {
    var s = scratch({ T: [[NOV11, 100], [NOV15, 200], [NOV15 + 10, 400]] });
    s.set('D1', NOV11, 'T').set('D2', NOV15, 'T');
    eq(s.f('=SUMIFS(B1:B3,A1:A3,">="&D1,A1:A3,"<="&D2)', 'Z1', 'T'), 300);
  }, { ref: 'P12' });

  t('SUMIF over a ROW range works -- the transposed volumes file', function () {
    var s = scratch({ T: [['x', 'y', 'x'], [1, 2, 4]] });
    eq(s.f('=SUMIF(T!1:1,"x",T!2:2)', 'Z5', 'T'), 5);
  });

  /* =========================================================================
     Lookups
     ========================================================================= */

  G('lookups');

  function lookSheet() {
    // A product table with the key in column A and three attributes after it.
    return scratch({
      PASTE: [['K100', 'Fillet', 180, 'Aldi'],
        ['K200', 'Thigh', 220, 'Tesco'],
        ['K300', 'Wing', 90, 'M&S']],
      T: [[null]]
    });
  }

  t('VLOOKUP exact finds a row', function () {
    eq(lookSheet().f('=VLOOKUP("K200",PASTE!$A:$D,2,FALSE)', 'Z1', 'T'), 'Thigh');
  });

  t('VLOOKUP col_index_num counts from the block, not from column A', function () {
    var s = lookSheet();
    // The block starts at column B, so the key must be in B and index 2 is C.
    eq(s.f('=VLOOKUP("Thigh",PASTE!$B:$D,2,FALSE)', 'Z1', 'T'), 220);
    // Index 2 of a block starting at A is B, which is a different column.
    eq(s.f('=VLOOKUP("K200",PASTE!$A:$D,2,FALSE)', 'Z2', 'T'), 'Thigh');
  });

  t('VLOOKUP exact, not found, is #N/A', function () {
    eq(lookSheet().f('=VLOOKUP("K999",PASTE!$A:$D,2,FALSE)', 'Z1', 'T'), '#N/A');
  });

  t('VLOOKUP with a col_index_num past the block is #REF!', function () {
    eq(lookSheet().f('=VLOOKUP("K200",PASTE!$A:$B,5,FALSE)', 'Z1', 'T'), '#REF!');
  });

  t('VLOOKUP with col_index_num 0 is #VALUE!', function () {
    eq(lookSheet().f('=VLOOKUP("K200",PASTE!$A:$D,0,FALSE)', 'Z1', 'T'), '#VALUE!');
  });

  t('VLOOKUP is case-insensitive on text keys', function () {
    eq(lookSheet().f('=VLOOKUP("k200",PASTE!$A:$D,2,FALSE)', 'Z1', 'T'), 'Thigh');
  });

  t('VLOOKUP will NOT match a number against numeric text -- the classic failure', function () {
    var s = scratch({ PASTE: [[null, 'Fillet']], T: [[null]] });
    s.set('A1', "'5002874", 'PASTE');           // text that looks like a number
    eq(s.f('=VLOOKUP(5002874,PASTE!$A:$B,2,FALSE)', 'Z1', 'T'), '#N/A');
    var v = s.val('Z1', 'T');
    truthy(/STORED AS TEXT/.test(v.detail || ''), 'and it explains why');
  });

  t('the same lookup works once the key is turned into text', function () {
    var s = scratch({ PASTE: [[null, 'Fillet']], T: [[null]] });
    s.set('A1', "'5002874", 'PASTE');
    eq(s.f('=VLOOKUP(5002874&"",PASTE!$A:$B,2,FALSE)', 'Z1', 'T'), 'Fillet');
  });

  t('VLOOKUP approximate takes the last value not greater than the key', function () {
    var s = scratch({ B: [[0, 'low'], [10, 'mid'], [20, 'high']], T: [[null]] });
    eq(s.f('=VLOOKUP(15,B!$A:$B,2,TRUE)', 'Z1', 'T'), 'mid');
    eq(s.f('=VLOOKUP(15,B!$A:$B,2)', 'Z2', 'T'), 'mid', 'TRUE is the default');
  });

  t('VLOOKUP approximate below the first key is #N/A', function () {
    var s = scratch({ B: [[10, 'mid'], [20, 'high']], T: [[null]] });
    eq(s.f('=VLOOKUP(5,B!$A:$B,2,TRUE)', 'Z1', 'T'), '#N/A');
  });

  t('XLOOKUP defaults to an exact match', function () {
    eq(lookSheet().f('=XLOOKUP("K300",PASTE!$A:$A,PASTE!$C:$C)', 'Z1', 'T'), 90);
  });

  t('XLOOKUP if_not_found replaces #N/A', function () {
    eq(lookSheet().f('=XLOOKUP("K999",PASTE!$A:$A,PASTE!$C:$C,"missing")', 'Z1', 'T'), 'missing');
  });

  t('XLOOKUP match_mode -1 takes the next smaller', function () {
    var s = scratch({ B: [[0, 'low'], [10, 'mid'], [20, 'high']], T: [[null]] });
    eq(s.f('=XLOOKUP(15,B!$A:$A,B!$B:$B,,-1)', 'Z1', 'T'), 'mid');
  });

  t('XLOOKUP match_mode 1 takes the next larger', function () {
    var s = scratch({ B: [[0, 'low'], [10, 'mid'], [20, 'high']], T: [[null]] });
    eq(s.f('=XLOOKUP(15,B!$A:$A,B!$B:$B,,1)', 'Z1', 'T'), 'high');
  });

  t('XLOOKUP search_mode -1 searches from the end', function () {
    var s = scratch({ B: [['x', 1], ['x', 2]], T: [[null]] });
    eq(s.f('=XLOOKUP("x",B!$A:$A,B!$B:$B,,0,-1)', 'Z1', 'T'), 2);
  });

  t('XLOOKUP takes a key array built with & -- pattern P9, no helper column', function () {
    var s = scratch({
      M: [['K100', 'AM', 'first'], ['K100', 'PM', 'second'], ['K200', 'AM', 'third']],
      T: [[null]]
    });
    eq(s.f('=XLOOKUP("K100"&"PM",M!$A:$A&M!$B:$B,M!$C:$C)', 'Z1', 'T'), 'second');
  }, { ref: 'P9' });

  t('VLOOKUP needs a physical helper column for the same job', function () {
    var s = scratch({
      M: [['K100', 'AM', '=A1&B1', 'first'],
        ['K100', 'PM', '=A2&B2', 'second'],
        ['K200', 'AM', '=A3&B3', 'third']],
      T: [[null]]
    });
    eq(s.f('=VLOOKUP("K100"&"PM",M!$C:$D,2,FALSE)', 'Z1', 'T'), 'second');
  }, { ref: 'P9' });

  t('MATCH exact returns a 1-based position', function () {
    eq(lookSheet().f('=MATCH("K300",PASTE!$A$1:$A$3,0)', 'Z1', 'T'), 3);
  });

  t('MATCH type 1 needs an ascending list', function () {
    var s = scratch({ B: [[0], [10], [20]], T: [[null]] });
    eq(s.f('=MATCH(15,B!$A$1:$A$3,1)', 'Z1', 'T'), 2);
  });

  t('MATCH type -1 needs a descending list', function () {
    var s = scratch({ B: [[20], [10], [0]], T: [[null]] });
    eq(s.f('=MATCH(15,B!$A$1:$A$3,-1)', 'Z1', 'T'), 1);
  });

  t('INDEX walks a one-dimensional block with one index', function () {
    eq(lookSheet().f('=INDEX(PASTE!$B$1:$B$3,2)', 'Z1', 'T'), 'Thigh');
  });

  t('INDEX with row and column', function () {
    eq(lookSheet().f('=INDEX(PASTE!$A$1:$D$3,3,4)', 'Z1', 'T'), 'M&S');
  });

  t('INDEX out of range is #REF!', function () {
    eq(lookSheet().f('=INDEX(PASTE!$A$1:$A$3,9)', 'Z1', 'T'), '#REF!');
  });

  t('INDEX + MATCH is a lookup that can look leftwards', function () {
    eq(lookSheet().f('=INDEX(PASTE!$A$1:$A$3,MATCH("Wing",PASTE!$B$1:$B$3,0))', 'Z1', 'T'), 'K300');
  });

  t('OFFSET shifts a reference', function () {
    var s = scratch({ T: [[1, 2], [3, 4]] });
    eq(s.f('=OFFSET(A1,1,1)', 'Z1', 'T'), 4);
    eq(s.f('=SUM(OFFSET(A1,0,0,2,2))', 'Z2', 'T'), 10);
  });

  t('INDIRECT turns text into a reference', function () {
    var s = scratch({ T: [[null, null, 99]] });
    eq(s.f('=INDIRECT("C1")', 'Z1', 'T'), 99);
  });

  /* =========================================================================
     Array / Ctrl+Shift+Enter behaviour (needed for P15 and P16)
     ========================================================================= */

  G('array formulas');

  t('boolean x boolean multiplies out as AND inside an array formula', function () {
    var s = scratch({ T: [[1, 5], [2, 6], [3, 7]] });
    eq(s.cse('=SUM((A1:A3>=2)*(B1:B3<=6))', 'Z1', 'T'), 1);
  });

  t('ROW(range) generates a list of row numbers in array context', function () {
    var s = scratch({ T: [[1], [2], [3]] });
    eq(s.cse('=SUM(ROW(A1:A3))', 'Z1', 'T'), 6, '1+2+3');
  });

  t('SMALL(IF(condition,ROW(range)),k) picks the kth matching row -- P15 core', function () {
    var s = scratch({ T: [['x'], ['y'], ['x'], ['x']] });
    eq(s.cse('=SMALL(IF(A1:A4="x",ROW(A1:A4)),1)', 'Z1', 'T'), 1);
    eq(s.cse('=SMALL(IF(A1:A4="x",ROW(A1:A4)),2)', 'Z2', 'T'), 3);
    eq(s.cse('=SMALL(IF(A1:A4="x",ROW(A1:A4)),3)', 'Z3', 'T'), 4);
  }, { ref: 'P15' });

  t('the FALSE that IF returns for non-matching rows is simply skipped by SMALL', function () {
    var s = scratch({ T: [['x'], ['y']] });
    eq(s.cse('=SMALL(IF(A1:A2="x",ROW(A1:A2)),2)', 'Z1', 'T'), '#NUM!',
      'only one match, so there is no 2nd');
  });

  t('MATCH(1,(a>=x)*(b<=y),0) finds a straddling row -- P16 core', function () {
    var s = scratch({
      Calendars: [['WK45', NOV11 - 2, NOV11 + 4], ['WK46', NOV11 + 5, NOV11 + 11]],
      T: [[NOV11]]
    });
    eq(s.cse('=MATCH(1,(A1>=Calendars!$B$1:$B$2)*(A1<=Calendars!$C$1:$C$2),0)', 'Z1', 'T'), 1);
    eq(s.cse('=INDEX(Calendars!$A$1:$A$2,MATCH(1,(A1>=Calendars!$B$1:$B$2)*' +
      '(A1<=Calendars!$C$1:$C$2),0))', 'Z2', 'T'), 'WK45');
  }, { ref: 'P16' });

  t('an ordinary formula does NOT get array behaviour -- it intersects instead', function () {
    var s = scratch({ T: [[1, 5], [2, 6], [3, 7]] });
    // Entered normally in row 2, (A1:A3>=2) collapses to A2>=2.
    var a = XT.R.rc('D2');
    s.wb.setInput('T', a.r, a.c, '=(A1:A3>=2)*1');
    s.wb.recalc();
    eq(s.val('D2', 'T'), 1, 'A2>=2 is TRUE, times 1');
  });

  /* =========================================================================
     Number formats (§7.8)
     ========================================================================= */

  G('§7.8 number formats');

  var fmtCases = [
    [NOV15, 'dd/mm/yyyy', '15/11/2026'],
    [NOV15, 'm/d/yyyy', '11/15/2026'],
    [NOV15, 'yyyy', '2026'],
    [NOV15, 'mmmm', 'November'],
    [NOV15, 'ddd', 'Sun'],
    [NOV15, 'dddd', 'Sunday'],
    [0.5, '[$-F400]h:mm:ss AM/PM', '12:00:00 PM'],
    [0.25, 'hh:mm:ss', '06:00:00'],
    [0.125, '0.000', '0.125'],
    [0.1234, '0.0%', '12.3%'],
    [1234.5, '#,##0.00', '1,234.50'],
    [1234.5, '£0.00', '£1234.50'],
    [-5, '0;[Red]0', '5', 'the negative section drops the sign'],
    [0.000012345, '0.0000E+00', '1.2345E-05'],
    [42, 'General', '42'],
    [1 / 3, 'General', '0.333333333333333']
  ];
  fmtCases.forEach(function (c) {
    t('format ' + show(c[0]) + ' as "' + c[1] + '" -> "' + c[2] + '"' +
      (c[3] ? '  (' + c[3] + ')' : ''), function () {
      eq(F.format(c[0], c[1]), c[2]);
    });
  });

  t('a format never changes the value, only the display', function () {
    var s = scratch({ T: [[{ v: NOV15, fmt: 'dddd' }]] });
    eq(s.val('A1', 'T'), NOV15);
    eq(s.wb.display('T', 0, 0), 'Sunday');
    eq(s.f('=A1+1', 'Z1', 'T'), NOV15 + 1);
  });

  t('a text section passes text through: d/m/yy;@', function () {
    eq(F.format('hello', 'd/m/yy;@'), 'hello');
  });

  t('numbers right-align and text left-aligns -- the quiet tell', function () {
    eq(F.alignOf(5), 'right');
    eq(F.alignOf('5'), 'left');
  });

  /* =========================================================================
     The twenty patterns, each worked on a tiny synthetic sheet
     ========================================================================= */

  G('patterns P1-P20');

  t('P1 bare pass-through', function () {
    var s = scratch({ SamplesinWB: [['scan-1']], scanMODWB: [[null]] });
    eq(s.f('=SamplesinWB!A1', 'A1', 'scanMODWB'), 'scan-1');
  });

  t('P2 zero-suppressing mirror hides an export padding zero', function () {
    var s = scratch({ 'Despatch Paste': [[0], [7]], 'Despatch Hidden': [[null]] });
    eq(s.f('=IF(\'Despatch Paste\'!A1>0,\'Despatch Paste\'!A1,"")', 'A1', 'Despatch Hidden'), '');
    eq(s.f('=IF(\'Despatch Paste\'!A2>0,\'Despatch Paste\'!A2,"")', 'A2', 'Despatch Hidden'), 7);
  });

  t('P2 caveat: the same guard lets TEXT through, because text > 0', function () {
    var s = scratch({ 'Despatch Paste': [[null]], 'Despatch Hidden': [[null]] });
    s.set('A1', "'0", 'Despatch Paste');       // a zero stored as text
    eq(s.f('=IF(\'Despatch Paste\'!A1>0,\'Despatch Paste\'!A1,"")', 'A1', 'Despatch Hidden'), '0',
      'the text "0" passes a >0 test');
  }, { ref: 'D5' });

  t('P3 blank-guarded mirror', function () {
    var s = scratch({ LVALL: [[null, 'line-1'], [null, 'line-2']], Sheet4: [[null]] });
    s.set('E1', 'anchor', 'Sheet4');
    eq(s.f('=IF(E1="","",LVALL!B1)', 'F1', 'Sheet4'), 'line-1');
    eq(s.f('=IF(E2="","",LVALL!B2)', 'F2', 'Sheet4'), '');
  });

  t('P4 belt-and-braces mirror behaves the same as P3 -- the IFERROR is dead weight', function () {
    var s = scratch({ 'Complaints Paste': [['x'], [null]], R: [[null]] });
    eq(s.f('=IFERROR(IF(\'Complaints Paste\'!A1="","",\'Complaints Paste\'!A1),"")', 'A1', 'R'), 'x');
    eq(s.f('=IFERROR(IF(\'Complaints Paste\'!A2="","",\'Complaints Paste\'!A2),"")', 'A2', 'R'), '');
  }, { ref: 'D12' });

  t('P5 the guarded join', function () {
    var s = lookSheet();
    eq(s.f('=IFERROR(VLOOKUP("K100",PASTE!$A:$DN,2,FALSE),"")', 'Z1', 'T'), 'Fillet');
    eq(s.f('=IFERROR(VLOOKUP("K999",PASTE!$A:$DN,2,FALSE),"")', 'Z2', 'T'), '');
  });

  t('P6 log-line slicing by character count', function () {
    var s = scratch({
      LVALL: [[null, null, 'JOB 7781 PACKLINE 2 SLEEVE CHECK FLASHLABEL; Invalid: Mismatch'],
        [null, null, 'JOB 7782 PACKLINE 2 SLEEVE CHECK OK']],
      Sheet4: [[null]]
    });
    eq(s.f('=IF(RIGHT(LVALL!C1,29)="FLASHLABEL; Invalid: Mismatch","Missing Sleeve",' +
      'IF(RIGHT(LVALL!C1,8)="Mismatch","Mismatch",""))', 'C1', 'Sheet4'), 'Missing Sleeve');
    eq(s.f('=IF(LEFT(LVALL!C2,3)="JOB",LEFT(LVALL!C2,8),"")', 'D2', 'Sheet4'), 'JOB 7782');
  });

  t('P7 field extraction by delimiter, including a negative instance', function () {
    var s = scratch({ T: [['JOB 7781 PACKLINE 2 SLEEVE']] });
    eq(s.f('=_xlfn.TEXTAFTER(A1," ",1)', 'Z1', 'T'), '7781 PACKLINE 2 SLEEVE');
    eq(s.f('=_xlfn.TEXTBEFORE(A1," ",2)', 'Z2', 'T'), 'JOB 7781');
    eq(s.f('=_xlfn.TEXTAFTER(A1," ",-1)', 'Z3', 'T'), 'SLEEVE');
    eq(s.f('=IFERROR(VALUE(_xlfn.TEXTBEFORE(_xlfn.TEXTAFTER(A1," ",1)," ",1)),"")', 'Z4', 'T'), 7781);
    eq(s.f('=IFERROR(_xlfn.TEXTAFTER(A1," ",9),"")', 'Z5', 'T'), '', 'past the end, guarded');
  });

  t('P8 week-ending date lookup', function () {
    var sat = NOV11 + (7 - D.weekday(NOV11, 1));
    var s = scratch({
      Calendars: [[null, null, null, null, null, null, null, sat, 'P9', 'WK46']],
      T: [[NOV11]]
    });
    eq(s.f('=IF(A1="","-",VLOOKUP((A1-WEEKDAY(A1,1))+7,Calendars!$H:$J,3,FALSE))', 'Z1', 'T'),
      'WK46');
  });

  t('P9 two keys at once, both dialects', function () {
    var s = scratch({
      M: [['K1', 'AM', '=A1&B1', 'early'], ['K1', 'PM', '=A2&B2', 'late']],
      T: [[null]]
    });
    eq(s.f('=XLOOKUP("K1"&"PM",M!$A:$A&M!$B:$B,M!$D:$D)', 'Z1', 'T'), 'late');
    eq(s.f('=VLOOKUP("K1"&"PM",M!$C:$D,2,FALSE)', 'Z2', 'T'), 'late');
  });

  t('P10 XLOOKUP column fan-out: one key, many columns, filled right', function () {
    var s = scratch({
      'MIMS SAVE FILE': [['Fillet', 'RAW', 180, 'K100'], ['Thigh', 'RAW', 220, 'K200']],
      FILTERWBTODAY: [[null, 'K200']],
      OUT: [[null]]
    });
    // The real formula keys on $B2 of the filter sheet and pulls column A, B, C in turn.
    var base = '=IFERROR(_xlfn.XLOOKUP(FILTERWBTODAY!$B1,\'MIMS SAVE FILE\'!$D:$D,' +
      '\'MIMS SAVE FILE\'!A:A),"")';
    eq(s.f(base, 'A1', 'OUT'), 'Thigh');
    // Filling right must move the RETURN column but not the key column.
    var right = P.translateText(base, 0, 1);
    eq(right, '=IFERROR(_xlfn.XLOOKUP(FILTERWBTODAY!$B1,\'MIMS SAVE FILE\'!$D:$D,' +
      '\'MIMS SAVE FILE\'!B:B),"")');
    eq(s.f(right, 'B1', 'OUT'), 'RAW');
  });

  t('P11 fallback keys, ending in a message to a human', function () {
    var s = scratch({
      'Volumes Paste': [['AA1', 'BB1', 'CC1', null, null, null, 'vol-1']],
      R: [[null]]
    });
    // Key 1 missing, key 2 present: the second branch should fire.
    s.set('K1', 'nope', 'R').set('M1', 'BB1', 'R').set('O1', 'CC1', 'R');
    s.set('L1', '=IF(K1="","",IF(COUNTIF(\'Volumes Paste\'!$A:$A,K1)>=1,1,""))', 'R');
    s.set('N1', '=IF(M1="","",IF(COUNTIF(\'Volumes Paste\'!$B:$B,M1)>=1,1,""))', 'R');
    s.set('P1', '=IF(O1="","",IF(COUNTIF(\'Volumes Paste\'!$C:$C,O1)>=1,1,""))', 'R');
    eq(s.val('L1', 'R'), '');
    eq(s.val('N1', 'R'), 1);
    var q = '=IFERROR(IF(L1=1,VLOOKUP(K1,\'Volumes Paste\'!$A:$G,7,FALSE),' +
      'IF(N1=1,VLOOKUP(M1,\'Volumes Paste\'!$B:$G,6,FALSE),' +
      'IF(P1=1,VLOOKUP(O1,\'Volumes Paste\'!$C:$G,5,FALSE),' +
      '"Check"&" "&"-"&" "&O1))),"Check"&" "&"-"&" "&O1)';
    eq(s.f(q, 'Q1', 'R'), 'vol-1');
    // Now break all three keys and check it speaks to the operator.
    s.set('M1', 'nope2', 'R').set('O1', 'nope3', 'R');
    s.wb.recalc();
    eq(s.val('Q1', 'R'), 'Check - nope3');
  });

  t('P12 date-window SUMIFS on a transposed sheet', function () {
    var s = scratch({
      'Volumes Paste': [
        [null, NOV11, NOV15, NOV15 + 10],      // row 1: the date header row
        [null, 100, 200, 400]                   // row 2: the volumes
      ],
      Extract: [[null]]
    });
    s.set('K1', NOV11, 'Extract').set('K2', NOV15, 'Extract');
    eq(s.f('=SUMIFS(\'Volumes Paste\'!2:2,\'Volumes Paste\'!$1:$1,">=" &K$1,' +
      '\'Volumes Paste\'!$1:$1,"<="&K$2)', 'L2', 'Extract'), 300);
  });

  t('P13 multi-criteria COUNTIFS, and it must survive being filled right', function () {
    var s = scratch({
      'Micro paste': [
        ['SCU', NOV11, null, null, null, null, null, 'TVC', null, null, null, null, 'C100', null, null, null, null, null, null, null, null, null, null, null, null, null, 'Yes'],
        ['SCU', NOV11, null, null, null, null, null, 'TVC', null, null, null, null, 'C100', null, null, null, null, null, null, null, null, null, null, null, null, null, 'No'],
        ['SCU', NOV11, null, null, null, null, null, 'ENT', null, null, null, null, 'C100', null, null, null, null, null, null, null, null, null, null, null, null, null, 'Yes']
      ],
      Rep: [[null]]
    });
    // D86 = test code column M value, C86 = column H value, Q23 = column B value
    s.set('D86', 'C100', 'Rep').set('C86', 'TVC', 'Rep').set('Q23', NOV11, 'Rep');
    var f = '=COUNTIFS(\'Micro paste\'!$M:$M,$D$86,\'Micro paste\'!$H:$H,$C$86,' +
      '\'Micro paste\'!$B:$B,Rep!$Q$23,\'Micro paste\'!$AA:$AA,"Yes")';
    eq(s.f(f, 'F88', 'Rep'), 1);
  });

  t('P14 manufacture a unique key with an expanding COUNTIF', function () {
    var s = scratch({ H: [['TVC'], ['TVC'], ['ENT'], ['TVC']] });
    s.set('B1', '=A1&COUNTIF($A$1:A1,A1)', 'H');
    s.set('B2', '=A2&COUNTIF($A$1:A2,A2)', 'H');
    s.set('B3', '=A3&COUNTIF($A$1:A3,A3)', 'H');
    s.set('B4', '=A4&COUNTIF($A$1:A4,A4)', 'H');
    eq(s.val('B1', 'H'), 'TVC1');
    eq(s.val('B2', 'H'), 'TVC2');
    eq(s.val('B3', 'H'), 'ENT1');
    eq(s.val('B4', 'H'), 'TVC3');
  });

  t('P14 and the fill translation that makes it work', function () {
    eq(P.translateText('=A1&COUNTIF($A$1:A1,A1)', 1, 0), '=A2&COUNTIF($A$1:A2,A2)');
  });

  t('P15 get the nth match, as a real array formula', function () {
    var s = scratch({
      F: [['K1', 'AM', 'first'], ['K2', 'AM', 'other'], ['K1', 'AM', 'second'], ['K1', 'PM', 'third']],
      OUT: [[null]]
    });
    var f = '=IFERROR(INDEX(F!$C$1:$C$4,SMALL(IF((F!$A$1:$A$4="K1")*(F!$B$1:$B$4="AM"),' +
      'ROW(F!$A$1:$A$4)),%)),"")';
    eq(s.cse(f.replace('%', '1'), 'Z1', 'OUT'), 'first');
    eq(s.cse(f.replace('%', '2'), 'Z2', 'OUT'), 'second');
    eq(s.cse(f.replace('%', '3'), 'Z3', 'OUT'), '', 'only two matches, guarded to ""');
  });

  t('P16 boundary-match lookup, and the modern one-liner beside it', function () {
    var s = scratch({
      Calendars: [[null, null, null, null, null, 'WK45', NOV11 - 3, NOV11 + 3],
        [null, null, null, null, null, 'WK46', NOV11 + 4, NOV11 + 10]],
      E: [[null, null, NOV11]]
    });
    eq(s.cse('=INDEX(Calendars!F:F,MATCH(1,(C1>=Calendars!G:G)*(C1<=Calendars!H:H),0))',
      'Z1', 'E'), 'WK45');
    eq(s.f('=XLOOKUP(C1,Calendars!$G$1:$G$2,Calendars!$F$1:$F$2,,-1)', 'Z2', 'E'), 'WK45',
      'the approximate-match alternative');
  });

  t('P17 make the sheet check itself', function () {
    var s = scratch({ R: [['A', 'B', 'C']] });
    s.set('Q1', '=A1&"_"&B1&"_"&C1', 'R');
    s.set('D1', 'A_B_C', 'R');
    s.set('R1', '=IF(D1=Q1,1,0)', 'R');
    s.set('S1', '=IF(R1=1,"DATA MATCH","PLEASE REVIEW AND AMEND")', 'R');
    eq(s.val('S1', 'R'), 'DATA MATCH');
    s.set('D1', 'A_B_X', 'R');
    s.wb.recalc();
    eq(s.val('S1', 'R'), 'PLEASE REVIEW AND AMEND');
  });

  t('P18 nested IF ladder, and its empty-final-argument trap', function () {
    var s = scratch({ T: [['Asda', 'u', 'v']] });
    eq(s.f('=IF(A1="Aldi",B1,IF(A1="Tesco",C1,IF(A1="Asda",C1,)))', 'Z1', 'T'), 'v');
    eq(s.f('=IF(A1="Aldi",B1,IF(A1="Tesco",C1,IF(A1="Nobody",C1,)))', 'Z2', 'T'), 0,
      'no branch matched, and the empty final argument gives 0, which PRINTS');
  }, { ref: 'D4' });

  t('P18 refactored into a lookup table gives the same answer, readably', function () {
    var s = scratch({
      REF: [['Aldi', 'u'], ['Tesco', 'v'], ['Asda', 'v']],
      T: [['Asda']]
    });
    eq(s.f('=IFNA(XLOOKUP(A1,REF!$A:$A,REF!$B:$B),"")', 'Z1', 'T'), 'v');
  });

  t('P19 spine generators', function () {
    var s = scratch({ T: [[NOV11]] });
    s.set('B1', '=+A1+1', 'T');
    eq(s.val('B1', 'T'), NOV11 + 1);
    var a = XT.R.rc('A7');
    s.wb.setInput('T', a.r, a.c, '=ROW()');
    s.wb.recalc();
    eq(s.val('A7', 'T'), 7);
  });

  t('P19 self-limiting counter stops at its ceiling', function () {
    var s = scratch({ T: [[1]] });
    s.set('L1', 3, 'T');
    s.set('A2', '=IF(A1="","",IF((A1+1)<=$L$1,(A1+1),""))', 'T');
    s.set('A3', '=IF(A2="","",IF((A2+1)<=$L$1,(A2+1),""))', 'T');
    s.set('A4', '=IF(A3="","",IF((A3+1)<=$L$1,(A3+1),""))', 'T');
    eq(s.val('A2', 'T'), 2);
    eq(s.val('A3', 'T'), 3);
    eq(s.val('A4', 'T'), '', 'it stops rather than running past the ceiling');
  });

  t('P20 rate finishers, including the flag-then-average idiom', function () {
    var s = scratch({ T: [[600, 'WK46', 'SCU'], [400, 'WK46', 'SCU'], [900, 'WK46', 'SCU']] });
    s.set('D1', '=IF(A1>500,1,0)', 'T');
    s.set('D2', '=IF(A2>500,1,0)', 'T');
    s.set('D3', '=IF(A3>500,1,0)', 'T');
    eq(s.f('=SUMIFS(D1:D3,B1:B3,"WK46",C1:C3,"SCU")/COUNTIFS(B1:B3,"WK46",C1:C3,"SCU")',
      'Z1', 'T'), 2 / 3);
    eq(s.f('=AVERAGEIFS(A1:A3,B1:B3,"WK46")', 'Z2', 'T'), 1900 / 3);
    eq(s.f('=IFERROR((A1/0)*100,"")', 'Z3', 'T'), '', 'divide-by-zero guarded');
  });

  /* =========================================================================
     Stacking -- the thing the trainer is really for
     ========================================================================= */

  G('stacking (composition)');

  t('a stack built one wrapper at a time gives the same answer at each step', function () {
    var s = scratch({
      PASTE: [['K100', 'Fillet'], ['K200', 'Thigh']],
      WORK: [['K200'], [null]]
    });
    // 1. the bare join
    eq(s.f('=VLOOKUP(A1,PASTE!$A:$B,2,FALSE)', 'B1', 'WORK'), 'Thigh');
    // 2. guard the "not found" case
    eq(s.f('=IFERROR(VLOOKUP(A1,PASTE!$A:$B,2,FALSE),"")', 'C1', 'WORK'), 'Thigh');
    // 3. do not even look up a blank key
    eq(s.f('=IF(A1="","",IFERROR(VLOOKUP(A1,PASTE!$A:$B,2,FALSE),""))', 'D1', 'WORK'), 'Thigh');
    // and the blank row below behaves at every level
    eq(s.f('=VLOOKUP(A2,PASTE!$A:$B,2,FALSE)', 'B2', 'WORK'), '#N/A');
    eq(s.f('=IFERROR(VLOOKUP(A2,PASTE!$A:$B,2,FALSE),"")', 'C2', 'WORK'), '');
    eq(s.f('=IF(A2="","",IFERROR(VLOOKUP(A2,PASTE!$A:$B,2,FALSE),""))', 'D2', 'WORK'), '');
  });

  t('shape.layers peels a stack apart', function () {
    var r = S.layers('=IFERROR(IF(A1="","",VLOOKUP(A1,PASTE!$A:$B,2,FALSE)),"")');
    truthy(r.ok, 'it parsed');
    eq(r.layers[0].fn, 'IFERROR');
    eq(r.layers[1].fn, 'IF');
    eq(r.layers[2].fn, 'VLOOKUP');
    eq(r.depth, 3);
  });

  t('nestDepth counts how deep the IFs go', function () {
    eq(S.nestDepth('=IF(A1=1,1,IF(A1=2,2,IF(A1=3,3,"")))', 'IF'), 3);
  });

  t('functionsUsed lists every function in the stack', function () {
    var u = S.functionsUsed('=IFERROR(VALUE(TEXTBEFORE(A1," ",1)),"")')
      .map(function (x) { return x.fn; }).sort();
    eq(u.join(','), 'IFERROR,TEXTBEFORE,VALUE');
  });

  /* =========================================================================
     Shape corpus: the real formulas from the scan must all parse
     ========================================================================= */

  G('shape corpus (400 real formulas)');

  t('every untruncated formula in the top-400 corpus tokenizes and parses', function () {
    var corpus = (XT.data && XT.data.shapeCorpus) || [];
    truthy(corpus.length > 0, 'the corpus loaded');
    var bad = [], n = 0;
    corpus.forEach(function (row) {
      if (row.truncated) return;
      n++;
      var r = P.parse('=' + row.formula);
      if (!r.ok) bad.push(row.formula.slice(0, 90) + '   <<' + r.error + '>>');
    });
    truthy(n >= 390, 'tested ' + n + ' of ' + corpus.length + ' rows');
    if (bad.length) {
      fail(bad.length + ' of ' + n + ' failed to parse:\n  ' + bad.slice(0, 12).join('\n  '));
    }
  });

  t('the two truncated corpus rows are the scanner’s fault, not the parser’s', function () {
    var corpus = (XT.data && XT.data.shapeCorpus) || [];
    var trunc = corpus.filter(function (r) { return r.truncated; });
    eq(trunc.length, XT.data.shapeCorpusTruncated || 0);
    trunc.forEach(function (row) {
      var opens = (row.formula.match(/\(/g) || []).length;
      var closes = (row.formula.match(/\)/g) || []).length;
      truthy(opens !== closes || row.formula.length >= 300,
        'the formula really is cut off: ' + row.formula.slice(0, 40));
    });
  });

  t('the normaliser reproduces the catalog shape for every untruncated row', function () {
    var corpus = (XT.data && XT.data.shapeCorpus) || [];
    var bad = [], n = 0;
    corpus.forEach(function (row) {
      if (row.truncated) return;
      n++;
      var got = S.normalise(row.formula);
      if (got !== row.shape) {
        bad.push('formula: ' + row.formula.slice(0, 70) +
          '\n      catalog: ' + row.shape + '\n      engine:  ' + got);
      }
    });
    if (bad.length) {
      fail(bad.length + ' of ' + n + ' shapes differ:\n  ' + bad.slice(0, 8).join('\n  '));
    }
  });

  t('every corpus formula whose functions are all implemented also evaluates', function () {
    var corpus = (XT.data && XT.data.shapeCorpus) || [];
    var s = scratch({ Sheet1: [[1]] });
    var crashed = [];
    corpus.forEach(function (row) {
      var fns = S.functionsUsed(row.formula);
      for (var i = 0; i < fns.length; i++) {
        var f = XT.FN[fns[i].fn];
        if (!f || f.unsupported) return;        // out of scope: skip
      }
      try {
        var v = s.f('=' + row.formula, 'AZ1', 'Sheet1');
        if (v === undefined) crashed.push(row.formula.slice(0, 80) + ' -> undefined');
      } catch (e) {
        crashed.push(row.formula.slice(0, 80) + ' -> threw ' + e.message);
      }
    });
    if (crashed.length) {
      fail(crashed.length + ' formulas crashed the evaluator:\n  ' + crashed.slice(0, 10).join('\n  '));
    }
  });

  t('at least 20 distinct functions appear across the corpus', function () {
    var corpus = (XT.data && XT.data.shapeCorpus) || [];
    var seen = {};
    corpus.forEach(function (row) {
      S.functionsUsed(row.formula).forEach(function (f) { seen[f.fn] = 1; });
    });
    truthy(Object.keys(seen).length >= 20, 'saw ' + Object.keys(seen).length + ' functions');
  });

  /* =========================================================================
     Pattern recognition
     ========================================================================= */

  G('pattern recognition');

  (XT.data.patterns || []).forEach(function (p) {
    t(p.id + ' recognises its own catalog example', function () {
      var m = S.match('=' + p.example.replace(/^=/, '').split('   then   ')[0]);
      var ids = m.hits.map(function (h) { return h.id; });
      truthy(ids.indexOf(p.id) >= 0,
        p.id + ' not matched. shape was: ' + m.clean + '  (hits: ' + ids.join(',') + ')');
    });
  });

  /* =========================================================================
     Registry sanity
     ========================================================================= */

  G('registry');

  t('every function the catalog §12.3 list names is registered', function () {
    var want = ('IF IFERROR IFNA VLOOKUP XLOOKUP INDEX MATCH COUNTIF COUNTIFS SUMIF SUMIFS ' +
      'AVERAGEIFS SUM COUNT COUNTA COUNTBLANK MIN MAX AND OR NOT TRUE FALSE LEFT RIGHT MID ' +
      'LEN TRIM VALUE TEXT CONCAT SUBSTITUTE CHAR FIND SEARCH TEXTBEFORE TEXTAFTER TODAY NOW ' +
      'DATE WEEKDAY MONTH YEAR DAY HOUR MINUTE SECOND TIME ROW ROWS COLUMN SMALL LARGE RANK ' +
      'SUMPRODUCT ISBLANK ISERROR ISNUMBER ISTEXT NA ROUND INT ABS').split(/\s+/);
    var missing = want.filter(function (n) { return !XT.FN[n] || XT.FN[n].unsupported; });
    if (missing.length) fail('not registered: ' + missing.join(', '));
  });

  t('every function registered in the estate’s top-50 list is reachable', function () {
    var estate = ('IF IFERROR VLOOKUP RIGHT LEFT WEEKDAY XLOOKUP COUNTIF SUMIF COUNTIFS ' +
      'SUMIFS SUM ROW MATCH INDEX TEXTAFTER VALUE TEXTBEFORE MIN AND COUNT TODAY MAX TEXT ' +
      'HOUR MINUTE TIME SECOND SUMPRODUCT AVERAGEIFS CHAR SUBSTITUTE LEN RANK INDIRECT SMALL ' +
      'LOOKUP NOT ISTEXT ISERROR COUNTA ISBLANK FIND OR COUNTBLANK AVERAGE IFNA MONTH INT ' +
      'TRANSPOSE UPPER CONCATENATE WEEKNUM OFFSET NOW').split(/\s+/);
    var missing = estate.filter(function (n) { return !XT.FN[n]; });
    if (missing.length) fail('not registered at all: ' + missing.join(', '));
  });

  t('the §8 out-of-scope functions are declared, not silently missing', function () {
    ['FILTER', 'SORT', 'SEQUENCE', 'LET', 'LAMBDA', 'UNIQUE', 'IFS', 'SWITCH', 'CHOOSE']
      .forEach(function (n) {
        truthy(XT.FN[n], n + ' should be registered');
        truthy(XT.FN[n].unsupported, n + ' should be marked unsupported');
      });
  });

  t('an out-of-scope function explains itself instead of giving a bare #NAME?', function () {
    var v = scratch().f('=FILTER(A1:A3,A1:A3>1)');
    eq(v, '#NAME?');
    truthy(/not implemented/.test(v.detail || ''), 'it says why');
  });

  t('wrong argument counts are reported with the signature', function () {
    var v = scratch().f('=VLOOKUP(1)');
    eq(v, '#VALUE!');
    truthy(/at least 3/.test(v.detail || ''), 'it says how many it needs');
  });

  t('every registered function has a signature, a doc line and a group', function () {
    var bad = [];
    XT.functionList().forEach(function (f) {
      if (!f.sig || !f.doc || !f.group) bad.push(f.name);
    });
    if (bad.length) fail('incomplete registry entries: ' + bad.join(', '));
  });

})(window.XT = window.XT || {});
