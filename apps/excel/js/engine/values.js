/* =============================================================================
   Excel Trainer — engine/values.js
   The value model. Read this file first; everything else depends on it.

   Excel has FIVE scalar value types, and the difference between two of them is
   the single most important thing in this whole program (catalog §7.4):

      number    42, 0.5, 46234 (a date is just a number)
      text      "ABC", "", "10"      <- note that "" IS a text value
      boolean   TRUE / FALSE
      error     #N/A #REF! #VALUE! #DIV/0! #NAME? #NUM! #NULL!
      EMPTY     a cell nobody has ever typed into

   EMPTY is NOT "" and NOT 0. A cell holding =IF(A1="","",A1) that takes the
   "" branch contains TEXT of zero length. A cell nobody touched contains EMPTY.
   ISBLANK can tell them apart; the = operator cannot. 1.2 million formulas in
   the source estate end in ,"") so this distinction is load-bearing.

   Representation in JS:
      number   -> JS number
      text     -> JS string
      boolean  -> JS true/false
      error    -> XErr instance
      EMPTY    -> the EMPTY singleton
      array    -> XArr   (a block of values, from a range or an array formula)
      range    -> XRange (a *reference*; some functions need the reference
                          itself, not its values -- ISBLANK, ROW, COUNTBLANK)
   ============================================================================= */
(function (XT) {
  'use strict';

  var V = {};
  XT.V = V;

  /* ---------------------------------------------------------------- EMPTY --- */

  // A unique object. Compare with ===, never with == or truthiness.
  var EMPTY = { __empty: true };
  V.EMPTY = EMPTY;
  V.isEmpty = function (v) { return v === EMPTY; };

  /* ---------------------------------------------------------------- errors -- */

  function XErr(code, detail) {
    this.code = code;        // the string Excel shows, e.g. "#REF!"
    this.detail = detail || null;  // our own plain-language explanation
  }
  XErr.prototype.toString = function () { return this.code; };
  V.XErr = XErr;

  V.ERR = {
    NULL: '#NULL!', DIV0: '#DIV/0!', VALUE: '#VALUE!', REF: '#REF!',
    NAME: '#NAME?', NUM: '#NUM!', NA: '#N/A'
  };
  var ERRCODES = ['#NULL!', '#DIV/0!', '#VALUE!', '#REF!', '#NAME?', '#NUM!', '#N/A'];
  V.ERRCODES = ERRCODES;

  // err('#VALUE!', 'LEFT needs a number of characters') -> XErr
  function err(code, detail) { return new XErr(code, detail); }
  V.err = err;
  V.isErr = function (v) { return v instanceof XErr; };

  // Shorthands used all over the function library.
  V.na = function (d) { return err('#N/A', d); };
  V.value = function (d) { return err('#VALUE!', d); };
  V.ref = function (d) { return err('#REF!', d); };
  V.div0 = function (d) { return err('#DIV/0!', d); };
  V.num = function (d) { return err('#NUM!', d); };
  V.name = function (d) { return err('#NAME?', d); };

  /* ---------------------------------------------------------------- arrays -- */

  /* A rectangular block of scalar values. Produced by:
       - materialising a range  (A1:C10 -> 10x3 array)
       - an array constant      {1,2;3,4}
       - array arithmetic       (A1:A9=B1:B9) -> 9x1 array of booleans
     Arrays only "spill" in modern Excel. In this trainer they exist only
     INSIDE a formula, exactly as in legacy Ctrl+Shift+Enter Excel (§7.9). */
  function XArr(rows) {
    this.rows = rows;                              // array of arrays
    this.nr = rows.length;
    this.nc = rows.length ? rows[0].length : 0;
  }
  XArr.prototype.get = function (r, c) {
    var row = this.rows[r];
    if (!row) return EMPTY;
    var v = row[c];
    return v === undefined ? EMPTY : v;
  };
  // Excel broadcasts a 1-wide or 1-tall array across the other operand.
  XArr.prototype.at = function (r, c) {
    return this.get(this.nr === 1 ? 0 : r, this.nc === 1 ? 0 : c);
  };
  XArr.prototype.flat = function () {
    var out = [];
    for (var r = 0; r < this.nr; r++) for (var c = 0; c < this.nc; c++) out.push(this.get(r, c));
    return out;
  };
  XArr.of = function (v) { return new XArr([[v]]); };
  XArr.col = function (list) { return new XArr(list.map(function (v) { return [v]; })); };
  V.XArr = XArr;
  V.isArr = function (v) { return v instanceof XArr; };

  /* ---------------------------------------------------------------- ranges -- */

  /* A reference, kept un-evaluated. r1/c1/r2/c2 are 0-based inclusive.
     wholeCol / wholeRow record that the user wrote $A:$DN or 17:17, so the
     evaluator can clip to the sheet's used range instead of walking a million
     rows (§7.9, defect D13). */
  function XRange(sheet, r1, c1, r2, c2, flags) {
    this.sheet = sheet;                 // sheet NAME (string)
    this.r1 = Math.min(r1, r2); this.r2 = Math.max(r1, r2);
    this.c1 = Math.min(c1, c2); this.c2 = Math.max(c1, c2);
    flags = flags || {};
    this.wholeCol = !!flags.wholeCol;   // written as A:A
    this.wholeRow = !!flags.wholeRow;   // written as 1:1
    this.abs = flags.abs || null;       // which $ the user typed, for display
  }
  XRange.prototype.isSingle = function () { return this.r1 === this.r2 && this.c1 === this.c2; };
  XRange.prototype.height = function () { return this.r2 - this.r1 + 1; };
  XRange.prototype.width = function () { return this.c2 - this.c1 + 1; };
  V.XRange = XRange;
  V.isRange = function (v) { return v instanceof XRange; };

  /* ------------------------------------------------- number -> text (General) */

  /* Excel renders a number as text with at most 15 significant digits.
     This is why  =1/3&""  gives "0.333333333333333"  and  =(0.1+0.2)&""
     gives "0.3" and not "0.30000000000000004" (§7.5). */
  function numToText(n) {
    if (n === 0) return '0';
    if (!isFinite(n)) return '#NUM!';
    if (Number.isInteger(n) && Math.abs(n) < 1e15) return String(n);
    var a = Math.abs(n);
    if (a >= 1e15 || a < 1e-9) {
      // General format falls back to scientific notation at the extremes.
      var s = n.toExponential(8).replace(/e([+-])(\d)$/, 'E$10$2').replace('e', 'E');
      return s.replace(/\.?0+E/, 'E');
    }
    var t = n.toPrecision(15);
    if (t.indexOf('e') >= 0) t = Number(t).toString();
    if (t.indexOf('.') >= 0) t = t.replace(/0+$/, '').replace(/\.$/, '');
    return t;
  }
  V.numToText = numToText;

  /* ------------------------------------------------- text -> number coercion */

  /* Used when text meets an arithmetic operator. Deliberately strict about
     whitespace: it trims ONLY real spaces and tabs. A non-breaking space
     (CHAR(160)) is left in place, so "5"&CHAR(160) still fails -- which is
     exactly defect D7 from the catalog, and we want learners to meet it. */
  function parseNumeric(s) {
    if (typeof s !== 'string') return null;
    var t = s.replace(/^[ \t]+|[ \t]+$/g, '');
    if (t === '') return null;
    var pct = false;
    if (/%$/.test(t)) { pct = true; t = t.slice(0, -1).replace(/[ \t]+$/, ''); }
    var neg = false;
    if (/^\(.*\)$/.test(t)) { neg = true; t = t.slice(1, -1); }   // (1.23) accounting
    t = t.replace(/^([+-])?[£$€]/, '$1');                          // currency prefix
    if (/^[+-]?[\d,]*\d(\.\d+)?([eE][+-]?\d+)?$/.test(t)) {
      var n = Number(t.replace(/,/g, ''));
      if (isNaN(n)) return null;
      if (pct) n /= 100;
      return neg ? -n : n;
    }
    return null;
  }
  V.parseNumeric = parseNumeric;

  /* -------------------------------------------------------------- coercions */

  // Coerce to a number for arithmetic. Returns a number or an XErr.
  function toNum(v) {
    if (v instanceof XErr) return v;
    if (v === EMPTY) return 0;
    if (typeof v === 'number') return v;
    if (typeof v === 'boolean') return v ? 1 : 0;
    if (typeof v === 'string') {
      var n = parseNumeric(v);
      if (n === null) {
        return err('#VALUE!', 'the text ' + JSON.stringify(v.length > 24 ? v.slice(0, 24) + '…' : v) +
          ' cannot be read as a number');
      }
      return n;
    }
    return err('#VALUE!', 'that value cannot be used as a number');
  }
  V.toNum = toNum;

  // Coerce to text for & and the text functions.
  function toText(v) {
    if (v instanceof XErr) return v;
    if (v === EMPTY) return '';
    if (typeof v === 'string') return v;
    if (typeof v === 'number') return numToText(v);
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    return err('#VALUE!', 'that value cannot be used as text');
  }
  V.toText = toText;

  // Coerce to boolean for IF / AND / OR.
  function toBool(v) {
    if (v instanceof XErr) return v;
    if (v === EMPTY) return false;
    if (typeof v === 'boolean') return v;
    if (typeof v === 'number') return v !== 0;
    if (typeof v === 'string') {
      var u = v.toUpperCase();
      if (u === 'TRUE') return true;
      if (u === 'FALSE') return false;
      var n = parseNumeric(v);
      if (n !== null) return n !== 0;
      return err('#VALUE!', 'the text ' + JSON.stringify(v) + ' is not TRUE or FALSE');
    }
    return err('#VALUE!', 'that value cannot be used as a condition');
  }
  V.toBool = toBool;

  /* ------------------------------------------------------------- comparison */

  /* Excel's cross-type ordering, which surprises everyone:
        every number  <  every text  <  FALSE  <  TRUE
     So  ="" > 0  is TRUE, because "" is text and text beats any number.
     That is why the estate's 472,402  =IF(Src!A6>0,…)  blank-guards do not
     mean what they look like they mean (catalog defect D5). */
  function rank(v) {
    if (typeof v === 'number') return 0;
    if (typeof v === 'string') return 1;
    if (typeof v === 'boolean') return 2;
    return 0;
  }

  function compare(a, b, op) {
    if (a instanceof XErr) return a;
    if (b instanceof XErr) return b;

    // An EMPTY operand takes the "zero" of whatever it is compared with.
    if (a === EMPTY && b === EMPTY) { a = 0; b = 0; }
    else if (a === EMPTY) a = (typeof b === 'string') ? '' : (typeof b === 'boolean' ? false : 0);
    else if (b === EMPTY) b = (typeof a === 'string') ? '' : (typeof a === 'boolean' ? false : 0);

    var ra = rank(a), rb = rank(b), cmp;
    if (ra !== rb) {
      cmp = ra < rb ? -1 : 1;
    } else if (ra === 1) {
      // Text comparison is case-INSENSITIVE in Excel. "abc" = "ABC" is TRUE.
      var ua = a.toUpperCase(), ub = b.toUpperCase();
      cmp = ua === ub ? 0 : (ua < ub ? -1 : 1);
    } else if (ra === 2) {
      cmp = (a === b) ? 0 : (a === false ? -1 : 1);
    } else {
      cmp = (a === b) ? 0 : (a < b ? -1 : 1);
    }

    switch (op) {
      case '=': return cmp === 0;
      case '<>': return cmp !== 0;
      case '<': return cmp < 0;
      case '<=': return cmp <= 0;
      case '>': return cmp > 0;
      case '>=': return cmp >= 0;
    }
    return err('#VALUE!', 'unknown comparison ' + op);
  }
  V.compare = compare;

  /* ------------------------------------------------------- display helpers */

  // What the cell shows when it has no special number format.
  V.toDisplay = function (v) {
    if (v instanceof XErr) return v.code;
    if (v === EMPTY) return '';
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (typeof v === 'number') return numToText(v);
    return v;
  };

  // A short type label for the inspector panels.
  V.typeName = function (v) {
    if (v instanceof XErr) return 'error';
    if (v === EMPTY) return 'EMPTY';
    if (v instanceof XArr) return 'array ' + v.nr + '×' + v.nc;
    if (v instanceof XRange) return 'range';
    if (typeof v === 'number') return 'number';
    if (typeof v === 'boolean') return 'boolean';
    if (typeof v === 'string') return v === '' ? 'text ""' : 'text';
    return typeof v;
  };

})(window.XT = window.XT || {});
