/* =============================================================================
   Excel Trainer — engine/refs.js
   A1-style references: parsing, formatting, and -- the important one --
   TRANSLATION, which is what happens when you fill a formula down or right.

   Fill-down is the main interaction in this trainer (catalog §7.9): the source
   estate has 3.07 million cells that are fill-down children of 66,542 parent
   formulas, and almost no dynamic arrays. So the rule that a $ freezes a
   coordinate and a bare letter/number moves is not a detail here; it is the
   whole skill.
   ============================================================================= */
(function (XT) {
  'use strict';

  var R = {};
  XT.R = R;

  R.MAXROW = 1048576;
  R.MAXCOL = 16384;      // XFD

  /* ------------------------------------------------- column letters <-> index */

  // "A" -> 0, "Z" -> 25, "AA" -> 26, "DN" -> 117
  function colToNum(s) {
    var n = 0;
    s = s.toUpperCase();
    for (var i = 0; i < s.length; i++) n = n * 26 + (s.charCodeAt(i) - 64);
    return n - 1;
  }
  // 0 -> "A", 117 -> "DN"
  function numToCol(n) {
    var s = '';
    n = n + 1;
    while (n > 0) {
      var r = (n - 1) % 26;
      s = String.fromCharCode(65 + r) + s;
      n = Math.floor((n - 1) / 26);
    }
    return s;
  }
  R.colToNum = colToNum;
  R.numToCol = numToCol;

  /* --------------------------------------------------------------- addresses */

  /* An address is {r, c, ra, ca} where ra/ca say whether the ROW / COLUMN was
     written absolute (with a $). That flag is all that F4 cycles through, and
     all that fill translation consults. */
  var A1_RE = /^(\$?)([A-Za-z]{1,3})(\$?)(\d{1,7})$/;

  function parseAddr(s) {
    var m = A1_RE.exec(s);
    if (!m) return null;
    return {
      c: colToNum(m[2]), r: parseInt(m[4], 10) - 1,
      ca: m[1] === '$', ra: m[3] === '$'
    };
  }
  function formatAddr(a) {
    return (a.ca ? '$' : '') + numToCol(a.c) + (a.ra ? '$' : '') + (a.r + 1);
  }
  R.parseAddr = parseAddr;
  R.formatAddr = formatAddr;
  R.A1_RE = A1_RE;

  // "B7" -> {r:6,c:1} ignoring $; convenience for tests and fixtures.
  R.rc = function (s) { var a = parseAddr(s); return a ? { r: a.r, c: a.c } : null; };

  /* ------------------------------------------------------------- translation */

  /* Shift one address by (dr, dc), honouring its $ flags.
     Returns null if the result falls off the sheet, which is how Excel
     produces #REF! when you fill a formula up past row 1. */
  function shiftAddr(a, dr, dc) {
    var r = a.ra ? a.r : a.r + dr;
    var c = a.ca ? a.c : a.c + dc;
    if (r < 0 || c < 0 || r >= R.MAXROW || c >= R.MAXCOL) return null;
    return { r: r, c: c, ra: a.ra, ca: a.ca };
  }
  R.shiftAddr = shiftAddr;

  /* --------------------------------------------- whole columns and whole rows */

  var COLRANGE_RE = /^(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})$/;   //  $A:$DN
  var ROWRANGE_RE = /^(\$?)(\d{1,7}):(\$?)(\d{1,7})$/;               //  17:17
  R.COLRANGE_RE = COLRANGE_RE;
  R.ROWRANGE_RE = ROWRANGE_RE;

  /* -------------------------------------------------------- sheet name quoting */

  // Excel quotes a sheet name when it is not a bare identifier.
  // 'MIMS SAVE FILE'!$M:$M  -- spaces, &, dashes and leading digits all force quotes.
  function needsQuotes(name) {
    return !/^[A-Za-z_][A-Za-z0-9_.]*$/.test(name);
  }
  R.needsQuotes = needsQuotes;
  R.quoteSheet = function (name) {
    return needsQuotes(name) ? "'" + name.replace(/'/g, "''") + "'" : name;
  };
  R.unquoteSheet = function (tok) {
    if (tok.charAt(0) === "'") return tok.slice(1, -1).replace(/''/g, "'");
    return tok;
  };

  /* --------------------------------------------------------- range iteration */

  /* Walk a range, but CLIPPED to the sheet's used area.

     This exists because the house style in the source workbooks is the whole
     column -- VLOOKUP(x, PASTE!$A:$DN, 2, FALSE) appears 237,873 times. Taken
     literally that is 118 columns x 1,048,576 rows = 124 million cells per
     formula. Excel itself clips; so do we. Catalog defect D13 is exactly the
     cost of not clipping. */
  /* A WHOLE column or row is clipped hard to the used area: the user wrote
     "everything", and walking a million rows per formula is what makes a real
     workbook take forty seconds to open (defect D13).

     An EXPLICIT range is NOT clipped, because its size is part of its meaning:
     COUNTBLANK(A1:A5) must see five cells even if only three are filled, and
     COUNTIF(A1:A5,"") must find the blank ones. It is only capped at a size
     that keeps the trainer responsive. */
  R.CAP_ROWS = 20000;
  R.CAP_COLS = 500;

  R.clip = function (rng, used) {
    var r1 = rng.r1, r2 = rng.r2, c1 = rng.c1, c2 = rng.c2;
    if (rng.wholeCol) { r1 = 0; r2 = Math.max(0, used.maxRow); }
    if (rng.wholeRow) { c1 = 0; c2 = Math.max(0, used.maxCol); }
    if (r2 - r1 + 1 > R.CAP_ROWS) r2 = r1 + R.CAP_ROWS - 1;
    if (c2 - c1 + 1 > R.CAP_COLS) c2 = c1 + R.CAP_COLS - 1;
    return { r1: r1, r2: r2, c1: c1, c2: c2 };
  };

})(window.XT = window.XT || {});
