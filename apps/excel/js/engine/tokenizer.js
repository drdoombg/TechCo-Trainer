/* =============================================================================
   Excel Trainer — engine/tokenizer.js
   Turns formula text into a flat list of tokens. No structure yet; that is the
   parser's job. Every token carries its start/end offset in the original text
   so the Evaluate-Formula stepper can underline exactly the piece it is on.

   The awkward cases this has to get right, all taken from real formulas in the
   source estate:

     'MIMS SAVE FILE'!$M:$M      quoted sheet name, then a WHOLE COLUMN
     'Volumes Paste'!$13:$13     quoted sheet name, then a WHOLE ROW
     [1]CODES!$G$4:$I$4975       an external workbook reference
     'QTR 2 Old'!#REF!           a sheet-qualified #REF! sitting in live code
     _xlfn.XLOOKUP(...)          the on-disk spelling of a modern function
     Table4[[#This Row],[SUPPLIER]]   a structured reference, nested brackets
     =+K13+1                     a leading unary + (a Lotus habit; 19,683 cells)
     0/1/1900                     three numbers and two divisions, not a date
   ============================================================================= */
(function (XT) {
  'use strict';

  var T = {};
  XT.T = T;

  // Error literals have to be matched before the operators, because #N/A
  // contains a slash and #NAME? contains a question mark.
  var ERR_RE = /^#(?:REF!|VALUE!|DIV\/0!|NAME\?|NUM!|NULL!|N\/A|GETTING_DATA)/i;

  // A sheet (or external workbook) prefix, up to and including the "!".
  var SHEET_RE = /^(\[\d+\])?('(?:[^']|'')*'|[A-Za-z_][A-Za-z0-9_.]*)!/;

  // Whole column ($A:$DN) and whole row (17:17). Both must be tried before the
  // plain cell / number rules, and both need a lookahead so that A1:B2 and
  // 17:17.5 are not mis-tokenised.
  var COLR_RE = /^(\$?)([A-Za-z]{1,3}):(\$?)([A-Za-z]{1,3})(?![A-Za-z0-9_$])/;
  var ROWR_RE = /^(\$?)(\d{1,7}):(\$?)(\d{1,7})(?![\d.])/;

  var CELL_RE = /^\$?[A-Za-z]{1,3}\$?\d{1,7}(?![A-Za-z0-9_$])/;
  var NUM_RE = /^(?:\d+(?:\.\d*)?|\.\d+)(?:[eE][+-]?\d+)?/;
  var IDENT_RE = /^[A-Za-z_\\][A-Za-z0-9_.\\]*/;

  // Table1[...] with one level of nested brackets, which covers
  // [[#This Row],[Col]] and the simpler [Col].
  var TABLE_RE = /^([A-Za-z_][A-Za-z0-9_.]*)(\[(?:[^\[\]]|\[[^\[\]]*\])*\])/;

  var OPS3 = [];
  var OPS2 = ['<=', '>=', '<>'];
  var OPS1 = ['=', '<', '>', '+', '-', '*', '/', '^', '&', '%', ':'];

  /* A token is {t, v, p, q} where
       t = type, v = text or parsed value, p = start offset, q = end offset. */
  function tok(t, v, p, q) { return { t: t, v: v, p: p, q: q }; }

  /* Throws {name:'FormulaError', message, pos} on something it cannot read.
     The parser lets that bubble up to the cell, which then shows #NAME?. */
  function fail(msg, pos) {
    var e = new Error(msg);
    e.name = 'FormulaError';
    e.pos = pos;
    throw e;
  }
  T.fail = fail;

  T.tokenize = function (src) {
    var s = String(src);
    // A leading "=" is the cell's way of saying "this is a formula"; drop it.
    if (s.charAt(0) === '=') s = s.slice(1);

    var out = [], i = 0, n = s.length, m;

    while (i < n) {
      var ch = s.charAt(i);

      // ---- whitespace -------------------------------------------------------
      // Excel's space is technically the intersection operator. The estate
      // never uses it that way (it writes  ">=" & $B$1  with stray spaces), so
      // we treat space as insignificant.
      if (ch === ' ' || ch === '\t' || ch === '\n' || ch === '\r') { i++; continue; }

      var rest = s.slice(i);

      // ---- string literal ---------------------------------------------------
      if (ch === '"') {
        var j = i + 1, buf = '';
        for (;;) {
          if (j >= n) fail('unterminated text literal', i);
          if (s.charAt(j) === '"') {
            if (s.charAt(j + 1) === '"') { buf += '"'; j += 2; continue; }
            j++; break;
          }
          buf += s.charAt(j); j++;
        }
        out.push(tok('str', buf, i, j));
        i = j; continue;
      }

      // ---- error literal ----------------------------------------------------
      if (ch === '#') {
        m = ERR_RE.exec(rest);
        if (m) {
          var code = m[0].toUpperCase().replace('#N/A', '#N/A');
          out.push(tok('err', code, i, i + m[0].length));
          i += m[0].length; continue;
        }
        // A trailing # is the spill operator (A1#). Not supported; make it loud.
        fail('unrecognised # construct', i);
      }

      // ---- sheet / workbook prefix ------------------------------------------
      if ((m = SHEET_RE.exec(rest))) {
        out.push(tok('sheet', { ext: m[1] || null, name: XT.R.unquoteSheet(m[2]) },
          i, i + m[0].length));
        i += m[0].length; continue;
      }

      // ---- structured table reference ---------------------------------------
      if (/[A-Za-z_]/.test(ch) && (m = TABLE_RE.exec(rest))) {
        out.push(tok('table', { table: m[1], spec: m[2] }, i, i + m[0].length));
        i += m[0].length; continue;
      }

      // ---- whole column -----------------------------------------------------
      if ((m = COLR_RE.exec(rest))) {
        out.push(tok('colrange', { c1: XT.R.colToNum(m[2]), c2: XT.R.colToNum(m[4]),
          a1: m[1] === '$', a2: m[3] === '$' }, i, i + m[0].length));
        i += m[0].length; continue;
      }

      // ---- whole row --------------------------------------------------------
      if ((m = ROWR_RE.exec(rest))) {
        out.push(tok('rowrange', { r1: parseInt(m[2], 10) - 1, r2: parseInt(m[4], 10) - 1,
          a1: m[1] === '$', a2: m[3] === '$' }, i, i + m[0].length));
        i += m[0].length; continue;
      }

      // ---- cell address -----------------------------------------------------
      if ((ch === '$' || /[A-Za-z]/.test(ch)) && (m = CELL_RE.exec(rest))) {
        var addr = XT.R.parseAddr(m[0]);
        if (addr) { out.push(tok('cell', addr, i, i + m[0].length)); i += m[0].length; continue; }
      }

      // ---- number -----------------------------------------------------------
      if (/[0-9.]/.test(ch) && (m = NUM_RE.exec(rest))) {
        out.push(tok('num', parseFloat(m[0]), i, i + m[0].length));
        i += m[0].length; continue;
      }

      // ---- identifier: function name, TRUE/FALSE, or a defined name ---------
      if (/[A-Za-z_\\]/.test(ch) && (m = IDENT_RE.exec(rest))) {
        var raw = m[0];
        // Look past spaces for the "(" that makes it a function call.
        var k = i + raw.length;
        while (k < n && s.charAt(k) === ' ') k++;
        var isCall = s.charAt(k) === '(';
        out.push(tok(isCall ? 'func' : 'ident', raw, i, i + raw.length));
        i += raw.length; continue;
      }

      // ---- brackets and separators -----------------------------------------
      if (ch === '(') { out.push(tok('(', '(', i, i + 1)); i++; continue; }
      if (ch === ')') { out.push(tok(')', ')', i, i + 1)); i++; continue; }
      if (ch === '{') { out.push(tok('{', '{', i, i + 1)); i++; continue; }
      if (ch === '}') { out.push(tok('}', '}', i, i + 1)); i++; continue; }
      if (ch === ',') { out.push(tok(',', ',', i, i + 1)); i++; continue; }
      if (ch === ';') { out.push(tok(';', ';', i, i + 1)); i++; continue; }

      // ---- operators --------------------------------------------------------
      var two = s.substr(i, 2);
      if (OPS2.indexOf(two) >= 0) { out.push(tok('op', two, i, i + 2)); i += 2; continue; }
      if (OPS1.indexOf(ch) >= 0) { out.push(tok('op', ch, i, i + 1)); i++; continue; }

      fail('unexpected character ' + JSON.stringify(ch), i);
    }

    out.push(tok('eof', null, n, n));
    return out;
  };

})(window.XT = window.XT || {});
