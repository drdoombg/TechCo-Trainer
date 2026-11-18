/* =============================================================================
   Excel Trainer — engine/criteria.js
   One criteria engine, shared by COUNTIF, COUNTIFS, SUMIF, SUMIFS and
   AVERAGEIFS. Worth its own file because the *IFS family is how the whole KPI
   reporting layer of the source estate is built: COUNTIFS appears in 39 of the
   152 workbooks, more than any other aggregator.

   The thing that catches everyone out is that a criterion is a STRING with an
   operator buried inside it, and the string is usually assembled on the fly:

       =SUMIFS('Volumes Paste'!17:17,
               'Volumes Paste'!$13:$13, ">=" & K$14,
               'Volumes Paste'!$13:$13, "<=" & K$15)

   ">=" & 46234  produces the text  ">=46234", and SUMIFS then has to take that
   apart again. So the criterion is parsed here, not by the parser.

   Criteria matching is also LOOSER than the = operator:
       A1 holds the text "5".     =A1=5          -> FALSE   (types differ)
                                  =COUNTIF(A1,5) -> 1       (criteria coerce)
   Both behaviours are correct, and both are tested in tests.html.
   ============================================================================= */
(function (XT) {
  'use strict';

  var C = {};
  XT.C = C;
  var V = XT.V;

  /* --------------------------------------------------------------- parsing -- */

  /* parse(criterionValue) -> {
         op       : '=' '<>' '<' '<=' '>' '>='
         num      : a number, when the operand is numeric
         text     : the operand as text, otherwise
         isBlank  : true for the ""  criterion
         re       : a RegExp when the operand contains * or ?
     } */
  C.parse = function (crit) {
    if (V.isErr(crit)) return { err: crit };

    var op = '=', operand = crit;

    if (typeof crit === 'string') {
      var m = /^(<>|>=|<=|=|>|<)([\s\S]*)$/.exec(crit);
      if (m) { op = m[1]; operand = m[2]; }
    } else if (V.isEmpty(crit)) {
      // COUNTIF(range, A1) where A1 is itself empty behaves like the 0 criterion.
      operand = 0;
    }

    var out = { op: op };

    if (operand === '' ) {
      // "" on its own, or "=" with nothing after it: match blank-or-empty-text.
      out.isBlank = true;
      out.text = '';
      return out;
    }

    if (typeof operand === 'number') { out.num = operand; return out; }
    if (typeof operand === 'boolean') { out.bool = operand; return out; }

    var n = V.parseNumeric(operand);
    if (n !== null) { out.num = n; out.text = operand; return out; }

    var u = String(operand).toUpperCase();
    if (u === 'TRUE') { out.bool = true; return out; }
    if (u === 'FALSE') { out.bool = false; return out; }

    out.text = String(operand);
    if (/[*?]/.test(out.text) && (op === '=' || op === '<>')) {
      out.re = wildcardToRegExp(out.text);
    }
    return out;
  };

  /* Excel wildcards:  * any run of characters,  ? any single character,
     and ~ escapes the next * ? or ~ so you can look for a literal one. */
  function wildcardToRegExp(pat) {
    var out = '', i = 0;
    while (i < pat.length) {
      var ch = pat.charAt(i);
      if (ch === '~') {
        var nx = pat.charAt(i + 1);
        if (nx === '*' || nx === '?' || nx === '~') { out += esc(nx); i += 2; continue; }
        out += esc('~'); i++; continue;
      }
      if (ch === '*') { out += '[\\s\\S]*'; i++; continue; }
      if (ch === '?') { out += '[\\s\\S]'; i++; continue; }
      out += esc(ch); i++;
    }
    return new RegExp('^' + out + '$', 'i');
  }
  function esc(s) { return s.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }
  C.wildcardToRegExp = wildcardToRegExp;

  /* --------------------------------------------------------------- matching - */

  /* match(cellValue, parsedCriterion) -> true | false
     cellValue may be EMPTY, number, text, boolean or an error. */
  C.match = function (v, c) {
    if (c.err) return false;
    if (V.isErr(v)) {
      // Only an explicit comparison against the same error text matches.
      return (c.op === '=' || c.op === '<>') && c.text !== undefined
        ? ((v.code.toUpperCase() === c.text.toUpperCase()) === (c.op === '='))
        : false;
    }

    var blank = V.isEmpty(v) || v === '';

    // The "" criterion, and its negation.
    if (c.isBlank) {
      if (c.op === '=') return blank;
      if (c.op === '<>') return !blank;
    }

    // A blank cell matches nothing else -- COUNTIF(A:A,0) ignores empty cells.
    if (V.isEmpty(v)) return c.op === '<>' ? true : false;

    // -- numeric criterion -----------------------------------------------------
    if (c.num !== undefined) {
      var cv = (typeof v === 'number') ? v
        : (typeof v === 'string' ? V.parseNumeric(v) : (typeof v === 'boolean' ? null : null));
      if (cv === null || cv === undefined) {
        // Text that is not a number can still fail-or-pass a <> test.
        return c.op === '<>';
      }
      return cmpNum(cv, c.num, c.op);
    }

    // -- boolean criterion -----------------------------------------------------
    if (c.bool !== undefined) {
      if (typeof v !== 'boolean') return c.op === '<>';
      return (v === c.bool) === (c.op !== '<>');
    }

    // -- text criterion --------------------------------------------------------
    if (typeof v === 'number' || typeof v === 'boolean') {
      // A number never equals a text criterion, but ordering still applies:
      // numbers sort below text in Excel.
      if (c.op === '=') return false;
      if (c.op === '<>') return true;
      return cmpText(V.toText(v), c.text, c.op, typeof v !== 'string');
    }

    if (c.re) {
      var hit = c.re.test(v);
      return c.op === '<>' ? !hit : hit;
    }

    return cmpText(String(v), c.text, c.op, false);
  };

  function cmpNum(a, b, op) {
    switch (op) {
      case '=': return a === b;
      case '<>': return a !== b;
      case '>': return a > b;
      case '>=': return a >= b;
      case '<': return a < b;
      case '<=': return a <= b;
    }
    return false;
  }

  // numericCell: the cell held a number, so it sorts BELOW any text criterion.
  function cmpText(a, b, op, numericCell) {
    if (numericCell) {
      switch (op) {
        case '<': case '<=': return true;
        case '>': case '>=': return false;
        default: return false;
      }
    }
    var ua = a.toUpperCase(), ub = String(b).toUpperCase();
    var cmp = ua === ub ? 0 : (ua < ub ? -1 : 1);
    return cmpNum(cmp, 0, op);
  }

  /* ------------------------------------------------------ the shared workhorse */

  /* Walk a criteria-range / criterion list in lock-step and call back with the
     offset of every row (or column) that satisfies every pair.

     ranges : array of {values: XArr-like accessor, nr, nc}
     crits  : array of parsed criteria, same length
     The ranges must all be the same shape -- Excel returns #VALUE! otherwise,
     and so do we. */
  C.forEachMatch = function (ranges, crits, cb) {
    if (!ranges.length) return null;
    var nr = ranges[0].nr, nc = ranges[0].nc;
    for (var i = 1; i < ranges.length; i++) {
      if (ranges[i].nr !== nr || ranges[i].nc !== nc) {
        return V.value('the criteria ranges are different shapes: ' +
          nr + '×' + nc + ' against ' + ranges[i].nr + '×' + ranges[i].nc);
      }
    }
    for (var r = 0; r < nr; r++) {
      for (var c = 0; c < nc; c++) {
        var ok = true;
        for (var k = 0; k < ranges.length; k++) {
          if (!C.match(ranges[k].get(r, c), crits[k])) { ok = false; break; }
        }
        if (ok) cb(r, c);
      }
    }
    return null;
  };

})(window.XT = window.XT || {});
