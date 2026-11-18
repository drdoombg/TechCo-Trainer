/* =============================================================================
   Excel Trainer — engine/shape.js
   Two related tools for reading formulas rather than writing them.

   1. NORMALISE. Strip a formula down to its shape, which is what the catalog
      counted:
          =IFERROR(VLOOKUP(C18,PASTE!$A:$DN,2,FALSE),"")
       -> IFERROR(VLOOKUP(R,SH!C:C,N,FALSE),"S")
      Sheet names become SH!, cell addresses become R, numbers become N, text
      literals become "S". Two formulas with the same shape are the same move,
      whatever cells they point at. That is how 4.2 million formulas collapsed
      into a few hundred recognisable patterns.

   2. LAYERS. Peel a formula apart into the stack of wrappers it is made of:
          IFERROR( IF( ISBLANK(key), "", VLOOKUP(key, table, 2, FALSE) ), "" )
          layer 1  IFERROR   catch anything that goes wrong
          layer 2  IF        do not look up a blank key
          layer 3  VLOOKUP   the actual join
      Being able to see that stack -- and build it one wrapper at a time -- is
      the skill the trainer is really after. Individual functions are easy; it
      is the stacking that is hard.
   ============================================================================= */
(function (XT) {
  'use strict';

  var S = {};
  XT.S = S;

  /* ------------------------------------------------------------- normalise -- */

  var STR = /"[^"]*"/g;
  var SHEET = /(\[\d+\])?('[^']*'|[A-Za-z_][A-Za-z0-9_. ]*)!/g;
  var COLR = /\$?[A-Z]{1,3}:\$?[A-Z]{1,3}/g;
  var CELL = /\$?[A-Z]{1,3}\$?\d{1,7}/g;
  // No lookbehind: not every browser engine has it, and a regex literal with
  // unsupported syntax is a parse-time error that would take the whole file out.
  var NUM = /\d+(\.\d+)?/g;
  function numNotPartOfName(s) {
    return s.replace(NUM, function (m, _d, off, whole) {
      var prev = off > 0 ? whole.charAt(off - 1) : '';
      if (/[A-Za-z0-9_.$]/.test(prev)) return m;
      return 'N';
    });
  }

  /* Deliberately a text transform in this exact order, not an AST walk, so that
     it produces byte-identical strings to the Python scanner that built the
     catalog. The shapes in scan/an/shapes.json can therefore be compared
     against these directly, and tests.html does exactly that.

     normalise(src)                -> catalog-faithful, keeps the _xlfn. prefix
     normalise(src, {strip:true})  -> also drops _xlfn., for pattern matching */
  S.normalise = function (src, opts) {
    var s = String(src || '');
    if (s.charAt(0) === '=') s = s.slice(1);
    s = s.replace(STR, '"S"');
    s = s.replace(SHEET, 'SH!');
    s = s.replace(COLR, 'C:C');
    s = s.replace(CELL, 'R');
    s = numNotPartOfName(s);
    if (opts && opts.strip) s = s.replace(/_xlfn\./g, '');
    s = s.replace(/\s+/g, '');
    return s;
  };

  /* ------------------------------------------------------------- matching --- */

  /* Which catalog patterns does this formula look like?
     Each pattern in data/patterns.js carries one or more `shapes` (exact
     normalised strings) and/or `test` regexes against the normalised form.
     Returns the matches, most specific first. */
  S.match = function (src) {
    var pats = (XT.data && XT.data.patterns) || [];
    var shape = S.normalise(src);
    var clean = S.normalise(src, { strip: true });
    var hits = [];
    pats.forEach(function (p) {
      var score = 0;
      if (p.shapes && (p.shapes.indexOf(shape) >= 0 || p.shapes.indexOf(clean) >= 0)) score = 100;
      else if (p.test) {
        for (var i = 0; i < p.test.length; i++) {
          if (new RegExp(p.test[i]).test(clean)) { score = Math.max(score, 60 - i); }
        }
      }
      if (score) hits.push({ id: p.id, title: p.title, score: score, pattern: p });
    });
    hits.sort(function (a, b) { return b.score - a.score; });
    return { shape: shape, clean: clean, hits: hits };
  };

  /* --------------------------------------------------------------- layers --- */

  /* Peel the stack. Walks down the "spine" of the formula -- at each level, the
     argument that itself contains the most work -- and reports one layer per
     function on the way down. */
  var ROLES = {
    IFERROR: 'guard: swallow any error',
    IFNA: 'guard: swallow "not found" only',
    IF: 'route: choose between two answers',
    VLOOKUP: 'join: find a row, return a column',
    XLOOKUP: 'join: find a key, return its partner',
    INDEX: 'pick: take an item by position',
    MATCH: 'locate: find the position of a value',
    COUNTIF: 'count: how many match one test',
    COUNTIFS: 'count: how many match every test',
    SUMIF: 'total: add what matches one test',
    SUMIFS: 'total: add what matches every test',
    AVERAGEIFS: 'average: mean of what matches',
    SUM: 'total',
    VALUE: 'convert: text to number',
    TEXT: 'convert: number to formatted text',
    LEFT: 'slice: from the start',
    RIGHT: 'slice: from the end',
    MID: 'slice: from the middle',
    TEXTBEFORE: 'slice: up to a delimiter',
    TEXTAFTER: 'slice: past a delimiter',
    TRIM: 'clean: strip spaces',
    SUBSTITUTE: 'clean: swap characters out',
    WEEKDAY: 'date: which day of the week',
    TIME: 'date: build a time of day',
    TODAY: 'date: today',
    SMALL: 'pick: the nth smallest',
    ROW: 'position: row number',
    AND: 'test: all of these',
    OR: 'test: any of these',
    NOT: 'test: the opposite',
    ISBLANK: 'test: truly empty?',
    ISERROR: 'test: is it an error?',
    LEN: 'measure: how many characters',
    SUMPRODUCT: 'count/total: the old multi-criteria way',
    RANK: 'order: position in a list',
    MIN: 'extreme: smallest',
    MAX: 'extreme: largest'
  };
  S.ROLES = ROLES;

  function weight(node) {
    if (!node) return 0;
    switch (node.t) {
      case 'func': {
        var w = 1;
        node.args.forEach(function (a) { w += weight(a); });
        return w;
      }
      case 'bin': return weight(node.l) + weight(node.r);
      case 'un': case 'post': case 'paren': return weight(node.x);
      case 'arr': {
        var t = 0;
        node.rows.forEach(function (r) { r.forEach(function (c) { t += weight(c); }); });
        return t;
      }
      default: return 0;
    }
  }

  S.layers = function (src) {
    var r = XT.P.parse(src);
    if (!r.ok) return { ok: false, error: r.error };
    var out = [], node = r.ast, depth = 0, guard = 0;

    while (node && depth < 40) {
      if (node.t === 'paren') { node = node.x; continue; }
      if (node.t === 'un' || node.t === 'post') { node = node.x; continue; }

      if (node.t === 'func') {
        out.push({
          depth: depth,
          fn: node.name,
          role: ROLES[node.name] || 'step',
          text: XT.P.unparse(node),
          args: node.args.map(function (a) { return XT.P.unparse(a); })
        });
        // Descend into the heaviest argument: that is where the work is.
        var best = null, bw = -1;
        node.args.forEach(function (a) {
          var w = weight(a);
          if (w > bw) { bw = w; best = a; }
        });
        if (bw <= 0) break;
        node = best;
        depth++;
        continue;
      }

      if (node.t === 'bin') {
        out.push({
          depth: depth, fn: node.op, role: 'operator', text: XT.P.unparse(node),
          args: [XT.P.unparse(node.l), XT.P.unparse(node.r)]
        });
        var lw = weight(node.l), rw = weight(node.r);
        if (lw <= 0 && rw <= 0) break;
        node = lw >= rw ? node.l : node.r;
        depth++;
        continue;
      }
      break;
    }
    if (guard) { /* reserved */ }
    return { ok: true, layers: out, depth: out.length };
  };

  /* ---------------------------------------------------------- AST summaries */

  // Every function name used anywhere in the formula.
  S.functionsUsed = function (src) {
    var r = XT.P.parse(src);
    if (!r.ok) return [];
    var seen = {};
    (function walk(n) {
      if (!n) return;
      if (n.t === 'func') { seen[n.name] = (seen[n.name] || 0) + 1; n.args.forEach(walk); }
      else if (n.t === 'bin') { walk(n.l); walk(n.r); }
      else if (n.t === 'un' || n.t === 'post' || n.t === 'paren') walk(n.x);
      else if (n.t === 'arr') n.rows.forEach(function (row) { row.forEach(walk); });
    })(r.ast);
    return Object.keys(seen).map(function (k) { return { fn: k, n: seen[k] }; });
  };

  // How deeply nested the IFs go. The estate reaches six.
  S.nestDepth = function (src, fnName) {
    var r = XT.P.parse(src);
    if (!r.ok) return 0;
    var want = (fnName || '').toUpperCase();
    function walk(n, d) {
      if (!n) return d;
      var here = d;
      if (n.t === 'func') {
        var nd = (!want || n.name === want) ? d + 1 : d;
        var m = nd;
        n.args.forEach(function (a) { m = Math.max(m, walk(a, nd)); });
        return m;
      }
      if (n.t === 'bin') return Math.max(walk(n.l, d), walk(n.r, d));
      if (n.t === 'un' || n.t === 'post' || n.t === 'paren') return walk(n.x, d);
      if (n.t === 'arr') {
        var mm = d;
        n.rows.forEach(function (row) { row.forEach(function (c) { mm = Math.max(mm, walk(c, d)); }); });
        return mm;
      }
      return here;
    }
    return walk(r.ast, 0);
  };

  // Every reference in the formula, as text, with its $ pattern. Used by the
  // grader to check anchoring, and by the grid to outline referenced ranges.
  S.references = function (src) {
    var r = XT.P.parse(src);
    if (!r.ok) return [];
    var out = [];
    (function walk(n) {
      if (!n) return;
      if (n.t === 'ref' || n.t === 'range') { out.push({ node: n, text: XT.P.unparse(n) }); return; }
      if (n.t === 'func') { n.args.forEach(walk); return; }
      if (n.t === 'bin') { walk(n.l); walk(n.r); return; }
      if (n.t === 'un' || n.t === 'post' || n.t === 'paren') { walk(n.x); return; }
      if (n.t === 'arr') n.rows.forEach(function (row) { row.forEach(walk); });
    })(r.ast);
    return out;
  };

  // Does the formula contain a literal that looks like it should have been a
  // reference? The grader uses this to fail hard-coded answers.
  S.literals = function (src) {
    var r = XT.P.parse(src);
    if (!r.ok) return [];
    var out = [];
    (function walk(n, inside) {
      if (!n) return;
      if (n.t === 'num' || n.t === 'str') { out.push({ t: n.t, v: n.v, inside: inside }); return; }
      if (n.t === 'func') { n.args.forEach(function (a) { walk(a, n.name); }); return; }
      if (n.t === 'bin') { walk(n.l, inside); walk(n.r, inside); return; }
      if (n.t === 'un' || n.t === 'post' || n.t === 'paren') { walk(n.x, inside); return; }
      if (n.t === 'arr') n.rows.forEach(function (row) { row.forEach(function (c) { walk(c, inside); }); });
    })(r.ast, null);
    return out;
  };

})(window.XT = window.XT || {});
