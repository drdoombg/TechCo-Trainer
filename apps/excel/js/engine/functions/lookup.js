/* =============================================================================
   Excel Trainer — engine/functions/lookup.js
   VLOOKUP, XLOOKUP, INDEX, MATCH, LOOKUP, OFFSET, INDIRECT, TRANSPOSE.

   The join. VLOOKUP appears in 689,179 cells, XLOOKUP in 128,323. Together with
   IF and IFERROR they are the entire enrichment layer of the source estate.

   One behaviour matters more than any other here, so it is spelled out:

       VLOOKUP AND XLOOKUP DO NOT COERCE TYPES ON AN EXACT MATCH.

   If the key column holds the TEXT "5002874" (which is what a CSV import or a
   barcode scanner gives you) and you look up the NUMBER 5002874, the answer is
   #N/A. COUNTIF would have matched it; VLOOKUP will not. That single asymmetry
   is behind a large share of real-world "my lookup is broken" reports, and the
   trainer reproduces it deliberately -- several fixtures ship with a key column
   half text and half number, exactly as the real exports arrive.
   ============================================================================= */
(function (XT) {
  'use strict';

  var V = XT.V, C = XT.C, def = XT.defineFunction;

  /* ------------------------------------------------- exact-match comparison */

  /* How VLOOKUP/XLOOKUP/MATCH(…,0) decide "same".
       - text vs text    : case-insensitive, and the KEY may contain * and ?
       - number vs number: strict equality
       - anything else   : no match
     EMPTY behaves as 0 on the key side and matches only empty cells. */
  function exactEq(key, cell, allowWild) {
    if (V.isErr(cell)) return false;
    if (V.isEmpty(key)) return V.isEmpty(cell) || cell === 0;
    if (typeof key === 'number') return typeof cell === 'number' && cell === key;
    if (typeof key === 'boolean') return typeof cell === 'boolean' && cell === key;
    if (typeof key === 'string') {
      if (typeof cell !== 'string') return false;
      if (allowWild && /[*?]/.test(key)) return C.wildcardToRegExp(key).test(cell);
      return key.toUpperCase() === cell.toUpperCase();
    }
    return false;
  }

  // Ordering for approximate matches. Returns -1 / 0 / 1, or null if the two
  // values are not comparable (different types).
  function order(a, b) {
    if (V.isEmpty(a)) a = 0;
    if (V.isEmpty(b)) return null;
    if (typeof a === 'number' && typeof b === 'number') return a === b ? 0 : (a < b ? -1 : 1);
    if (typeof a === 'string' && typeof b === 'string') {
      var ua = a.toUpperCase(), ub = b.toUpperCase();
      return ua === ub ? 0 : (ua < ub ? -1 : 1);
    }
    return null;
  }

  // Pull a one-dimensional list out of a view, remembering its orientation.
  function vector(view) {
    var vertical = view.nc === 1 || view.nr > 1;
    var n = (view.nc === 1) ? view.nr : (view.nr === 1 ? view.nc : view.nr);
    return {
      n: n,
      vertical: view.nc === 1 || (view.nr > 1 && view.nc > 1),
      get: function (i) { return (view.nc === 1) ? view.get(i, 0) : view.get(0, i); }
    };
  }

  /* -------------------------------------------------------------- VLOOKUP --- */

  def({
    name: 'VLOOKUP', min: 3, max: 4, group: 'lookup', arrayArgs: [1],
    sig: 'VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])',
    doc: 'Finds a value in the FIRST column of a block and returns a value from ' +
      'the nth column of the same row.',
    estate: 'The house style is a whole-column block and an explicit exact match: ' +
      'VLOOKUP(C18, PASTE!$A:$DN, 2, FALSE), wrapped in IFERROR(…,""). ' +
      'col_index_num counts from the first column of the block, NOT from column A ' +
      '-- the classic off-by-one.',
    pats: ['P5', 'P9', 'P11', 'P18'],
    impl: function (args, ctx) {
      var key = ctx.scalar(args[0]);
      if (V.isErr(key)) return key;
      var view = ctx.asView(args[1]);
      if (V.isErr(view)) return view;
      var idx = ctx.int(args[2]);
      if (V.isErr(idx)) return idx;

      var approx = true;
      if (args.length >= 4 && args[3] !== undefined) {
        var b = ctx.bool(args[3]);
        if (V.isErr(b)) return b;
        approx = b;
      }

      if (idx < 1) return V.value('col_index_num must be 1 or more; it was ' + idx);
      if (idx > view.nc) {
        return V.ref('col_index_num is ' + idx + ' but the block is only ' +
          view.nc + ' column' + (view.nc === 1 ? '' : 's') + ' wide');
      }

      var found = -1;
      if (!approx) {
        for (var r = 0; r < view.nr; r++) {
          if (exactEq(key, view.get(r, 0), true)) { found = r; break; }
        }
        if (found < 0) {
          return V.na('VLOOKUP looked for ' + describeKey(key) +
            ' in the first column of the block and did not find it' + typeHint(key, view));
        }
      } else {
        // Approximate: the block must be sorted ascending. Take the last row
        // whose key is <= the lookup value.
        for (var i = 0; i < view.nr; i++) {
          var o = order(key, view.get(i, 0));
          if (o === null) continue;
          if (o >= 0) found = i; else break;
        }
        if (found < 0) {
          return V.na('VLOOKUP with an approximate match found nothing smaller than ' +
            describeKey(key) + '. The first column must be sorted ascending.');
        }
      }
      return view.get(found, idx - 1);
    }
  });

  function describeKey(k) {
    if (V.isEmpty(k)) return 'an empty value';
    if (typeof k === 'string') return 'the text ' + JSON.stringify(k);
    if (typeof k === 'number') return 'the number ' + V.numToText(k);
    return String(k);
  }

  /* If the key is a number and the column is full of numeric TEXT (or the other
     way round) say so. This is the single most common cause of a #N/A in the
     kind of workbook this trainer is modelled on. */
  function typeHint(key, view) {
    var kt = typeof key;
    var sawText = false, sawNum = false;
    var limit = Math.min(view.nr, 200);
    for (var r = 0; r < limit; r++) {
      var v = view.get(r, 0);
      if (typeof v === 'string' && v !== '') { sawText = true; if (V.parseNumeric(v) !== null) sawNum = true; }
      else if (typeof v === 'number') sawNum = true;
    }
    if (kt === 'number' && sawText && sawNum) {
      return '. The key column looks like numbers STORED AS TEXT; a number will ' +
        'never match text, even when they look identical.';
    }
    if (kt === 'string' && V.parseNumeric(key) !== null && sawNum && !sawText) {
      return '. You are looking up text that looks like a number, in a column of ' +
        'real numbers. Wrap the key in VALUE(), or fix the source.';
    }
    return '.';
  }

  /* -------------------------------------------------------------- XLOOKUP --- */

  def({
    name: 'XLOOKUP', min: 3, max: 6, group: 'lookup', arrayArgs: [1, 2],
    sig: 'XLOOKUP(lookup_value, lookup_array, return_array, [if_not_found], ' +
      '[match_mode], [search_mode])',
    doc: 'Finds a value in one array and returns the matching item from another. ' +
      'Exact match by default.',
    estate: 'Used as a column fan-out: the same key, 37 separate XLOOKUPs, one per ' +
      'output column. Also the clean way to join on two keys at once, because the ' +
      'lookup_array can be built with & instead of needing a helper column: ' +
      'XLOOKUP(B2&G2, M:M&V:V, J:J).',
    pats: ['P9', 'P10'],
    impl: function (args, ctx) {
      var key = ctx.scalar(args[0]);
      if (V.isErr(key)) return key;
      var lv = ctx.asView(args[1]);
      if (V.isErr(lv)) return lv;
      var rv = ctx.asView(args[2]);
      if (V.isErr(rv)) return rv;

      var notFound = (args.length >= 4 && args.nodes && args.nodes[3] &&
        args.nodes[3].t !== 'missing') ? ctx.scalar(args[3]) : undefined;
      var mode = args.length >= 5 ? ctx.int(args[4]) : 0;
      if (V.isErr(mode)) return mode;
      var search = args.length >= 6 ? ctx.int(args[5]) : 1;
      if (V.isErr(search)) return search;
      if (mode === undefined || isNaN(mode)) mode = 0;

      var L = vector(lv);
      if (L.n === 0) return V.na('the lookup_array is empty');

      var best = -1;
      if (mode === 0 || mode === 2) {
        var wild = (mode === 2);
        if (search === -1) {
          for (var i = L.n - 1; i >= 0; i--) if (exactEq(key, L.get(i), wild)) { best = i; break; }
        } else {
          for (var j = 0; j < L.n; j++) if (exactEq(key, L.get(j), wild)) { best = j; break; }
        }
      } else if (mode === -1) {
        // exact, or the next SMALLER item
        var bo = null;
        for (var k = 0; k < L.n; k++) {
          var o = order(key, L.get(k));
          if (o === null) continue;
          if (o === 0) { best = k; break; }
          if (o > 0) {
            var cand = L.get(k);
            if (bo === null || order(cand, bo) > 0) { bo = cand; best = k; }
          }
        }
      } else if (mode === 1) {
        // exact, or the next LARGER item
        var bo2 = null;
        for (var m = 0; m < L.n; m++) {
          var o2 = order(key, L.get(m));
          if (o2 === null) continue;
          if (o2 === 0) { best = m; break; }
          if (o2 < 0) {
            var cand2 = L.get(m);
            if (bo2 === null || order(cand2, bo2) < 0) { bo2 = cand2; best = m; }
          }
        }
      } else {
        return V.value('match_mode must be 0, -1, 1 or 2');
      }

      if (best < 0) {
        if (notFound !== undefined) return notFound;
        return V.na('XLOOKUP looked for ' + describeKey(key) +
          ' in the lookup_array and did not find it' +
          (lv.nc === 1 ? typeHint(key, lv) : '.'));
      }

      // Return array must line up with the lookup array.
      if (rv.nr === lv.nr && rv.nc === lv.nc) {
        return (rv.nc === 1) ? rv.get(best, 0) : rv.get(0, best);
      }
      if (lv.nc === 1 && rv.nr === lv.nr) {
        // A multi-column return_array: hand back the whole matched row.
        var row = [];
        for (var c = 0; c < rv.nc; c++) row.push(rv.get(best, c));
        return new V.XArr([row]);
      }
      if (lv.nr === 1 && rv.nc === lv.nc) {
        var col = [];
        for (var r2 = 0; r2 < rv.nr; r2++) col.push([rv.get(r2, best)]);
        return new V.XArr(col);
      }
      return V.value('the lookup_array (' + lv.nr + '×' + lv.nc +
        ') and return_array (' + rv.nr + '×' + rv.nc + ') are not the same shape');
    }
  });

  /* ---------------------------------------------------------------- MATCH --- */

  def({
    name: 'MATCH', min: 2, max: 3, group: 'lookup', arrayArgs: [1],
    sig: 'MATCH(lookup_value, lookup_array, [match_type])',
    doc: 'Returns the POSITION of a value in a list, counting from 1.',
    estate: 'Almost always paired with INDEX. The clever use in the estate is ' +
      'MATCH(1,(C2>=Calendars!G:G)*(C2<=Calendars!H:H),0) -- find the row where a ' +
      'date falls between a start and an end column (pattern P16).',
    pats: ['P15', 'P16'],
    impl: function (args, ctx) {
      var key = ctx.scalar(args[0]);
      if (V.isErr(key)) return key;
      var view = ctx.asView(args[1]);
      if (V.isErr(view)) return view;
      var type = args.length >= 3 ? ctx.int(args[2]) : 1;
      if (V.isErr(type)) return type;

      var L = vector(view);
      var i, o, best = -1;

      if (type === 0) {
        for (i = 0; i < L.n; i++) if (exactEq(key, L.get(i), true)) return i + 1;
        return V.na('MATCH did not find ' + describeKey(key) + ' in the list');
      }
      if (type === 1) {
        for (i = 0; i < L.n; i++) {
          o = order(key, L.get(i));
          if (o === null) continue;
          if (o >= 0) best = i; else break;
        }
        if (best < 0) return V.na('MATCH found nothing less than or equal to ' +
          describeKey(key) + '. With match_type 1 the list must be sorted ascending.');
        return best + 1;
      }
      if (type === -1) {
        for (i = 0; i < L.n; i++) {
          o = order(key, L.get(i));
          if (o === null) continue;
          if (o <= 0) best = i; else break;
        }
        if (best < 0) return V.na('MATCH found nothing greater than or equal to ' +
          describeKey(key) + '. With match_type -1 the list must be sorted descending.');
        return best + 1;
      }
      return V.value('match_type must be 1, 0 or -1');
    }
  });

  /* ---------------------------------------------------------------- INDEX --- */

  def({
    name: 'INDEX', min: 2, max: 3, group: 'lookup', arrayArgs: [0],
    sig: 'INDEX(array, row_num, [column_num])',
    doc: 'Returns the item at a given position in a block.',
    estate: 'INDEX(Calendars!F:F, MATCH(…)) to read a week label, and the ' +
      'INDEX/SMALL/IF array formula that retrieves the nth matching row rather ' +
      'than just the first (pattern P15).',
    pats: ['P15', 'P16'],
    impl: function (args, ctx) {
      var view = ctx.asView(args[0]);
      if (V.isErr(view)) return view;
      var rn = ctx.int(args[1]);
      if (V.isErr(rn)) return rn;
      var cn = args.length >= 3 ? ctx.int(args[2]) : null;
      if (V.isErr(cn)) return cn;

      // One-dimensional block with a single index: walk along it.
      if (cn === null && (view.nr === 1 || view.nc === 1)) {
        var L = vector(view);
        if (rn === 0) return view;                       // the whole vector
        if (rn < 1 || rn > L.n) {
          return V.ref('INDEX was asked for item ' + rn + ' of a list of ' + L.n);
        }
        return L.get(rn - 1);
      }

      if (cn === null) cn = 1;
      if (rn === 0 && cn === 0) return view;
      if (rn === 0) {
        if (cn < 1 || cn > view.nc) return V.ref('INDEX: column ' + cn + ' is outside the block');
        var col = [];
        for (var r = 0; r < view.nr; r++) col.push([view.get(r, cn - 1)]);
        return new V.XArr(col);
      }
      if (cn === 0) {
        if (rn < 1 || rn > view.nr) return V.ref('INDEX: row ' + rn + ' is outside the block');
        var row = [];
        for (var c = 0; c < view.nc; c++) row.push(view.get(rn - 1, c));
        return new V.XArr([row]);
      }
      if (rn < 1 || rn > view.nr || cn < 1 || cn > view.nc) {
        return V.ref('INDEX was asked for row ' + rn + ', column ' + cn +
          ' of a block that is ' + view.nr + '×' + view.nc);
      }
      return view.get(rn - 1, cn - 1);
    }
  });

  /* --------------------------------------------------------------- LOOKUP --- */

  def({
    name: 'LOOKUP', min: 2, max: 3, group: 'lookup', arrayArgs: [1, 2],
    sig: 'LOOKUP(lookup_value, lookup_vector, [result_vector])',
    doc: 'The old approximate-match lookup. Needs a sorted list.',
    estate: '355 cells. Superseded by VLOOKUP and XLOOKUP; kept so old formulas run.',
    impl: function (args, ctx) {
      var key = ctx.scalar(args[0]);
      if (V.isErr(key)) return key;
      var lv = ctx.asView(args[1]);
      if (V.isErr(lv)) return lv;
      var L = vector(lv), best = -1, i, o;
      for (i = 0; i < L.n; i++) {
        o = order(key, L.get(i));
        if (o === null) continue;
        if (o >= 0) best = i; else break;
      }
      if (best < 0) return V.na('LOOKUP found nothing less than or equal to ' + describeKey(key));
      if (args.length >= 3) {
        var rv = ctx.asView(args[2]);
        if (V.isErr(rv)) return rv;
        var Rv = vector(rv);
        if (best >= Rv.n) return V.ref('the result_vector is shorter than the lookup_vector');
        return Rv.get(best);
      }
      return L.get(best);
    }
  });

  /* --------------------------------------------------------------- OFFSET --- */

  def({
    name: 'OFFSET', min: 3, max: 5, group: 'lookup',
    sig: 'OFFSET(reference, rows, cols, [height], [width])',
    doc: 'Returns a reference shifted from a starting reference.',
    estate: '7 cells, e.g. OFFSET(AG3,,MATCH(S3,AH3:AM3,0)) to pick a retailer column.',
    impl: function (args, ctx) {
      if (!V.isRange(args[0])) {
        return V.value('OFFSET needs a reference as its first argument');
      }
      var base = args[0];
      var dr = args.nodes && args.nodes[1] && args.nodes[1].t === 'missing' ? 0 : ctx.int(args[1]);
      if (V.isErr(dr)) return dr;
      var dc = args.nodes && args.nodes[2] && args.nodes[2].t === 'missing' ? 0 : ctx.int(args[2]);
      if (V.isErr(dc)) return dc;
      var h = (args.length >= 4 && args.nodes[3] && args.nodes[3].t !== 'missing')
        ? ctx.int(args[3]) : base.height();
      if (V.isErr(h)) return h;
      var w = (args.length >= 5 && args.nodes[4] && args.nodes[4].t !== 'missing')
        ? ctx.int(args[4]) : base.width();
      if (V.isErr(w)) return w;

      var r1 = base.r1 + dr, c1 = base.c1 + dc;
      if (r1 < 0 || c1 < 0 || h < 1 || w < 1) {
        return V.ref('OFFSET moved outside the sheet');
      }
      return new V.XRange(base.sheet, r1, c1, r1 + h - 1, c1 + w - 1);
    }
  });

  /* ------------------------------------------------------------- INDIRECT --- */

  def({
    name: 'INDIRECT', min: 1, max: 2, group: 'lookup',
    sig: 'INDIRECT(ref_text, [a1])',
    doc: 'Turns a piece of text into a real reference.',
    estate: '554 cells in one workbook, building a sheet name from a cell. Powerful ' +
      'and brittle: nothing can trace through it, and it cannot see a closed workbook.',
    impl: function (args, ctx) {
      var t = ctx.text(args[0]);
      if (V.isErr(t)) return t;
      var parsed = XT.P.parse(t);
      if (!parsed.ok) return V.ref('INDIRECT could not read ' + JSON.stringify(t) + ' as a reference');
      var v = ctx.ev(parsed.ast);
      if (!V.isRange(v) && !V.isArr(v)) {
        return V.ref('INDIRECT needs text that describes a cell or range');
      }
      return v;
    }
  });

  /* ------------------------------------------------------------ TRANSPOSE --- */

  def({
    name: 'TRANSPOSE', min: 1, max: 1, group: 'lookup', arrayArgs: [0],
    sig: 'TRANSPOSE(array)', doc: 'Flips rows into columns.',
    estate: '3 cells.',
    impl: function (args, ctx) {
      var view = ctx.asView(args[0]);
      if (V.isErr(view)) return view;
      var rows = [];
      for (var c = 0; c < view.nc; c++) {
        var row = [];
        for (var r = 0; r < view.nr; r++) row.push(view.get(r, c));
        rows.push(row);
      }
      return new V.XArr(rows);
    }
  });

  XT.lookupHelpers = { exactEq: exactEq, order: order, vector: vector };

})(window.XT = window.XT || {});
