/* =============================================================================
   Excel Trainer — engine/functions/logic.js
   IF, IFERROR, IFNA, AND, OR, NOT, TRUE, FALSE.

   These three carry the estate: IF appears in 2,235,460 cells across 73 of the
   152 workbooks, IFERROR in 750,917. Between them they are the blank-guard and
   the error-swallower that every pattern from P2 to P11 is built out of.

   IFERROR here does one extra thing that Excel does not: when it swallows an
   error it records what it swallowed. The error inspector reads that back, so a
   cell quietly showing "" can be made to confess. That is the whole of catalog
   defects D1 and D2 -- a VLOOKUP into a deleted column, permanently broken,
   wrapped in IFERROR and therefore invisible for years.
   ============================================================================= */
(function (XT) {
  'use strict';

  var V = XT.V, def = XT.defineFunction;
  function isBlock(v) { return V.isRange(v) || V.isArr(v); }

  /* ------------------------------------------------------------------- IF --- */

  def({
    name: 'IF', min: 2, max: 3, lazy: true, group: 'logic',
    sig: 'IF(logical_test, value_if_true, [value_if_false])',
    doc: 'Picks one of two values depending on a test.',
    estate: 'The workhorse: blank-guarding (IF(A1="","",…)), routing a row to a ' +
      'label, and nesting up to six deep as a lookup table written as code.',
    pats: ['P2', 'P3', 'P4', 'P11', 'P17', 'P18'],
    impl: function (nodes, ctx) {
      var cond = ctx.ev(nodes[0]);

      // Array condition inside a Ctrl+Shift+Enter formula: go elementwise.
      // This is what makes SMALL(IF((a=x)*(b=y),ROW(r)),k) work -- pattern P15.
      if (ctx.arrayMode && isBlock(cond)) {
        var cv = ctx.asView(cond);
        if (V.isErr(cv)) return cv;
        var tv = nodes[1] && nodes[1].t !== 'missing' ? ctx.asView(ctx.ev(nodes[1])) : null;
        var fv = nodes[2] && nodes[2].t !== 'missing' ? ctx.asView(ctx.ev(nodes[2])) : null;
        var rows = [];
        for (var i = 0; i < cv.nr; i++) {
          var row = [];
          for (var j = 0; j < cv.nc; j++) {
            var b = V.toBool(cv.get(i, j));
            if (V.isErr(b)) { row.push(b); continue; }
            if (b) row.push(tv ? tv.at(i, j) : true);
            else row.push(fv ? fv.at(i, j) : false);
          }
          rows.push(row);
        }
        return new V.XArr(rows);
      }

      var test = V.toBool(ctx.scalar(cond));
      if (V.isErr(test)) return test;

      if (test) {
        if (!nodes[1] || nodes[1].t === 'missing') return 0;
        return ctx.scalar(ctx.ev(nodes[1]));
      }
      // No third argument at all -> FALSE.
      // A third argument written as nothing, =IF(x,1,) -> 0.
      // Those are different, and catalog defect D4 depends on which one it is.
      if (nodes.length < 3) return false;
      if (nodes[2].t === 'missing') return 0;
      return ctx.scalar(ctx.ev(nodes[2]));
    }
  });

  /* -------------------------------------------------------------- IFERROR --- */

  def({
    name: 'IFERROR', min: 2, max: 2, lazy: true, group: 'logic',
    sig: 'IFERROR(value, value_if_error)',
    doc: 'Returns value, or value_if_error if value is any error.',
    estate: 'Almost always IFERROR(…,"") wrapped round a lookup, to turn "not ' +
      'found" into a blank cell. It catches EVERY error, including the #REF! ' +
      'of a structurally broken formula -- which is how breakage stays hidden.',
    pats: ['P4', 'P5', 'P10', 'P11'],
    impl: function (nodes, ctx) {
      var v = ctx.ev(nodes[0]);

      if (ctx.arrayMode && isBlock(v)) {
        var view = ctx.asView(v);
        if (V.isErr(view)) { return ctx.scalar(ctx.ev(nodes[1])); }
        var alt = null, rows = [];
        for (var i = 0; i < view.nr; i++) {
          var row = [];
          for (var j = 0; j < view.nc; j++) {
            var cell = view.get(i, j);
            if (V.isErr(cell)) {
              if (alt === null) alt = ctx.scalar(ctx.ev(nodes[1]));
              remember(ctx, nodes[0], cell);
              row.push(alt);
            } else row.push(cell);
          }
          rows.push(row);
        }
        return new V.XArr(rows);
      }

      v = ctx.scalar(v);
      if (V.isErr(v)) {
        remember(ctx, nodes[0], v);
        return ctx.scalar(ctx.ev(nodes[1]));
      }
      return v;
    }
  });

  /* ----------------------------------------------------------------- IFNA --- */

  def({
    name: 'IFNA', min: 2, max: 2, lazy: true, group: 'logic',
    sig: 'IFNA(value, value_if_na)',
    doc: 'Returns value, or value_if_na only when value is #N/A.',
    estate: 'Used 45 times against 750,917 uses of IFERROR -- and it is almost ' +
      'always the right one. IFNA catches "I looked and it is not there" while ' +
      'letting #REF! and #VALUE! shout, so a broken formula stays visible.',
    pats: ['P5'],
    impl: function (nodes, ctx) {
      var v = ctx.scalar(ctx.ev(nodes[0]));
      if (V.isErr(v) && v.code === '#N/A') {
        remember(ctx, nodes[0], v);
        return ctx.scalar(ctx.ev(nodes[1]));
      }
      return v;
    }
  });

  /* Record an error that a guard has just hidden, so the inspector can show it.
     Excel has no equivalent; this is the trainer teaching against the grain. */
  function remember(ctx, node, e) {
    if (!ctx.hidden) return;
    ctx.hidden.push({
      from: node ? XT.P.unparse(node) : '?',
      code: e.code,
      detail: e.detail || null
    });
  }

  /* ------------------------------------------------------- AND / OR / NOT --- */

  /* Note: AND and OR do NOT short-circuit in Excel. Every argument is
     evaluated, so an error anywhere in the list wins. */
  def({
    name: 'AND', min: 1, max: 255, group: 'logic',
    sig: 'AND(logical1, [logical2], …)',
    doc: 'TRUE only if every argument is TRUE.',
    estate: 'Used inside a date-window test, and notably multiplied by 1 to turn ' +
      'the boolean into 1/0: IFERROR(IF(N28=$U$17,AND(Q28>=$W$13,O28<=$W$17),"")*1,0)',
    pats: ['P12'],
    impl: function (args, ctx) {
      var any = false, out = true;
      for (var i = 0; i < args.length; i++) {
        var vals = flattenLogical(args[i], ctx);
        if (V.isErr(vals)) return vals;
        for (var j = 0; j < vals.length; j++) { any = true; if (!vals[j]) out = false; }
      }
      if (!any) return V.value('AND found no TRUE/FALSE values to test');
      return out;
    }
  });

  def({
    name: 'OR', min: 1, max: 255, group: 'logic',
    sig: 'OR(logical1, [logical2], …)',
    doc: 'TRUE if any argument is TRUE.',
    estate: 'Rare: 100 cells. Nested IF is used instead.',
    impl: function (args, ctx) {
      var any = false, out = false;
      for (var i = 0; i < args.length; i++) {
        var vals = flattenLogical(args[i], ctx);
        if (V.isErr(vals)) return vals;
        for (var j = 0; j < vals.length; j++) { any = true; if (vals[j]) out = true; }
      }
      if (!any) return V.value('OR found no TRUE/FALSE values to test');
      return out;
    }
  });

  def({
    name: 'NOT', min: 1, max: 1, group: 'logic',
    sig: 'NOT(logical)', doc: 'Reverses TRUE and FALSE.',
    estate: '346 cells, usually NOT(ISERROR(SEARCH(…))) inside conditional formatting.',
    impl: function (args, ctx) {
      var b = ctx.bool(args[0]);
      return V.isErr(b) ? b : !b;
    }
  });

  /* Ranges contribute only their TRUE/FALSE and numeric cells; text and blanks
     inside a range are ignored, exactly as in Excel. A direct text argument,
     however, is an error. */
  function flattenLogical(v, ctx) {
    var out = [];
    if (isBlock(v)) {
      var view = ctx.asView(v);
      if (V.isErr(view)) return view;
      var all = view.flat();
      for (var i = 0; i < all.length; i++) {
        var x = all[i];
        if (V.isErr(x)) return x;
        if (typeof x === 'boolean') out.push(x);
        else if (typeof x === 'number') out.push(x !== 0);
        // text and EMPTY inside a range are skipped
      }
      return out;
    }
    var b = V.toBool(ctx.scalar(v));
    if (V.isErr(b)) return b;
    out.push(b);
    return out;
  }

  /* ------------------------------------------------------- TRUE / FALSE ----- */

  def({
    name: 'TRUE', min: 0, max: 0, group: 'logic', sig: 'TRUE()',
    doc: 'The boolean TRUE.', impl: function () { return true; }
  });
  def({
    name: 'FALSE', min: 0, max: 0, group: 'logic', sig: 'FALSE()',
    doc: 'The boolean FALSE.', impl: function () { return false; }
  });

})(window.XT = window.XT || {});
