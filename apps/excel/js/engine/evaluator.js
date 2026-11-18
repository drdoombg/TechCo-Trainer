/* =============================================================================
   Excel Trainer — engine/evaluator.js
   Walks the AST and produces a value. One case per node type; read it beside
   parser.js and the whole language is accounted for.

   Three things worth understanding before you read the code:

   1. A RANGE IS A VALUE HERE. evalNode on a range node returns an XRange -- the
      reference, not its contents. Functions that want the reference (ISBLANK,
      ROW, COUNTBLANK, OFFSET) get it; functions that want contents ask ctx for
      a view. The collapse to a single value happens in ctx.scalar, which is
      Excel's implicit intersection.

   2. ARRAY CONTEXT IS A FLAG, NOT A TYPE. ctx.arrayMode is true inside a
      Ctrl+Shift+Enter cell, and temporarily true while evaluating the argument
      positions a function declares in arrayArgs. Operators consult it to
      decide between "broadcast elementwise" and "intersect to one value".

   3. ERRORS TRAVEL. Every operator checks its operands for an error first and
      returns it unchanged, left before right -- so #REF! from a deleted column
      arrives at the top of the formula with its identity intact, which is what
      lets IFERROR hide it (defects D1/D2) and lets the error inspector dig it
      back out.
   ============================================================================= */
(function (XT) {
  'use strict';

  var V = XT.V, R = XT.R;

  /* ------------------------------------------------------------ entry point */

  XT.evalNode = function (node, ctx) {
    if (!node) return V.EMPTY;

    // The stepper hooks in here: it records (node, value) for every node it
    // sees, which is enough to replay an evaluation one step at a time.
    var v = evalInner(node, ctx);
    if (ctx.trace) ctx.trace.push({ node: node, value: v });
    return v;
  };

  function evalInner(node, ctx) {
    switch (node.t) {

      case 'num': return node.v;
      case 'str': return node.v;
      case 'bool': return node.v;

      // A #REF! typed into, or left behind in, a live formula. 17 of the top
      // 400 formula shapes in the source estate contain one.
      case 'err': return V.err(node.v, 'this error is written into the formula itself');

      // An argument written as nothing: =TEXT(C1,)  or  =IF(x,1,)
      case 'missing': return V.EMPTY;

      case 'paren': return ctx.ev(node.x);

      case 'arr': return evalArrayConst(node, ctx);

      case 'un': return evalUnary(node, ctx);
      case 'post': {
        var x = ctx.num(ctx.ev(node.x));
        return V.isErr(x) ? x : x / 100;
      }

      case 'bin': return evalBinary(node, ctx);

      case 'ref': return evalRef(node, ctx);
      case 'range': return evalRange(node, ctx);
      case 'name': return evalName(node, ctx);
      case 'table': return evalTable(node, ctx);

      case 'func': return evalCall(node, ctx);
    }
    return V.value('the engine does not know this kind of expression');
  }

  /* ----------------------------------------------------------------- unary -- */

  /* Unary minus coerces to a number. Unary PLUS does not: it is a no-op that
     leaves text alone, which is why the Lotus-style  =+K13  works on anything.
     Flagged in OPEN_QUESTIONS.md for confirmation. */
  function evalUnary(node, ctx) {
    var x = ctx.ev(node.x);
    if (node.op === '+') return ctx.scalar(x);
    var n = ctx.num(x);
    return V.isErr(n) ? n : -n;
  }

  /* ---------------------------------------------------------------- binary -- */

  var CMP = { '=': 1, '<>': 1, '<': 1, '<=': 1, '>': 1, '>=': 1 };

  function evalBinary(node, ctx) {
    var l = ctx.ev(node.l), r = ctx.ev(node.r);

    // In array context, a range or array operand makes the whole operation
    // elementwise. This is what builds (C2>=Calendars!G:G)*(C2<=Calendars!H:H)
    // in pattern P16, and 'MIMS'!$M:$M & 'MIMS'!V:V in pattern P9.
    if (ctx.arrayMode && (isBlock(l) || isBlock(r))) {
      return broadcast(l, r, node.op, ctx);
    }

    l = ctx.scalar(l); r = ctx.scalar(r);
    return applyOp(l, r, node.op);
  }

  function isBlock(v) { return V.isRange(v) || V.isArr(v); }

  function broadcast(l, r, op, ctx) {
    var lv = isBlock(l) ? ctx.asView(l) : null;
    var rv = isBlock(r) ? ctx.asView(r) : null;
    if (V.isErr(lv)) return lv;
    if (V.isErr(rv)) return rv;
    var nr = Math.max(lv ? lv.nr : 1, rv ? rv.nr : 1);
    var nc = Math.max(lv ? lv.nc : 1, rv ? rv.nc : 1);
    // Guard against a runaway: two whole columns clipped to a big used range
    // is fine, but a cross product is not.
    if (nr * nc > 400000) {
      return V.num('that array operation is too large for the trainer (' + nr + '×' + nc + ')');
    }
    var rows = [];
    for (var i = 0; i < nr; i++) {
      var row = [];
      for (var j = 0; j < nc; j++) {
        var a = lv ? lv.at(i, j) : l;
        var b = rv ? rv.at(i, j) : r;
        row.push(applyOp(a, b, op));
      }
      rows.push(row);
    }
    return new V.XArr(rows);
  }

  function applyOp(l, r, op) {
    if (V.isErr(l)) return l;
    if (V.isErr(r)) return r;

    if (CMP[op]) return V.compare(l, r, op);

    if (op === '&') {
      var a = V.toText(l); if (V.isErr(a)) return a;
      var b = V.toText(r); if (V.isErr(b)) return b;
      return a + b;
    }

    var x = V.toNum(l); if (V.isErr(x)) return x;
    var y = V.toNum(r); if (V.isErr(y)) return y;
    switch (op) {
      case '+': return x + y;
      case '-': return x - y;
      case '*': return x * y;
      case '/':
        if (y === 0) return V.div0('dividing ' + V.numToText(x) + ' by zero');
        return x / y;
      case '^': {
        var p = Math.pow(x, y);
        if (!isFinite(p) || isNaN(p)) return V.num('that power is out of range');
        return p;
      }
    }
    return V.value('unknown operator ' + op);
  }
  XT.applyOp = applyOp;

  /* ------------------------------------------------------------ array const */

  function evalArrayConst(node, ctx) {
    var rows = node.rows.map(function (row) {
      return row.map(function (n) { return ctx.scalar(ctx.ev(n)); });
    });
    var w = 0;
    rows.forEach(function (r) { if (r.length > w) w = r.length; });
    rows.forEach(function (r) { while (r.length < w) r.push(V.EMPTY); });
    return new V.XArr(rows);
  }

  /* ------------------------------------------------------------- references */

  /* An external reference. We parse it -- that is required, because 20 of the
     152 source workbooks contain them -- but we cannot follow it, so it reports
     a #REF! that says so. That is the honest answer and it is also the lesson
     behind defects D1 and D10. */
  function externalError(node) {
    return V.ref('this points into another workbook (' + node.ext +
      ') which is not open, so the link cannot be followed');
  }

  /* A single cell evaluates to a 1x1 RANGE, not to its contents.

     That looks pedantic until you need ISBLANK(A1), ROW(A1), COUNTBLANK(A1) or
     OFFSET(A1,2,0), all of which need the reference itself. Everything that
     wants the contents goes through ctx.scalar / ctx.num / ctx.text, which
     collapse a 1x1 range to its value automatically. */
  function evalRef(node, ctx) {
    if (node.ext) return externalError(node);
    var sheetName = node.sheet || ctx.sheet;
    if (!ctx.wb.sheet(sheetName)) {
      return V.ref('there is no sheet called ' + JSON.stringify(sheetName));
    }
    return new V.XRange(sheetName, node.addr.r, node.addr.c, node.addr.r, node.addr.c,
      { abs: { ra: node.addr.ra, ca: node.addr.ca } });
  }

  function evalRange(node, ctx) {
    if (node.ext) return externalError(node);
    var sheetName = node.sheet || ctx.sheet;
    if (!ctx.wb.sheet(sheetName)) {
      return V.ref('there is no sheet called ' + JSON.stringify(sheetName));
    }
    if (node.kind === 'col') {
      return new V.XRange(sheetName, 0, node.c1, R.MAXROW - 1, node.c2, { wholeCol: true });
    }
    if (node.kind === 'row') {
      return new V.XRange(sheetName, node.r1, 0, node.r2, R.MAXCOL - 1, { wholeRow: true });
    }
    return new V.XRange(sheetName, node.a.r, node.a.c, node.b.r, node.b.c);
  }

  /* ----------------------------------------------------------- defined names */

  /* Defined names in the source estate come in two flavours: tidy picklist
     ranges (Shift -> picklists!$G$2:$G$10) and inherited wreckage pointing at
     #REF!. Both resolve here; the broken ones produce #REF! with an
     explanation, which is what the Name Manager lesson needs. */
  function evalName(node, ctx) {
    var def = ctx.wb.names[String(node.name).toUpperCase()];
    if (!def) {
      return V.name('there is no function or defined name called ' + JSON.stringify(node.name));
    }
    if (def.broken || !def.ast) {
      return V.ref('the defined name ' + node.name + ' points at ' + def.text +
        ', which is not a usable reference');
    }
    return ctx.ev(def.ast);
  }

  /* --------------------------------------------------- structured references */

  /* Table4[[#This Row],[SUPPLIER]] and Table1[Barcode].
     Structured references turn up mixed with plain A1 refs inside a single
     formula in the source files, so they cannot be skipped. */
  function evalTable(node, ctx) {
    var t = ctx.wb.tables[String(node.table).toUpperCase()];
    if (!t) return V.name('there is no table called ' + JSON.stringify(node.table));

    var inner = node.spec.slice(1, -1);
    var parts = inner.match(/\[[^\[\]]*\]/g);
    var items = parts ? parts.map(function (s) { return s.slice(1, -1); }) : [inner];

    var thisRow = false, headers = false, totals = false, colName = null;
    items.forEach(function (it) {
      var u = it.toLowerCase();
      if (u === '#this row') thisRow = true;
      else if (u === '#headers') headers = true;
      else if (u === '#totals') totals = true;
      else if (u === '#all' || u === '#data') { /* default */ }
      else colName = it;
    });

    if (colName === null) {
      return new V.XRange(t.sheet, t.r1 + 1, t.c1, t.r2, t.c2);
    }

    var idx = -1;
    for (var i = 0; i < t.headers.length; i++) {
      if (String(t.headers[i]).toLowerCase() === colName.toLowerCase()) { idx = i; break; }
    }
    if (idx < 0) {
      return V.ref('the table ' + t.name + ' has no column called ' + JSON.stringify(colName));
    }
    var c = t.c1 + idx;

    if (headers) return new V.XRange(t.sheet, t.r1, c, t.r1, c);
    if (totals) return new V.XRange(t.sheet, t.r2, c, t.r2, c);
    if (thisRow) {
      // The row the formula is on. Outside the table, Excel gives #VALUE!.
      if (ctx.row < t.r1 + 1 || ctx.row > t.r2) {
        return V.value('[#This Row] only works inside the table ' + t.name);
      }
      return new V.XRange(t.sheet, ctx.row, c, ctx.row, c);
    }
    return new V.XRange(t.sheet, t.r1 + 1, c, t.r2, c);
  }

  /* --------------------------------------------------------- function calls */

  function evalCall(node, ctx) {
    var f = XT.FN[node.name];
    if (!f) {
      return V.name('there is no function called ' + node.name +
        '. Check the spelling, or it may be outside this trainer’s scope.');
    }

    var nargs = node.args.length;
    // A single trailing "missing" is how =SUM() style empty calls arrive.
    if (nargs === 1 && node.args[0].t === 'missing') nargs = 0;

    if (nargs < f.min) {
      return V.value(f.name + ' needs at least ' + f.min + ' argument' +
        (f.min === 1 ? '' : 's') + '; it was given ' + nargs + '. ' + (f.sig || ''));
    }
    if (nargs > f.max) {
      return V.value(f.name + ' takes at most ' + f.max + ' arguments; it was given ' +
        nargs + '. ' + (f.sig || ''));
    }

    // Lazy functions choose for themselves what to evaluate.
    if (f.lazy) return f.impl(node.args, ctx, node);

    var vals = [], i;
    for (i = 0; i < node.args.length; i++) {
      if (f.arrayArgs.indexOf(i) >= 0) {
        // Force array context for this argument position only.
        var saved = ctx.arrayMode;
        ctx.arrayMode = true;
        try { vals.push(ctx.ev(node.args[i])); } finally { ctx.arrayMode = saved; }
      } else {
        vals.push(ctx.ev(node.args[i]));
      }
    }

    // Errors in arguments propagate before the function runs, unless it has
    // asked to see them (IS* functions, and anything error-aware).
    if (!f.seesErrors) {
      for (i = 0; i < vals.length; i++) {
        if (V.isErr(vals[i])) return vals[i];
      }
    }

    vals.nodes = node.args;        // a few functions want the source text
    return f.impl(vals, ctx, node);
  }

})(window.XT = window.XT || {});
