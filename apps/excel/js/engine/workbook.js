/* =============================================================================
   Excel Trainer — engine/workbook.js
   The workbook, the cells, and the evaluator.

   A cell is  {input, ast, value, fmt, cse}
       input  exactly what was typed, including the leading "="
       ast    the parsed formula, or null for a literal
       value  the last computed value (number | text | boolean | error | EMPTY)
       fmt    a number-format code, display only
       cse    true if it was entered with Ctrl+Shift+Enter (a legacy array formula)

   Recalculation is a memoised depth-first walk rather than an explicit
   topological sort. Asking for a cell's value computes whatever it depends on,
   in whatever order that turns out to be, and a cell already in progress is a
   circular reference. That is fewer moving parts than a dependency graph and,
   at trainer scale (hundreds of rows, not hundreds of thousands), just as fast.

   What Excel does with a circle is show 0 and warn in the status bar. We do the
   same, and name the cells in the warning, because "why is this 0" is a
   question worth being able to answer.
   ============================================================================= */
(function (XT) {
  'use strict';

  var V = XT.V, R = XT.R, P = XT.P, F = XT.F;

  /* ============================================================== Sheet ===== */

  function Sheet(name, opts) {
    opts = opts || {};
    this.name = name;
    this.hidden = !!opts.hidden;
    this.note = opts.note || '';          // shown in the tab tooltip & layer badge
    this.layer = opts.layer || null;      // 1..5, the five-layer role (catalog §3)
    this.cells = Object.create(null);     // "r,c" -> cell
    this.maxRow = -1;
    this.maxCol = -1;
    this.colWidth = Object.create(null);  // c -> px
    this.frozen = opts.frozen || null;    // {r, c} freeze panes
  }
  Sheet.prototype.key = function (r, c) { return r + ',' + c; };

  Sheet.prototype.get = function (r, c) { return this.cells[r + ',' + c] || null; };

  Sheet.prototype.ensure = function (r, c) {
    var k = r + ',' + c, cell = this.cells[k];
    if (!cell) {
      cell = { input: '', ast: null, value: V.EMPTY, fmt: null, cse: false, state: 2 };
      this.cells[k] = cell;
      if (r > this.maxRow) this.maxRow = r;
      if (c > this.maxCol) this.maxCol = c;
    }
    return cell;
  };

  Sheet.prototype.used = function () {
    return { maxRow: Math.max(0, this.maxRow), maxCol: Math.max(0, this.maxCol) };
  };

  Sheet.prototype.recomputeExtent = function () {
    var mr = -1, mc = -1;
    for (var k in this.cells) {
      var cell = this.cells[k];
      if (cell.input === '' && cell.value === V.EMPTY) continue;
      var p = k.split(','), r = +p[0], c = +p[1];
      if (r > mr) mr = r;
      if (c > mc) mc = c;
    }
    this.maxRow = mr; this.maxCol = mc;
  };

  XT.Sheet = Sheet;

  /* ============================================================ Workbook ==== */

  function Workbook(name) {
    this.name = name || 'Workbook';
    this.sheets = [];
    this.byName = Object.create(null);
    this.names = Object.create(null);     // defined names: UPPER -> {text, ast, sheet}
    this.tables = Object.create(null);    // UPPER -> {sheet,r1,c1,r2,c2,headers:[]}
    this.circular = [];                   // ["Sheet1!A1", …] after a recalc
    this.evalCount = 0;
  }

  Workbook.prototype.addSheet = function (name, opts) {
    var s = new Sheet(name, opts);
    this.sheets.push(s);
    this.byName[name.toLowerCase()] = s;
    return s;
  };
  Workbook.prototype.sheet = function (name) {
    return this.byName[String(name).toLowerCase()] || null;
  };
  Workbook.prototype.removeSheet = function (name) {
    var s = this.sheet(name);
    if (!s) return;
    this.sheets.splice(this.sheets.indexOf(s), 1);
    delete this.byName[name.toLowerCase()];
  };

  /* A defined name that parses but points at an error -- the inherited wreckage
     the catalog found, where boom, xxxxx and jgjgjgjhgjhg all resolve to #REF! --
     is recorded as BROKEN, not as a valid name that happens to produce an error.
     That is what lets the Name Manager flag it (defect D8). */
  Workbook.prototype.defineName = function (name, refText, sheetScope) {
    var r = P.parse(refText);
    var broken = !r.ok || (r.ast && r.ast.t === 'err');
    this.names[name.toUpperCase()] =
      { text: refText, ast: r.ok ? r.ast : null, sheet: sheetScope || null, broken: broken };
  };

  Workbook.prototype.defineTable = function (name, sheet, ref, headers) {
    var m = /^(\$?[A-Za-z]{1,3}\$?\d+):(\$?[A-Za-z]{1,3}\$?\d+)$/.exec(ref);
    if (!m) return;
    var a = R.parseAddr(m[1]), b = R.parseAddr(m[2]);
    this.tables[name.toUpperCase()] = {
      name: name, sheet: sheet,
      r1: a.r, c1: a.c, r2: b.r, c2: b.c,
      headers: headers || []
    };
  };

  /* ------------------------------------------------------------ set a cell -- */

  /* setInput("Sheet1", 0, 0, "=IFERROR(VLOOKUP(…),\"\")")
     Stores the text, parses it if it starts with "=", and leaves the value to
     the next recalc. Typing a date or a number is interpreted here, once,
     exactly as Excel does on entry. */
  Workbook.prototype.setInput = function (sheetName, r, c, input, opts) {
    opts = opts || {};
    var sh = this.sheet(sheetName);
    if (!sh) return null;
    var cell = sh.ensure(r, c);
    cell.input = (input === null || input === undefined) ? '' : String(input);
    cell.cse = !!opts.cse;
    if (opts.fmt !== undefined) cell.fmt = opts.fmt;
    cell.parseError = null;
    cell.ast = null;

    if (cell.input.charAt(0) === '=') {
      var res = P.parse(cell.input);
      if (res.ok) cell.ast = res.ast;
      else { cell.parseError = res.error; cell.value = V.name('this formula will not parse: ' + res.error); }
      // Mark it dirty, so the next read computes it even without a full recalc.
      cell.state = 0;
    } else {
      // A literal. Work out what Excel would store.
      var lit = interpretLiteral(cell.input, cell);
      cell.value = lit.value;
      cell.state = 2;
      if (lit.fmt && cell.fmt === null) cell.fmt = lit.fmt;
    }
    if (cell.input !== '') {
      if (r > sh.maxRow) sh.maxRow = r;
      if (c > sh.maxCol) sh.maxCol = c;
    }
    return cell;
  };

  /* Literal entry. The order of these tests is what makes "007" stay text and
     "15/11/2026" become a date -- and what makes a pasted export full of
     numbers-stored-as-text keep its disguise. */
  function interpretLiteral(s, cell) {
    if (s === '') return { value: V.EMPTY };

    // Leading apostrophe forces text, as in Excel.
    if (s.charAt(0) === "'") return { value: s.slice(1) };

    var u = s.toUpperCase();
    if (u === 'TRUE') return { value: true };
    if (u === 'FALSE') return { value: false };
    if (V.ERRCODES.indexOf(u) >= 0) return { value: V.err(u, 'typed in directly') };

    // Dates and times, UK order.
    var dser = XT.D.parseDateText(s);
    if (dser !== null) {
      var hasTime = /[ T]\d{1,2}:\d{2}/.test(s);
      var dateOnly = /^\d{1,2}[/]\d{1,2}[/]\d{2,4}$|^\d{4}-\d{1,2}-\d{1,2}$/.test(s);
      if (dateOnly) return { value: dser, fmt: 'dd/mm/yyyy' };
      if (hasTime) return { value: dser, fmt: 'dd/mm/yyyy hh:mm' };
      return { value: dser, fmt: 'hh:mm:ss' };
    }

    // A plain number. NOTE: a value with a leading zero that changes meaning
    // ("007") is NOT a number to Excel either -- it becomes text. Excel decides
    // by whether the text round-trips, and so do we.
    var n = V.parseNumeric(s);
    if (n !== null && V.numToText(n) === s.replace(/^[ \t]+|[ \t]+$/g, '')) {
      return { value: n };
    }
    if (n !== null && /^[+-]?[\d,]*\d(\.\d+)?$/.test(s.trim()) && !/^0\d/.test(s.trim())) {
      return { value: n };           // 1,234 and 1.50 are numbers
    }
    if (n !== null && /%$/.test(s.trim())) return { value: n, fmt: '0.0%' };

    return { value: s };             // everything else stays text, warts and all
  }
  XT.interpretLiteral = interpretLiteral;

  Workbook.prototype.clearCell = function (sheetName, r, c) {
    var sh = this.sheet(sheetName);
    if (!sh) return;
    delete sh.cells[r + ',' + c];
  };

  /* ============================================================ evaluation == */

  /* RangeView: a lazy window onto a block of cells, with (0,0) at its top-left.
     Functions see only nr/nc/get, so an XArr and a RangeView are
     interchangeable -- which is why COUNTIF works equally on A:A and on a
     computed array. */
  function RangeView(wb, sheet, r1, c1, r2, c2) {
    this.wb = wb; this.sh = sheet;
    this.r1 = r1; this.c1 = c1;
    this.nr = r2 - r1 + 1; this.nc = c2 - c1 + 1;
  }
  RangeView.prototype.get = function (r, c) {
    return this.wb.cellValue(this.sh, this.r1 + r, this.c1 + c);
  };
  RangeView.prototype.at = function (r, c) {
    return this.get(this.nr === 1 ? 0 : r, this.nc === 1 ? 0 : c);
  };
  RangeView.prototype.flat = function () {
    var out = [];
    for (var r = 0; r < this.nr; r++) for (var c = 0; c < this.nc; c++) out.push(this.get(r, c));
    return out;
  };
  XT.RangeView = RangeView;

  // The value of one cell, computing it first if necessary.
  Workbook.prototype.cellValue = function (sheet, r, c) {
    var sh = (typeof sheet === 'string') ? this.sheet(sheet) : sheet;
    if (!sh) return V.ref('that sheet does not exist');
    var cell = sh.get(r, c);
    if (!cell) return V.EMPTY;
    if (!cell.ast) return cell.value;
    return this.evalCell(sh, r, c, cell);
  };

  Workbook.prototype.evalCell = function (sh, r, c, cell) {
    if (cell.state === 2) return cell.value;
    if (cell.state === 1) {
      // A cell that depends on itself. Excel shows 0 and warns; so do we.
      var addr = sh.name + '!' + R.numToCol(c) + (r + 1);
      if (this.circular.indexOf(addr) < 0) this.circular.push(addr);
      cell.circular = true;
      return 0;
    }
    cell.state = 1;
    cell.circular = false;
    var ctx = this.makeCtx(sh, r, c, cell.cse);
    // Collects any error that IFERROR/IFNA swallows while this cell evaluates,
    // so the error inspector can show what a quietly-blank cell is hiding.
    ctx.hidden = [];
    var val;
    try {
      val = XT.evalNode(cell.ast, ctx);
      if (!cell.cse) val = ctx.scalar(val);
      else if (V.isArr(val)) val = val.get(0, 0);   // a CSE cell shows its top-left
      if (V.isRange(val)) val = ctx.scalar(val);
    } catch (e) {
      if (e && e.name === 'FormulaError') val = V.name(e.message);
      else val = V.value('the engine could not evaluate this: ' + (e && e.message));
    }
    /* If this cell turned out to depend on itself, Excel shows 0 and warns in
       the status bar rather than showing whatever the partial pass computed.
       Match that: the half-finished answer is worse than useless. */
    if (cell.circular) val = 0;

    cell.value = val;
    cell.hidden = ctx.hidden;
    cell.state = 2;
    this.evalCount++;
    return val;
  };

  /* Reset every formula cell so the next read recomputes it. Cheap: a trainer
     workbook is small, and a full recalc keeps the mental model simple. */
  Workbook.prototype.recalc = function () {
    this.circular = [];
    this.evalCount = 0;
    var i, k;
    for (i = 0; i < this.sheets.length; i++) {
      var cells = this.sheets[i].cells;
      for (k in cells) if (cells[k].ast) cells[k].state = 0;
    }
    for (i = 0; i < this.sheets.length; i++) {
      var sh = this.sheets[i], cs = sh.cells;
      for (k in cs) {
        if (!cs[k].ast) continue;
        var p = k.split(',');
        this.evalCell(sh, +p[0], +p[1], cs[k]);
      }
    }
    return this;
  };

  /* ------------------------------------------------------ evaluation context */

  Workbook.prototype.makeCtx = function (sh, row, col, arrayMode) {
    var wb = this;
    var ctx = {
      wb: wb,
      sheetObj: sh,
      sheet: sh.name,
      row: row,
      col: col,
      arrayMode: !!arrayMode,
      trace: null          // the stepper hangs a collector here
    };

    // Evaluate a node in this context.
    ctx.ev = function (node) { return XT.evalNode(node, ctx); };

    // Resolve an XRange to a concrete, clipped view.
    ctx.view = function (rng) {
      var target = wb.sheet(rng.sheet || sh.name);
      if (!target) return V.ref('that sheet does not exist');
      var cl = R.clip(rng, target.used());
      return new RangeView(wb, target, cl.r1, cl.c1, cl.r2, cl.c2);
    };

    /* Turn whatever we have into a single scalar value.

       This is Excel's IMPLICIT INTERSECTION. A range used where one value is
       wanted collapses to the cell on the formula's own row (or column). If
       there is no overlap, #VALUE!. Array formulas skip this entirely. */
    ctx.scalar = function (v) {
      if (V.isRange(v)) {
        var view = ctx.view(v);
        if (V.isErr(view)) return view;
        if (view.nr === 1 && view.nc === 1) return view.get(0, 0);
        if (view.nc === 1 && row >= view.r1 && row < view.r1 + view.nr) {
          return view.get(row - view.r1, 0);
        }
        if (view.nr === 1 && col >= view.c1 && col < view.c1 + view.nc) {
          return view.get(0, col - view.c1);
        }
        return V.value('a range of ' + view.nr + '×' + view.nc +
          ' cells was used where one value is needed, and it does not line up with this cell');
      }
      if (V.isArr(v)) return v.nr || v.nc ? v.get(0, 0) : V.EMPTY;
      return v;
    };

    // Coercion shortcuts that also scalarise. Errors pass straight through.
    ctx.num = function (v) { return V.toNum(ctx.scalar(v)); };
    ctx.text = function (v) { return V.toText(ctx.scalar(v)); };
    ctx.bool = function (v) { return V.toBool(ctx.scalar(v)); };
    ctx.int = function (v) {
      var n = ctx.num(v);
      return V.isErr(n) ? n : Math.trunc(n);
    };

    /* A view over anything that can act as a range: an XRange, an array, or a
       single scalar (which becomes 1x1, so =SUMIF('QTR 1'!K16,"<3") works). */
    ctx.asView = function (v) {
      if (V.isRange(v)) return ctx.view(v);
      if (V.isArr(v)) return v;
      if (V.isErr(v)) return v;
      return V.XArr.of(v);
    };

    return ctx;
  };

  /* ---------------------------------------------------------- bulk loading -- */

  /* load({sheets:[{name, hidden, layer, cols:{}, rows:[[...]]}], names:{}, tables:[]})
     Fixtures use this. A cell may be a bare value, or {v, f, fmt, cse} where
     f is a formula string. Strings beginning with "=" are treated as formulas. */
  Workbook.prototype.load = function (def) {
    var wb = this;
    (def.sheets || []).forEach(function (sd) {
      var sh = wb.sheet(sd.name) || wb.addSheet(sd.name, sd);
      if (sd.colWidth) for (var c in sd.colWidth) sh.colWidth[c] = sd.colWidth[c];
      (sd.rows || []).forEach(function (row, ri) {
        if (!row) return;
        row.forEach(function (v, ci) {
          if (v === null || v === undefined) return;
          var r = (sd.startRow || 0) + ri, c = (sd.startCol || 0) + ci;
          if (typeof v === 'object' && !(v instanceof Date)) {
            if (v.f !== undefined) wb.setInput(sd.name, r, c, v.f, { cse: v.cse, fmt: v.fmt });
            else {
              var cell = wb.setInput(sd.name, r, c, '', { fmt: v.fmt });
              cell.value = v.v === undefined ? V.EMPTY : v.v;
              cell.input = literalInputFor(v.v, v.fmt);
            }
          } else if (typeof v === 'string' && v.charAt(0) === '=') {
            wb.setInput(sd.name, r, c, v);
          } else {
            var cl = wb.setInput(sd.name, r, c, '');
            cl.value = v;
            cl.input = literalInputFor(v, null);
          }
        });
      });
      sh.recomputeExtent();
    });
    if (def.names) for (var nm in def.names) wb.defineName(nm, def.names[nm]);
    (def.tables || []).forEach(function (t) {
      wb.defineTable(t.name, t.sheet, t.ref, t.headers);
    });
    return wb;
  };

  // What the formula bar should show for a loaded literal.
  function literalInputFor(v, fmt) {
    if (v === undefined || v === null || v === V.EMPTY) return '';
    if (typeof v === 'number' && fmt && F.isDateFormat(fmt)) return F.format(v, fmt);
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    return String(v);
  }

  /* --------------------------------------------------------- introspection -- */

  /* Display text for a cell, honouring its number format.

     ONE SPECIAL CASE, and it is the reason half the estate exists:
     a FORMULA that reads an empty cell shows **0**, not nothing.

         A1 is empty.   =A1   shows 0.

     The value is still EMPTY -- it concatenates as "", it is ignored by SUM,
     and =A1="" is TRUE -- but on screen it is a zero. Which is exactly why
     472,402 cells in the real files are wrapped in IF(src>0,src,"") and
     428,034 more in IF(anchor="","",src): a bare mirror of a half-empty export
     fills the sheet with zeros that print on labels and look like data.
     Pattern P2 and P3 are the fix for this one display rule. */
  Workbook.prototype.display = function (sheet, r, c) {
    var sh = (typeof sheet === 'string') ? this.sheet(sheet) : sheet;
    if (!sh) return '';
    var cell = sh.get(r, c);
    if (!cell) return '';
    var v = this.cellValue(sh, r, c);
    if (cell.ast && V.isEmpty(v)) return F.format(0, cell.fmt);
    return F.format(v, cell.fmt);
  };

  // Everything the inspector panel wants to know about one cell.
  Workbook.prototype.describe = function (sheetName, r, c) {
    var sh = this.sheet(sheetName);
    var cell = sh && sh.get(r, c);
    var addr = R.numToCol(c) + (r + 1);
    if (!cell) return { addr: addr, sheet: sheetName, empty: true, value: V.EMPTY, type: 'EMPTY' };
    var v = this.cellValue(sh, r, c);
    return {
      addr: addr, sheet: sheetName, cell: cell,
      input: cell.input, value: v, type: V.typeName(v),
      display: F.format(v, cell.fmt), fmt: cell.fmt,
      isFormula: !!cell.ast, cse: cell.cse,
      parseError: cell.parseError || null,
      circular: !!cell.circular,
      hidden: cell.hidden || []
    };
  };

  XT.Workbook = Workbook;
  XT.newWorkbook = function (name) { return new Workbook(name); };

})(window.XT = window.XT || {});
