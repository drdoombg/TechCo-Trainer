/* =============================================================================
   Excel Trainer — ui/panels.js
   The learning instruments. These are the reason the trainer exists rather than
   just using Excel: they make the inside of a formula visible.

     STEPPER    Excel's Evaluate Formula, one substitution at a time. The
                innermost finished piece is replaced by its value, repeatedly,
                until the whole formula is one answer. Watching a stack collapse
                from the inside out is the fastest way to understand it.

     STACK      The same formula as a list of layers -- guard, route, join --
                which is how the estate's formulas are actually built and the
                skill the trainer is really after.

     INSPECTOR  What this cell is, what type it is, and for an error, WHICH
                PART produced it. Crucially it also shows errors that an
                IFERROR has swallowed, so a cell quietly showing "" can be made
                to confess (catalog defects D1 and D2).

     TRACER     Precedents and dependents, including across sheets.

     SHAPE      The formula normalised to the catalog's notation, with the
                matching pattern number, so real files become readable.
   ============================================================================= */
(function (XT) {
  'use strict';

  var V = XT.V, R = XT.R, P = XT.P, S = XT.S, F = XT.F;
  var Panels = {};
  XT.Panels = Panels;

  function esc(s) {
    return String(s === null || s === undefined ? '' : s)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  }
  Panels.esc = esc;

  function showVal(v) {
    if (V.isErr(v)) return v.code;
    if (V.isEmpty(v)) return '(empty)';
    if (typeof v === 'string') return v === '' ? '"" (empty text)' : '"' + v + '"';
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (V.isArr(v)) return 'array ' + v.nr + '×' + v.nc;
    if (V.isRange(v)) return P.litText(v);
    return V.numToText(v);
  }
  Panels.showVal = showVal;

  /* =========================================================== the stepper == */

  function isLiteralNode(n) {
    return n.t === 'num' || n.t === 'str' || n.t === 'bool' || n.t === 'err' ||
      n.t === 'lit' || n.t === 'missing';
  }

  function childrenOf(n) {
    switch (n.t) {
      case 'func': return n.args;
      case 'bin': return [n.l, n.r];
      case 'un': case 'post': case 'paren': return [n.x];
      case 'arr': {
        var out = [];
        n.rows.forEach(function (row) { row.forEach(function (c) { out.push(c); }); });
        return out;
      }
      default: return [];
    }
  }

  function replaceChild(parent, oldN, newN) {
    switch (parent.t) {
      case 'func':
        parent.args = parent.args.map(function (a) { return a === oldN ? newN : a; });
        return;
      case 'bin':
        if (parent.l === oldN) parent.l = newN;
        if (parent.r === oldN) parent.r = newN;
        return;
      case 'un': case 'post': case 'paren':
        if (parent.x === oldN) parent.x = newN;
        return;
      case 'arr':
        parent.rows = parent.rows.map(function (row) {
          return row.map(function (c) { return c === oldN ? newN : c; });
        });
        return;
    }
  }

  function clone(n) {
    if (Array.isArray(n)) return n.map(clone);
    if (!n || typeof n !== 'object') return n;
    var o = {};
    Object.keys(n).forEach(function (k) {
      var v = n[k];
      o[k] = (v && typeof v === 'object') ? clone(v) : v;
    });
    return o;
  }

  /* Find the next thing to evaluate: the deepest, leftmost node that is ready,
     meaning all its children are already literals. A reference or a name is
     always ready. That ordering is what makes the display read inside-out. */
  function findReducible(root) {
    var found = null;
    (function walk(n, parent) {
      if (found) return;
      if (isLiteralNode(n)) return;
      var kids = childrenOf(n);
      for (var i = 0; i < kids.length; i++) {
        walk(kids[i], n);
        if (found) return;
      }
      var ready = kids.every(isLiteralNode);
      if (ready || n.t === 'ref' || n.t === 'range' || n.t === 'name' || n.t === 'table') {
        found = { node: n, parent: parent };
      }
    })(root, null);
    return found;
  }

  /* steps(wb, sheet, r, c) -> [{formula, focus, value, note}]
     `focus` is the text of the sub-expression being replaced at that step. */
  Panels.steps = function (wb, sheetName, r, c, maxSteps) {
    var sh = wb.sheet(sheetName);
    var cell = sh && sh.get(r, c);
    if (!cell || !cell.ast) return { ok: false, reason: 'This cell does not contain a formula.' };

    var work = clone(cell.ast);
    var ctx = wb.makeCtx(sh, r, c, cell.cse);
    ctx.hidden = [];
    var out = [{ formula: '=' + P.unparse(work), focus: null, value: null,
      note: 'the formula as written' }];
    var limit = maxSteps || 60;

    for (var i = 0; i < limit; i++) {
      if (isLiteralNode(work)) break;
      var hit = findReducible(work);
      if (!hit) break;
      var focusText = P.unparse(hit.node);
      var val;
      try { val = ctx.ev(hit.node); } catch (e) { val = V.name(e.message); }
      // A bare reference collapses to its contents for display purposes.
      if (hit.node.t !== 'range' || !V.isRange(val)) val = ctx.scalar(val);
      else val = ctx.scalar(val);
      var lit = { t: 'lit', val: val, p: hit.node.p, q: hit.node.q };
      if (hit.parent) replaceChild(hit.parent, hit.node, lit);
      else work = lit;
      out.push({
        formula: '=' + P.unparse(work),
        focus: focusText,
        value: val,
        note: noteFor(hit.node, val)
      });
    }
    return { ok: true, steps: out, hidden: ctx.hidden, cse: cell.cse };
  };

  function noteFor(node, val) {
    if (node.t === 'ref' || node.t === 'range') return 'read ' + P.unparse(node);
    if (node.t === 'name') return 'resolved the name ' + node.name;
    if (node.t === 'table') return 'resolved ' + P.unparse(node);
    if (node.t === 'bin') return 'applied ' + node.op;
    if (node.t === 'un') return 'applied unary ' + node.op;
    if (node.t === 'post') return 'applied %';
    if (node.t === 'paren') return 'brackets';
    if (node.t === 'func') {
      var f = XT.FN[node.name];
      var role = (S.ROLES && S.ROLES[node.name]) || (f && f.doc) || '';
      return 'ran ' + node.name + (role ? ' — ' + role : '');
    }
    return '';
  }

  Panels.renderStepper = function (wb, sheetName, r, c, upto) {
    var res = Panels.steps(wb, sheetName, r, c);
    if (!res.ok) return '<p class="xt-muted">' + esc(res.reason) + '</p>';
    var n = res.steps.length;
    var at = (upto === undefined || upto === null) ? n - 1 : Math.max(0, Math.min(upto, n - 1));
    var h = [];

    h.push('<div class="xt-step-nav">' +
      '<button data-step="0" title="back to the start">⏮</button>' +
      '<button data-step="' + (at - 1) + '"' + (at <= 0 ? ' disabled' : '') + '>◀ back</button>' +
      '<span class="xt-step-count">step ' + at + ' of ' + (n - 1) + '</span>' +
      '<button data-step="' + (at + 1) + '"' + (at >= n - 1 ? ' disabled' : '') + '>next ▶</button>' +
      '<button data-step="' + (n - 1) + '" title="jump to the answer">⏭</button>' +
      '</div>');

    if (res.cse) {
      h.push('<p class="xt-badge-cse">entered with Ctrl+Shift+Enter — ' +
        'arguments are evaluated as arrays</p>');
    }

    res.steps.forEach(function (st, i) {
      if (i > at) return;
      var cls = 'xt-step' + (i === at ? ' xt-step-now' : '');
      h.push('<div class="' + cls + '">');
      h.push('<div class="xt-step-i">' + (i === 0 ? 'start' : i) + '</div>');
      h.push('<div class="xt-step-body">');
      h.push('<code class="xt-step-f">' + esc(st.formula) + '</code>');
      if (st.focus) {
        h.push('<div class="xt-step-note">' + esc(st.note) + '<br>' +
          '<code class="xt-step-sub">' + esc(st.focus) + '</code> → ' +
          '<strong>' + esc(showVal(st.value)) + '</strong>' +
          ' <span class="xt-type">' + esc(V.typeName(st.value)) + '</span></div>');
      } else {
        h.push('<div class="xt-step-note">' + esc(st.note) + '</div>');
      }
      h.push('</div></div>');
    });

    if (res.hidden && res.hidden.length) {
      h.push('<div class="xt-hidden-box"><h4>Errors this formula swallowed</h4>');
      res.hidden.forEach(function (x) {
        h.push('<p><code>' + esc(x.code) + '</code> from <code>' + esc(x.from) + '</code>' +
          (x.detail ? '<br><span class="xt-muted">' + esc(x.detail) + '</span>' : '') + '</p>');
      });
      h.push('<p class="xt-muted">A guard hid these. The cell looks fine. ' +
        'This is catalog defect D1/D2 — and it is why IFNA is often the better guard.</p>');
      h.push('</div>');
    }
    return h.join('');
  };

  /* ============================================================= the stack == */

  Panels.renderStack = function (wb, sheetName, r, c) {
    var sh = wb.sheet(sheetName);
    var cell = sh && sh.get(r, c);
    if (!cell || !cell.input) return '<p class="xt-muted">Select a cell with something in it.</p>';
    if (!cell.ast) {
      return '<p class="xt-muted">This cell holds a value, not a formula. ' +
        'Nothing to take apart.</p>';
    }
    var res = S.layers(cell.input);
    if (!res.ok) return '<p class="xt-muted">' + esc(res.error) + '</p>';

    var h = ['<p class="xt-muted">Read it from the top down to see what it does, ' +
      'from the bottom up to see how it was built.</p>'];
    h.push('<ol class="xt-stack">');
    res.layers.forEach(function (L) {
      h.push('<li style="--d:' + L.depth + '">' +
        '<span class="xt-stack-fn">' + esc(L.fn) + '</span>' +
        '<span class="xt-stack-role">' + esc(L.role) + '</span>' +
        '<code class="xt-stack-text">' + esc(L.text) + '</code>' +
        '</li>');
    });
    h.push('</ol>');

    var used = S.functionsUsed(cell.input);
    if (used.length) {
      h.push('<h4>Functions in this cell</h4><div class="xt-chiprow">');
      used.forEach(function (u) {
        var f = XT.FN[u.fn];
        h.push('<button class="xt-chip" data-fn="' + esc(u.fn) + '" title="' +
          esc(f ? f.doc : '') + '">' + esc(u.fn) + (u.n > 1 ? ' ×' + u.n : '') + '</button>');
      });
      h.push('</div>');
    }
    var depth = S.nestDepth(cell.input, 'IF');
    if (depth >= 3) {
      h.push('<p class="xt-warn">The IFs nest ' + depth + ' deep. ' +
        'Past about three, a two-column lookup table is easier to read and to change ' +
        '— that is the level 12 refactor.</p>');
    }
    return h.join('');
  };

  /* ========================================================== the inspector = */

  /* Find the node that first produced the error a cell is showing. Runs the
     same reduction as the stepper and stops at the first error. */
  Panels.blame = function (wb, sheetName, r, c) {
    var sh = wb.sheet(sheetName);
    var cell = sh && sh.get(r, c);
    if (!cell || !cell.ast) return null;
    var work = clone(cell.ast);
    var ctx = wb.makeCtx(sh, r, c, cell.cse);
    ctx.hidden = [];
    for (var i = 0; i < 120; i++) {
      if (isLiteralNode(work)) break;
      var hit = findReducible(work);
      if (!hit) break;
      var val;
      try { val = ctx.scalar(ctx.ev(hit.node)); } catch (e) { val = V.name(e.message); }
      if (V.isErr(val)) {
        return { text: P.unparse(hit.node), err: val, node: hit.node };
      }
      var lit = { t: 'lit', val: val };
      if (hit.parent) replaceChild(hit.parent, hit.node, lit); else work = lit;
    }
    return null;
  };

  var ERR_PLAIN = {
    '#N/A': 'A lookup looked and did not find it. The value genuinely is not there, ' +
      'or the key does not match the way you think it does (number against text is ' +
      'the usual culprit).',
    '#REF!': 'A reference points at nothing. A column was deleted, a sheet was ' +
      'renamed or removed, or the formula was filled off the edge of the sheet. ' +
      'This one is STRUCTURAL — it will not fix itself when the data changes.',
    '#VALUE!': 'Something was the wrong type: text where a number was needed, or a ' +
      'range where one value was needed.',
    '#DIV/0!': 'A division by zero, or by an empty cell.',
    '#NAME?': 'A name the engine does not recognise — a misspelled function, a ' +
      'deleted defined name, or text that should have been in quotes.',
    '#NUM!': 'A number out of range, or an item asked for that does not exist ' +
      '(SMALL for the 3rd of 2 matches, for instance).',
    '#NULL!': 'Two ranges that were supposed to overlap do not.'
  };

  Panels.renderInspector = function (wb, sheetName, r, c) {
    var d = wb.describe(sheetName, r, c);
    var h = [];
    h.push('<table class="xt-kv">');
    h.push('<tr><td>cell</td><td><code>' + esc(sheetName) + '!' + esc(d.addr) + '</code></td></tr>');
    h.push('<tr><td>shows</td><td>' + (d.display === '' ? '<em class="xt-muted">nothing</em>'
      : '<code>' + esc(d.display) + '</code>') + '</td></tr>');
    h.push('<tr><td>value</td><td><code>' + esc(showVal(d.value)) + '</code></td></tr>');
    h.push('<tr><td>type</td><td><span class="xt-type">' + esc(d.type) + '</span>' +
      (d.type === 'text ""'
        ? ' <span class="xt-muted">&mdash; looks blank, is not. ISBLANK says FALSE.</span>'
        : '') +
      (d.type === 'EMPTY'
        ? ' <span class="xt-muted">&mdash; truly empty. ISBLANK says TRUE.</span>'
        : '') +
      '</td></tr>');
    if (d.fmt) h.push('<tr><td>format</td><td><code>' + esc(d.fmt) + '</code></td></tr>');
    if (d.isFormula) {
      h.push('<tr><td>formula</td><td><code>' + esc(d.input) + '</code>' +
        (d.cse ? ' <span class="xt-type">array</span>' : '') + '</td></tr>');
    }
    h.push('</table>');

    if (typeof d.value === 'number' && d.fmt && F.isDateFormat(d.fmt)) {
      h.push('<p class="xt-muted">That is the number ' + V.numToText(d.value) +
        ' shown as a date. Dates are numbers; the format is only how it looks.</p>');
    }
    if (typeof d.value === 'string' && V.parseNumeric(d.value) !== null) {
      h.push('<p class="xt-warn">This cell holds <strong>text that looks like a ' +
        'number</strong>. It will not add up, and VLOOKUP will not match it against a ' +
        'real number. Check it with <code>=ISNUMBER(' + esc(d.addr) + ')</code>.</p>');
    }
    if (typeof d.value === 'string' && / /.test(d.value)) {
      h.push('<p class="xt-warn">There is a <strong>non-breaking space, CHAR(160)</strong>, ' +
        'in this cell. TRIM will not remove it and VALUE will refuse it. ' +
        'Use <code>SUBSTITUTE(x,CHAR(160),"")</code>. This is catalog defect D7.</p>');
    }
    if (typeof d.value === 'string' && /^[ \t]|[ \t]$/.test(d.value)) {
      h.push('<p class="xt-warn">This cell has leading or trailing whitespace, which is ' +
        'why an exact-match lookup on it fails. <code>TRIM()</code> fixes it.</p>');
    }
    if (d.parseError) {
      h.push('<p class="xt-bad">This formula will not parse: ' + esc(d.parseError) + '</p>');
    }
    if (d.circular) {
      h.push('<p class="xt-bad">Circular reference: this cell depends on itself, so ' +
        'Excel shows 0 and gives up. Find the loop with Trace precedents.</p>');
    }

    if (V.isErr(d.value)) {
      var blame = Panels.blame(wb, sheetName, r, c);
      h.push('<div class="xt-err-box">');
      h.push('<h4>' + esc(d.value.code) + '</h4>');
      h.push('<p>' + esc(ERR_PLAIN[d.value.code] || '') + '</p>');
      if (d.value.detail) h.push('<p class="xt-detail">' + esc(d.value.detail) + '</p>');
      if (blame) {
        h.push('<p>The part that produced it:<br><code>' + esc(blame.text) + '</code></p>');
        if (blame.err.detail && blame.err.detail !== d.value.detail) {
          h.push('<p class="xt-detail">' + esc(blame.err.detail) + '</p>');
        }
      }
      h.push('</div>');
    }

    if (d.hidden && d.hidden.length) {
      h.push('<div class="xt-hidden-box"><h4>Hidden underneath</h4>');
      h.push('<p>This cell looks fine, but a guard swallowed ' + d.hidden.length +
        ' error' + (d.hidden.length === 1 ? '' : 's') + ':</p>');
      d.hidden.forEach(function (x) {
        h.push('<p><code>' + esc(x.code) + '</code> from <code>' + esc(x.from) + '</code>' +
          (x.detail ? '<br><span class="xt-muted">' + esc(x.detail) + '</span>' : '') + '</p>');
      });
      h.push('<p class="xt-muted">If that error is <code>#N/A</code> it may be ' +
        'legitimate — "not found" is a real answer. If it is <code>#REF!</code> or ' +
        '<code>#VALUE!</code> the formula is broken and the guard is hiding it. ' +
        'Swap IFERROR for IFNA and the difference becomes visible.</p>');
      h.push('</div>');
    }
    return h.join('');
  };

  /* ============================================================= the tracer = */

  // Resolve one reference node to a concrete block of cells on a named sheet.
  function resolveRef(wb, node, hostSheet) {
    var sheet = node.sheet || hostSheet;
    var sh = wb.sheet(sheet);
    if (!sh) return null;
    var used = sh.used();
    if (node.t === 'ref') {
      return { sheet: sheet, r1: node.addr.r, r2: node.addr.r, c1: node.addr.c, c2: node.addr.c };
    }
    if (node.kind === 'cells') {
      return { sheet: sheet,
        r1: Math.min(node.a.r, node.b.r), r2: Math.max(node.a.r, node.b.r),
        c1: Math.min(node.a.c, node.b.c), c2: Math.max(node.a.c, node.b.c) };
    }
    if (node.kind === 'col') {
      return { sheet: sheet, r1: 0, r2: Math.max(0, used.maxRow),
        c1: Math.min(node.c1, node.c2), c2: Math.max(node.c1, node.c2), whole: 'column' };
    }
    return { sheet: sheet, c1: 0, c2: Math.max(0, used.maxCol),
      r1: Math.min(node.r1, node.r2), r2: Math.max(node.r1, node.r2), whole: 'row' };
  }
  Panels.resolveRef = resolveRef;

  Panels.precedents = function (wb, sheetName, r, c) {
    var sh = wb.sheet(sheetName);
    var cell = sh && sh.get(r, c);
    if (!cell || !cell.ast) return [];
    var out = [];
    S.references(cell.input).forEach(function (ref) {
      if (ref.node.ext) { out.push({ external: ref.node.ext, text: ref.text }); return; }
      var box = resolveRef(wb, ref.node, sheetName);
      if (box) { box.text = ref.text; out.push(box); }
    });
    return out;
  };

  Panels.dependents = function (wb, sheetName, r, c) {
    var out = [];
    wb.sheets.forEach(function (sh) {
      Object.keys(sh.cells).forEach(function (k) {
        var cell = sh.cells[k];
        if (!cell.ast) return;
        var p = k.split(','), cr = +p[0], cc = +p[1];
        if (sh.name === sheetName && cr === r && cc === c) return;
        var hit = false;
        S.references(cell.input).forEach(function (ref) {
          if (hit || ref.node.ext) return;
          var box = resolveRef(wb, ref.node, sh.name);
          if (!box) return;
          if (box.sheet.toLowerCase() !== sheetName.toLowerCase()) return;
          if (r >= box.r1 && r <= box.r2 && c >= box.c1 && c <= box.c2) hit = true;
        });
        if (hit) out.push({ sheet: sh.name, r: cr, c: cc, input: cell.input });
      });
    });
    return out;
  };

  // Build the mark map the grid paints for the tracer.
  Panels.traceMarks = function (wb, sheetName, r, c, mode) {
    var marks = {};
    if (mode === 'precedents' || mode === 'both') {
      Panels.precedents(wb, sheetName, r, c).forEach(function (b) {
        if (b.external || b.sheet.toLowerCase() !== sheetName.toLowerCase()) return;
        for (var rr = b.r1; rr <= b.r2; rr++) {
          for (var cc = b.c1; cc <= b.c2; cc++) marks[rr + ',' + cc] = 'xt-trace-pre';
        }
      });
    }
    if (mode === 'dependents' || mode === 'both') {
      Panels.dependents(wb, sheetName, r, c).forEach(function (d) {
        if (d.sheet.toLowerCase() !== sheetName.toLowerCase()) return;
        marks[d.r + ',' + d.c] = 'xt-trace-dep';
      });
    }
    return marks;
  };

  Panels.renderTracer = function (wb, sheetName, r, c) {
    var pre = Panels.precedents(wb, sheetName, r, c);
    var dep = Panels.dependents(wb, sheetName, r, c);
    var addr = R.numToCol(c) + (r + 1);
    var h = [];

    h.push('<h4>This cell reads from</h4>');
    if (!pre.length) h.push('<p class="xt-muted">Nothing — it has no references.</p>');
    else {
      h.push('<ul class="xt-tracelist">');
      pre.forEach(function (b) {
        if (b.external) {
          h.push('<li class="xt-bad"><code>' + esc(b.text) + '</code> — another ' +
            'workbook ' + esc(b.external) + ', which is not open. A link like this is ' +
            'one of the ways a file quietly breaks (defect D10).</li>');
          return;
        }
        var cross = b.sheet.toLowerCase() !== sheetName.toLowerCase();
        h.push('<li' + (cross ? ' class="xt-cross"' : '') + '>' +
          '<code>' + esc(b.text) + '</code>' +
          (cross ? ' <span class="xt-type">sheet ' + esc(b.sheet) + '</span>' : '') +
          (b.whole ? ' <span class="xt-type">whole ' + b.whole + ', clipped to ' +
            ((b.r2 - b.r1 + 1) * (b.c2 - b.c1 + 1)) + ' cells</span>' : '') +
          '</li>');
      });
      h.push('</ul>');
    }

    h.push('<h4>Cells that read this one</h4>');
    if (!dep.length) h.push('<p class="xt-muted">Nothing depends on ' + esc(addr) +
      '. Changing it is safe.</p>');
    else {
      h.push('<ul class="xt-tracelist">');
      dep.slice(0, 40).forEach(function (d) {
        var cross = d.sheet.toLowerCase() !== sheetName.toLowerCase();
        h.push('<li' + (cross ? ' class="xt-cross"' : '') + '>' +
          '<a href="#" data-goto="' + esc(d.sheet) + '!' + R.numToCol(d.c) + (d.r + 1) + '">' +
          esc(d.sheet) + '!' + R.numToCol(d.c) + (d.r + 1) + '</a> ' +
          '<code>' + esc(d.input.length > 60 ? d.input.slice(0, 60) + '…' : d.input) +
          '</code></li>');
      });
      if (dep.length > 40) h.push('<li class="xt-muted">…and ' + (dep.length - 40) + ' more</li>');
      h.push('</ul>');
    }
    return h.join('');
  };

  /* ============================================================== the shape = */

  Panels.renderShape = function (wb, sheetName, r, c) {
    var sh = wb.sheet(sheetName);
    var cell = sh && sh.get(r, c);
    if (!cell || !cell.ast) {
      return '<p class="xt-muted">Select a cell containing a formula. This panel ' +
        'strips it down to its shape, which is how the catalog recognised a few ' +
        'hundred patterns inside four million formulas.</p>';
    }
    var m = S.match(cell.input);
    var h = [];
    h.push('<h4>Shape</h4>');
    h.push('<code class="xt-shape">' + esc(m.shape) + '</code>');
    h.push('<p class="xt-muted"><code>SH!</code> a sheet reference &middot; ' +
      '<code>R</code> a cell address &middot; <code>C:C</code> a whole column &middot; ' +
      '<code>N</code> a number &middot; <code>"S"</code> a text literal</p>');

    if (!m.hits.length) {
      h.push('<p class="xt-muted">No catalog pattern matches this shape. That is ' +
        'not a criticism — the twenty patterns cover about 85% of the formula ' +
        'mass, not all of it.</p>');
      return h.join('');
    }

    m.hits.slice(0, 3).forEach(function (hit, i) {
      var p = hit.pattern;
      var mass = XT.data.patternMass ? (100 * p.cells / XT.data.patternMass) : 0;
      h.push('<div class="xt-pat' + (i === 0 ? ' xt-pat-top' : '') + '">');
      h.push('<div class="xt-pat-head"><span class="xt-pid">' + esc(p.id) + '</span> ' +
        esc(p.title) + (hit.score === 100 ? ' <span class="xt-type">exact</span>' : '') + '</div>');
      h.push('<p>' + esc(p.oneLine) + '</p>');
      h.push('<p class="xt-muted">' + p.cells.toLocaleString() + ' cells in ' + p.files +
        ' of 152 workbooks — about ' + mass.toFixed(1) + '% of the estate’s formula mass.</p>');
      if (p.why) h.push('<p><strong>Why it exists.</strong> ' + esc(p.why) + '</p>');
      if (p.watch) h.push('<p class="xt-warn"><strong>Watch out.</strong> ' + esc(p.watch) + '</p>');
      if (p.teaches) {
        h.push('<div class="xt-chiprow">' + p.teaches.map(function (x) {
          return '<span class="xt-chip xt-chip-flat">' + esc(x) + '</span>';
        }).join('') + '</div>');
      }
      h.push('</div>');
    });
    return h.join('');
  };

  /* ====================================================== function reference */

  var GROUP_ORDER = ['logic', 'lookup', 'math', 'text', 'date', 'info', 'unsupported'];
  var GROUP_NAME = {
    logic: 'Logic and guards', lookup: 'Lookup and join', math: 'Aggregate and arithmetic',
    text: 'Text', date: 'Date and time', info: 'Information', unsupported: 'Deliberately not here'
  };

  Panels.renderFunctions = function (filter, openName) {
    var q = String(filter || '').toUpperCase();
    var byGroup = {};
    XT.functionList().forEach(function (f) {
      if (q && f.name.indexOf(q) < 0 && String(f.doc).toUpperCase().indexOf(q) < 0) return;
      (byGroup[f.group] = byGroup[f.group] || []).push(f);
    });
    var h = [];
    GROUP_ORDER.forEach(function (g) {
      var list = byGroup[g];
      if (!list || !list.length) return;
      h.push('<h4>' + esc(GROUP_NAME[g] || g) + '</h4>');
      list.forEach(function (f) {
        var open = openName && openName.toUpperCase() === f.name;
        h.push('<details class="xt-fn' + (f.unsupported ? ' xt-fn-off' : '') + '"' +
          (open ? ' open' : '') + ' data-fn="' + esc(f.name) + '">');
        h.push('<summary><code>' + esc(f.name) + '</code> <span class="xt-muted">' +
          esc(f.doc) + '</span></summary>');
        h.push('<p class="xt-sig"><code>' + esc(f.sig) + '</code></p>');
        if (f.estate) h.push('<p><strong>In your files.</strong> ' + esc(f.estate) + '</p>');
        if (f.pats && f.pats.length) {
          h.push('<div class="xt-chiprow">' + f.pats.map(function (pid) {
            var p = XT.data.pattern && XT.data.pattern(pid);
            return '<button class="xt-chip" data-pat="' + esc(pid) + '">' + esc(pid) +
              (p ? ' ' + esc(p.title) : '') + '</button>';
          }).join('') + '</div>');
        }
        h.push('</details>');
      });
    });
    if (!h.length) h.push('<p class="xt-muted">Nothing matches that.</p>');
    return h.join('');
  };

  /* =================================================== the pattern reference */

  Panels.renderPatterns = function (openId) {
    var h = ['<p class="xt-muted">The twenty moves that make up this estate, ' +
      'in order of how much of it they account for.</p>'];
    (XT.data.patterns || []).forEach(function (p) {
      var mass = XT.data.patternMass ? (100 * p.cells / XT.data.patternMass) : 0;
      h.push('<details class="xt-fn" data-pat="' + esc(p.id) + '"' +
        (openId === p.id ? ' open' : '') + '>');
      h.push('<summary><span class="xt-pid">' + esc(p.id) + '</span> ' + esc(p.title) +
        ' <span class="xt-muted">' + mass.toFixed(1) + '%</span></summary>');
      h.push('<p>' + esc(p.oneLine) + '</p>');
      h.push('<p class="xt-sig"><code>' + esc(p.example) + '</code></p>');
      if (p.alsoExample) h.push('<p class="xt-sig"><code>' + esc(p.alsoExample) + '</code></p>');
      h.push('<p class="xt-muted">' + p.cells.toLocaleString() + ' cells in ' + p.files +
        ' workbooks &middot; level ' + p.level + ' &middot; layer ' + p.layer + '</p>');
      if (p.why) h.push('<p><strong>Why.</strong> ' + esc(p.why) + '</p>');
      if (p.watch) h.push('<p class="xt-warn"><strong>Watch out.</strong> ' + esc(p.watch) + '</p>');
      h.push('</details>');
    });
    return h.join('');
  };

})(window.XT = window.XT || {});
