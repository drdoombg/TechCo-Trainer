/* =============================================================================
   Excel Trainer — trainer/grader.js
   Two checks, and BOTH have to pass (catalog §12.7).

     VALUES  the target cells hold the right answers, graded strictly: "",
             EMPTY and 0 are three different results and a near miss is a miss.

     SHAPE   the formula is built the right way. Hard-coding the answer fails.
             So does a formula that happens to work in the first cell and breaks
             when filled, which is the most common real mistake and the one a
             value-only grader would wave through.

   Expected values are not written out by hand. Each exercise carries a
   `solution`, and the grader rebuilds the scenario from the same seed, applies
   that solution, and reads the answers off it. The fixtures are generated, so
   this is the only way the two can never drift apart -- and it means an exercise
   can be reseeded for fresh practice data without re-authoring anything.

   Feedback is specific or it is useless. "Wrong" teaches nothing;
   "your COUNTIFS works in B5 but breaks when filled right, because column C is
   not locked" teaches the thing.
   ============================================================================= */
(function (XT) {
  'use strict';

  var V = XT.V, R = XT.R, P = XT.P, S = XT.S;
  var G = {};
  XT.Grader = G;

  /* --------------------------------------------------------------- helpers -- */

  function addr(cell) { return cell; }
  function rc(a) { return R.rc(a); }

  // Expand "A2:A20" or "A2" into a list of {r,c,a}.
  function cellList(spec) {
    var m = /^([A-Za-z]{1,3}\d{1,7}):([A-Za-z]{1,3}\d{1,7})$/.exec(spec);
    if (!m) {
      var one = rc(spec);
      return one ? [{ r: one.r, c: one.c, a: spec.toUpperCase() }] : [];
    }
    var a = rc(m[1]), b = rc(m[2]), out = [];
    for (var r = Math.min(a.r, b.r); r <= Math.max(a.r, b.r); r++) {
      for (var c = Math.min(a.c, b.c); c <= Math.max(a.c, b.c); c++) {
        out.push({ r: r, c: c, a: R.numToCol(c) + (r + 1) });
      }
    }
    return out;
  }
  G.cellList = cellList;

  function describeValue(v) {
    if (V.isErr(v)) return v.code;
    if (V.isEmpty(v)) return 'nothing (a truly empty cell)';
    if (typeof v === 'string') return v === '' ? '"" (empty text)' : '"' + v + '"';
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    return V.numToText(v);
  }
  G.describeValue = describeValue;

  // Strict equality, including the EMPTY / "" / 0 distinction.
  function sameValue(a, b) {
    if (V.isErr(a) || V.isErr(b)) return V.isErr(a) && V.isErr(b) && a.code === b.code;
    if (V.isEmpty(a) || V.isEmpty(b)) return V.isEmpty(a) && V.isEmpty(b);
    if (typeof a !== typeof b) return false;
    if (typeof a === 'number') return Math.abs(a - b) < 1e-9;
    return a === b;
  }
  G.sameValue = sameValue;

  /* --------------------------------------------------- the reference solution */

  /* Rebuild the exercise's workbook from scratch and apply its model answer.
     `solution` is a map of sheet!cell (or just cell, on the exercise's sheet)
     to formula text; `fill` entries say how far that formula should run. */
  G.referenceWorkbook = function (ex) {
    var wb = G.exerciseWorkbook(ex);
    if (!wb) return null;
    G.applyFormulas(wb, ex.solution, ex.sheet);
    wb.recalc();
    return wb;
  };

  /* The workbook an exercise STARTS from: the scenario, plus any `setup`
     formulas that stand in for earlier layers. Setup is what keeps every
     exercise independent -- you can attempt level 6 without having done
     level 1, because the mirror it needs is already there. */
  G.exerciseWorkbook = function (ex) {
    var wb = XT.data.makeWorkbook(ex.scenario, { seed: ex.seed, today: ex.today });
    if (!wb) return null;
    G.applyFormulas(wb, ex.setup, ex.sheet);
    wb.recalc();
    return wb;
  };

  G.applySolution = function (wb, ex) { G.applyFormulas(wb, ex.solution, ex.sheet); };

  G.applyFormulas = function (wb, sol, defaultSheet) {
    sol = sol || {};
    var ex = { sheet: defaultSheet };
    Object.keys(sol).forEach(function (key) {
      var spec = sol[key];
      var parts = splitTarget(key, ex.sheet);
      var text = typeof spec === 'string' ? spec : spec.f;
      var cse = typeof spec === 'object' && spec.cse;
      var cells = cellList(parts.range);
      if (!cells.length) return;
      var first = cells[0];
      cells.forEach(function (cl) {
        var t = text;
        if (cl !== first && text.charAt(0) === '=') {
          t = P.translateText(text, cl.r - first.r, cl.c - first.c);
        }
        wb.setInput(parts.sheet, cl.r, cl.c, t, { cse: cse });
      });
      var sh = wb.sheet(parts.sheet);
      if (sh) sh.recomputeExtent();
    });
  };

  function splitTarget(key, defaultSheet) {
    var i = key.lastIndexOf('!');
    if (i < 0) return { sheet: defaultSheet, range: key };
    return { sheet: key.slice(0, i).replace(/^'|'$/g, ''), range: key.slice(i + 1) };
  }
  G.splitTarget = splitTarget;

  /* One entry point. An exercise's `mode` picks which kind of grading it gets:
       'cells'   (the default) values + shape on the target cells
       'layers'  the structural five-layer check
       'defects' the find-and-fix check for the debugging level            */
  G.run = function (wb, ex) {
    if (ex.mode === 'layers') return G.gradeLayers(wb, ex);
    if (ex.mode === 'defects') return G.gradeDefects(wb, ex);
    return G.grade(wb, ex);
  };

  /* ================================================================= grading */

  /* grade(wb, ex) -> {
         pass, checks: [{id, ok, title, detail}], score
     } */
  G.grade = function (wb, ex) {
    var checks = [];
    var ref = G.referenceWorkbook(ex);
    if (!ref) {
      return { pass: false, checks: [{ id: 'setup', ok: false, title: 'Exercise not set up',
        detail: 'The scenario for this exercise could not be rebuilt.' }] };
    }

    var targets = [];
    (ex.targets || []).forEach(function (t) {
      var parts = splitTarget(t.range || t, ex.sheet);
      cellList(parts.range).forEach(function (cl) {
        targets.push({ sheet: parts.sheet, r: cl.r, c: cl.c, a: cl.a, spec: t });
      });
    });

    /* ---- 1. values -------------------------------------------------------- */
    var wrong = [], blank = 0;
    targets.forEach(function (t) {
      var got = wb.cellValue(t.sheet, t.r, t.c);
      var want = ref.cellValue(t.sheet, t.r, t.c);
      if (V.isEmpty(got) && !V.isEmpty(want)) { blank++; return; }
      if (!sameValue(got, want)) wrong.push({ t: t, got: got, want: want });
    });

    if (blank === targets.length && targets.length) {
      checks.push({ id: 'values', ok: false, title: 'Nothing entered yet',
        detail: 'Put a formula in ' + targets[0].sheet + '!' + targets[0].a +
          (targets.length > 1 ? ' and fill it down to ' + targets[targets.length - 1].a : '') + '.' });
    } else if (blank) {
      checks.push({ id: 'values', ok: false, title: blank + ' cell' +
        (blank === 1 ? ' is' : 's are') + ' still empty',
        detail: 'The formula has to run the whole way down. Drag the fill handle, or ' +
          'select the block and press Ctrl+D.' });
    } else if (wrong.length) {
      var w = wrong[0];
      var detail = w.t.sheet + '!' + w.t.a + ' shows ' + describeValue(w.got) +
        ' but should show ' + describeValue(w.want) + '.';
      detail += typeHint(w.got, w.want);
      if (wrong.length > 1) {
        detail += ' ' + (wrong.length - 1) + ' other cell' +
          (wrong.length === 2 ? '' : 's') + ' disagree' + (wrong.length === 2 ? 's' : '') + ' too';
        // If only the later rows are wrong, the fill is the problem, not the formula.
        var firstWrongIdx = targets.indexOf(w.t);
        if (firstWrongIdx > 0) {
          detail += ' — and the first ' + firstWrongIdx + ' are right, which almost ' +
            'always means a reference that should have been anchored with $ (or one that ' +
            'should not have been).';
        }
        detail += '.';
      }
      checks.push({ id: 'values', ok: false, title: 'The answers do not match', detail: detail });
    } else {
      checks.push({ id: 'values', ok: true, title: 'Values correct',
        detail: targets.length + ' cell' + (targets.length === 1 ? '' : 's') + ' check out, ' +
          'including the blank ones.' });
    }

    /* ---- 2. shape --------------------------------------------------------- */
    var rules = ex.shapeRules || {};
    var inputs = targets.map(function (t) {
      var sh = wb.sheet(t.sheet);
      var cell = sh && sh.get(t.r, t.c);
      return { t: t, input: cell ? cell.input : '', cell: cell };
    }).filter(function (x) { return x.input !== ''; });

    if (inputs.length) {
      G.shapeChecks(rules, inputs, targets, ex, ref).forEach(function (c) { checks.push(c); });
    }

    var pass = checks.every(function (c) { return c.ok; }) && targets.length > 0 && inputs.length > 0;
    return { pass: pass, checks: checks, ref: ref };
  };

  /* --------------------------------------------------------- the shape rules */

  G.shapeChecks = function (rules, inputs, targets, ex, ref) {
    var out = [];
    var first = inputs[0];
    var fns = {};
    inputs.forEach(function (x) {
      S.functionsUsed(x.input).forEach(function (f) { fns[f.fn] = true; });
    });

    // --- every target must actually be a formula ---------------------------
    if (rules.mustBeFormula !== false) {
      var literals = inputs.filter(function (x) { return x.input.charAt(0) !== '='; });
      if (literals.length) {
        out.push({ id: 'formula', ok: false, title: 'That is a typed-in value, not a formula',
          detail: literals[0].t.a + ' contains ' + JSON.stringify(literals[0].input) +
            ' typed in directly. The point is to compute it, so that it still works ' +
            'tomorrow when the data changes. Start with = .' });
        return out;
      }
    }

    // --- no hard-coded answers --------------------------------------------
    if (rules.mustNotHardcode) {
      var caught = null;
      inputs.forEach(function (x) {
        if (caught) return;
        var want = ref.cellValue(x.t.sheet, x.t.r, x.t.c);
        S.literals(x.input).forEach(function (lit) {
          if (caught) return;
          if (sameValue(lit.v, want) && !(typeof want === 'string' && want === '')) {
            caught = { x: x, lit: lit, want: want };
          }
        });
      });
      if (caught) {
        out.push({ id: 'hardcode', ok: false, title: 'The answer is typed into the formula',
          detail: caught.x.t.a + ' contains the literal ' + describeValue(caught.lit.v) +
            ', which happens to be the right answer for that row. It will be wrong for ' +
            'the next row, and wrong next week. Look the value up instead.' });
      }
    }

    // --- required / forbidden functions -----------------------------------
    (rules.mustUse || []).forEach(function (name) {
      var alts = String(name).split('|');
      var hit = alts.some(function (a) { return fns[a.toUpperCase()]; });
      if (!hit) {
        out.push({ id: 'use-' + name, ok: false,
          title: 'Use ' + alts.join(' or ') + ' for this one',
          detail: alts.length > 1
            ? 'Either would do. Both are in the real files; ' + alts[0] +
              ' is the one this exercise is about.'
            : XT.FN[alts[0].toUpperCase()]
              ? XT.FN[alts[0].toUpperCase()].sig
              : '' });
      }
    });
    (rules.mustNotUse || []).forEach(function (name) {
      if (fns[name.toUpperCase()]) {
        out.push({ id: 'avoid-' + name, ok: false, title: 'Not with ' + name + ', this time',
          detail: rules.mustNotUseWhy || 'This exercise is about doing it a different way.' });
      }
    });

    // --- the fill has to be a real fill -----------------------------------
    /* The most valuable structural check there is. Every target after the first
       must be exactly what you get by filling the first one. If it is not, the
       learner typed each cell by hand, or the anchoring is inconsistent. */
    if (rules.mustFill !== false && inputs.length > 1) {
      var bad = null;
      inputs.forEach(function (x) {
        if (bad || x === first) return;
        var expect = P.translateText(first.input, x.t.r - first.t.r, x.t.c - first.t.c);
        if (norm(expect) !== norm(x.input)) bad = { x: x, expect: expect };
      });
      if (bad) {
        out.push({ id: 'fill', ok: false, title: 'These were not filled from one formula',
          detail: bad.x.t.a + ' contains\n    ' + bad.x.input +
            '\nbut filling ' + first.t.a + ' down to it gives\n    ' + bad.expect +
            '\nOne formula, filled, is the whole idea: it is how 3 million cells in your ' +
            'real files were made, and it is the only version you can maintain.' });
      }
    }

    // --- anchoring --------------------------------------------------------
    /* anchorPattern is a regex that at least one reference in the formula must
       match, e.g. "^\\$[A-Z]+:\\$[A-Z]+$" for a locked whole-column block. */
    if (rules.anchorPattern) {
      var re = new RegExp(rules.anchorPattern);
      var refs = S.references(first.input).map(function (x) {
        return x.text.replace(/^(\[\d+\])?('[^']*'|[A-Za-z_][A-Za-z0-9_.]*)!/, '');
      });
      if (!refs.some(function (t) { return re.test(t); })) {
        out.push({ id: 'anchor', ok: false, title: 'Check your $ signs',
          detail: rules.anchorWhy ||
            ('None of the references in ' + first.t.a + ' are anchored the way this needs. ' +
             'References found: ' + (refs.join(', ') || 'none') + '. ' +
             'Put the cursor on a reference while editing and press F4 to cycle the $.') });
      }
    }

    // --- must point at a particular sheet ---------------------------------
    (rules.mustReference || []).forEach(function (sheetName) {
      var found = false;
      S.references(first.input).forEach(function (x) {
        var sn = x.node.sheet;
        if (sn && sn.toLowerCase() === sheetName.toLowerCase()) found = true;
      });
      if (!found) {
        out.push({ id: 'ref-' + sheetName, ok: false,
          title: 'This should read from ‘' + sheetName + '’',
          detail: 'The formula in ' + first.t.a + ' does not reference that sheet. ' +
            'Each layer should point at the one before it.' });
      }
    });

    // --- nesting depth ----------------------------------------------------
    if (rules.maxNestDepth) {
      var d = S.nestDepth(first.input, 'IF');
      if (d > rules.maxNestDepth) {
        out.push({ id: 'depth', ok: false, title: 'That is deeper than it needs to be',
          detail: 'The IFs nest ' + d + ' deep; this can be done in ' +
            rules.maxNestDepth + '. ' + (rules.depthWhy || '') });
      }
    }

    // --- array entry ------------------------------------------------------
    if (rules.mustBeArray) {
      var notArr = inputs.filter(function (x) { return !x.cell || !x.cell.cse; });
      if (notArr.length) {
        out.push({ id: 'cse', ok: false, title: 'This one needs Ctrl+Shift+Enter',
          detail: notArr[0].t.a + ' was entered as an ordinary formula. Without array ' +
            'entry the ranges collapse to a single row (implicit intersection) and the ' +
            'answer is quietly wrong. Edit the cell and press Ctrl+Shift+Enter.' });
      }
    }

    // --- pattern recognition ----------------------------------------------
    if (rules.requirePattern) {
      var m = S.match(first.input);
      var ids = m.hits.map(function (h) { return h.id; });
      if (ids.indexOf(rules.requirePattern) < 0) {
        var p = XT.data.pattern(rules.requirePattern);
        out.push({ id: 'pattern', ok: false,
          title: 'Not quite the shape of ' + rules.requirePattern,
          detail: 'Your formula normalises to\n    ' + m.clean +
            '\nand ' + rules.requirePattern + ' (' + (p ? p.title : '') + ') looks like\n    ' +
            (p && p.shapes ? p.shapes[0] : '') +
            '\nThe answer may still be right — this check is about learning to ' +
            'recognise the move.' });
      }
    }

    if (!out.length) {
      out.push({ id: 'shape', ok: true, title: 'Built the right way',
        detail: 'One formula, correctly anchored, filled through the block.' });
    }
    return out;
  };

  function norm(s) { return String(s).replace(/\s+/g, '').toUpperCase(); }

  /* ------------------------------------------------- the type-mismatch hint */

  /* When a value is wrong, say WHY in the vocabulary of this dialect. These
     three mistakes account for most of what goes wrong in the real files. */
  function typeHint(got, want) {
    if (V.isErr(got) && !V.isErr(want)) {
      if (got.code === '#N/A') {
        return ' A #N/A means the lookup looked and did not find it. ' +
          'Check the key: a number will never match text that looks like the same number.';
      }
      if (got.code === '#REF!') {
        return ' A #REF! is structural — a reference pointing at nothing. ' +
          'Check the column index against the width of the block you gave it.';
      }
      if (got.code === '#VALUE!') {
        return ' A #VALUE! usually means text where a number was wanted. ' +
          'Try VALUE() on the slice, or TRIM() if there is stray whitespace.';
      }
      return ' ' + (got.detail || '');
    }
    if (typeof got === 'string' && got === '' && typeof want === 'number') {
      return ' Your guard is firing when it should not. Check which cell the IF is testing.';
    }
    if (V.isEmpty(got) && typeof want === 'string' && want === '') {
      return ' Close: the cell is truly empty, but it should hold "" — empty TEXT. ' +
        'They look identical on screen and behave differently (ISBLANK tells them apart).';
    }
    if (typeof got === 'string' && typeof want === 'number') {
      return ' You have text where a number is wanted. If it came out of a slice, wrap it ' +
        'in VALUE().';
    }
    if (typeof got === 'number' && typeof want === 'string') {
      return ' You have a number where text is wanted.';
    }
    if (typeof got === 'string' && typeof want === 'string' &&
      got.replace(/\s+$/, '') === want.replace(/\s+$/, '')) {
      return ' The text matches except for whitespace. Something has a trailing space ' +
        '— TRIM() it.';
    }
    return '';
  }

  /* ----------------------------------------------- structural layer grading */

  /* Level 17 is graded on SHAPE OF WORKBOOK, not on any one cell: did the
     learner keep the five layers on separate sheets, each pointing only at the
     one before it? (catalog §3) */
  G.gradeLayers = function (wb, ex) {
    var order = ex.layerOrder || [];
    var checks = [];
    var index = {};
    order.forEach(function (name, i) { index[name.toLowerCase()] = i; });

    order.forEach(function (name, i) {
      var sh = wb.sheet(name);
      if (!sh) {
        checks.push({ id: 'layer-' + i, ok: false, title: 'Missing layer: ' + name,
          detail: 'The five-layer build needs a sheet called ' + name + '.' });
        return;
      }
      var offenders = [];
      var empty = true;
      Object.keys(sh.cells).forEach(function (k) {
        var cell = sh.cells[k];
        if (cell.input !== '') empty = false;
        if (!cell.ast) return;
        S.references(cell.input).forEach(function (ref) {
          var target = (ref.node.sheet || name).toLowerCase();
          if (!(target in index)) return;               // reference sheets are fine
          var j = index[target];
          if (j === i) return;                          // within its own layer is fine
          if (j === i - 1) return;                      // the layer before: correct
          var p = k.split(',');
          offenders.push({
            a: R.numToCol(+p[1]) + (+p[0] + 1),
            to: ref.node.sheet, text: ref.text, jump: j - i
          });
        });
      });
      if (empty && i > 0) {
        checks.push({ id: 'layer-' + i, ok: false, title: name + ' is empty',
          detail: 'Every layer has to do its own job. ' + name + ' has nothing in it.' });
      } else if (offenders.length) {
        var o = offenders[0];
        checks.push({ id: 'layer-' + i, ok: false, title: name + ' reaches past its layer',
          detail: name + '!' + o.a + ' reads ' + o.text + ' from ‘' + o.to +
            '’, which is ' + (o.jump < 0 ? Math.abs(o.jump) + ' layer(s) further back'
              : o.jump + ' layer(s) further forward') + '. Each layer should point only ' +
            'at the one before it. That is what makes the chain traceable — and what ' +
            'stops a re-paste breaking something three sheets away.' });
      } else {
        checks.push({ id: 'layer-' + i, ok: true, title: name + ' is clean',
          detail: 'Reads only from the layer before it.' });
      }
    });

    return { pass: checks.every(function (c) { return c.ok; }), checks: checks };
  };

  /* -------------------------------------------------------- defect grading -- */

  /* Level 18: the workbook arrives broken and the task is to fix it. A fix
     counts when the named cell no longer shows the defect's signature. */
  G.gradeDefects = function (wb, ex) {
    var checks = [];
    (ex.defects || []).forEach(function (d) {
      var parts = splitTarget(d.cell, ex.sheet);
      var cl = cellList(parts.range)[0];
      var sh = wb.sheet(parts.sheet);
      var cell = sh && sh.get(cl.r, cl.c);
      var val = sh ? wb.cellValue(parts.sheet, cl.r, cl.c) : V.EMPTY;
      var ok = true, why = '';

      if (d.test === 'no-error') {
        ok = !V.isErr(val);
        why = ok ? '' : 'still showing ' + val.code;
      } else if (d.test === 'no-hidden-error') {
        ok = !(cell && cell.hidden && cell.hidden.length);
        why = ok ? '' : 'still hiding ' + cell.hidden[0].code + ' behind a guard';
      } else if (d.test === 'not-false') {
        ok = val !== false && val !== 0;
        why = ok ? '' : 'still returns ' + describeValue(val);
      } else if (d.test === 'is-number') {
        ok = typeof val === 'number';
        why = ok ? '' : 'still ' + V.typeName(val);
      } else if (d.test === 'formula-lacks') {
        ok = !cell || !cell.input || cell.input.toUpperCase().indexOf(d.text.toUpperCase()) < 0;
        why = ok ? '' : 'the formula still contains ' + d.text;
      } else if (d.test === 'formula-has') {
        ok = !!(cell && cell.input && cell.input.toUpperCase().indexOf(d.text.toUpperCase()) >= 0);
        why = ok ? '' : 'the formula does not contain ' + d.text + ' yet';
      }

      checks.push({ id: d.id, ok: ok,
        title: (ok ? 'Fixed: ' : 'Still broken: ') + d.id + ' — ' + d.title,
        detail: ok ? (d.fixedNote || '') : (parts.sheet + '!' + cl.a + ' ' + why + '. ' + (d.hint || '')) });
    });
    return { pass: checks.every(function (c) { return c.ok; }), checks: checks };
  };

})(window.XT = window.XT || {});
