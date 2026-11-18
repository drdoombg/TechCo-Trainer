/* =============================================================================
   Excel Trainer — engine/functions/math.js
   Aggregates, the *IF/*IFS family, and the handful of arithmetic functions the
   source estate actually uses.

   Two rules run through all of it, and they are the ones people get wrong:

   RULE 1 -- A RANGE AND A DIRECT ARGUMENT ARE TREATED DIFFERENTLY.
       SUM("2")        = 2    a direct text argument is coerced
       SUM(A1) where A1 holds the text "2"  = 0    text inside a range is ignored
       COUNT("2")      = 1    direct
       COUNT(A1)       = 0    text in a range is not a number
   This is not a quirk to work around; it is why a column of numbers-stored-as-
   text silently totals zero, which is one of the defects the trainer teaches.

   RULE 2 -- COUNTIF-STYLE CRITERIA ARE LOOSER THAN THE = OPERATOR.
       A1 holds the text "5".
       =A1=5            FALSE
       =COUNTIF(A1,5)   1
   Both are correct Excel. criteria.js owns that comparison.

   COUNTIFS is the most WIDELY used aggregator in the estate -- 39 of 152
   workbooks -- because every dashboard number is a count of paste rows matching
   a test code, a site and a week (pattern P13).
   ============================================================================= */
(function (XT) {
  'use strict';

  var V = XT.V, C = XT.C, def = XT.defineFunction;

  function isBlock(v) { return V.isRange(v) || V.isArr(v); }

  /* ------------------------------------------------------------- collectors */

  /* Walk the argument list, separating "came from a range" from "written
     directly", because Rule 1 above depends on the difference.
     cb(value, fromRange) is called for every scalar found.
     Returns an error if one is met inside a range (Excel propagates those). */
  function walk(args, ctx, cb) {
    for (var i = 0; i < args.length; i++) {
      var a = args[i];
      if (isBlock(a)) {
        var view = ctx.asView(a);
        if (V.isErr(view)) return view;
        for (var r = 0; r < view.nr; r++) {
          for (var c = 0; c < view.nc; c++) {
            var v = view.get(r, c);
            if (V.isErr(v)) return v;
            cb(v, true);
          }
        }
      } else {
        if (V.isErr(a)) return a;
        cb(a, false);
      }
    }
    return null;
  }

  /* ------------------------------------------------------------------- SUM --- */

  def({
    name: 'SUM', min: 1, max: 255, group: 'math',
    sig: 'SUM(number1, [number2], …)',
    doc: 'Adds up numbers. Text and blanks inside a range are ignored.',
    estate: '23,631 cells -- modest, because this estate reorganises rather than ' +
      'calculates. Mostly row totals, SUM(D17:G17).',
    pats: ['P19'],
    impl: function (args, ctx) {
      var total = 0, bad = null;
      var e = walk(args, ctx, function (v, fromRange) {
        if (typeof v === 'number') { total += v; return; }
        if (fromRange) return;                      // text/bool/blank in a range: skip
        if (typeof v === 'boolean') { total += v ? 1 : 0; return; }
        if (V.isEmpty(v)) return;
        var n = V.toNum(v);
        if (V.isErr(n)) { bad = n; return; }
        total += n;
      });
      if (e) return e;
      if (bad) return bad;
      return total;
    }
  });

  def({
    name: 'COUNT', min: 1, max: 255, group: 'math',
    sig: 'COUNT(value1, [value2], …)',
    doc: 'Counts how many values are numbers. Text in a range does not count.',
    estate: '4,119 cells, usually COUNT(L16:L54) to find how many results exist ' +
      'before dividing by it.',
    impl: function (args, ctx) {
      var n = 0;
      var e = walk(args, ctx, function (v, fromRange) {
        if (typeof v === 'number') { n++; return; }
        if (fromRange) return;
        if (typeof v === 'boolean') { n++; return; }
        if (typeof v === 'string' && V.parseNumeric(v) !== null) n++;
      });
      return e || n;
    }
  });

  def({
    name: 'COUNTA', min: 1, max: 255, group: 'math',
    sig: 'COUNTA(value1, [value2], …)',
    doc: 'Counts cells that are not truly empty. A cell holding "" DOES count.',
    estate: '176 cells, and always as a data-height probe: ' +
      'IF(ROW()<COUNTA($B:$B)-49,"",…) builds a rolling last-50-rows window.',
    pats: ['P19'],
    seesErrors: true,
    impl: function (args, ctx) {
      var n = 0;
      walk2(args, ctx, function (v) { if (!V.isEmpty(v)) n++; });
      return n;
    }
  });

  def({
    name: 'COUNTBLANK', min: 1, max: 1, group: 'math',
    sig: 'COUNTBLANK(range)',
    doc: 'Counts cells that are empty OR hold "". Both count as blank here.',
    estate: '74 cells. The interesting part is what it reveals: a cell holding ' +
      '"" is blank to COUNTBLANK but not to COUNTA, which is the clearest way to ' +
      'see the difference between EMPTY and empty text.',
    pats: ['P3'],
    seesErrors: true,
    impl: function (args, ctx) {
      var view = ctx.asView(args[0]);
      if (V.isErr(view)) return view;
      var n = 0, all = view.flat();
      for (var i = 0; i < all.length; i++) {
        if (V.isEmpty(all[i]) || all[i] === '') n++;
      }
      return n;
    }
  });

  // Like walk, but does not stop on errors (for COUNTA / COUNTBLANK).
  function walk2(args, ctx, cb) {
    for (var i = 0; i < args.length; i++) {
      var a = args[i];
      if (isBlock(a)) {
        var view = ctx.asView(a);
        if (V.isErr(view)) { cb(view); continue; }
        var all = view.flat();
        for (var k = 0; k < all.length; k++) cb(all[k]);
      } else cb(a);
    }
  }

  def({
    name: 'AVERAGE', min: 1, max: 255, group: 'math',
    sig: 'AVERAGE(number1, [number2], …)',
    doc: 'The mean of the numbers. #DIV/0! if there are none.',
    estate: 'Only 48 cells -- the estate prefers AVERAGEIFS, or the ' +
      'flag-then-divide idiom (pattern P20).',
    impl: function (args, ctx) {
      var t = 0, n = 0;
      var e = walk(args, ctx, function (v, fromRange) {
        if (typeof v === 'number') { t += v; n++; return; }
        if (fromRange) return;
        if (typeof v === 'boolean') { t += v ? 1 : 0; n++; }
      });
      if (e) return e;
      if (!n) return V.div0('AVERAGE found no numbers to average');
      return t / n;
    }
  });

  function extreme(name, pick, blurb) {
    def({
      name: name, min: 1, max: 255, group: 'math',
      sig: name + '(number1, [number2], …)',
      doc: blurb,
      estate: 'Used to find the earliest or latest date in a pasted block: ' +
        'MIN(\'Transaction Listings Paste\'!I:I). Dates are numbers, so MIN works on them.',
      impl: function (args, ctx) {
        var best = null;
        var e = walk(args, ctx, function (v, fromRange) {
          var n = null;
          if (typeof v === 'number') n = v;
          else if (!fromRange && typeof v === 'boolean') n = v ? 1 : 0;
          if (n === null) return;
          if (best === null || pick(n, best)) best = n;
        });
        if (e) return e;
        return best === null ? 0 : best;     // Excel returns 0 when nothing qualifies
      }
    });
  }
  extreme('MIN', function (a, b) { return a < b; }, 'The smallest number. Returns 0 if there are none.');
  extreme('MAX', function (a, b) { return a > b; }, 'The largest number. Returns 0 if there are none.');

  /* ------------------------------------------------- the *IF / *IFS family -- */

  /* Build a view over sum_range that is ALIGNED to the criteria range: same
     shape, anchored at sum_range's top-left. That is Excel's documented
     behaviour and it is why =SUMIF(A:A,">0",B:B) works even though B:B is a
     different block from A:A. */
  function aligned(ctx, critView, sumVal) {
    if (!V.isRange(sumVal)) return ctx.asView(sumVal);
    var sh = ctx.wb.sheet(sumVal.sheet);
    if (!sh) return V.ref('there is no sheet called ' + JSON.stringify(sumVal.sheet));
    return new XT.RangeView(ctx.wb, sh, sumVal.r1, sumVal.c1,
      sumVal.r1 + critView.nr - 1, sumVal.c1 + critView.nc - 1);
  }

  def({
    name: 'COUNTIF', min: 2, max: 2, group: 'math', arrayArgs: [0],
    sig: 'COUNTIF(range, criteria)',
    doc: 'Counts cells matching one condition.',
    estate: '71,070 cells. Two jobs: an existence check before a join, and the ' +
      'expanding-range trick COUNTIF($C5:C$16,C5) that numbers repeats 1,2,3 so a ' +
      'unique key can be manufactured (pattern P14).',
    pats: ['P11', 'P14'],
    impl: function (args, ctx) {
      var view = ctx.asView(args[0]);
      if (V.isErr(view)) return view;
      var crit = C.parse(ctx.scalar(args[1]));
      if (crit.err) return crit.err;
      var n = 0, all = view.flat();
      for (var i = 0; i < all.length; i++) if (C.match(all[i], crit)) n++;
      return n;
    }
  });

  def({
    name: 'COUNTIFS', min: 2, max: 254, group: 'math',
    arrayArgs: [0, 2, 4, 6, 8, 10, 12, 14],
    sig: 'COUNTIFS(criteria_range1, criteria1, [criteria_range2, criteria2], …)',
    doc: 'Counts rows that satisfy every condition at once.',
    estate: 'The most widespread aggregator in the estate: 39 of 152 workbooks. ' +
      'Every KPI number is one of these -- count of paste rows matching test code ' +
      'AND site AND week AND a Yes/No flag. Mixed anchoring ($A$1 vs B$1 vs $A1) ' +
      'is what lets one formula fill both right and down (pattern P13).',
    pats: ['P13'],
    impl: function (args, ctx) {
      if (args.length % 2) return V.value('COUNTIFS needs the arguments in pairs: ' +
        'a range then a criterion, a range then a criterion');
      var views = [], crits = [];
      for (var i = 0; i < args.length; i += 2) {
        var v = ctx.asView(args[i]);
        if (V.isErr(v)) return v;
        var c = C.parse(ctx.scalar(args[i + 1]));
        if (c.err) return c.err;
        views.push(v); crits.push(c);
      }
      var n = 0;
      var e = C.forEachMatch(views, crits, function () { n++; });
      return e || n;
    }
  });

  def({
    name: 'SUMIF', min: 2, max: 3, group: 'math', arrayArgs: [0, 2],
    sig: 'SUMIF(range, criteria, [sum_range])',
    doc: 'Adds the cells that match one condition.',
    estate: '44,081 cells. Often with a text criterion on a row: ' +
      'SUMIF(\'Bone QA Check Paste\'!$2:$2, AKF2, \'Bone QA Check Paste\'!$16:$16).',
    pats: ['P12'],
    impl: function (args, ctx) {
      var view = ctx.asView(args[0]);
      if (V.isErr(view)) return view;
      var crit = C.parse(ctx.scalar(args[1]));
      if (crit.err) return crit.err;
      var sv = view;
      if (args.length >= 3) {
        sv = aligned(ctx, view, args[2]);
        if (V.isErr(sv)) return sv;
      }
      var total = 0;
      for (var r = 0; r < view.nr; r++) {
        for (var c = 0; c < view.nc; c++) {
          if (!C.match(view.get(r, c), crit)) continue;
          var v = sv.get ? sv.get(r, c) : V.EMPTY;
          if (V.isErr(v)) return v;
          if (typeof v === 'number') total += v;
        }
      }
      return total;
    }
  });

  function ifsAggregate(name, finish, blurb, estate, pats) {
    def({
      name: name, min: 3, max: 255, group: 'math',
      arrayArgs: [0, 1, 3, 5, 7, 9, 11, 13, 15],
      sig: name + '(' + (name === 'SUMIFS' ? 'sum_range' : 'average_range') +
        ', criteria_range1, criteria1, …)',
      doc: blurb, estate: estate, pats: pats,
      impl: function (args, ctx) {
        if (args.length % 2 === 0) {
          return V.value(name + ' needs a range to total, then range/criterion pairs');
        }
        var views = [], crits = [];
        for (var i = 1; i < args.length; i += 2) {
          var v = ctx.asView(args[i]);
          if (V.isErr(v)) return v;
          var c = C.parse(ctx.scalar(args[i + 1]));
          if (c.err) return c.err;
          views.push(v); crits.push(c);
        }
        var target = aligned(ctx, views[0], args[0]);
        if (V.isErr(target)) return target;
        var total = 0, n = 0, bad = null;
        var e = C.forEachMatch(views, crits, function (r, c) {
          var v = target.get(r, c);
          if (V.isErr(v)) { bad = v; return; }
          if (typeof v === 'number') { total += v; n++; }
        });
        if (e) return e;
        if (bad) return bad;
        return finish(total, n);
      }
    });
  }

  ifsAggregate('SUMIFS', function (t) { return t; },
    'Adds the cells that satisfy every condition.',
    'The period-window total: SUMIFS(\'Volumes Paste\'!17:17, \'Volumes Paste\'!$13:$13, ' +
    '">=" & K$14, \'Volumes Paste\'!$13:$13, "<=" & K$15). Note the criteria are ' +
    'STRINGS assembled with & -- that is where most people come unstuck, and note ' +
    'the ranges are ROWS, because the pasted volumes file has dates across the top.',
    ['P12']);

  ifsAggregate('AVERAGEIFS', function (t, n) {
    return n ? t / n : V.div0('AVERAGEIFS found no matching numbers to average');
  },
    'The mean of the cells that satisfy every condition.',
    'Mean result per week per site: AVERAGEIFS($U:$U,$V:$V,$V2,$AA:$AA,$AA2).',
    ['P20']);

  /* ---------------------------------------------------------- SUMPRODUCT --- */

  def({
    name: 'SUMPRODUCT', min: 1, max: 255, group: 'math',
    arrayArgs: [0, 1, 2, 3, 4, 5],
    sig: 'SUMPRODUCT(array1, [array2], …)',
    doc: 'Multiplies the arrays together element by element, then adds the results.',
    estate: '1,763 cells. This is how multi-criteria counting was done before ' +
      'COUNTIFS existed: SUMPRODUCT((a=x)*(b=y)). Worth knowing because old ' +
      'workbooks are full of it.',
    impl: function (args, ctx) {
      var views = [];
      for (var i = 0; i < args.length; i++) {
        var v = ctx.asView(args[i]);
        if (V.isErr(v)) return v;
        views.push(v);
      }
      var nr = views[0].nr, nc = views[0].nc;
      for (var k = 1; k < views.length; k++) {
        if (views[k].nr !== nr || views[k].nc !== nc) {
          return V.value('SUMPRODUCT needs all its arrays to be the same shape');
        }
      }
      var total = 0;
      for (var r = 0; r < nr; r++) {
        for (var c = 0; c < nc; c++) {
          var p = 1;
          for (var j = 0; j < views.length; j++) {
            var v2 = views[j].get(r, c);
            if (V.isErr(v2)) return v2;
            var n = (typeof v2 === 'number') ? v2
              : (typeof v2 === 'boolean' ? (v2 ? 1 : 0) : 0);   // text counts as 0
            p *= n;
          }
          total += p;
        }
      }
      return total;
    }
  });

  /* ------------------------------------------------------- SMALL / LARGE --- */

  function nth(name, fromEnd) {
    def({
      name: name, min: 2, max: 2, group: 'math', arrayArgs: [0],
      sig: name + '(array, k)',
      doc: 'The ' + (fromEnd ? 'kth largest' : 'kth smallest') + ' number in a block.',
      estate: 'SMALL is the engine of the "get the nth match" array formula: ' +
        'SMALL(IF(condition, ROW(range)), k). Non-numbers -- including the FALSE ' +
        'that IF returns for non-matching rows -- are simply skipped, which is ' +
        'what makes the trick work (pattern P15).',
      pats: ['P15'],
      impl: function (args, ctx) {
        var view = ctx.asView(args[0]);
        if (V.isErr(view)) return view;
        var k = ctx.int(args[1]);
        if (V.isErr(k)) return k;
        var nums = [], all = view.flat();
        for (var i = 0; i < all.length; i++) {
          if (V.isErr(all[i])) return all[i];
          if (typeof all[i] === 'number') nums.push(all[i]);
        }
        if (k < 1 || k > nums.length) {
          return V.num(name + ' was asked for item ' + k + ' of ' + nums.length + ' numbers');
        }
        nums.sort(function (a, b) { return a - b; });
        return fromEnd ? nums[nums.length - k] : nums[k - 1];
      }
    });
  }
  nth('SMALL', false);
  nth('LARGE', true);

  def({
    name: 'RANK', min: 2, max: 3, group: 'math', arrayArgs: [1],
    sig: 'RANK(number, ref, [order])',
    doc: 'The position of a number in a list. Largest is 1 unless order is non-zero.',
    estate: '896 cells, for "worst 5" tables.',
    impl: function (args, ctx) {
      var x = ctx.num(args[0]);
      if (V.isErr(x)) return x;
      var view = ctx.asView(args[1]);
      if (V.isErr(view)) return view;
      var asc = false;
      if (args.length >= 3) {
        var o = ctx.num(args[2]);
        if (V.isErr(o)) return o;
        asc = o !== 0;
      }
      var all = view.flat(), better = 0, seen = false;
      for (var i = 0; i < all.length; i++) {
        if (typeof all[i] !== 'number') continue;
        if (all[i] === x) seen = true;
        if (asc ? all[i] < x : all[i] > x) better++;
      }
      if (!seen) return V.na('RANK could not find ' + V.numToText(x) + ' in the list');
      return better + 1;
    }
  });

  /* ------------------------------------------------------------- arithmetic */

  function simple(name, min, max, fn, sig, doc, estate) {
    def({
      name: name, min: min, max: max, group: 'math', sig: sig, doc: doc,
      estate: estate || '',
      impl: function (args, ctx) {
        var ns = [];
        for (var i = 0; i < args.length; i++) {
          var n = ctx.num(args[i]);
          if (V.isErr(n)) return n;
          ns.push(n);
        }
        return fn.apply(null, ns);
      }
    });
  }

  simple('ROUND', 1, 2, function (x, d) {
    d = d === undefined ? 0 : Math.trunc(d);
    var m = Math.pow(10, d);
    return (x < 0 ? -1 : 1) * Math.round(Math.abs(x) * m) / m;    // Excel rounds half away from zero
  }, 'ROUND(number, num_digits)', 'Rounds to a number of decimal places.',
    'Only 9 cells in the whole estate. Rounding is not what this work is about.');

  simple('ROUNDUP', 1, 2, function (x, d) {
    d = d === undefined ? 0 : Math.trunc(d);
    var m = Math.pow(10, d);
    return (x < 0 ? -1 : 1) * Math.ceil(Math.abs(x) * m) / m;
  }, 'ROUNDUP(number, num_digits)', 'Rounds away from zero.');

  simple('ROUNDDOWN', 1, 2, function (x, d) {
    d = d === undefined ? 0 : Math.trunc(d);
    var m = Math.pow(10, d);
    return (x < 0 ? -1 : 1) * Math.floor(Math.abs(x) * m) / m;
  }, 'ROUNDDOWN(number, num_digits)', 'Rounds towards zero.');

  simple('INT', 1, 1, function (x) { return Math.floor(x); },
    'INT(number)', 'Rounds down to a whole number. On a date-and-time serial this ' +
    'throws the time away and leaves the date.',
    '39 cells. INT(serial) is the standard way to strip a time off a timestamp.');

  simple('ABS', 1, 1, function (x) { return Math.abs(x); },
    'ABS(number)', 'The size of a number, ignoring its sign.');

  simple('MOD', 2, 2, function (x, y) {
    if (y === 0) return V.div0('MOD cannot divide by zero');
    return x - y * Math.floor(x / y);
  }, 'MOD(number, divisor)', 'The remainder after division.');

})(window.XT = window.XT || {});
