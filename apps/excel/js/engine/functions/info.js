/* =============================================================================
   Excel Trainer — engine/functions/info.js
   The IS* family, NA(), the ROW/COLUMN group, and the honest list of functions
   this trainer deliberately does not implement.

   ISBLANK is the only way to tell EMPTY from "" (catalog §7.4), so it earns its
   place even at 139 uses. ROW() matters more than it looks: 18,468 cells use it,
   mostly as a position generator inside the array formula that retrieves the
   nth matching row (pattern P15), and 17,238 as a bare =ROW() row ordinal.

   The unsupported list at the bottom is not laziness. The catalog measured the
   source estate and found NO FILTER, NO SORT, NO SEQUENCE, NO LET, NO LAMBDA,
   two uses of UNIQUE, and no IFS / SWITCH / CHOOSE at all. Dynamic arrays are
   not part of this dialect; fill-down is. Registering them anyway means a
   formula pasted from somewhere else gets a straight answer about why it will
   not run here, instead of a bare #NAME?.
   ============================================================================= */
(function (XT) {
  'use strict';

  var V = XT.V, def = XT.defineFunction;

  /* ------------------------------------------------------------- IS family -- */

  def({
    name: 'ISBLANK', min: 1, max: 1, group: 'info', seesErrors: true,
    sig: 'ISBLANK(value)',
    doc: 'TRUE only for a cell nobody has typed into. A cell holding "" is NOT blank.',
    estate: '139 cells. Under-used, given that 1.2 million formulas in the estate ' +
      'produce "" and every one of them makes a cell that LOOKS blank and is not. ' +
      'ISBLANK is the only function that can tell the difference.',
    pats: ['P3'],
    impl: function (args, ctx) {
      var v = ctx.scalar(args[0]);
      return V.isEmpty(v);
    }
  });

  def({
    name: 'ISERROR', min: 1, max: 1, group: 'info', seesErrors: true,
    sig: 'ISERROR(value)', doc: 'TRUE for any error value.',
    estate: '193 cells, mostly NOT(ISERROR(SEARCH(…))) in conditional formatting.',
    impl: function (args, ctx) { return V.isErr(ctx.scalar(args[0])); }
  });

  def({
    name: 'ISNA', min: 1, max: 1, group: 'info', seesErrors: true,
    sig: 'ISNA(value)', doc: 'TRUE only for #N/A.',
    estate: 'The discriminating version of ISERROR: "not found" rather than "broken".',
    impl: function (args, ctx) {
      var v = ctx.scalar(args[0]);
      return V.isErr(v) && v.code === '#N/A';
    }
  });

  def({
    name: 'ISNUMBER', min: 1, max: 1, group: 'info', seesErrors: true,
    sig: 'ISNUMBER(value)', doc: 'TRUE if the value is a number. Dates are numbers.',
    estate: '12 cells. Deserves far more use: it is the quickest way to prove a ' +
      'column of "numbers" pasted out of a system is actually text.',
    impl: function (args, ctx) { return typeof ctx.scalar(args[0]) === 'number'; }
  });

  def({
    name: 'ISTEXT', min: 1, max: 1, group: 'info', seesErrors: true,
    sig: 'ISTEXT(value)', doc: 'TRUE if the value is text. "" counts as text.',
    estate: '238 cells.',
    impl: function (args, ctx) { return typeof ctx.scalar(args[0]) === 'string'; }
  });

  def({
    name: 'ISNONTEXT', min: 1, max: 1, group: 'info', seesErrors: true,
    sig: 'ISNONTEXT(value)', doc: 'TRUE if the value is not text. Blank counts as non-text.',
    impl: function (args, ctx) { return typeof ctx.scalar(args[0]) !== 'string'; }
  });

  def({
    name: 'ISLOGICAL', min: 1, max: 1, group: 'info', seesErrors: true,
    sig: 'ISLOGICAL(value)', doc: 'TRUE if the value is TRUE or FALSE.',
    impl: function (args, ctx) { return typeof ctx.scalar(args[0]) === 'boolean'; }
  });

  def({
    name: 'NA', min: 0, max: 0, group: 'info',
    sig: 'NA()', doc: 'Produces #N/A on purpose.',
    estate: '17 cells. Used to mark "no data" so that it cannot be mistaken for zero.',
    impl: function () { return V.na('this cell was told to report #N/A on purpose'); }
  });

  /* ---------------------------------------------------- ROW / COLUMN group -- */

  def({
    name: 'ROW', min: 0, max: 1, group: 'info', arrayArgs: [0],
    sig: 'ROW([reference])',
    doc: 'The row number of a reference, or of this cell if you give it nothing.',
    estate: '18,468 cells. Two distinct jobs: a bare =ROW() as a row ordinal ' +
      '(17,238 cells), and ROW(range) inside an array formula to generate a list ' +
      'of row numbers to pick from with SMALL (pattern P15).',
    pats: ['P15', 'P19'],
    impl: function (args, ctx) {
      if (args.length === 0 || (args.nodes && args.nodes[0] && args.nodes[0].t === 'missing')) {
        return ctx.row + 1;
      }
      var a = args[0];
      if (!V.isRange(a)) return V.value('ROW needs a cell or a range');
      if (!ctx.arrayMode || a.r1 === a.r2) return a.r1 + 1;
      // In array context, hand back the whole column of row numbers, clipped to
      // the used range so it lines up with the condition array beside it.
      var view = ctx.view(a);
      if (V.isErr(view)) return view;
      var rows = [];
      for (var i = 0; i < view.nr; i++) rows.push([view.r1 + i + 1]);
      return new V.XArr(rows);
    }
  });

  def({
    name: 'COLUMN', min: 0, max: 1, group: 'info',
    sig: 'COLUMN([reference])',
    doc: 'The column number of a reference, or of this cell. Column A is 1.',
    estate: 'Not used in the estate, but the natural partner to ROW() when a ' +
      'formula has to know where it is.',
    impl: function (args, ctx) {
      if (args.length === 0 || (args.nodes && args.nodes[0] && args.nodes[0].t === 'missing')) {
        return ctx.col + 1;
      }
      if (!V.isRange(args[0])) return V.value('COLUMN needs a cell or a range');
      return args[0].c1 + 1;
    }
  });

  def({
    name: 'ROWS', min: 1, max: 1, group: 'info', arrayArgs: [0],
    sig: 'ROWS(array)', doc: 'How many rows a block has.',
    estate: '12 cells.',
    impl: function (args, ctx) {
      var view = ctx.asView(args[0]);
      return V.isErr(view) ? view : view.nr;
    }
  });

  def({
    name: 'COLUMNS', min: 1, max: 1, group: 'info', arrayArgs: [0],
    sig: 'COLUMNS(array)', doc: 'How many columns a block has.',
    impl: function (args, ctx) {
      var view = ctx.asView(args[0]);
      return V.isErr(view) ? view : view.nc;
    }
  });

  /* -------------------------------------------------- knowingly unsupported */

  var U = XT.declareUnsupported;

  U('FILTER', 'The source estate contains no FILTER at all. Its job is done there ' +
    'by a helper column plus the INDEX/SMALL/IF array formula (pattern P15), which ' +
    'this trainer does implement.', ['P15']);
  U('SORT', 'Not used anywhere in the source estate. Sorting is done through the ' +
    'ribbon, not a formula.');
  U('SORTBY', 'Not used in the source estate.');
  U('SEQUENCE', 'Not used in the source estate; =+K13+1 filled down does this job ' +
    'there (pattern P19).', ['P19']);
  U('UNIQUE', 'Two uses in 152 workbooks. The estate manufactures unique keys with ' +
    'COUNTIF on an expanding range instead (pattern P14).', ['P14']);
  U('LET', 'Not used in the source estate.');
  U('LAMBDA', 'Not used in the source estate.');
  U('IFS', 'Not used in the source estate: nested IF does this work there, up to ' +
    'six deep (pattern P18).', ['P18']);
  U('SWITCH', 'Not used in the source estate.');
  U('CHOOSE', 'Not used in the source estate.');
  U('XMATCH', 'Not used in the source estate; MATCH is.');
  U('TAKE', '12 cells, inside a dynamic-array formula that is outside this trainer’s scope.');
  U('ANCHORARRAY', '12 cells. An internal spill-reference helper, not a function to learn.');
  U('IMAGE', '87 cells, on QR-code label sheets. It fetches a picture from a URL, ' +
    'which an offline trainer cannot do.');
  U('NETWORKDAYS', '11 cells in one workbook.');
  U('SUBTOTAL', 'Not used in the source estate.');
  U('VLOOKUPS', 'Not a real function. You may be thinking of VLOOKUP.');

})(window.XT = window.XT || {});
