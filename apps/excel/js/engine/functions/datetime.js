/* =============================================================================
   Excel Trainer — engine/functions/datetime.js
   Dates, times and the week grid.

   WEEKDAY is the surprise of this group: 177,467 cells, all of them inside one
   idiom that rounds a date to its week-ending day:

       (B4 - WEEKDAY(B4,1)) + 7

   Read it as: "take the date, step back to the Saturday before it, then step
   forward seven days." The result is the week-ending Saturday, which is then
   looked up in a Calendars sheet to get the retailer period. WEEKNUM is used
   ONCE in 152 workbooks, because retailer weeks do not line up with ISO weeks,
   so the calendar has to be a table, not a calculation. That is pattern P8.

   The TIME/HOUR/MINUTE/SECOND cluster (about 2,775 cells each) exists for one
   purpose: comparing a timestamp against a shift cut-off. You cannot compare a
   full timestamp to a time of day, so the time is rebuilt from its parts:

       IF( MANUFTIME < TIME(HOUR($R$1),MINUTE($R$1),SECOND($R$1)), …back-date… )
   ============================================================================= */
(function (XT) {
  'use strict';

  var V = XT.V, D = XT.D, def = XT.defineFunction;

  def({
    name: 'TODAY', min: 0, max: 0, group: 'date', volatile: true,
    sig: 'TODAY()',
    doc: 'Today’s date, as a serial number with no time part.',
    estate: '3,872 cells in 24 workbooks, driving "today’s samples" filters. ' +
      'In the trainer today is FROZEN so exercises grade the same on any day.',
    impl: function () { return D.today(); }
  });

  def({
    name: 'NOW', min: 0, max: 0, group: 'date', volatile: true,
    sig: 'NOW()', doc: 'The current date AND time.',
    estate: '5 cells, on cover sheets.',
    impl: function () { return D.now(); }
  });

  def({
    name: 'DATE', min: 3, max: 3, group: 'date',
    sig: 'DATE(year, month, day)',
    doc: 'Builds a date from its parts. Out-of-range months and days roll over, ' +
      'so DATE(2026,13,1) is January 2027.',
    estate: 'Not used in the estate -- which is why 0/1/1900 appears instead, and ' +
      'that is arithmetic, not a date (defect D3). DATE(1900,1,1) is what was meant.',
    pats: ['D3'],
    impl: function (args, ctx) {
      var y = ctx.int(args[0]); if (V.isErr(y)) return y;
      var m = ctx.int(args[1]); if (V.isErr(m)) return m;
      var d = ctx.int(args[2]); if (V.isErr(d)) return d;
      if (y < 0 || y > 9999) return V.num('DATE needs a year between 0 and 9999');
      if (y < 1900) y += 1900;              // Excel's two-digit year rule
      var s = D.ymdToSerial(y, m, d);
      if (s < 0) return V.num('that date is before 1900 and has no serial number');
      return s;
    }
  });

  def({
    name: 'WEEKDAY', min: 1, max: 2, group: 'date',
    sig: 'WEEKDAY(serial_number, [return_type])',
    doc: 'Which day of the week a date is. With return_type 1 (the default), ' +
      'Sunday is 1 and Saturday is 7.',
    estate: '177,467 cells, essentially all of them inside ' +
      '(date - WEEKDAY(date,1)) + 7, which gives the week-ending Saturday.',
    pats: ['P8'],
    impl: function (args, ctx) {
      var s = ctx.num(args[0]); if (V.isErr(s)) return s;
      if (s < 0) return V.num('a date serial cannot be negative');
      var t = args.length >= 2 ? ctx.int(args[1]) : 1;
      if (V.isErr(t)) return t;
      var w = D.weekday(s, t);
      if (w === null) return V.num('return_type ' + t + ' is not one Excel recognises');
      return w;
    }
  });

  function part(name, pick, sig, doc, estate) {
    def({
      name: name, min: 1, max: 1, group: 'date', sig: sig, doc: doc, estate: estate || '',
      impl: function (args, ctx) {
        var s = ctx.num(args[0]); if (V.isErr(s)) return s;
        if (s < 0) return V.num('a date serial cannot be negative');
        return pick(s);
      }
    });
  }

  part('YEAR', function (s) { return D.serialToYMD(s).y; },
    'YEAR(serial_number)', 'The year of a date.');
  part('MONTH', function (s) { return D.serialToYMD(s).m; },
    'MONTH(serial_number)', 'The month of a date, 1 to 12.',
    '41 cells. Note MONTH(G2+2) in the Calendars sheet -- the +2 nudges the week ' +
    'into the month it is reported under.');
  part('DAY', function (s) { return D.serialToYMD(s).d; },
    'DAY(serial_number)', 'The day of the month, 1 to 31.');
  part('HOUR', function (s) { return D.serialToHMS(s).h; },
    'HOUR(serial_number)', 'The hour of a time, 0 to 23.',
    '2,777 cells, always feeding TIME() to rebuild a shift cut-off.');
  part('MINUTE', function (s) { return D.serialToHMS(s).mi; },
    'MINUTE(serial_number)', 'The minutes of a time, 0 to 59.');
  part('SECOND', function (s) { return D.serialToHMS(s).s; },
    'SECOND(serial_number)', 'The seconds of a time, 0 to 59.');

  def({
    name: 'TIME', min: 3, max: 3, group: 'date',
    sig: 'TIME(hour, minute, second)',
    doc: 'Builds a time as a fraction of a day. TIME(12,0,0) is 0.5.',
    estate: '2,773 cells. Always TIME(HOUR(x),MINUTE(x),SECOND(x)) -- stripping the ' +
      'date off a timestamp so the time of day can be compared on its own.',
    impl: function (args, ctx) {
      var h = ctx.int(args[0]); if (V.isErr(h)) return h;
      var m = ctx.int(args[1]); if (V.isErr(m)) return m;
      var s = ctx.int(args[2]); if (V.isErr(s)) return s;
      var frac = D.hmsToSerial(h, m, s);
      // Excel wraps past 24 hours rather than erroring.
      frac = frac - Math.floor(frac);
      if (frac < 0) return V.num('TIME cannot make a negative time');
      return frac;
    }
  });

  def({
    name: 'WEEKNUM', min: 1, max: 2, group: 'date',
    sig: 'WEEKNUM(serial_number, [return_type])',
    doc: 'The ISO-ish week number of a date.',
    estate: 'Used ONCE in 152 workbooks, and that is the lesson: the business runs ' +
      'on retailer calendars that do not match these week numbers, so the estate ' +
      'uses a Calendars lookup table instead (pattern P8). Reach for this and you ' +
      'will get the wrong week.',
    pats: ['P8'],
    impl: function (args, ctx) {
      var s = ctx.num(args[0]); if (V.isErr(s)) return s;
      var t = args.length >= 2 ? ctx.int(args[1]) : 1;
      if (V.isErr(t)) return t;
      var ymd = D.serialToYMD(s);
      var jan1 = D.ymdToSerial(ymd.y, 1, 1);
      var startDow = (t === 2) ? D.weekday(jan1, 2) : D.weekday(jan1, 1);
      return Math.floor((s - jan1 + startDow - 1) / 7) + 1;
    }
  });

})(window.XT = window.XT || {});
