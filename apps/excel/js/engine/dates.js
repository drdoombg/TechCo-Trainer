/* =============================================================================
   Excel Trainer — engine/dates.js
   Dates in Excel are just numbers. Serial 1 is 1900-01-01, and a time is the
   fractional part of a day, so 0.5 is noon.

   Two consequences the catalog leans on:

     1. (B4 - WEEKDAY(B4,1)) + 7   rounds a date to its week-ending Saturday.
        That pattern appears 177,447 times in the source estate (pattern P8),
        because the business runs on retailer calendars, not ISO weeks --
        WEEKNUM is used exactly once in 152 workbooks.

     2. 0/1/1900 is NOT a date. Excel parses it as the arithmetic 0 / 1 / 1900,
        which is 0. A live workbook contains =IF(D2=0/1/1900,"",…) and that
        test can only ever be true for a zero. That is catalog defect D3, and
        we reproduce it by doing nothing special -- the parser just divides.

   The 1900 leap-year bug: Excel believes 1900-02-29 existed (serial 60). Every
   serial above 60 is therefore one greater than a naive day count. We honour
   it so that serials copied out of a real workbook land on the right day.
   ============================================================================= */
(function (XT) {
  'use strict';

  var D = {};
  XT.D = D;

  var MS = 86400000;
  // 1899-12-31 in UTC. Serial 1 is one day after it.
  var EPOCH = Date.UTC(1899, 11, 31);

  /* ------------------------------------------------------ serial <-> calendar */

  // Whole-day part of a serial -> {y, m, d} (m is 1-12).
  D.serialToYMD = function (serial) {
    var days = Math.floor(serial);
    if (days > 59) days -= 1;             // step over the phantom 1900-02-29
    var dt = new Date(EPOCH + days * MS);
    return { y: dt.getUTCFullYear(), m: dt.getUTCMonth() + 1, d: dt.getUTCDate() };
  };

  D.ymdToSerial = function (y, m, d) {
    // Excel lets DATE() roll over out-of-range months and days.
    var dt = Date.UTC(y, m - 1, d);
    var days = Math.round((dt - EPOCH) / MS);
    if (days > 59) days += 1;
    return days;
  };

  // Fractional part of a serial -> {h, mi, s}
  D.serialToHMS = function (serial) {
    var frac = serial - Math.floor(serial);
    var total = Math.round(frac * 86400);
    if (total >= 86400) total -= 86400;
    return { h: Math.floor(total / 3600), mi: Math.floor((total % 3600) / 60), s: total % 60 };
  };

  D.hmsToSerial = function (h, mi, s) {
    return (h * 3600 + mi * 60 + s) / 86400;
  };

  /* ------------------------------------------------------------------ WEEKDAY */

  /* Serial 1 is a Sunday as far as Excel is concerned, so type 1 is simply
       ((serial - 1) mod 7) + 1
     Sanity check: 2000-01-01 is serial 36526 and was a Saturday.
       ((36526-1) mod 7) + 1 = 6 + 1 = 7 = Saturday.  Correct. */
  D.weekday = function (serial, type) {
    var s = Math.floor(serial);
    var sun = ((s - 1) % 7 + 7) % 7;          // 0 = Sunday
    switch (type === undefined ? 1 : type) {
      case 1: return sun + 1;                  // Sun=1 .. Sat=7   (the only one used)
      case 2: return ((sun + 6) % 7) + 1;      // Mon=1 .. Sun=7
      case 3: return (sun + 6) % 7;            // Mon=0 .. Sun=6
      case 11: return ((sun + 6) % 7) + 1;
      case 12: return ((sun + 5) % 7) + 1;
      case 13: return ((sun + 4) % 7) + 1;
      case 14: return ((sun + 3) % 7) + 1;
      case 15: return ((sun + 2) % 7) + 1;
      case 16: return ((sun + 1) % 7) + 1;
      case 17: return sun + 1;
      default: return null;
    }
  };

  /* --------------------------------------------------------- today / fixtures */

  /* The trainer freezes "today" so that exercises grade deterministically.
     Lessons that use TODAY() therefore behave the same on any day, which also
     means a saved answer stays correct tomorrow. Sandbox mode can unfreeze it. */
  D.frozenToday = null;          // a serial, or null for the real clock

  D.today = function () {
    if (D.frozenToday !== null) return D.frozenToday;
    var n = new Date();
    return D.ymdToSerial(n.getFullYear(), n.getMonth() + 1, n.getDate());
  };

  D.now = function () {
    if (D.frozenToday !== null) return D.frozenToday + D.hmsToSerial(9, 30, 0);
    var n = new Date();
    return D.ymdToSerial(n.getFullYear(), n.getMonth() + 1, n.getDate()) +
      D.hmsToSerial(n.getHours(), n.getMinutes(), n.getSeconds());
  };

  /* ------------------------------------------------------------ text -> date */

  /* Only VALUE() and the fixture loader use this. Ordinary arithmetic does NOT
     parse date text, which is faithful enough for the dialect we target.
     Accepts dd/mm/yyyy (UK -- the estate's default display), yyyy-mm-dd and
     dd/mm/yyyy hh:mm[:ss]. */
  D.parseDateText = function (s) {
    if (typeof s !== 'string') return null;
    var t = s.replace(/^[ \t]+|[ \t]+$/g, '');
    var m;
    var timePart = 0;
    var tm = /[ T](\d{1,2}):(\d{2})(?::(\d{2}))?$/.exec(t);
    if (tm) {
      timePart = D.hmsToSerial(+tm[1], +tm[2], +(tm[3] || 0));
      t = t.slice(0, tm.index);
    }
    if ((m = /^(\d{4})-(\d{1,2})-(\d{1,2})$/.exec(t))) {
      return D.ymdToSerial(+m[1], +m[2], +m[3]) + timePart;
    }
    if ((m = /^(\d{1,2})\/(\d{1,2})\/(\d{2}|\d{4})$/.exec(t))) {
      var y = +m[3];
      if (y < 100) y += (y < 30 ? 2000 : 1900);
      return D.ymdToSerial(y, +m[2], +m[1]) + timePart;      // day first: UK
    }
    if (tm && t === '') return timePart;                      // a bare time
    return null;
  };

  // Day and month names, for the TEXT() and number-format code.
  D.DAYS = ['Sunday', 'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday'];
  D.MONTHS = ['January', 'February', 'March', 'April', 'May', 'June',
    'July', 'August', 'September', 'October', 'November', 'December'];

})(window.XT = window.XT || {});
