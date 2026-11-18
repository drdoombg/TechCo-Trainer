/* =============================================================================
   Excel Trainer — engine/formats.js
   Number formats. These are DISPLAY ONLY -- they never change a cell's value.
   A date is the number 46341 whether it shows as 15/11/2026 or as "Sunday".

   The format codes implemented here are exactly the ones counted in the source
   estate (catalog §7.8), the common ones being:

       dd/mm/yyyy            73 workbooks   (UK order -- the trainer default)
       m/d/yyyy             226 occurrences (what a US-built export arrives as)
       [$-F400]h:mm:ss AM/PM   152          (Excel's "system time" code)
       ddd / dddd / mmmm                    day and month names
       yyyy.mm.dd hh:mm:ss                  the MIMS scanner timestamp
       £0.00  £#,##0.00  and the accounting form
       0.0%   0.0   0.000   0.0000E+00
       0;[Red]0            a two-section format: negatives in red
       @                   text placeholder
       d/m/yy;@            date, with text passed straight through

   The same engine serves the TEXT() function, so TEXT(AC2,"MMM-YY") -- the
   estate's period label, 3,346 cells -- goes through this code too.
   ============================================================================= */
(function (XT) {
  'use strict';

  var F = {};
  XT.F = F;
  var V = XT.V, D = XT.D;

  F.GENERAL = 'General';

  /* ------------------------------------------------------------ section split */

  // Split on ";" but not inside "quoted text" or [brackets].
  function splitSections(code) {
    var out = [], cur = '', q = false, br = 0;
    for (var i = 0; i < code.length; i++) {
      var ch = code.charAt(i);
      if (ch === '"') { q = !q; cur += ch; continue; }
      if (!q && ch === '[') { br++; cur += ch; continue; }
      if (!q && ch === ']') { br--; cur += ch; continue; }
      if (!q && !br && ch === ';') { out.push(cur); cur = ''; continue; }
      if (ch === '\\') { cur += ch + code.charAt(i + 1); i++; continue; }
      cur += ch;
    }
    out.push(cur);
    return out;
  }

  /* ------------------------------------------------------------- tokenising -- */

  /* A format section becomes a token list. Token types:
       lit   literal text
       d     date/time field, with .f = the field letters (d, dd, mmm, hh, ss, …)
       num   a run of digit placeholders 0 # ? . , %
       at    the @ text placeholder
       ampm  AM/PM marker  */
  function tokenizeFormat(sec) {
    var toks = [], i = 0, n = sec.length, color = null, condition = null;

    while (i < n) {
      var ch = sec.charAt(i);

      // [Red] [Blue] [$-F400] [>100]
      if (ch === '[') {
        var end = sec.indexOf(']', i);
        if (end < 0) { toks.push({ t: 'lit', v: sec.slice(i) }); break; }
        var body = sec.slice(i + 1, end);
        if (/^(Red|Blue|Green|Black|White|Yellow|Magenta|Cyan)$/i.test(body)) color = body.toLowerCase();
        else if (/^[<>=]/.test(body)) condition = body;
        // [$-F400] and other locale ids carry no visible output.
        i = end + 1; continue;
      }

      if (ch === '"') {
        var e = sec.indexOf('"', i + 1);
        if (e < 0) e = n;
        toks.push({ t: 'lit', v: sec.slice(i + 1, e) });
        i = e + 1; continue;
      }

      if (ch === '\\') { toks.push({ t: 'lit', v: sec.charAt(i + 1) || '' }); i += 2; continue; }
      if (ch === '_') { toks.push({ t: 'lit', v: ' ' }); i += 2; continue; }   // _( = a space
      if (ch === '*') { i += 2; continue; }                                     // fill char: ignore

      if (ch === '@') { toks.push({ t: 'at' }); i++; continue; }

      // AM/PM switches the hour field to 12-hour.
      if (/^(AM\/PM|A\/P)/i.test(sec.slice(i))) {
        var mm = /^(AM\/PM|A\/P)/i.exec(sec.slice(i));
        toks.push({ t: 'ampm', v: mm[0] });
        i += mm[0].length; continue;
      }

      // Date/time field letters, greedily.
      var dm = /^(y{1,4}|m{1,5}|d{1,4}|h{1,2}|s{1,2})/i.exec(sec.slice(i));
      if (dm) { toks.push({ t: 'd', f: dm[0] }); i += dm[0].length; continue; }

      // A run of number placeholders.
      var nm = /^[#0?.,%]+/.exec(sec.slice(i));
      if (nm) { toks.push({ t: 'num', v: nm[0] }); i += nm[0].length; continue; }

      // Scientific notation.
      var em = /^[Ee][+-]/.exec(sec.slice(i));
      if (em) { toks.push({ t: 'exp', v: em[0] }); i += 2; continue; }

      toks.push({ t: 'lit', v: ch }); i++;
    }

    // Resolve the m ambiguity: m is MINUTES when it sits next to an hour or a
    // second field, and MONTHS otherwise.  "hh:mm:ss" vs "dd/mm/yyyy".
    for (var k = 0; k < toks.length; k++) {
      if (toks[k].t === 'd' && /^m+$/i.test(toks[k].f)) {
        var prev = prevD(toks, k), nxt = nextD(toks, k);
        if ((prev && /^h/i.test(prev.f)) || (nxt && /^s/i.test(nxt.f))) toks[k].minute = true;
      }
    }
    return { toks: toks, color: color, condition: condition };
  }
  function prevD(toks, k) { for (var i = k - 1; i >= 0; i--) if (toks[i].t === 'd') return toks[i]; return null; }
  function nextD(toks, k) { for (var i = k + 1; i < toks.length; i++) if (toks[i].t === 'd') return toks[i]; return null; }

  function isDateSection(parsed) {
    for (var i = 0; i < parsed.toks.length; i++) if (parsed.toks[i].t === 'd') return true;
    return false;
  }

  /* ------------------------------------------------------------- number part */

  function pad(n, w) { n = String(n); while (n.length < w) n = '0' + n; return n; }

  // Render a number through a run of placeholders like  #,##0.00  or  0.0%
  function renderNumber(value, toks) {
    /* Collect the numeric picture. Two things have to be kept apart here:
         - the digits AFTER an E+ belong to the exponent, not to the mantissa,
           so 0.0000E+00 has four decimal places and not six;
         - a % both multiplies the value by 100 AND prints a % sign. */
    var picture = '', expPicture = '', expTok = null;
    toks.forEach(function (t) {
      if (t.t === 'exp') { expTok = t.v; return; }
      if (t.t === 'num') { if (expTok) expPicture += t.v; else picture += t.v; }
    });

    var pct = (picture.match(/%/g) || []).length;
    var hasComma = /[#0?],[#0?]/.test(picture) || /,[#0?]/.test(picture);
    var clean = picture.replace(/%/g, '').replace(/,/g, '');
    var dot = clean.indexOf('.');
    var before = dot < 0 ? clean : clean.slice(0, dot);
    var after = dot < 0 ? '' : clean.slice(dot + 1);

    var v = value * Math.pow(100, pct);
    var dec = (after.match(/[0#?]/g) || []).length;
    var suffix = pct ? new Array(pct + 1).join('%') : '';

    if (expTok) {
      var expDigits = Math.max(2, (expPicture.match(/[0#?]/g) || []).length);
      var s = v.toExponential(dec);
      var pm = /e([+-])(\d+)/.exec(s);
      var mant = s.slice(0, s.indexOf('e'));
      return mant + 'E' + (pm[1] === '-' ? '-' : (expTok.charAt(1) === '+' ? '+' : '')) +
        pad(pm[2], expDigits) + suffix;
    }

    var neg = v < 0;
    var av = Math.abs(v);
    var fixed = dec > 0 ? av.toFixed(dec) : String(Math.round(av));
    var parts = fixed.split('.');
    var ip = parts[0], fp = parts[1] || '';

    // Minimum integer digits, from the count of 0s in the picture.
    var minInt = (before.match(/0/g) || []).length;
    while (ip.length < minInt) ip = '0' + ip;
    if (minInt === 0 && ip === '0' && before.indexOf('#') >= 0) ip = '';   // #.## shows .5

    if (hasComma) ip = ip.replace(/\B(?=(\d{3})+(?!\d))/g, ',');

    var out = ip + (dec > 0 ? '.' + fp : '') + suffix;
    if (neg) out = '-' + out;
    return out;
  }

  /* --------------------------------------------------------------- date part */

  function renderDate(serial, toks) {
    if (serial < 0) return null;
    var ymd = D.serialToYMD(serial);
    var hms = D.serialToHMS(serial);
    var wd = D.weekday(serial, 1) - 1;                 // 0 = Sunday
    var use12 = toks.some(function (t) { return t.t === 'ampm'; });
    var h12 = hms.h % 12; if (h12 === 0) h12 = 12;

    var out = '';
    toks.forEach(function (t) {
      if (t.t === 'lit') { out += t.v; return; }
      if (t.t === 'at') { return; }
      if (t.t === 'ampm') { out += hms.h < 12 ? 'AM' : 'PM'; return; }
      if (t.t === 'num') { out += t.v.replace(/[#0?]/g, ''); return; }
      if (t.t !== 'd') return;
      var f = t.f.toLowerCase();
      if (f.charAt(0) === 'y') {
        out += (f.length <= 2) ? pad(ymd.y % 100, 2) : String(ymd.y);
      } else if (f.charAt(0) === 'd') {
        if (f.length === 1) out += String(ymd.d);
        else if (f.length === 2) out += pad(ymd.d, 2);
        else if (f.length === 3) out += D.DAYS[wd].slice(0, 3);
        else out += D.DAYS[wd];
      } else if (f.charAt(0) === 'h') {
        var hv = use12 ? h12 : hms.h;
        out += (f.length === 1) ? String(hv) : pad(hv, 2);
      } else if (f.charAt(0) === 's') {
        out += (f.length === 1) ? String(hms.s) : pad(hms.s, 2);
      } else if (f.charAt(0) === 'm') {
        if (t.minute) out += (f.length === 1) ? String(hms.mi) : pad(hms.mi, 2);
        else if (f.length === 1) out += String(ymd.m);
        else if (f.length === 2) out += pad(ymd.m, 2);
        else if (f.length === 3) out += D.MONTHS[ymd.m - 1].slice(0, 3);
        else if (f.length === 4) out += D.MONTHS[ymd.m - 1];
        else out += D.MONTHS[ymd.m - 1].charAt(0);
      }
    });
    return out;
  }

  /* ----------------------------------------------------------------- General */

  /* The General format is not a picture; it is "show the number as briefly as
     possible, up to about 11 characters in a default-width column". We use the
     simpler 15-significant-digit rule, which is what & does, because that is
     what learners will compare against. */
  function general(v) { return V.numToText(v); }

  /* -------------------------------------------------------------- public API */

  var cache = {};
  function compile(code) {
    if (cache[code]) return cache[code];
    var secs = splitSections(code).map(tokenizeFormat);
    var info = { secs: secs, date: secs.length && isDateSection(secs[0]) };
    cache[code] = info;
    return info;
  }
  F.compile = compile;

  // Is this format code a date/time picture? The grid uses it to decide
  // whether a typed "15/11/2026" should become a serial.
  F.isDateFormat = function (code) {
    if (!code || code === F.GENERAL) return false;
    return compile(code).date;
  };

  /* format(value, code) -> display string
     value may be a number, text, boolean, EMPTY or an error. */
  F.format = function (value, code) {
    if (V.isErr(value)) return value.code;
    if (V.isEmpty(value)) return '';
    if (!code || code === F.GENERAL) return V.toDisplay(value);

    var info = compile(code);
    var secs = info.secs;

    // Text goes to the 4th section if there is one, otherwise straight through.
    if (typeof value === 'string' || typeof value === 'boolean') {
      if (typeof value === 'boolean') return value ? 'TRUE' : 'FALSE';
      var tsec = secs[3] || null;
      if (!tsec) return value;
      var out = '';
      tsec.toks.forEach(function (t) {
        if (t.t === 'at') out += value;
        else if (t.t === 'lit') out += t.v;
      });
      return out;
    }

    // Pick the section: positive ; negative ; zero ; text
    var sec = secs[0], n = value;
    if (secs.length >= 2 && value < 0) { sec = secs[1]; n = Math.abs(value); }
    if (secs.length >= 3 && value === 0) { sec = secs[2]; }

    if (isDateSection(sec)) {
      var r = renderDate(n, sec.toks);
      return r === null ? '#######' : r;
    }

    // A numeric picture: walk the tokens, substituting the number once.
    var done = false, s = '';
    for (var i = 0; i < sec.toks.length; i++) {
      var t = sec.toks[i];
      if (t.t === 'num' || t.t === 'exp') {
        if (!done) { s += renderNumber(n, sec.toks); done = true; }
        continue;
      }
      if (t.t === 'lit') s += t.v;
      if (t.t === 'at') s += V.numToText(n);
    }
    if (!done) s += general(n);
    return s;
  };

  // The colour a format asks for, e.g. 0;[Red]0 on a negative number.
  F.colorFor = function (value, code) {
    if (!code || code === F.GENERAL || typeof value !== 'number') return null;
    var secs = compile(code).secs;
    var sec = secs[0];
    if (secs.length >= 2 && value < 0) sec = secs[1];
    if (secs.length >= 3 && value === 0) sec = secs[2];
    return sec.color || null;
  };

  // Right-align numbers and dates, left-align text: Excel's default alignment,
  // and a quiet tell that a "number" is really text.
  F.alignOf = function (value) {
    if (V.isErr(value)) return 'center';
    if (typeof value === 'number') return 'right';
    if (typeof value === 'boolean') return 'center';
    return 'left';
  };

})(window.XT = window.XT || {});
