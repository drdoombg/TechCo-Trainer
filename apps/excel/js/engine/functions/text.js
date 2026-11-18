/* =============================================================================
   Excel Trainer — engine/functions/text.js
   Text surgery. In the Technical Co-ordinator's folder this is the whole job:
   RIGHT appears in 204,153 cells and LEFT in 200,242, because a vision system
   emits one line of text per pack and somebody has to turn it into columns.

   The best lesson in the entire catalog lives in this file. One real workbook
   contains BOTH of these, solving the same problem:

       =IF(RIGHT(LVALL!C1,29)="FLASHLABEL; Invalid: Mismatch","Missing Sleeve",
           IF(RIGHT(LVALL!C1,8)="Mismatch","Mismatch",""))        <- counting characters

       =IFERROR(_xlfn.TEXTAFTER(F1," ",3),"")                     <- counting fields

   The first breaks the day the vendor changes a message. The second does not.
   Showing them side by side is what level 11 is for.

   Note TRIM: it removes ordinary spaces, CHAR(32), and nothing else. A
   non-breaking space, CHAR(160), survives TRIM and survives VALUE -- which is
   why one workbook needs SUBSTITUTE(D1,CHAR(160)," ") before it can add
   anything up. That is catalog defect D7, and it is reproduced faithfully here.
   ============================================================================= */
(function (XT) {
  'use strict';

  var V = XT.V, F = XT.F, def = XT.defineFunction;

  /* ------------------------------------------------- LEFT / RIGHT / MID ---- */

  def({
    name: 'LEFT', min: 1, max: 2, group: 'text',
    sig: 'LEFT(text, [num_chars])',
    doc: 'The first n characters of some text. n defaults to 1.',
    estate: 'Used to test a prefix and keep a fixed-width slice: ' +
      'IF(LEFT(LVALL!C1,3)="JOB", LEFT(LVALL!C1,45), "").',
    pats: ['P6'],
    impl: function (args, ctx) {
      var s = ctx.text(args[0]); if (V.isErr(s)) return s;
      var n = args.length >= 2 ? ctx.int(args[1]) : 1;
      if (V.isErr(n)) return n;
      if (n < 0) return V.value('LEFT cannot take a negative number of characters');
      return s.slice(0, n);
    }
  });

  def({
    name: 'RIGHT', min: 1, max: 2, group: 'text',
    sig: 'RIGHT(text, [num_chars])',
    doc: 'The last n characters of some text. n defaults to 1.',
    estate: 'The estate’s log classifier: RIGHT(line, 29) = "FLASHLABEL; ' +
      'Invalid: Mismatch". Counting 29 characters by hand is why it is brittle.',
    pats: ['P6'],
    impl: function (args, ctx) {
      var s = ctx.text(args[0]); if (V.isErr(s)) return s;
      var n = args.length >= 2 ? ctx.int(args[1]) : 1;
      if (V.isErr(n)) return n;
      if (n < 0) return V.value('RIGHT cannot take a negative number of characters');
      return n >= s.length ? s : s.slice(s.length - n);
    }
  });

  def({
    name: 'MID', min: 3, max: 3, group: 'text',
    sig: 'MID(text, start_num, num_chars)',
    doc: 'Characters from the middle of some text, counting from 1.',
    estate: 'Not used in the estate, which is itself informative: the work is ' +
      'done from the ends inwards, or by delimiter.',
    impl: function (args, ctx) {
      var s = ctx.text(args[0]); if (V.isErr(s)) return s;
      var st = ctx.int(args[1]); if (V.isErr(st)) return st;
      var n = ctx.int(args[2]); if (V.isErr(n)) return n;
      if (st < 1) return V.value('MID needs start_num of 1 or more');
      if (n < 0) return V.value('MID cannot take a negative number of characters');
      return s.substr(st - 1, n);
    }
  });

  /* --------------------------------------------------------- LEN / TRIM ---- */

  def({
    name: 'LEN', min: 1, max: 1, group: 'text',
    sig: 'LEN(text)', doc: 'How many characters are in some text.',
    estate: 'Used as a blank test inside conditional formatting: LEN(TRIM(A2))=0. ' +
      'Also the fastest way to prove a "blank-looking" cell is not blank.',
    impl: function (args, ctx) {
      var s = ctx.text(args[0]);
      return V.isErr(s) ? s : s.length;
    }
  });

  def({
    name: 'TRIM', min: 1, max: 1, group: 'text',
    sig: 'TRIM(text)',
    doc: 'Removes spaces from the ends, and collapses runs of spaces inside, to one.',
    estate: 'Barely used as a function, though badly needed: the exports arrive ' +
      'with trailing spaces. Note it only removes CHAR(32). A non-breaking space, ' +
      'CHAR(160), goes straight through it.',
    pats: ['D7'],
    impl: function (args, ctx) {
      var s = ctx.text(args[0]);
      if (V.isErr(s)) return s;
      // Deliberately " " and not \s: a tab, a newline and CHAR(160) all survive.
      return s.replace(/^ +| +$/g, '').replace(/ {2,}/g, ' ');
    }
  });

  /* ------------------------------------------------------- UPPER / LOWER --- */

  def({
    name: 'UPPER', min: 1, max: 1, group: 'text', sig: 'UPPER(text)',
    doc: 'Text in capitals.', estate: '2 cells.',
    impl: function (args, ctx) { var s = ctx.text(args[0]); return V.isErr(s) ? s : s.toUpperCase(); }
  });
  def({
    name: 'LOWER', min: 1, max: 1, group: 'text', sig: 'LOWER(text)',
    doc: 'Text in lower case.',
    impl: function (args, ctx) { var s = ctx.text(args[0]); return V.isErr(s) ? s : s.toLowerCase(); }
  });

  /* ---------------------------------------------------------------- VALUE -- */

  def({
    name: 'VALUE', min: 1, max: 1, group: 'text',
    sig: 'VALUE(text)',
    doc: 'Turns text that looks like a number (or a date) into a real number.',
    estate: '7,915 cells, always immediately after a slice: ' +
      'IFERROR(VALUE(TEXTBEFORE(G1," ",1)),""). Slicing text gives you text, even ' +
      'when it looks like a number, so it has to be converted before it can be added.',
    pats: ['P7'],
    impl: function (args, ctx) {
      var v = ctx.scalar(args[0]);
      if (V.isErr(v)) return v;
      if (typeof v === 'number') return v;
      if (typeof v === 'boolean') return V.value('VALUE cannot convert TRUE or FALSE');
      if (V.isEmpty(v)) return 0;
      var n = V.parseNumeric(v);
      if (n !== null) return n;
      var d = XT.D.parseDateText(v);
      if (d !== null) return d;
      return V.value('VALUE could not read ' + JSON.stringify(preview(v)) +
        ' as a number' + nbspHint(v));
    }
  });

  function preview(s) { return s.length > 30 ? s.slice(0, 30) + '…' : s; }

  /* The single most useful diagnostic message in this whole program. */
  function nbspHint(s) {
    if (/ /.test(s)) {
      return '. It contains a NON-BREAKING SPACE, CHAR(160) — invisible, and ' +
        'TRIM will not remove it. Use SUBSTITUTE(text,CHAR(160),"") first.';
    }
    if (/^[ \t]|[ \t]$/.test(s)) {
      return '. It has leading or trailing whitespace; try TRIM first.';
    }
    if (/,/.test(s)) return '. It contains a comma; VALUE wants plain digits.';
    return '';
  }

  /* ----------------------------------------------------------------- TEXT -- */

  def({
    name: 'TEXT', min: 1, max: 2, group: 'text',
    sig: 'TEXT(value, format_text)',
    doc: 'Formats a number as text using a number-format code.',
    estate: '3,346 cells in 30 workbooks, and nearly always the period label: ' +
      'TEXT(AC2,"MMM-YY") turning a date into "Nov-26". Remember the result is ' +
      'TEXT, so it will not sort or compare as a date any more.',
    pats: ['P12'],
    impl: function (args, ctx) {
      var v = ctx.scalar(args[0]);
      if (V.isErr(v)) return v;
      var f = args.length >= 2 ? ctx.text(args[1]) : '';
      if (V.isErr(f)) return f;
      if (f === '') return V.toText(v);     // =TEXT(C1,) with an empty format
      return F.format(v, f);
    }
  });

  /* -------------------------------------------------- CONCAT / & helpers --- */

  def({
    name: 'CONCAT', min: 1, max: 255, group: 'text', arrayArgs: [0, 1, 2, 3, 4, 5],
    sig: 'CONCAT(text1, [text2], …)',
    doc: 'Joins text together. Ranges are allowed.',
    estate: 'The estate uses the & operator instead, everywhere.',
    impl: function (args, ctx) {
      var out = '';
      for (var i = 0; i < args.length; i++) {
        var a = args[i];
        if (V.isRange(a) || V.isArr(a)) {
          var view = ctx.asView(a);
          if (V.isErr(view)) return view;
          var all = view.flat();
          for (var k = 0; k < all.length; k++) {
            if (V.isErr(all[k])) return all[k];
            var t1 = V.toText(all[k]); if (V.isErr(t1)) return t1;
            out += t1;
          }
        } else {
          var t = V.toText(ctx.scalar(a)); if (V.isErr(t)) return t;
          out += t;
        }
      }
      return out;
    }
  });

  def({
    name: 'CONCATENATE', min: 1, max: 255, group: 'text',
    sig: 'CONCATENATE(text1, [text2], …)',
    doc: 'The older form of CONCAT. Does not accept ranges.',
    estate: '2 cells. The & operator won.',
    impl: function (args, ctx) {
      var out = '';
      for (var i = 0; i < args.length; i++) {
        var t = V.toText(ctx.scalar(args[i])); if (V.isErr(t)) return t;
        out += t;
      }
      return out;
    }
  });

  /* ----------------------------------------------- SUBSTITUTE / CHAR / CODE */

  def({
    name: 'SUBSTITUTE', min: 3, max: 4, group: 'text',
    sig: 'SUBSTITUTE(text, old_text, new_text, [instance_num])',
    doc: 'Replaces one piece of text with another. Case-sensitive.',
    estate: '1,564 cells, all of them SUBSTITUTE(D1,CHAR(160)," ") -- stripping ' +
      'non-breaking spaces out of a PLU export so that VALUE() will work on it.',
    pats: ['D7'],
    impl: function (args, ctx) {
      var s = ctx.text(args[0]); if (V.isErr(s)) return s;
      var oldT = ctx.text(args[1]); if (V.isErr(oldT)) return oldT;
      var newT = ctx.text(args[2]); if (V.isErr(newT)) return newT;
      if (oldT === '') return s;
      if (args.length < 4 || (args.nodes && args.nodes[3] && args.nodes[3].t === 'missing')) {
        return s.split(oldT).join(newT);
      }
      var inst = ctx.int(args[3]); if (V.isErr(inst)) return inst;
      if (inst < 1) return V.value('instance_num must be 1 or more');
      var at = -1, count = 0;
      for (;;) {
        at = s.indexOf(oldT, at + 1);
        if (at < 0) return s;
        if (++count === inst) return s.slice(0, at) + newT + s.slice(at + oldT.length);
      }
    }
  });

  def({
    name: 'CHAR', min: 1, max: 1, group: 'text',
    sig: 'CHAR(number)', doc: 'The character with a given code. CHAR(160) is a non-breaking space.',
    estate: '1,564 cells, only ever CHAR(160) inside a SUBSTITUTE.',
    pats: ['D7'],
    impl: function (args, ctx) {
      var n = ctx.int(args[0]); if (V.isErr(n)) return n;
      if (n < 1 || n > 65535) return V.value('CHAR needs a code between 1 and 255');
      return String.fromCharCode(n);
    }
  });

  def({
    name: 'CODE', min: 1, max: 1, group: 'text',
    sig: 'CODE(text)', doc: 'The numeric code of the first character. Useful for ' +
      'identifying an invisible character.',
    impl: function (args, ctx) {
      var s = ctx.text(args[0]); if (V.isErr(s)) return s;
      if (s === '') return V.value('CODE needs some text');
      return s.charCodeAt(0);
    }
  });

  /* ------------------------------------------------------- FIND / SEARCH --- */

  def({
    name: 'FIND', min: 2, max: 3, group: 'text',
    sig: 'FIND(find_text, within_text, [start_num])',
    doc: 'Where one piece of text starts inside another. Case-SENSITIVE. ' +
      '#VALUE! if it is not there.',
    estate: '103 cells.',
    impl: function (args, ctx) {
      var needle = ctx.text(args[0]); if (V.isErr(needle)) return needle;
      var hay = ctx.text(args[1]); if (V.isErr(hay)) return hay;
      var st = args.length >= 3 ? ctx.int(args[2]) : 1;
      if (V.isErr(st)) return st;
      if (st < 1) return V.value('start_num must be 1 or more');
      var at = hay.indexOf(needle, st - 1);
      if (at < 0) return V.value('FIND could not find ' + JSON.stringify(preview(needle)) +
        ' inside ' + JSON.stringify(preview(hay)));
      return at + 1;
    }
  });

  def({
    name: 'SEARCH', min: 2, max: 3, group: 'text',
    sig: 'SEARCH(find_text, within_text, [start_num])',
    doc: 'Like FIND, but case-INSENSITIVE and it understands * and ?.',
    estate: 'Used inside conditional formatting: NOT(ISERROR(SEARCH("p",D7))).',
    impl: function (args, ctx) {
      var needle = ctx.text(args[0]); if (V.isErr(needle)) return needle;
      var hay = ctx.text(args[1]); if (V.isErr(hay)) return hay;
      var st = args.length >= 3 ? ctx.int(args[2]) : 1;
      if (V.isErr(st)) return st;
      if (st < 1) return V.value('start_num must be 1 or more');
      var tail = hay.slice(st - 1);
      if (/[*?]/.test(needle)) {
        var re = new RegExp(needle.replace(/[.+^${}()|[\]\\]/g, '\\$&')
          .replace(/\*/g, '[\\s\\S]*').replace(/\?/g, '[\\s\\S]'), 'i');
        var m = re.exec(tail);
        if (!m) return V.value('SEARCH found no match for ' + JSON.stringify(preview(needle)));
        return st + m.index;
      }
      var at = tail.toUpperCase().indexOf(needle.toUpperCase());
      if (at < 0) return V.value('SEARCH could not find ' + JSON.stringify(preview(needle)) +
        ' inside ' + JSON.stringify(preview(hay)));
      return st + at;
    }
  });

  /* ------------------------------------------- TEXTBEFORE / TEXTAFTER ------ */

  /* These are the modern answer to pattern P6. instance_num may be NEGATIVE,
     which counts delimiters from the end -- TEXTAFTER(M2," ",-1) is "everything
     after the last space", and it appears in the estate exactly for that. */
  function occurrences(hay, delim, caseInsensitive) {
    var out = [], h = caseInsensitive ? hay.toUpperCase() : hay;
    var d = caseInsensitive ? delim.toUpperCase() : delim;
    if (d === '') return out;
    var at = h.indexOf(d);
    while (at >= 0) { out.push(at); at = h.indexOf(d, at + d.length); }
    return out;
  }

  function textSplitFn(name, after) {
    def({
      name: name, min: 2, max: 6, group: 'text',
      sig: name + '(text, delimiter, [instance_num], [match_mode], [match_end], [if_not_found])',
      doc: after
        ? 'Everything after the nth occurrence of a delimiter.'
        : 'Everything before the nth occurrence of a delimiter.',
      estate: 'Counting FIELDS instead of characters. TEXTAFTER(F1," ",3) takes the ' +
        '3rd space-delimited field; a negative instance counts from the end. This is ' +
        'the fix for the brittle RIGHT(…,29) pattern that sits in the same workbook.',
      pats: ['P7'],
      impl: function (args, ctx) {
        var s = ctx.text(args[0]); if (V.isErr(s)) return s;
        var d = ctx.text(args[1]); if (V.isErr(d)) return d;
        var inst = 1;
        if (args.length >= 3 && args.nodes && args.nodes[2] && args.nodes[2].t !== 'missing') {
          inst = ctx.int(args[2]); if (V.isErr(inst)) return inst;
        }
        var ci = false;
        if (args.length >= 4 && args.nodes && args.nodes[3] && args.nodes[3].t !== 'missing') {
          var mm = ctx.int(args[3]); if (V.isErr(mm)) return mm;
          ci = mm === 1;
        }
        var notFound = (args.length >= 6 && args.nodes && args.nodes[5] &&
          args.nodes[5].t !== 'missing') ? ctx.scalar(args[5]) : undefined;

        if (inst === 0) return V.value('instance_num cannot be 0');
        if (d === '') return after ? s : '';

        var occ = occurrences(s, d, ci);
        var k = inst > 0 ? inst - 1 : occ.length + inst;
        if (k < 0 || k >= occ.length) {
          if (notFound !== undefined) return notFound;
          return V.na(name + ' wanted delimiter number ' +
            (inst > 0 ? inst : '(' + inst + ' from the end)') + ' but ' +
            JSON.stringify(preview(s)) + ' contains ' + occ.length + ' of them');
        }
        var at = occ[k];
        return after ? s.slice(at + d.length) : s.slice(0, at);
      }
    });
  }
  textSplitFn('TEXTAFTER', true);
  textSplitFn('TEXTBEFORE', false);

})(window.XT = window.XT || {});
