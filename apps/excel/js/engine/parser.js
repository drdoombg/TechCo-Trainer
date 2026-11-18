/* =============================================================================
   Excel Trainer — engine/parser.js
   Recursive descent, one function per precedence level. Read it top to bottom
   and you have read Excel's operator precedence.

   Loosest binding at the top, tightest at the bottom:

        =  <>  <  <=  >  >=        comparison
        &                          concatenation
        +  -                       add / subtract
        *  /                       multiply / divide
        ^                          power
        +x  -x                     unary sign   (binds TIGHTER than ^ ,
                                                 so -2^2 is 4, not -4)
        x%                         percent
        :                          range          (tightest of all)

   Every node carries p/q, the span of source text it came from. The
   Evaluate-Formula stepper walks the tree and uses those offsets to underline
   the sub-expression it is about to replace with a value.

   Node shapes
     {t:'num',  v}                       {t:'str', v}
     {t:'bool', v}                       {t:'err', v}          #REF! etc
     {t:'missing'}                        an argument written as nothing: TEXT(C1,)
     {t:'ref',   sheet, ext, addr}        one cell
     {t:'range', sheet, ext, kind, ...}   kind = 'cells' | 'col' | 'row'
     {t:'name',  name}                    a defined name
     {t:'table', table, spec}             Table4[[#This Row],[SUPPLIER]]
     {t:'func',  name, args}
     {t:'bin',   op, l, r}   {t:'un', op, x}   {t:'post', op, x}
     {t:'paren', x}          {t:'arr', rows}
   ============================================================================= */
(function (XT) {
  'use strict';

  var P = {};
  XT.P = P;

  var CMP = ['=', '<>', '<', '<=', '>', '>='];

  function Parser(tokens, src) {
    this.toks = tokens;
    this.i = 0;
    this.src = src;
  }
  Parser.prototype.peek = function () { return this.toks[this.i]; };
  Parser.prototype.next = function () { return this.toks[this.i++]; };
  Parser.prototype.at = function (t, v) {
    var k = this.toks[this.i];
    return k.t === t && (v === undefined || k.v === v);
  };
  Parser.prototype.eat = function (t, v) {
    if (this.at(t, v)) return this.next();
    return null;
  };
  Parser.prototype.expect = function (t, v) {
    var k = this.eat(t, v);
    if (!k) XT.T.fail('expected ' + (v || t) + ' but found ' +
      (this.peek().t === 'eof' ? 'end of formula' : JSON.stringify(String(this.peek().v))),
      this.peek().p);
    return k;
  };

  function span(a, b) { return { p: a.p, q: b.q !== undefined ? b.q : b.p }; }
  function mk(node, p, q) { node.p = p; node.q = q; return node; }

  /* ----------------------------------------------------------- precedence 1 */

  Parser.prototype.parseExpr = function () { return this.parseCmp(); };

  Parser.prototype.parseCmp = function () {
    var l = this.parseCat();
    while (this.peek().t === 'op' && CMP.indexOf(this.peek().v) >= 0) {
      var op = this.next().v;
      var r = this.parseCat();
      l = mk({ t: 'bin', op: op, l: l, r: r }, l.p, r.q);
    }
    return l;
  };

  Parser.prototype.parseCat = function () {
    var l = this.parseAdd();
    while (this.at('op', '&')) {
      this.next();
      var r = this.parseAdd();
      l = mk({ t: 'bin', op: '&', l: l, r: r }, l.p, r.q);
    }
    return l;
  };

  Parser.prototype.parseAdd = function () {
    var l = this.parseMul();
    while (this.at('op', '+') || this.at('op', '-')) {
      var op = this.next().v;
      var r = this.parseMul();
      l = mk({ t: 'bin', op: op, l: l, r: r }, l.p, r.q);
    }
    return l;
  };

  Parser.prototype.parseMul = function () {
    var l = this.parsePow();
    while (this.at('op', '*') || this.at('op', '/')) {
      var op = this.next().v;
      var r = this.parsePow();
      l = mk({ t: 'bin', op: op, l: l, r: r }, l.p, r.q);
    }
    return l;
  };

  // Left-associative, so 2^3^2 is (2^3)^2 = 64.
  // Flagged in OPEN_QUESTIONS.md for confirmation against real Excel.
  Parser.prototype.parsePow = function () {
    var l = this.parseUnary();
    while (this.at('op', '^')) {
      this.next();
      var r = this.parseUnary();
      l = mk({ t: 'bin', op: '^', l: l, r: r }, l.p, r.q);
    }
    return l;
  };

  /* The leading unary + is what makes  =+K13+1  legal. It is a Lotus 1-2-3
     habit that survives in 19,683 cells of the source estate plus 14,568 more
     of the form =+Home!AS1, so it is not optional for us. */
  Parser.prototype.parseUnary = function () {
    if (this.at('op', '-') || this.at('op', '+')) {
      var t = this.next();
      var x = this.parseUnary();
      return mk({ t: 'un', op: t.v, x: x }, t.p, x.q);
    }
    return this.parsePostfix();
  };

  Parser.prototype.parsePostfix = function () {
    var x = this.parsePrimary();
    while (this.at('op', '%')) {
      var t = this.next();
      x = mk({ t: 'post', op: '%', x: x }, x.p, t.q);
    }
    return x;
  };

  /* --------------------------------------------------------------- primaries */

  Parser.prototype.parsePrimary = function () {
    var k = this.peek();

    if (k.t === 'num') { this.next(); return mk({ t: 'num', v: k.v }, k.p, k.q); }
    if (k.t === 'str') { this.next(); return mk({ t: 'str', v: k.v }, k.p, k.q); }
    if (k.t === 'err') { this.next(); return mk({ t: 'err', v: k.v }, k.p, k.q); }

    if (k.t === '(') {
      this.next();
      var inner = this.parseExpr();
      var close = this.expect(')');
      return mk({ t: 'paren', x: inner }, k.p, close.q);
    }

    if (k.t === '{') return this.parseArrayConst();

    if (k.t === 'func') return this.parseCall();

    if (k.t === 'table') {
      this.next();
      return mk({ t: 'table', table: k.v.table, spec: k.v.spec }, k.p, k.q);
    }

    // A reference, with or without a sheet prefix.
    if (k.t === 'sheet' || k.t === 'cell' || k.t === 'colrange' || k.t === 'rowrange') {
      return this.parseRef();
    }

    if (k.t === 'ident') {
      this.next();
      var up = String(k.v).toUpperCase();
      if (up === 'TRUE') return mk({ t: 'bool', v: true }, k.p, k.q);
      if (up === 'FALSE') return mk({ t: 'bool', v: false }, k.p, k.q);
      return mk({ t: 'name', name: k.v }, k.p, k.q);
    }

    XT.T.fail('unexpected ' + (k.t === 'eof' ? 'end of formula' : JSON.stringify(String(k.v))), k.p);
  };

  Parser.prototype.parseArrayConst = function () {
    var open = this.expect('{');
    var rows = [[]];
    if (!this.at('}')) {
      for (;;) {
        rows[rows.length - 1].push(this.parseExpr());
        if (this.eat(',')) continue;
        if (this.eat(';')) { rows.push([]); continue; }
        break;
      }
    }
    var close = this.expect('}');
    return mk({ t: 'arr', rows: rows }, open.p, close.q);
  };

  Parser.prototype.parseCall = function () {
    var nameTok = this.next();                // 'func' token
    this.expect('(');
    var args = [];
    if (!this.at(')')) {
      for (;;) {
        if (this.at(',') || this.at(')')) {
          // An argument written as nothing at all. Real formulas do this:
          //   =TEXT(C1,)                   1,564 cells
          //   =IF(S3="Asda",V3,)           the trailing-comma IF (defect D4)
          var here = this.peek();
          args.push(mk({ t: 'missing' }, here.p, here.p));
        } else {
          args.push(this.parseExpr());
        }
        if (this.eat(',')) continue;
        break;
      }
    }
    var close = this.expect(')');
    // Strip the on-disk _xlfn. prefix: _xlfn.XLOOKUP and XLOOKUP are one thing.
    var nm = String(nameTok.v).toUpperCase().replace(/^_XLFN\./, '').replace(/^_XLWS\./, '');
    return mk({ t: 'func', name: nm, raw: nameTok.v, args: args }, nameTok.p, close.q);
  };

  /* ---------------------------------------------------------------- refs ---- */

  /* Handles, in one place:
         A1          $A$1        A1:B2       A:A        $A:$DN       17:17
         Sheet1!A1   'Micro paste'!$M:$M     [1]CODES!$G$4:$I$4975
         'QTR 2 Old'!#REF!
     A sheet prefix applies to the whole range, which is how Excel writes it. */
  Parser.prototype.parseRef = function () {
    var start = this.peek();
    var sheet = null, ext = null;

    if (this.at('sheet')) {
      var sh = this.next();
      sheet = sh.v.name;
      ext = sh.v.ext;          // "[1]" etc; kept so we can report "link not available"
    }

    var k = this.peek();

    // A sheet-qualified error literal, e.g. 'QTR 2 Old'!#REF! -- 1,504 cells.
    if (k.t === 'err') { this.next(); return mk({ t: 'err', v: k.v }, start.p, k.q); }

    if (k.t === 'colrange') {
      this.next();
      return mk({ t: 'range', kind: 'col', sheet: sheet, ext: ext,
        c1: k.v.c1, c2: k.v.c2, a1: k.v.a1, a2: k.v.a2 }, start.p, k.q);
    }

    if (k.t === 'rowrange') {
      this.next();
      return mk({ t: 'range', kind: 'row', sheet: sheet, ext: ext,
        r1: k.v.r1, r2: k.v.r2, a1: k.v.a1, a2: k.v.a2 }, start.p, k.q);
    }

    if (k.t === 'cell') {
      this.next();
      // A range written out in full: A1:B2
      if (this.at('op', ':')) {
        this.next();
        // The right-hand side may repeat the sheet name; Excel allows it.
        if (this.at('sheet')) this.next();
        var k2 = this.peek();
        if (k2.t === 'err') { this.next(); return mk({ t: 'err', v: k2.v }, start.p, k2.q); }
        if (k2.t !== 'cell') XT.T.fail('expected a cell address after ":"', k2.p);
        this.next();
        return mk({ t: 'range', kind: 'cells', sheet: sheet, ext: ext,
          a: k.v, b: k2.v }, start.p, k2.q);
      }
      return mk({ t: 'ref', sheet: sheet, ext: ext, addr: k.v }, start.p, k.q);
    }

    // A sheet prefix in front of a defined name: picklists!Shift
    if (k.t === 'ident') {
      this.next();
      return mk({ t: 'name', name: k.v, sheet: sheet }, start.p, k.q);
    }

    XT.T.fail('expected a cell or range after the sheet name', k.p);
  };

  /* --------------------------------------------------------------- entry ---- */

  /* parse("=IFERROR(VLOOKUP(A2,PASTE!$A:$D,2,FALSE),\"\")")
       -> {ok:true, ast:…}
       -> {ok:false, error:"…", pos:n}   on a syntax error
     Callers that want a throw can use parseOrThrow. */
  P.parse = function (src) {
    try {
      var toks = XT.T.tokenize(src);
      var p = new Parser(toks, String(src));
      var ast = p.parseExpr();
      if (!p.at('eof')) {
        XT.T.fail('unexpected extra text', p.peek().p);
      }
      return { ok: true, ast: ast, src: String(src) };
    } catch (e) {
      if (e.name === 'FormulaError') return { ok: false, error: e.message, pos: e.pos };
      throw e;
    }
  };

  P.parseOrThrow = function (src) {
    var r = P.parse(src);
    if (!r.ok) {
      var e = new Error(r.error);
      e.name = 'FormulaError';
      e.pos = r.pos;
      throw e;
    }
    return r.ast;
  };

  /* --------------------------------------------- render an AST back to text  */

  /* Used by the fill handle (after translating addresses) and by the
     "formula shape" panel. Round-tripping through text keeps one source of
     truth for the formula and avoids a second, divergent writer. */
  P.unparse = function (node) {
    var u = P.unparse;
    switch (node.t) {
      case 'num': return XT.V.numToText(node.v);
      case 'str': return '"' + String(node.v).replace(/"/g, '""') + '"';
      // A value substituted back into the formula by the Evaluate stepper.
      case 'lit': return XT.P.litText(node.val);
      case 'bool': return node.v ? 'TRUE' : 'FALSE';
      case 'err': return node.v;
      case 'missing': return '';
      case 'paren': return '(' + u(node.x) + ')';
      case 'un': return node.op + u(node.x);
      case 'post': return u(node.x) + node.op;
      case 'bin': return u(node.l) + node.op + u(node.r);
      case 'arr':
        return '{' + node.rows.map(function (row) {
          return row.map(u).join(',');
        }).join(';') + '}';
      case 'func':
        // Keep the on-disk _xlfn. spelling if that is how it was written, so a
        // formula survives a round-trip through the fill handle unchanged.
        return (/^_xlfn\./i.test(node.raw || '') ? node.raw : node.name) +
          '(' + node.args.map(u).join(',') + ')';
      case 'name':
        return (node.sheet ? XT.R.quoteSheet(node.sheet) + '!' : '') + node.name;
      case 'table': return node.table + node.spec;
      case 'ref':
        return refPrefix(node) + XT.R.formatAddr(node.addr);
      case 'range':
        if (node.kind === 'col') {
          return refPrefix(node) + (node.a1 ? '$' : '') + XT.R.numToCol(node.c1) +
            ':' + (node.a2 ? '$' : '') + XT.R.numToCol(node.c2);
        }
        if (node.kind === 'row') {
          return refPrefix(node) + (node.a1 ? '$' : '') + (node.r1 + 1) +
            ':' + (node.a2 ? '$' : '') + (node.r2 + 1);
        }
        return refPrefix(node) + XT.R.formatAddr(node.a) + ':' + XT.R.formatAddr(node.b);
    }
    return '?';
  };

  /* How a computed value looks when written back into a formula. Excel's
     Evaluate Formula dialog does the same substitution. */
  P.litText = function (v) {
    var V = XT.V;
    if (V.isErr(v)) return v.code;
    if (V.isEmpty(v)) return '0';            // an empty cell reads as 0 mid-formula
    if (typeof v === 'string') return '"' + v.replace(/"/g, '""') + '"';
    if (typeof v === 'boolean') return v ? 'TRUE' : 'FALSE';
    if (typeof v === 'number') return V.numToText(v);
    if (V.isArr(v)) {
      var rows = v.rows.map(function (row) {
        return row.map(function (x) { return P.litText(x); }).join(',');
      });
      return '{' + rows.join(';') + '}';
    }
    if (V.isRange(v)) {
      return (v.sheet ? XT.R.quoteSheet(v.sheet) + '!' : '') +
        XT.R.numToCol(v.c1) + (v.r1 + 1) +
        (v.isSingle() ? '' : ':' + XT.R.numToCol(v.c2) + (v.r2 + 1));
    }
    return String(v);
  };

  function refPrefix(node) {
    var s = '';
    if (node.ext) s += node.ext;
    if (node.sheet) s += XT.R.quoteSheet(node.sheet) + '!';
    return s;
  }
  P.refPrefix = refPrefix;

  /* ------------------------------------------- translate for fill-down/right */

  /* Deep-copy an AST, shifting every reference by (dr, dc) unless its $ says
     otherwise. This IS fill-down. Whole-column and whole-row references never
     move, which is why =VLOOKUP(A2,PASTE!$A:$D,2,0) survives being filled a
     thousand rows while =VLOOKUP(A2,PASTE!A2:D500,2,0) quietly rots. */
  P.translate = function (node, dr, dc) {
    var tr = function (n) { return P.translate(n, dr, dc); };
    switch (node.t) {
      case 'num': case 'str': case 'bool': case 'err': case 'missing':
      case 'name': case 'table':
        return shallow(node);
      case 'paren': return mk2({ t: 'paren', x: tr(node.x) }, node);
      case 'un': return mk2({ t: 'un', op: node.op, x: tr(node.x) }, node);
      case 'post': return mk2({ t: 'post', op: node.op, x: tr(node.x) }, node);
      case 'bin': return mk2({ t: 'bin', op: node.op, l: tr(node.l), r: tr(node.r) }, node);
      case 'arr': return mk2({ t: 'arr', rows: node.rows.map(function (row) { return row.map(tr); }) }, node);
      case 'func': return mk2({ t: 'func', name: node.name, raw: node.raw, args: node.args.map(tr) }, node);

      case 'ref': {
        var a = XT.R.shiftAddr(node.addr, dr, dc);
        if (!a) return mk2({ t: 'err', v: '#REF!' }, node);
        return mk2({ t: 'ref', sheet: node.sheet, ext: node.ext, addr: a }, node);
      }
      case 'range': {
        if (node.kind === 'col') {
          var c1 = node.a1 ? node.c1 : node.c1 + dc;
          var c2 = node.a2 ? node.c2 : node.c2 + dc;
          if (c1 < 0 || c2 < 0) return mk2({ t: 'err', v: '#REF!' }, node);
          return mk2({ t: 'range', kind: 'col', sheet: node.sheet, ext: node.ext,
            c1: c1, c2: c2, a1: node.a1, a2: node.a2 }, node);
        }
        if (node.kind === 'row') {
          var r1 = node.a1 ? node.r1 : node.r1 + dr;
          var r2 = node.a2 ? node.r2 : node.r2 + dr;
          if (r1 < 0 || r2 < 0) return mk2({ t: 'err', v: '#REF!' }, node);
          return mk2({ t: 'range', kind: 'row', sheet: node.sheet, ext: node.ext,
            r1: r1, r2: r2, a1: node.a1, a2: node.a2 }, node);
        }
        var na = XT.R.shiftAddr(node.a, dr, dc);
        var nb = XT.R.shiftAddr(node.b, dr, dc);
        if (!na || !nb) return mk2({ t: 'err', v: '#REF!' }, node);
        return mk2({ t: 'range', kind: 'cells', sheet: node.sheet, ext: node.ext,
          a: na, b: nb }, node);
      }
    }
    return shallow(node);
  };

  function shallow(n) {
    var o = {};
    for (var k in n) if (Object.prototype.hasOwnProperty.call(n, k)) o[k] = n[k];
    return o;
  }
  function mk2(n, from) { n.p = from.p; n.q = from.q; return n; }

  /* Translate a whole formula string. This is what Ctrl+D calls.
     Returns the new text, or the original if it will not parse. */
  P.translateText = function (src, dr, dc) {
    var r = P.parse(src);
    if (!r.ok) return src;
    return '=' + P.unparse(P.translate(r.ast, dr, dc));
  };

})(window.XT = window.XT || {});
