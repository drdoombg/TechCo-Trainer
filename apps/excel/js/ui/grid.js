/* =============================================================================
   Excel Trainer — ui/grid.js
   The grid: selection, keyboard, editing, the fill handle, and paste.

   Priorities here come straight from the catalog. Fill-down with reference
   translation is the main interaction (§7.9) -- the source estate has 3.07
   million fill-down child cells and essentially no dynamic arrays -- and paste
   is "the trainer's defining gesture" (§12.6), so both get first-class
   treatment rather than being conveniences bolted on.

   Rendering is a plain <table>, built once per sheet and then updated in place.
   At trainer scale (a few hundred rows) that is simpler and quicker than
   virtualising, and it means every cell is a real DOM node the instruments can
   point an arrow at.
   ============================================================================= */
(function (XT) {
  'use strict';

  var V = XT.V, R = XT.R, F = XT.F, P = XT.P, S = XT.S;

  var MIN_ROWS = 40, MIN_COLS = 16, PAD_ROWS = 12, PAD_COLS = 4;
  var DEFAULT_W = 92;

  function el(tag, cls, text) {
    var e = document.createElement(tag);
    if (cls) e.className = cls;
    if (text !== undefined) e.textContent = text;
    return e;
  }

  function Grid(host, opts) {
    this.host = host;
    this.opts = opts || {};
    this.wb = null;
    this.sheetName = null;
    this.sel = { r: 0, c: 0, r2: 0, c2: 0 };
    this.editing = null;
    this.showFormulas = false;
    this.cellEls = [];
    this.nRows = MIN_ROWS;
    this.nCols = MIN_COLS;
    this.undoStack = [];
    this.redoStack = [];
    this.clip = null;              // internal copy buffer {cut, rows:[[input]]}
    this.highlight = [];           // ranges outlined while editing
    this.marks = {};               // "r,c" -> css class, for the tracer
    this.readOnlyCells = null;     // "r,c" -> true, used by exercises
    this.build();
  }

  /* ------------------------------------------------------------------ build */

  Grid.prototype.build = function () {
    var self = this;
    this.host.classList.add('xt-grid-host');
    this.host.innerHTML = '';

    this.scroller = el('div', 'xt-scroller');
    this.table = el('table', 'xt-grid');
    this.thead = el('thead');
    this.tbody = el('tbody');
    this.table.appendChild(this.thead);
    this.table.appendChild(this.tbody);
    this.scroller.appendChild(this.table);

    // The editor is one input, moved over whichever cell is being edited.
    this.editor = el('input', 'xt-editor');
    this.editor.setAttribute('spellcheck', 'false');
    this.editor.setAttribute('autocomplete', 'off');
    this.scroller.appendChild(this.editor);

    // The fill handle: the little square you drag.
    this.handle = el('div', 'xt-fill-handle');
    this.scroller.appendChild(this.handle);

    this.host.appendChild(this.scroller);

    this.scroller.addEventListener('mousedown', function (e) { self.onMouseDown(e); });
    this.scroller.addEventListener('mousemove', function (e) { self.onMouseMove(e); });
    document.addEventListener('mouseup', function (e) { self.onMouseUp(e); });
    this.scroller.addEventListener('dblclick', function (e) { self.onDblClick(e); });
    this.scroller.setAttribute('tabindex', '0');
    this.scroller.addEventListener('keydown', function (e) { self.onKeyDown(e); });
    this.editor.addEventListener('keydown', function (e) { self.onEditorKey(e); });
    this.editor.addEventListener('input', function () { self.onEditorInput(); });
    this.host.addEventListener('paste', function (e) { self.onPaste(e); });
    this.host.addEventListener('copy', function (e) { self.onCopy(e, false); });
    this.host.addEventListener('cut', function (e) { self.onCopy(e, true); });
  };

  /* ------------------------------------------------------------- workbook -- */

  Grid.prototype.setWorkbook = function (wb, sheetName) {
    this.wb = wb;
    this.sheetName = sheetName || (wb.sheets[0] && wb.sheets[0].name);
    this.undoStack = []; this.redoStack = [];
    this.sel = { r: 0, c: 0, r2: 0, c2: 0 };
    this.layout();
    this.refresh();
  };

  Grid.prototype.setSheet = function (name) {
    if (!this.wb.sheet(name)) return;
    this.commitEdit();
    this.sheetName = name;
    this.sel = { r: 0, c: 0, r2: 0, c2: 0 };
    this.marks = {};
    this.layout();
    this.refresh();
    this.emit('select');
  };

  Grid.prototype.sheet = function () { return this.wb && this.wb.sheet(this.sheetName); };

  Grid.prototype.emit = function (name) {
    if (this.opts['on' + name.charAt(0).toUpperCase() + name.slice(1)]) {
      this.opts['on' + name.charAt(0).toUpperCase() + name.slice(1)](this);
    }
  };

  /* --------------------------------------------------------------- layout -- */

  // Work out how many rows and columns to draw, then build the table skeleton.
  Grid.prototype.layout = function () {
    var sh = this.sheet();
    if (!sh) return;
    var used = sh.used();
    this.nRows = Math.max(MIN_ROWS, used.maxRow + 1 + PAD_ROWS);
    this.nCols = Math.max(MIN_COLS, used.maxCol + 1 + PAD_COLS);

    this.thead.innerHTML = '';
    this.tbody.innerHTML = '';
    this.cellEls = [];

    var hr = el('tr');
    var corner = el('th', 'xt-corner');
    hr.appendChild(corner);
    for (var c = 0; c < this.nCols; c++) {
      var th = el('th', 'xt-colhead', R.numToCol(c));
      th.style.width = (sh.colWidth[c] || DEFAULT_W) + 'px';
      th.dataset.c = c;
      hr.appendChild(th);
    }
    this.thead.appendChild(hr);

    var frag = document.createDocumentFragment();
    for (var r = 0; r < this.nRows; r++) {
      var tr = el('tr');
      var rh = el('th', 'xt-rowhead', String(r + 1));
      rh.dataset.r = r;
      tr.appendChild(rh);
      var row = [];
      for (var c2 = 0; c2 < this.nCols; c2++) {
        var td = el('td', 'xt-cell');
        td.dataset.r = r; td.dataset.c = c2;
        tr.appendChild(td);
        row.push(td);
      }
      this.cellEls.push(row);
      frag.appendChild(tr);
    }
    this.tbody.appendChild(frag);
  };

  /* -------------------------------------------------------------- refresh -- */

  /* Recalculate, then push values into the DOM. Called after every edit.
     Only touches a cell's text when it has actually changed, so the browser
     does almost no work on a big sheet where one number moved. */
  Grid.prototype.refresh = function (skipRecalc) {
    var sh = this.sheet();
    if (!sh) return;
    if (!skipRecalc) this.wb.recalc();

    var used = sh.used();
    if (used.maxRow + 1 > this.nRows || used.maxCol + 1 > this.nCols) {
      this.layout();
    }

    for (var r = 0; r < this.nRows; r++) {
      for (var c = 0; c < this.nCols; c++) {
        var td = this.cellEls[r][c];
        var cell = sh.get(r, c);
        var txt = '', cls = 'xt-cell', title = '';

        if (cell && (cell.input !== '' || cell.ast)) {
          var val = this.wb.cellValue(sh, r, c);
          if (this.showFormulas) {
            txt = cell.input;
            cls += ' xt-showf';
          } else {
            // wb.display, not F.format: it knows that a formula reading an empty
            // cell shows 0 rather than nothing. See the note in workbook.js.
            txt = this.wb.display(sh, r, c);
            cls += ' xt-al-' + F.alignOf(V.isEmpty(val) && cell.ast ? 0 : val);
          }
          if (V.isErr(val)) { cls += ' xt-err'; title = val.code + (val.detail ? ' — ' + val.detail : ''); }
          if (cell.ast) cls += ' xt-formula';
          if (cell.cse) cls += ' xt-cse';
          if (cell.hidden && cell.hidden.length) cls += ' xt-hidden-err';
          if (typeof val === 'string' && val === '') cls += ' xt-emptytext';
          if (cell.circular) { cls += ' xt-err'; title = 'circular reference'; }
        }
        if (this.readOnlyCells && this.readOnlyCells[r + ',' + c]) cls += ' xt-locked';
        if (this.marks[r + ',' + c]) cls += ' ' + this.marks[r + ',' + c];

        if (td.textContent !== txt) td.textContent = txt;
        if (td.className !== cls) td.className = cls;
        if (td.title !== title) td.title = title;
      }
    }
    this.paintSelection();
    this.emit('refresh');
  };

  /* ------------------------------------------------------------ selection -- */

  Grid.prototype.normSel = function () {
    var s = this.sel;
    return {
      r1: Math.min(s.r, s.r2), r2: Math.max(s.r, s.r2),
      c1: Math.min(s.c, s.c2), c2: Math.max(s.c, s.c2)
    };
  };

  Grid.prototype.paintSelection = function () {
    var n = this.normSel();
    for (var r = 0; r < this.nRows; r++) {
      for (var c = 0; c < this.nCols; c++) {
        var td = this.cellEls[r][c];
        var inSel = r >= n.r1 && r <= n.r2 && c >= n.c1 && c <= n.c2;
        td.classList.toggle('xt-sel', inSel);
        td.classList.toggle('xt-active', r === this.sel.r && c === this.sel.c);
      }
    }
    // Header highlight, so you can see where you are without hunting.
    var ths = this.thead.querySelectorAll('.xt-colhead');
    for (var i = 0; i < ths.length; i++) {
      ths[i].classList.toggle('xt-head-on', i >= n.c1 && i <= n.c2);
    }
    var rhs = this.tbody.querySelectorAll('.xt-rowhead');
    for (var j = 0; j < rhs.length; j++) {
      rhs[j].classList.toggle('xt-head-on', j >= n.r1 && j <= n.r2);
    }
    this.placeHandle();
  };

  Grid.prototype.placeHandle = function () {
    var n = this.normSel();
    var td = this.cellEls[n.r2] && this.cellEls[n.r2][n.c2];
    if (!td || this.editing) { this.handle.style.display = 'none'; return; }
    this.handle.style.display = 'block';
    this.handle.style.left = (td.offsetLeft + td.offsetWidth - 4) + 'px';
    this.handle.style.top = (td.offsetTop + td.offsetHeight - 4) + 'px';
  };

  Grid.prototype.select = function (r, c, extend) {
    r = Math.max(0, Math.min(r, this.nRows - 1));
    c = Math.max(0, Math.min(c, this.nCols - 1));
    if (extend) { this.sel.r2 = r; this.sel.c2 = c; }
    else { this.sel = { r: r, c: c, r2: r, c2: c }; }
    this.paintSelection();
    this.scrollIntoView(r, c);
    this.emit('select');
  };

  Grid.prototype.scrollIntoView = function (r, c) {
    var td = this.cellEls[r] && this.cellEls[r][c];
    if (!td) return;
    var sc = this.scroller;
    var left = td.offsetLeft, top = td.offsetTop;
    var w = td.offsetWidth, h = td.offsetHeight;
    var headW = 44, headH = 24;
    if (left - headW < sc.scrollLeft) sc.scrollLeft = Math.max(0, left - headW - 2);
    if (left + w > sc.scrollLeft + sc.clientWidth) sc.scrollLeft = left + w - sc.clientWidth + 2;
    if (top - headH < sc.scrollTop) sc.scrollTop = Math.max(0, top - headH - 2);
    if (top + h > sc.scrollTop + sc.clientHeight) sc.scrollTop = top + h - sc.clientHeight + 2;
  };

  Grid.prototype.addr = function () {
    var n = this.normSel();
    var a = R.numToCol(this.sel.c) + (this.sel.r + 1);
    if (n.r1 === n.r2 && n.c1 === n.c2) return a;
    return a + ':' + R.numToCol(n.c2) + (n.r2 + 1);
  };

  /* ---------------------------------------------------------------- mouse -- */

  Grid.prototype.cellFromEvent = function (e) {
    var td = e.target.closest ? e.target.closest('td.xt-cell') : null;
    if (!td) return null;
    return { r: +td.dataset.r, c: +td.dataset.c };
  };

  Grid.prototype.onMouseDown = function (e) {
    var self = this;

    // Starting a fill drag.
    if (e.target === this.handle) {
      e.preventDefault();
      this.filling = { from: this.normSel(), to: this.normSel() };
      return;
    }

    // Column / row header click selects the whole column / row.
    var th = e.target.closest && e.target.closest('th');
    if (th && th.dataset.c !== undefined) {
      this.commitEdit();
      this.sel = { r: 0, c: +th.dataset.c, r2: this.nRows - 1, c2: +th.dataset.c };
      this.paintSelection(); this.emit('select');
      this.scroller.focus();
      return;
    }
    if (th && th.dataset.r !== undefined) {
      this.commitEdit();
      this.sel = { r: +th.dataset.r, c: 0, r2: +th.dataset.r, c2: this.nCols - 1 };
      this.paintSelection(); this.emit('select');
      this.scroller.focus();
      return;
    }

    var pos = this.cellFromEvent(e);
    if (!pos) return;

    // Click-to-reference: if we are mid-formula, clicking a cell inserts its
    // address instead of moving the selection. Same as Excel.
    if (this.editing && this.editorWantsReference()) {
      e.preventDefault();
      this.insertReference(pos.r, pos.c, e.shiftKey);
      return;
    }

    this.commitEdit();
    this.dragging = true;
    this.select(pos.r, pos.c, e.shiftKey);
    this.scroller.focus();
  };

  Grid.prototype.onMouseMove = function (e) {
    if (this.filling) {
      var pos = this.cellFromEvent(e);
      if (!pos) return;
      var from = this.filling.from;
      // Fill runs in one direction only: down or right, whichever moved more.
      var dr = pos.r - from.r2, dc = pos.c - from.c2;
      if (Math.abs(dr) >= Math.abs(dc)) {
        this.filling.to = { r1: from.r1, r2: Math.max(from.r2, pos.r), c1: from.c1, c2: from.c2 };
      } else {
        this.filling.to = { r1: from.r1, r2: from.r2, c1: from.c1, c2: Math.max(from.c2, pos.c) };
      }
      this.previewFill();
      return;
    }
    if (this.dragging) {
      var p2 = this.cellFromEvent(e);
      if (p2) this.select(p2.r, p2.c, true);
    }
  };

  Grid.prototype.previewFill = function () {
    var t = this.filling.to;
    for (var r = 0; r < this.nRows; r++) {
      for (var c = 0; c < this.nCols; c++) {
        var on = r >= t.r1 && r <= t.r2 && c >= t.c1 && c <= t.c2;
        this.cellEls[r][c].classList.toggle('xt-fillpreview', on);
      }
    }
  };

  Grid.prototype.onMouseUp = function () {
    if (this.filling) {
      var t = this.filling.to, f = this.filling.from;
      this.filling = null;
      for (var r = 0; r < this.nRows; r++) {
        for (var c = 0; c < this.nCols; c++) this.cellEls[r][c].classList.remove('xt-fillpreview');
      }
      if (t.r2 > f.r2) this.fillFrom(f, 'down', t.r2);
      else if (t.c2 > f.c2) this.fillFrom(f, 'right', t.c2);
    }
    this.dragging = false;
  };

  Grid.prototype.onDblClick = function (e) {
    // Double-clicking the fill handle fills to the end of the adjacent data,
    // which is how anyone who uses Excel daily actually fills a column.
    if (e.target === this.handle) { this.fillToEdge(); return; }
    var pos = this.cellFromEvent(e);
    if (pos) { this.select(pos.r, pos.c); this.beginEdit(); }
  };

  /* ----------------------------------------------------------------- keys -- */

  Grid.prototype.onKeyDown = function (e) {
    if (this.editing) return;
    var k = e.key;
    var mod = e.ctrlKey || e.metaKey;

    if (mod && (k === 'z' || k === 'Z')) { e.preventDefault(); this.undo(); return; }
    if (mod && (k === 'y' || k === 'Y')) { e.preventDefault(); this.redo(); return; }
    if (mod && (k === 'd' || k === 'D')) { e.preventDefault(); this.fillSelection('down'); return; }
    if (mod && (k === 'r' || k === 'R')) { e.preventDefault(); this.fillSelection('right'); return; }
    if (mod && k === '`') { e.preventDefault(); this.toggleFormulas(); return; }
    if (mod && (k === 'c' || k === 'x' || k === 'v')) return;   // let the native events fire

    switch (k) {
      case 'ArrowUp': e.preventDefault(); this.select(this.sel.r - 1, this.sel.c, e.shiftKey); return;
      case 'ArrowDown': e.preventDefault(); this.select(this.sel.r + 1, this.sel.c, e.shiftKey); return;
      case 'ArrowLeft': e.preventDefault(); this.select(this.sel.r, this.sel.c - 1, e.shiftKey); return;
      case 'ArrowRight': e.preventDefault(); this.select(this.sel.r, this.sel.c + 1, e.shiftKey); return;
      case 'Home': e.preventDefault(); this.select(this.sel.r, 0, e.shiftKey); return;
      case 'PageDown': e.preventDefault(); this.select(this.sel.r + 20, this.sel.c, e.shiftKey); return;
      case 'PageUp': e.preventDefault(); this.select(this.sel.r - 20, this.sel.c, e.shiftKey); return;
      case 'Enter': e.preventDefault(); this.select(this.sel.r + (e.shiftKey ? -1 : 1), this.sel.c); return;
      case 'Tab': e.preventDefault(); this.select(this.sel.r, this.sel.c + (e.shiftKey ? -1 : 1)); return;
      case 'F2': e.preventDefault(); this.beginEdit(); return;
      case 'Delete': case 'Backspace': e.preventDefault(); this.clearSelection(); return;
      case 'Escape': return;
    }

    // Any printable character starts an edit, replacing what was there.
    if (k.length === 1 && !mod && !e.altKey) {
      e.preventDefault();
      this.beginEdit('');
      this.editor.value = k;
      this.onEditorInput();
    }
  };

  /* -------------------------------------------------------------- editing -- */

  Grid.prototype.beginEdit = function (seed) {
    var sh = this.sheet();
    if (!sh) return;
    var r = this.sel.r, c = this.sel.c;
    if (this.readOnlyCells && this.readOnlyCells[r + ',' + c]) {
      this.emit('lockedEdit');
      return;
    }
    var cell = sh.get(r, c);
    var td = this.cellEls[r][c];
    this.editing = { r: r, c: c };
    this.editor.value = seed !== undefined ? seed : (cell ? cell.input : '');
    this.editor.style.display = 'block';
    this.editor.style.left = td.offsetLeft + 'px';
    this.editor.style.top = td.offsetTop + 'px';
    this.editor.style.minWidth = td.offsetWidth + 'px';
    this.editor.style.height = td.offsetHeight + 'px';
    this.editor.focus();
    this.editor.setSelectionRange(this.editor.value.length, this.editor.value.length);
    this.handle.style.display = 'none';
    this.emit('editStart');
    this.onEditorInput();
  };

  Grid.prototype.commitEdit = function (move, cse) {
    if (!this.editing) return;
    var e = this.editing;
    var text = this.editor.value;
    this.editing = null;
    this.editor.style.display = 'none';
    this.clearHighlights();
    this.pushUndo([{ r: e.r, c: e.c }]);
    this.wb.setInput(this.sheetName, e.r, e.c, text, { cse: !!cse });
    this.sheet().recomputeExtent();
    this.refresh();
    this.emit('editEnd');
    if (move === 'down') this.select(e.r + 1, e.c);
    else if (move === 'up') this.select(e.r - 1, e.c);
    else if (move === 'right') this.select(e.r, e.c + 1);
    else if (move === 'left') this.select(e.r, e.c - 1);
    else this.paintSelection();
    this.scroller.focus();
  };

  Grid.prototype.cancelEdit = function () {
    if (!this.editing) return;
    this.editing = null;
    this.editor.style.display = 'none';
    this.clearHighlights();
    this.paintSelection();
    this.scroller.focus();
    this.emit('editEnd');
  };

  Grid.prototype.onEditorKey = function (e) {
    var k = e.key;
    if (this.opts.onAutocompleteKey && this.opts.onAutocompleteKey(e, this)) return;

    if (k === 'Enter') {
      e.preventDefault();
      // Ctrl+Shift+Enter enters a legacy array formula, as in Excel.
      this.commitEdit(e.shiftKey ? 'up' : 'down', (e.ctrlKey || e.metaKey) && e.shiftKey);
      return;
    }
    if (k === 'Tab') { e.preventDefault(); this.commitEdit(e.shiftKey ? 'left' : 'right'); return; }
    if (k === 'Escape') { e.preventDefault(); this.cancelEdit(); return; }
    if (k === 'F4') { e.preventDefault(); this.cycleAnchor(); return; }
  };

  Grid.prototype.onEditorInput = function () {
    this.showHighlights(this.editor.value);
    this.emit('editInput');
  };

  /* F4 cycles the $ signs on the reference the cursor is sitting in:
       A1  ->  $A$1  ->  A$1  ->  $A1  ->  A1
     This is the single most useful keystroke in Excel and almost nobody is
     taught it, so it is wired up from level 3 onwards. */
  Grid.prototype.cycleAnchor = function () {
    var text = this.editor.value;
    var pos = this.editor.selectionStart;
    var re = /\$?[A-Za-z]{1,3}\$?\d{1,7}/g, m;
    while ((m = re.exec(text))) {
      var start = m.index, end = start + m[0].length;
      // The cursor may be just after the reference, which still counts.
      if (pos >= start && pos <= end) {
        var a = R.parseAddr(m[0]);
        if (!a) break;
        var state = (a.ca ? 2 : 0) + (a.ra ? 1 : 0);   // 0 none, 3 both, 1 row, 2 col
        var next = { 0: { ca: true, ra: true }, 3: { ca: false, ra: true },
          1: { ca: true, ra: false }, 2: { ca: false, ra: false } }[state];
        a.ca = next.ca; a.ra = next.ra;
        var repl = R.formatAddr(a);
        this.editor.value = text.slice(0, start) + repl + text.slice(end);
        this.editor.setSelectionRange(start + repl.length, start + repl.length);
        this.onEditorInput();
        return;
      }
    }
  };

  /* Is the editor in a state where clicking a cell should insert a reference?
     True when the text is a formula and ends with an operator, a comma or an
     open bracket -- i.e. Excel is waiting for an operand. */
  Grid.prototype.editorWantsReference = function () {
    var t = this.editor.value;
    if (t.charAt(0) !== '=') return false;
    return /[=+\-*/^&<>(,:]\s*$/.test(t.slice(0, this.editor.selectionStart));
  };

  Grid.prototype.insertReference = function (r, c, extend) {
    var ref = R.numToCol(c) + (r + 1);
    if (this.sheetName !== this.editing.sheetName && false) ref = ref;   // same sheet
    var pos = this.editor.selectionStart;
    var t = this.editor.value;
    this.editor.value = t.slice(0, pos) + ref + t.slice(pos);
    this.editor.setSelectionRange(pos + ref.length, pos + ref.length);
    this.editor.focus();
    this.onEditorInput();
  };

  /* Outline every range the formula being typed refers to, each in its own
     colour, matching the colour the reference text gets in the formula bar. */
  Grid.prototype.showHighlights = function (text) {
    this.clearHighlights();
    if (!text || text.charAt(0) !== '=') return;
    var refs = S.references(text);
    var self = this;
    this.highlight = [];
    refs.forEach(function (ref, i) {
      var n = ref.node;
      var sheet = n.sheet || self.sheetName;
      if (sheet.toLowerCase() !== self.sheetName.toLowerCase()) return;
      var r1, r2, c1, c2;
      if (n.t === 'ref') { r1 = r2 = n.addr.r; c1 = c2 = n.addr.c; }
      else if (n.kind === 'cells') {
        r1 = Math.min(n.a.r, n.b.r); r2 = Math.max(n.a.r, n.b.r);
        c1 = Math.min(n.a.c, n.b.c); c2 = Math.max(n.a.c, n.b.c);
      } else if (n.kind === 'col') {
        r1 = 0; r2 = Math.min(self.nRows - 1, self.sheet().used().maxRow);
        c1 = Math.min(n.c1, n.c2); c2 = Math.max(n.c1, n.c2);
      } else {
        c1 = 0; c2 = Math.min(self.nCols - 1, self.sheet().used().maxCol);
        r1 = Math.min(n.r1, n.r2); r2 = Math.max(n.r1, n.r2);
      }
      var cls = 'xt-ref' + (i % 6);
      for (var r = Math.max(0, r1); r <= Math.min(r2, self.nRows - 1); r++) {
        for (var c = Math.max(0, c1); c <= Math.min(c2, self.nCols - 1); c++) {
          self.cellEls[r][c].classList.add(cls);
          self.highlight.push([r, c, cls]);
        }
      }
    });
  };

  Grid.prototype.clearHighlights = function () {
    var self = this;
    (this.highlight || []).forEach(function (h) {
      var td = self.cellEls[h[0]] && self.cellEls[h[0]][h[1]];
      if (td) td.classList.remove(h[2]);
    });
    this.highlight = [];
  };

  /* ----------------------------------------------------------------- fill -- */

  /* THE important one. Copy the formula (or value) from the source block and
     translate every reference by how far each target cell has moved.
     P.translate does the work; this just decides which cells get what. */
  Grid.prototype.fillFrom = function (from, dir, to) {
    var sh = this.sheet();
    var touched = [], i, r, c;

    if (dir === 'down') {
      for (r = from.r2 + 1; r <= to; r++) {
        for (c = from.c1; c <= from.c2; c++) {
          var srcR = from.r1 + ((r - from.r1) % (from.r2 - from.r1 + 1));
          touched.push({ r: r, c: c, srcR: srcR, srcC: c });
        }
      }
    } else {
      for (c = from.c2 + 1; c <= to; c++) {
        for (r = from.r1; r <= from.r2; r++) {
          var srcC = from.c1 + ((c - from.c1) % (from.c2 - from.c1 + 1));
          touched.push({ r: r, c: c, srcR: r, srcC: srcC });
        }
      }
    }
    if (!touched.length) return;

    this.pushUndo(touched);
    for (i = 0; i < touched.length; i++) {
      var t = touched[i];
      if (this.readOnlyCells && this.readOnlyCells[t.r + ',' + t.c]) continue;
      var src = sh.get(t.srcR, t.srcC);
      var input = src ? src.input : '';
      var cse = src ? src.cse : false;
      var fmt = src ? src.fmt : null;
      var text = input;
      if (input.charAt(0) === '=') {
        text = P.translateText(input, t.r - t.srcR, t.c - t.srcC);
      }
      this.wb.setInput(this.sheetName, t.r, t.c, text, { cse: cse, fmt: fmt });
    }
    sh.recomputeExtent();
    // Extend the selection over what was just filled, as Excel does.
    if (dir === 'down') this.sel = { r: from.r1, c: from.c1, r2: to, c2: from.c2 };
    else this.sel = { r: from.r1, c: from.c1, r2: from.r2, c2: to };
    this.layout();
    this.refresh();
    this.emit('filled');
  };

  // Ctrl+D / Ctrl+R: fill the top row (or left column) of the selection
  // through the rest of it.
  Grid.prototype.fillSelection = function (dir) {
    var n = this.normSel();
    if (dir === 'down') {
      if (n.r2 <= n.r1) return;
      this.fillFrom({ r1: n.r1, r2: n.r1, c1: n.c1, c2: n.c2 }, 'down', n.r2);
    } else {
      if (n.c2 <= n.c1) return;
      this.fillFrom({ r1: n.r1, r2: n.r2, c1: n.c1, c2: n.c1 }, 'right', n.c2);
    }
  };

  /* Double-click the handle: fill down as far as the neighbouring column has
     data. Looks left first, then right. */
  Grid.prototype.fillToEdge = function () {
    var sh = this.sheet();
    var n = this.normSel();
    var probe = n.c1 > 0 ? n.c1 - 1 : n.c2 + 1;
    var last = n.r2;
    for (var r = n.r2 + 1; r < this.nRows; r++) {
      var cell = sh.get(r, probe);
      var v = cell ? this.wb.cellValue(sh, r, probe) : V.EMPTY;
      if (V.isEmpty(v) || v === '') break;
      last = r;
    }
    if (last > n.r2) this.fillFrom(n, 'down', last);
  };

  Grid.prototype.toggleFormulas = function () {
    this.showFormulas = !this.showFormulas;
    this.host.classList.toggle('xt-formulas-visible', this.showFormulas);
    this.refresh(true);
    this.emit('select');
  };

  /* ---------------------------------------------------------------- clear -- */

  Grid.prototype.clearSelection = function () {
    var n = this.normSel(), touched = [];
    for (var r = n.r1; r <= n.r2; r++) {
      for (var c = n.c1; c <= n.c2; c++) {
        if (this.readOnlyCells && this.readOnlyCells[r + ',' + c]) continue;
        touched.push({ r: r, c: c });
      }
    }
    this.pushUndo(touched);
    for (var i = 0; i < touched.length; i++) {
      this.wb.setInput(this.sheetName, touched[i].r, touched[i].c, '');
    }
    this.sheet().recomputeExtent();
    this.refresh();
  };

  /* ---------------------------------------------------------------- paste -- */

  Grid.prototype.onCopy = function (e, cut) {
    if (this.editing) return;
    var sh = this.sheet(), n = this.normSel();
    var lines = [], inputs = [];
    for (var r = n.r1; r <= n.r2; r++) {
      var vals = [], ins = [];
      for (var c = n.c1; c <= n.c2; c++) {
        var cell = sh.get(r, c);
        vals.push(cell ? F.format(this.wb.cellValue(sh, r, c), cell.fmt) : '');
        ins.push(cell ? cell.input : '');
      }
      lines.push(vals.join('\t'));
      inputs.push(ins);
    }
    this.clip = { rows: inputs, from: n, cut: !!cut, sheet: this.sheetName };
    if (e.clipboardData) {
      e.clipboardData.setData('text/plain', lines.join('\n'));
      e.preventDefault();
    }
    if (cut) this.clearSelection();
  };

  /* PASTE -- the trainer's defining gesture (§12.6).

     Two things are deliberate here:

     1. Dirty data stays dirty. A pasted "007" stays text, a trailing space is
        kept, and a non-breaking space survives. The real exports arrive like
        that and the whole point is to meet it.
     2. Pasting OVER formulas destroys them, visibly. That is how a learner
        finds out what re-pasting into the wrong sheet costs. "Paste values"
        and "paste" are two separate commands, as in Excel.                   */
  Grid.prototype.onPaste = function (e) {
    if (this.editing) return;
    var text = e.clipboardData && e.clipboardData.getData('text/plain');
    if (!text) {
      // An internal copy with no system clipboard: use our own buffer.
      if (this.clip) { this.pasteInputs(this.clip.rows); e.preventDefault(); }
      return;
    }
    e.preventDefault();
    this.pasteText(text);
  };

  Grid.prototype.pasteText = function (text, asValues) {
    var rows = String(text).replace(/\r\n/g, '\n').replace(/\r/g, '\n');
    // A single trailing newline is an artefact of copying whole rows.
    rows = rows.replace(/\n$/, '').split('\n').map(function (line) { return line.split('\t'); });
    this.pasteInputs(rows, asValues);
  };

  Grid.prototype.pasteInputs = function (rows, asValues) {
    var n = this.normSel(), touched = [], r, c;
    for (r = 0; r < rows.length; r++) {
      for (c = 0; c < rows[r].length; c++) {
        touched.push({ r: n.r1 + r, c: n.c1 + c });
      }
    }
    this.pushUndo(touched);
    var destroyed = 0;
    for (r = 0; r < rows.length; r++) {
      for (c = 0; c < rows[r].length; c++) {
        var tr = n.r1 + r, tc = n.c1 + c;
        if (this.readOnlyCells && this.readOnlyCells[tr + ',' + tc]) continue;
        var existing = this.sheet().get(tr, tc);
        if (existing && existing.ast) destroyed++;
        var v = rows[r][c];
        if (asValues && v.charAt(0) === '=') v = "'" + v;
        this.wb.setInput(this.sheetName, tr, tc, v);
      }
    }
    this.sheet().recomputeExtent();
    this.sel = { r: n.r1, c: n.c1, r2: n.r1 + rows.length - 1,
      c2: n.c1 + (rows[0] ? rows[0].length - 1 : 0) };
    this.layout();
    this.refresh();
    this.lastPasteDestroyed = destroyed;
    this.emit('pasted');
  };

  /* ----------------------------------------------------------- undo / redo -- */

  Grid.prototype.snapshot = function (cells) {
    var sh = this.sheet(), out = [];
    cells.forEach(function (p) {
      var cell = sh.get(p.r, p.c);
      out.push({ r: p.r, c: p.c, input: cell ? cell.input : '',
        cse: cell ? cell.cse : false, fmt: cell ? cell.fmt : null });
    });
    return { sheet: this.sheetName, cells: out };
  };

  Grid.prototype.pushUndo = function (cells) {
    this.undoStack.push(this.snapshot(cells));
    if (this.undoStack.length > 120) this.undoStack.shift();
    this.redoStack = [];
  };

  Grid.prototype.applySnapshot = function (snap) {
    var self = this;
    var before = this.snapshot(snap.cells);
    before.sheet = snap.sheet;
    snap.cells.forEach(function (c) {
      self.wb.setInput(snap.sheet, c.r, c.c, c.input, { cse: c.cse, fmt: c.fmt });
    });
    this.wb.sheet(snap.sheet).recomputeExtent();
    this.refresh();
    return before;
  };

  Grid.prototype.undo = function () {
    var snap = this.undoStack.pop();
    if (!snap) return;
    if (snap.sheet !== this.sheetName) this.setSheet(snap.sheet);
    this.redoStack.push(this.applySnapshot(snap));
    this.emit('select');
  };

  Grid.prototype.redo = function () {
    var snap = this.redoStack.pop();
    if (!snap) return;
    if (snap.sheet !== this.sheetName) this.setSheet(snap.sheet);
    this.undoStack.push(this.applySnapshot(snap));
    this.emit('select');
  };

  /* ---------------------------------------------------------------- marks -- */

  // Used by the tracer to draw precedent / dependent rings.
  Grid.prototype.setMarks = function (map) {
    this.marks = map || {};
    this.refresh(true);
  };

  XT.Grid = Grid;

})(window.XT = window.XT || {});
