/* =============================================================================
   Excel Trainer — ui/app.js
   The shell: toolbar, formula bar, sheet tabs, status bar, the instrument panels
   on the right, and the lesson runner.
   ============================================================================= */
(function (XT) {
  'use strict';

  var V = XT.V, R = XT.R, P = XT.P, S = XT.S, D = XT.D;
  var Panels = XT.Panels, Grader = XT.Grader, Pr = XT.Progress;

  var App = {
    wb: null, grid: null, scenario: 'T1', exercise: null,
    panel: 'lesson', traceMode: null, acIndex: 0, acList: [],
    lessonOpen: null
  };
  XT.App = App;

  function $(id) { return document.getElementById(id); }
  function esc(s) { return Panels.esc(s); }

  /* ------------------------------------------------------------------ start */

  App.start = function () {
    App.grid = new XT.Grid($('grid'), {
      onSelect: onSelect, onRefresh: onRefresh, onEditStart: onEditStart,
      onEditInput: onEditInput, onEditEnd: onEditEnd, onFilled: afterChange,
      onPasted: onPasted, onLockedEdit: function () {
        setStatus('That cell is part of the exercise setup and is locked.', 'warn');
      },
      onAutocompleteKey: onAutocompleteKey
    });

    wireToolbar();
    wireFormulaBar();
    wirePanels();

    var last = Pr.state.lastExercise && XT.data.exercise(Pr.state.lastExercise);
    if (last) loadExercise(last.id);
    else loadScenario(XT.data.scenarioOrder[0]);
  };

  /* -------------------------------------------------------------- workbooks */

  function loadScenario(id, keepPanel) {
    App.scenario = id;
    App.exercise = null;
    App.wb = XT.data.makeWorkbook(id, {});
    App.grid.readOnlyCells = null;
    App.grid.setWorkbook(App.wb, App.wb.sheets[0].name);
    renderTabs();
    if (!keepPanel) showPanel('lesson');
    renderPanel();
    setStatus('Sandbox: ' + XT.data.scenarios[id].title + '. Nothing is graded here.');
  }
  App.loadScenario = loadScenario;

  function loadExercise(exId) {
    var ex = XT.data.exercise(exId);
    if (!ex) return;
    App.exercise = ex;
    App.scenario = ex.scenario;
    Pr.setLast(exId);
    Pr.markAttempted(exId);

    App.wb = Grader.exerciseWorkbook(ex);

    // Lock the cells the setup provided, so an exercise cannot be "solved" by
    // editing its own scaffolding.
    var locked = {};
    Object.keys(ex.setup || {}).forEach(function (key) {
      var parts = Grader.splitTarget(key, ex.sheet);
      if (parts.sheet !== ex.sheet) return;
      Grader.cellList(parts.range).forEach(function (cl) { locked[cl.r + ',' + cl.c] = true; });
    });
    // Target cells are never locked, even if setup also wrote them.
    (ex.targets || []).forEach(function (t) {
      var parts = Grader.splitTarget(t.range || t, ex.sheet);
      if (parts.sheet !== ex.sheet) return;
      Grader.cellList(parts.range).forEach(function (cl) { delete locked[cl.r + ',' + cl.c]; });
    });

    App.grid.readOnlyCells = locked;
    App.grid.setWorkbook(App.wb, ex.sheet);
    // Clear the target cells so the learner starts from blank.
    (ex.targets || []).forEach(function (t) {
      var parts = Grader.splitTarget(t.range || t, ex.sheet);
      Grader.cellList(parts.range).forEach(function (cl) {
        App.wb.setInput(parts.sheet, cl.r, cl.c, '');
      });
    });
    App.wb.recalc();
    var first = (ex.targets && ex.targets[0]) ? Grader.cellList(
      Grader.splitTarget(ex.targets[0].range || ex.targets[0], ex.sheet).range)[0] : null;
    App.grid.layout();
    App.grid.refresh();
    if (first) App.grid.select(first.r, first.c);
    renderTabs();
    showPanel('lesson');
    App.lastResult = null;
    setStatus('Level ' + ex.level + ' — ' + ex.title);
  }
  App.loadExercise = loadExercise;

  /* ---------------------------------------------------------------- toolbar */

  function wireToolbar() {
    $('btnCheck').onclick = check;
    $('btnShowMe').onclick = showMe;
    $('btnReset').onclick = function () {
      if (App.exercise) loadExercise(App.exercise.id); else loadScenario(App.scenario);
    };
    $('btnFormulas').onclick = function () {
      App.grid.toggleFormulas();
      this.classList.toggle('on', App.grid.showFormulas);
    };
    $('btnFillDown').onclick = function () { App.grid.fillSelection('down'); App.grid.scroller.focus(); };
    $('btnFillRight').onclick = function () { App.grid.fillSelection('right'); App.grid.scroller.focus(); };
    $('btnUndo').onclick = function () { App.grid.undo(); };
    $('btnPaste').onclick = openPaste;
    $('btnTheme').onclick = function () {
      var dark = document.documentElement.getAttribute('data-theme') === 'dark';
      document.documentElement.setAttribute('data-theme', dark ? 'light' : 'dark');
      try { localStorage.setItem('xt.theme', dark ? 'light' : 'dark'); } catch (e) {}
    };
    $('btnTrace').onclick = function () {
      App.traceMode = App.traceMode === 'both' ? null : 'both';
      this.classList.toggle('on', !!App.traceMode);
      applyTrace();
      if (App.traceMode) { App.sub.explain = 'trace'; showPanel('explain'); }
    };
    $('btnHint').onclick = function () {
      if (!App.exercise) { setStatus('Pick a task first — the Lesson panel lists them.'); return; }
      App.hintsShown = (App.hintsShown || 0) + 1;
      showPanel('lesson');
    };
    $('btnMore').onclick = function () {
      $('moreBar').classList.toggle('hidden');
      this.classList.toggle('on', !$('moreBar').classList.contains('hidden'));
    };

    var sel = $('scenarioPick');
    XT.data.scenarioOrder.forEach(function (id) {
      var o = document.createElement('option');
      o.value = id;
      o.textContent = XT.data.scenarios[id].title;
      sel.appendChild(o);
    });
    sel.value = App.scenario;
    sel.onchange = function () { loadScenario(this.value); };

    document.addEventListener('keydown', function (e) {
      if ((e.ctrlKey || e.metaKey) && e.key === 'Enter') { e.preventDefault(); check(); }
    });
  }

  function openPaste() {
    var m = $('pasteModal');
    m.classList.remove('hidden');
    var which = App.scenario === 'S1' ? 'samplesin' : 'samplesin';
    $('pasteText').value = XT.data.pasteBlock(which, { n: 10, seed: Date.now() % 9999 });
    $('pasteText').focus();
    $('pasteText').select();
  }
  $('pasteClose') && ($('pasteClose').onclick = function () { $('pasteModal').classList.add('hidden'); });

  /* ------------------------------------------------------------ formula bar */

  function wireFormulaBar() {
    var fi = $('finput');
    fi.addEventListener('focus', function () {
      var s = App.grid.sel;
      if (!App.grid.editing) App.grid.beginEdit();
      // Mirror the grid editor; the grid editor stays the source of truth.
      fi.value = App.grid.editor.value;
    });
    fi.addEventListener('input', function () {
      if (App.grid.editing) {
        App.grid.editor.value = fi.value;
        App.grid.onEditorInput();
      }
    });
    fi.addEventListener('keydown', function (e) {
      if (e.key === 'Enter') { e.preventDefault(); App.grid.commitEdit('down'); fi.blur(); }
      if (e.key === 'Escape') { e.preventDefault(); App.grid.cancelEdit(); fi.blur(); }
    });
    $('namebox').addEventListener('keydown', function (e) {
      if (e.key !== 'Enter') return;
      var v = this.value.trim();
      var m = /^(?:(.+)!)?([A-Za-z]{1,3}\d{1,7})$/.exec(v);
      if (!m) return;
      if (m[1] && App.wb.sheet(m[1].replace(/'/g, ''))) App.grid.setSheet(m[1].replace(/'/g, ''));
      var a = R.rc(m[2]);
      if (a) App.grid.select(a.r, a.c);
      App.grid.scroller.focus();
    });
  }

  function onSelect() {
    var g = App.grid;
    $('namebox').value = g.addr();
    var sh = g.sheet();
    var cell = sh && sh.get(g.sel.r, g.sel.c);
    $('finput').value = cell ? cell.input : '';
    if (App.traceMode) applyTrace();
    // Only the Explain tab follows the selection. The rest would flicker.
    if (App.panel === 'explain') renderPanel();
    updateStatusCell();
  }

  function onEditStart() { $('finput').value = App.grid.editor.value; }
  function onEditEnd() { hideAutocomplete(); $('fhint').textContent = ''; }

  function onEditInput() {
    var t = App.grid.editor.value;
    $('finput').value = t;
    showArgHint(t);
    showAutocomplete(t);
  }

  function onRefresh() {
    var circ = App.wb.circular;
    if (circ.length) {
      setStatus('Circular reference: ' + circ.slice(0, 3).join(', ') +
        (circ.length > 3 ? ' and ' + (circ.length - 3) + ' more' : '') +
        '. Those cells show 0 — Excel does the same.', 'bad');
    }
  }

  function onPasted() {
    var n = App.grid.lastPasteDestroyed;
    if (n) {
      setStatus('That paste overwrote ' + n + ' formula' + (n === 1 ? '' : 's') +
        '. In a real workbook that is how a re-paste silently breaks a sheet. Ctrl+Z undoes it.',
        'warn');
    }
    afterChange();
  }

  function afterChange() { if (App.traceMode) applyTrace(); }

  /* ------------------------------------------------- argument hint + autocomplete */

  function showArgHint(text) {
    var h = $('fhint');
    if (!text || text.charAt(0) !== '=') { h.textContent = ''; return; }
    var upto = text.slice(0, App.grid.editor.selectionStart || text.length);
    // Find the innermost unclosed function call.
    var stack = [], re = /([A-Za-z_][A-Za-z0-9_.]*)\s*\(|\(|\)|,/g, m;
    while ((m = re.exec(upto))) {
      if (m[0] === ')') stack.pop();
      else if (m[0] === ',') { if (stack.length) stack[stack.length - 1].arg++; }
      else stack.push({ name: m[1] || null, arg: 0 });
    }
    var top = null;
    for (var i = stack.length - 1; i >= 0; i--) if (stack[i].name) { top = stack[i]; break; }
    if (!top) { h.textContent = ''; return; }
    var f = XT.FN[top.name.toUpperCase().replace(/^_XLFN\./, '')];
    if (!f) { h.textContent = ''; return; }
    var parts = /\((.*)\)/.exec(f.sig);
    var args = parts ? parts[1].split(/,\s*/) : [];
    var html = f.name + '(' + args.map(function (a, k) {
      return k === top.arg ? '<b>' + esc(a) + '</b>' : esc(a);
    }).join(', ') + ')  — ' + esc(f.doc);
    h.innerHTML = html;
  }

  function showAutocomplete(text) {
    var pos = App.grid.editor.selectionStart || text.length;
    var m = /([A-Za-z]{2,})$/.exec(text.slice(0, pos));
    if (!m || text.charAt(0) !== '=') { hideAutocomplete(); return; }
    var names = XT.functionNames(m[1]);
    if (!names.length || (names.length === 1 && names[0] === m[1].toUpperCase())) {
      hideAutocomplete(); return;
    }
    App.acList = names.slice(0, 12);
    App.acIndex = 0;
    App.acPrefix = m[1];
    var box = $('ac');
    box.innerHTML = App.acList.map(function (n, i) {
      var f = XT.FN[n];
      return '<div class="' + (i === 0 ? 'on' : '') + '" data-n="' + n + '">' + n +
        '<span>' + esc(f ? f.doc : '') + '</span></div>';
    }).join('');
    var ed = App.grid.editor;
    box.style.display = 'block';
    var host = App.grid.host.getBoundingClientRect();
    box.style.left = (host.left + ed.offsetLeft - App.grid.scroller.scrollLeft) + 'px';
    box.style.top = (host.top + ed.offsetTop + ed.offsetHeight -
      App.grid.scroller.scrollTop + 2) + 'px';
    Array.prototype.forEach.call(box.children, function (d) {
      d.onmousedown = function (e) { e.preventDefault(); acceptAutocomplete(d.dataset.n); };
    });
  }
  function hideAutocomplete() { $('ac').style.display = 'none'; App.acList = []; }

  function acceptAutocomplete(name) {
    var ed = App.grid.editor;
    var pos = ed.selectionStart;
    var before = ed.value.slice(0, pos).replace(/[A-Za-z]{2,}$/, '');
    ed.value = before + name + '(' + ed.value.slice(pos);
    var np = before.length + name.length + 1;
    ed.setSelectionRange(np, np);
    hideAutocomplete();
    App.grid.onEditorInput();
    ed.focus();
  }

  function onAutocompleteKey(e) {
    if (!App.acList.length) return false;
    var box = $('ac');
    if (e.key === 'ArrowDown' || e.key === 'ArrowUp') {
      e.preventDefault();
      App.acIndex = (App.acIndex + (e.key === 'ArrowDown' ? 1 : -1) + App.acList.length) %
        App.acList.length;
      Array.prototype.forEach.call(box.children, function (d, i) {
        d.className = i === App.acIndex ? 'on' : '';
      });
      return true;
    }
    if (e.key === 'Tab' || (e.key === 'Enter' && App.acList.length)) {
      e.preventDefault();
      acceptAutocomplete(App.acList[App.acIndex]);
      return true;
    }
    if (e.key === 'Escape') { hideAutocomplete(); return true; }
    return false;
  }

  /* ------------------------------------------------------------------- tabs */

  function renderTabs() {
    var el = $('tabs');
    el.innerHTML = '';
    App.wb.sheets.forEach(function (sh) {
      var b = document.createElement('button');
      b.className = 'xt-tab' + (sh.name === App.grid.sheetName ? ' on' : '') +
        (sh.hidden ? ' hiddensheet' : '');
      b.title = sh.note || '';
      b.innerHTML = (sh.layer ? '<span class="lay">' + sh.layer + '</span>' : '') +
        esc(sh.name);
      b.onclick = function () { App.grid.setSheet(sh.name); renderTabs(); };
      el.appendChild(b);
    });
  }

  /* ----------------------------------------------------------------- status */

  function setStatus(msg, kind) {
    var el = $('statusMsg');
    el.textContent = msg || '';
    el.className = kind || '';
  }
  function updateStatusCell() {
    var g = App.grid;
    var d = App.wb.describe(g.sheetName, g.sel.r, g.sel.c);
    $('statusCell').textContent = d.type + (d.isFormula ? ' · formula' : '') +
      (d.cse ? ' · array' : '');
  }

  /* ----------------------------------------------------------------- trace */

  function applyTrace() {
    if (!App.traceMode) { App.grid.setMarks({}); return; }
    var g = App.grid;
    App.grid.setMarks(Panels.traceMarks(App.wb, g.sheetName, g.sel.r, g.sel.c, App.traceMode));
  }

  /* ---------------------------------------------------------------- panels */

  /* Five tabs, not ten. The instruments that explain a cell live together
     behind one tab, because you reach for them at the same moment; the
     reference material lives behind another, because you reach for that at a
     different moment. */
  var PANELS = [
    ['lesson', 'Task'],
    ['learn', 'Learn by formula'],
    ['explain', 'Explain this cell'],
    ['progress', 'Progress'],
    ['reference', 'Reference']
  ];

  var SUBS = {
    explain: [['step', 'Evaluate'], ['stack', 'Layers'], ['inspect', 'Inspector'],
      ['trace', 'Trace'], ['shape', 'Shape']],
    reference: [['functions', 'Functions'], ['patterns', 'Patterns'],
      ['scenarios', 'Toolkits']]
  };

  App.sub = { explain: 'step', reference: 'functions' };

  function wirePanels() {
    var el = $('ptabs');
    PANELS.forEach(function (p) {
      var b = document.createElement('button');
      b.className = 'xt-ptab';
      b.dataset.p = p[0];
      b.textContent = p[1];
      b.onclick = function () { showPanel(p[0]); };
      el.appendChild(b);
    });
    $('pbody').addEventListener('click', onPanelClick);
    showPanel('lesson');
  }

  function showPanel(name) {
    App.panel = name;
    Array.prototype.forEach.call($('ptabs').children, function (b) {
      b.classList.toggle('on', b.dataset.p === name);
    });
    App.stepAt = null;
    renderPanel();
  }
  App.showPanel = showPanel;

  function subBar(group) {
    var cur = App.sub[group];
    return '<div class="xt-sub">' + SUBS[group].map(function (s) {
      return '<button class="xt-stab' + (s[0] === cur ? ' on' : '') +
        '" data-sub="' + group + ':' + s[0] + '">' + s[1] + '</button>';
    }).join('') + '</div>';
  }

  function renderPanel() {
    var g = App.grid, body = $('pbody');
    var r = g.sel.r, c = g.sel.c;
    switch (App.panel) {
      case 'lesson': body.innerHTML = renderLesson(); break;
      case 'learn': body.innerHTML = renderLearn(); break;
      case 'progress': body.innerHTML = renderProgress(); break;

      case 'explain':
        body.innerHTML = subBar('explain') + renderExplain(g, r, c);
        break;

      case 'reference':
        body.innerHTML = subBar('reference') + renderReference();
        var q = $('fnq');
        if (q) {
          q.value = App.fnQuery || '';
          q.oninput = function () { App.fnQuery = this.value; renderPanel(); $('fnq').focus(); };
        }
        break;
    }
  }
  App.renderPanel = renderPanel;

  function renderExplain(g, r, c) {
    switch (App.sub.explain) {
      case 'stack': return Panels.renderStack(App.wb, g.sheetName, r, c);
      case 'inspect': return Panels.renderInspector(App.wb, g.sheetName, r, c);
      case 'trace': return Panels.renderTracer(App.wb, g.sheetName, r, c);
      case 'shape': return Panels.renderShape(App.wb, g.sheetName, r, c);
      default: return Panels.renderStepper(App.wb, g.sheetName, r, c, App.stepAt);
    }
  }

  function renderReference() {
    switch (App.sub.reference) {
      case 'patterns': return Panels.renderPatterns(App.patOpen);
      case 'scenarios': return renderScenarios();
      default:
        return '<input id="fnq" class="xt-namebox" ' +
          'style="width:100%;border:1px solid var(--line);border-radius:5px;margin-bottom:8px" ' +
          'placeholder="filter functions…">' + Panels.renderFunctions(App.fnQuery, App.fnOpen);
    }
  }

  function onPanelClick(e) {
    var t = e.target.closest('[data-step],[data-ex],[data-scen],[data-lvl],[data-pat],' +
      '[data-fn],[data-goto],[data-act],[data-sub],[data-lesson]');
    if (!t) return;
    if (t.dataset.sub) {
      var bits = t.dataset.sub.split(':');
      App.sub[bits[0]] = bits[1];
      App.stepAt = null;
      renderPanel();
      return;
    }
    if (t.dataset.lesson !== undefined) {
      var li = +t.dataset.lesson;
      App.lessonOpen = (App.lessonOpen === li) ? null : li;
      renderPanel();
      return;
    }
    if (t.dataset.step !== undefined) {
      App.stepAt = +t.dataset.step;
      renderPanel();
      return;
    }
    if (t.dataset.ex) { loadExercise(t.dataset.ex); return; }
    if (t.dataset.scen) { $('scenarioPick').value = t.dataset.scen; loadScenario(t.dataset.scen); return; }
    if (t.dataset.lvl) {
      var L = XT.data.levels[+t.dataset.lvl];
      if (L && L.exercises && L.exercises.length) loadExercise(L.exercises[0].id);
      return;
    }
    if (t.dataset.pat) {
      App.patOpen = t.dataset.pat; App.sub.reference = 'patterns';
      showPanel('reference'); return;
    }
    if (t.dataset.fn) {
      App.fnOpen = t.dataset.fn; App.fnQuery = t.dataset.fn;
      App.sub.reference = 'functions'; showPanel('reference'); return;
    }
    if (t.dataset.goto) {
      e.preventDefault();
      var m = /^(.*)!([A-Za-z]{1,3}\d+)$/.exec(t.dataset.goto);
      if (m) {
        App.grid.setSheet(m[1]); renderTabs();
        var a = R.rc(m[2]);
        App.grid.select(a.r, a.c);
      }
      return;
    }
    if (t.dataset.act === 'hint') {
      App.hintsShown = (App.hintsShown || 0) + 1;
      renderPanel();
      return;
    }
    if (t.dataset.act === 'resetProgress') {
      Pr.reset(); renderPanel(); return;
    }
  }

  /* ------------------------------------------------------- learn by formula */

  /* The second way through the same four worksheets: one function at a time,
     in the order the real toolkits use them. */
  function renderLearn() {
    var track = XT.data.formulaTrack || [];
    var all = XT.data.drills || [];
    var done = all.filter(function (d) { return Pr.isDone(d.id); }).length;
    var pct = all.length ? (100 * done / all.length) : 0;

    var h = ['<h3>Learn by formula</h3>',
      '<p class="xt-muted">One function at a time, on the four toolkit ' +
      'worksheets, in the order those files actually use them. Each lesson ends ' +
      'in a graded task on a real sheet.</p>',
      '<div class="xt-prog"><i style="width:' + pct.toFixed(1) + '%"></i></div>',
      '<p class="xt-muted">' + done + ' of ' + all.length + ' tasks passed.</p>'];

    track.forEach(function (L, i) {
      var ds = L.drills || [];
      var nDone = ds.filter(function (d) { return Pr.isDone(d.id); }).length;
      var st = nDone === ds.length ? 'mastered' : (nDone ? 'attempted' : 'untouched');
      var open = App.lessonOpen === i;
      h.push('<div class="xt-flesson ' + st + (open ? ' open' : '') + '">');
      h.push('<div class="hd" data-lesson="' + i + '">' +
        '<span class="mk">' + (nDone === ds.length ? '●' : (nDone ? '◐' : '○')) + '</span>' +
        '<span class="nm">' + esc(L.title) + '</span>' +
        '<span class="ct">' + (L.cells ? L.cells.toLocaleString() + ' cells' : '') + '</span>' +
        '</div>');
      if (open) {
        h.push('<div class="bd">');
        h.push('<p><strong>What it does.</strong> ' + esc(L.what) + '</p>');
        h.push('<p><strong>In your files.</strong> ' + esc(L.estate) + '</p>');
        h.push('<p class="xt-warn"><strong>The trap.</strong> ' + esc(L.trap) + '</p>');
        h.push('<p class="xt-muted">Worked in: <strong>' +
          esc(XT.data.scenarios[L.scenario].title) + '</strong></p>');
        ds.forEach(function (d) {
          h.push('<div class="xt-drow' + (Pr.isDone(d.id) ? ' done' : '') +
            '" data-ex="' + d.id + '">' +
            '<span class="mk">' + (Pr.isDone(d.id) ? '✓' : '→') + '</span>' +
            '<span><strong>' + esc(d.title) + '</strong>' +
            '<span class="xt-muted"> · ' + esc(d.sheet) + ' ' +
            esc((d.targets || []).join(', ')) + '</span></span></div>');
        });
        h.push('</div>');
      }
      h.push('</div>');
    });
    return h.join('');
  }

  /* ---------------------------------------------------------------- lesson */

  function renderLesson() {
    var ex = App.exercise;
    if (!ex) {
      var sc = XT.data.scenarios[App.scenario];
      var h = ['<h3>' + esc(sc.title) + '</h3>',
        '<p>' + esc(sc.job) + '</p>',
        '<p class="xt-muted">Nothing is graded in here. Every sheet is open and ' +
        'every function works — type a formula and watch the Explain tab take it ' +
        'apart.</p>',
        '<h4>Two ways to learn it</h4>',
        '<div class="xt-row">' +
        '<button class="xt-tool primary" data-act="startFormulas">Learn by formula</button>' +
        '<button class="xt-tool" data-act="startCourse">Follow the course</button></div>',
        '<p class="xt-muted"><strong>Learn by formula</strong> takes one function ' +
        'at a time — IF, then VLOOKUP, then IFERROR — in the order the four ' +
        'toolkits use them. Start here if you want the formulas.<br>' +
        '<strong>The course</strong> builds one pipeline end to end, landing sheet ' +
        'to printed output. Start here if you want to understand a whole workbook.</p>'];
      var out = h.join('');
      setTimeout(function () {
        var b = $('pbody').querySelector('[data-act="startCourse"]');
        if (b) b.onclick = function () { loadExercise(XT.data.exercises[0].id); };
        var f = $('pbody').querySelector('[data-act="startFormulas"]');
        if (f) f.onclick = function () { App.lessonOpen = 0; showPanel('learn'); };
      }, 0);
      return out;
    }

    var res = App.lastResult;
    var h2 = [];
    h2.push('<div class="xt-lesson">');
    if (ex.track === 'formula') {
      h2.push('<div class="xt-lvl">Learn by formula · <strong>' + esc(ex.fn) +
        '</strong> · ' + esc(ex.lessonTitle) + '</div>');
    } else {
      h2.push('<div class="xt-lvl">Course · level ' + ex.level + ' · ' +
        esc(ex.levelTitle) +
        (ex.patterns && ex.patterns.length ? ' · ' + ex.patterns.join(' ') : '') + '</div>');
    }
    h2.push('<h3>' + esc(ex.title) + '</h3>');
    var L = ex.track === 'formula' ? null : XT.data.levels[ex.level];
    if (L && L.intro) h2.push('<p class="xt-muted">' + esc(L.intro) + '</p>');
    h2.push('<div class="xt-task"><strong>Task.</strong> ' + esc(ex.task) + '</div>');
    h2.push('<p class="xt-muted">Sheet <code>' + esc(ex.sheet) + '</code>' +
      (ex.targets ? ' · cells <code>' + ex.targets.join(', ') + '</code>' : '') + '</p>');

    if (ex.mode === 'defects') {
      h2.push('<h4>The five faults</h4><ul class="xt-tracelist">');
      ex.defects.forEach(function (d) {
        h2.push('<li><code>' + esc(d.cell) + '</code> — ' + esc(d.title) +
          ' <span class="xt-type">' + esc(d.id) + '</span></li>');
      });
      h2.push('</ul>');
    }

    // Hints, one at a time.
    var shown = App.hintsShown || 0;
    (ex.hints || []).forEach(function (hint, i) {
      if (i < shown) h2.push('<div class="xt-hint"><strong>Hint ' + (i + 1) + '.</strong> ' +
        esc(hint) + '</div>');
    });
    if (shown < (ex.hints || []).length) {
      h2.push('<div class="xt-row"><button class="xt-tool" data-act="hint">Show hint ' +
        (shown + 1) + ' of ' + ex.hints.length + '</button></div>');
    }

    if (res) {
      h2.push('<h4>' + (res.pass ? 'Passed' : 'Not yet') + '</h4>');
      if (res.pass) h2.push('<p class="xt-ok">Both checks pass: the values are right and ' +
        'the formula is built so it keeps being right.</p>');
      res.checks.forEach(function (c) {
        h2.push('<div class="xt-check ' + (c.ok ? 'pass' : 'fail') + '">' +
          '<div class="mark">' + (c.ok ? '✓' : '✗') + '</div>' +
          '<div><div class="t">' + esc(c.title) + '</div>' +
          '<div class="d">' + esc(c.detail || '') + '</div></div></div>');
      });
    }

    if (ex.realWorldNote) {
      h2.push('<h4>In your real files</h4><p class="xt-muted">' + esc(ex.realWorldNote) + '</p>');
    }
    if (ex.defectLink) {
      h2.push('<p class="xt-warn">Related catalog defect: <strong>' + esc(ex.defectLink) +
        '</strong></p>');
    }

    // Navigation, within whichever track this task belongs to.
    var all = ex.track === 'formula' ? XT.data.drills : XT.data.exercises;
    var i = all.indexOf(ex);
    h2.push('<div class="xt-row">');
    if (i > 0) h2.push('<button class="xt-tool" data-ex="' + all[i - 1].id + '">◀ previous</button>');
    if (i < all.length - 1) h2.push('<button class="xt-tool primary" data-ex="' + all[i + 1].id +
      '">next ▶</button>');
    if (ex.track === 'formula') {
      h2.push('<button class="xt-tool" data-act="backToLearn">all formulas</button>');
      setTimeout(function () {
        var b = $('pbody').querySelector('[data-act="backToLearn"]');
        if (b) b.onclick = function () {
          App.lessonOpen = ex.lessonIndex;
          showPanel('learn');
        };
      }, 0);
    }
    h2.push('</div>');
    h2.push('</div>');
    return h2.join('');
  }

  function levelList() {
    return XT.data.levels.map(function (L, i) {
      var st = Pr.levelStatus(L);
      var mark = st === 'mastered' ? '●' : (st === 'attempted' ? '◐' : '○');
      return '<div class="xt-lvlrow ' + st + '" data-lvl="' + i + '">' +
        '<div class="n">' + L.id + '</div><div class="st">' + mark + '</div>' +
        '<div>' + esc(L.title) + (L.v2 ? ' <span class="xt-type">v2</span>' : '') +
        '<div class="xt-muted" style="font-size:11.5px">' +
        (L.exercises.length ? L.exercises.length + ' exercise' +
          (L.exercises.length === 1 ? '' : 's') : 'not built yet') +
        (L.patterns ? ' · ' + L.patterns.join(' ') : '') + '</div></div></div>';
    }).join('');
  }

  /* --------------------------------------------------------------- progress */

  function renderProgress() {
    var cov = Pr.coverage();
    var all = XT.data.drills || [];
    var fDone = all.filter(function (d) { return Pr.isDone(d.id); }).length;
    var fPct = all.length ? (100 * fDone / all.length) : 0;

    var h = ['<h3>Progress</h3>'];
    h.push('<h4>Learn by formula</h4>');
    h.push('<div class="xt-prog"><i style="width:' + fPct.toFixed(1) + '%"></i></div>');
    h.push('<p class="xt-muted">' + fDone + ' of ' + all.length + ' formula tasks ' +
      'passed, across ' + (XT.data.formulaTrack || []).length + ' lessons.</p>');

    h.push('<h4>The course</h4>');
    h.push('<p>You have covered <strong>' + cov.percent.toFixed(1) + '%</strong> of the ' +
      'formula mass of your real estate.</p>');
    h.push('<div class="xt-prog"><i style="width:' + cov.percent.toFixed(1) + '%"></i></div>');
    h.push('<p class="xt-muted">' + Math.round(cov.massCovered).toLocaleString() + ' of ' +
      cov.massTotal.toLocaleString() + ' measured formula cells, weighted by how often each ' +
      'pattern actually appears. ' + cov.exercisesDone + ' of ' + cov.exercises +
      ' exercises passed.</p>');

    h.push('<h4>Patterns</h4><div class="xt-pgrid">');
    (XT.data.patterns || []).forEach(function (p) {
      var st = Pr.patternStatus(p.id);
      var pct = (100 * p.cells / XT.data.patternMass).toFixed(1);
      h.push('<div class="xt-pcell ' + st + '" data-pat="' + p.id + '" title="' +
        esc(p.title) + '"><b>' + p.id + '</b>' + pct + '%</div>');
    });
    h.push('</div>');
    h.push('<h4>Levels</h4>' + levelList());
    h.push('<div class="xt-row"><button class="xt-tool" data-act="resetProgress">' +
      'Reset progress</button></div>');
    return h.join('');
  }

  /* -------------------------------------------------------------- scenarios */

  function renderScenarios() {
    var h = ['<h3>Toolkits</h3>',
      '<p class="xt-muted">The four toolkits you work in daily come first, then the ' +
      'other pipelines from the estate. All rebuilt with synthetic data: same sheet ' +
      'names, same column positions, same dirt.</p>'];
    XT.data.scenarioOrder.forEach(function (id) {
      var sc = XT.data.scenarios[id];
      h.push('<div class="xt-scen">');
      h.push('<h4>' + id + ' · ' + esc(sc.title) + '</h4>');
      h.push('<div class="src">' + esc(sc.source) + '</div>');
      h.push('<p>' + esc(sc.job) + '</p>');
      h.push('<p class="xt-muted">' + esc(sc.why) + '</p>');
      h.push('<div class="xt-layers">' + sc.layers.map(function (l, i) {
        return (i ? '<i>→</i>' : '') + '<span>' + esc(l) + '</span>';
      }).join('') + '</div>');
      h.push('<div class="xt-row"><button class="xt-tool" data-scen="' + id +
        '">Open in sandbox</button>');
      var exs = XT.data.exercises.filter(function (e) { return e.scenario === id; });
      if (exs.length) h.push('<button class="xt-tool" data-ex="' + exs[0].id + '">' +
        exs.length + ' course exercise' + (exs.length === 1 ? '' : 's') + '</button>');
      var drs = (XT.data.drills || []).filter(function (d) { return d.scenario === id; });
      if (drs.length) h.push('<button class="xt-tool" data-ex="' + drs[0].id + '">' +
        drs.length + ' formula task' + (drs.length === 1 ? '' : 's') + '</button>');
      h.push('</div></div>');
    });
    h.push('<h4>VBA boundary</h4><p class="xt-muted">In the real files a button runs a ' +
      'macro at several of these steps — 71 of your 152 workbooks contain VBA. ' +
      'Here you do those steps by hand, which is the point.</p>');
    return h.join('');
  }

  /* ------------------------------------------------------------- check / show */

  function check() {
    if (!App.exercise) { setStatus('Sandbox mode — nothing to check.'); return; }
    App.grid.commitEdit();
    App.wb.recalc();
    App.lastResult = Grader.run(App.wb, App.exercise);
    if (App.lastResult.pass) {
      Pr.markDone(App.exercise.id);
      setStatus('Passed — ' + (App.exercise.track === 'formula'
        ? App.exercise.fn + ': ' : 'level ' + App.exercise.level + ', ') +
        App.exercise.title, 'ok');
    } else {
      setStatus('Not yet. The Lesson panel says what is off.', 'warn');
    }
    showPanel('lesson');
  }
  App.check = check;

  function showMe() {
    var ex = App.exercise;
    if (!ex || !ex.solution) { setStatus('No model answer for this one.'); return; }
    Grader.applySolution(App.wb, ex);
    App.wb.recalc();
    App.grid.layout();
    App.grid.refresh();
    setStatus('Model answer filled in. Read it, then press Reset and do it yourself.', 'warn');
  }

  /* ----------------------------------------------------------------- paste */

  App.doPaste = function (asValues) {
    var text = $('pasteText').value;
    $('pasteModal').classList.add('hidden');
    App.grid.pasteText(text, asValues);
    App.grid.scroller.focus();
  };

  /* ------------------------------------------------------------------ boot */

  document.addEventListener('DOMContentLoaded', function () {
    try {
      var th = localStorage.getItem('xt.theme');
      if (th) document.documentElement.setAttribute('data-theme', th);
    } catch (e) {}
    App.start();
  });

})(window.XT = window.XT || {});
