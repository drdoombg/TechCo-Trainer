/* =============================================================================
   Excel Trainer — trainer/progress.js
   Mastery tracking in localStorage, wrapped in try/catch because file:// in a
   private window can refuse it and the trainer must still work.

   Coverage is WEIGHTED by the catalog's measured formula counts, so the headline
   number means something real: "you have covered 71% of this estate's formula
   mass", not "you have done 12 exercises".
   ============================================================================= */
(function (XT) {
  'use strict';

  var KEY = 'xt.progress.v1';
  var Pr = {};
  XT.Progress = Pr;

  var state = { done: {}, attempted: {}, lastExercise: null, openedPatterns: {} };

  function load() {
    try {
      var raw = localStorage.getItem(KEY);
      if (raw) {
        var o = JSON.parse(raw);
        if (o && typeof o === 'object') {
          state.done = o.done || {};
          state.attempted = o.attempted || {};
          state.lastExercise = o.lastExercise || null;
          state.openedPatterns = o.openedPatterns || {};
        }
      }
    } catch (e) { /* no storage: run in memory */ }
  }
  function save() {
    try { localStorage.setItem(KEY, JSON.stringify(state)); } catch (e) { /* ignore */ }
  }
  load();

  Pr.state = state;
  Pr.markAttempted = function (id) { state.attempted[id] = true; save(); };
  Pr.markDone = function (id) { state.done[id] = true; state.attempted[id] = true; save(); };
  Pr.setLast = function (id) { state.lastExercise = id; save(); };
  Pr.isDone = function (id) { return !!state.done[id]; };
  Pr.isAttempted = function (id) { return !!state.attempted[id]; };
  Pr.reset = function () {
    state = { done: {}, attempted: {}, lastExercise: null, openedPatterns: {} };
    Pr.state = state;
    save();
  };

  // 'mastered' | 'attempted' | 'untouched' for a level.
  Pr.levelStatus = function (level) {
    var exs = (level.exercises || []);
    if (!exs.length) return level.v2 ? 'v2' : 'untouched';
    var done = exs.filter(function (e) { return Pr.isDone(e.id); }).length;
    if (done === exs.length) return 'mastered';
    if (done || exs.some(function (e) { return Pr.isAttempted(e.id); })) return 'attempted';
    return 'untouched';
  };

  Pr.patternStatus = function (patternId) {
    var exs = XT.data.exercises.filter(function (e) {
      return (e.patterns || []).indexOf(patternId) >= 0;
    });
    if (!exs.length) return 'none';
    if (exs.some(function (e) { return Pr.isDone(e.id); })) return 'mastered';
    if (exs.some(function (e) { return Pr.isAttempted(e.id); })) return 'attempted';
    return 'untouched';
  };

  /* The weighted figure. A pattern counts as covered once any exercise that
     teaches it has been passed. */
  Pr.coverage = function () {
    var total = 0, covered = 0, attempted = 0;
    (XT.data.patterns || []).forEach(function (p) {
      total += p.cells;
      var st = Pr.patternStatus(p.id);
      if (st === 'mastered') covered += p.cells;
      else if (st === 'attempted') attempted += p.cells;
    });
    var exs = XT.data.exercises.length;
    var doneCount = XT.data.exercises.filter(function (e) { return Pr.isDone(e.id); }).length;
    return {
      massTotal: total,
      massCovered: covered,
      massAttempted: attempted,
      percent: total ? (100 * covered / total) : 0,
      exercises: exs,
      exercisesDone: doneCount
    };
  };

})(window.XT = window.XT || {});
