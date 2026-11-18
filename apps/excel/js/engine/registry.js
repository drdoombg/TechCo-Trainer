/* =============================================================================
   Excel Trainer — engine/registry.js
   The function table. Every function the trainer knows is registered here by
   the files in js/engine/functions/.

   An entry looks like:
     {
       name   : 'VLOOKUP',
       min, max : argument count limits
       arrayArgs: [1]        argument positions evaluated in ARRAY context
       lazy   : false        true = impl receives AST nodes, not values
       impl   : function(args, ctx)
       sig    : 'VLOOKUP(lookup_value, table_array, col_index_num, [range_lookup])'
       doc    : one plain sentence
       estate : what the source workbooks actually do with it
       pats   : ['P5','P9']  the catalog patterns it serves
       group  : 'lookup'
     }

   Two deliberate design points:

   ARRAY ARGUMENTS. Excel evaluates some arguments in array context whether or
   not the formula was entered with Ctrl+Shift+Enter. That is the only reason
       XLOOKUP(key, 'MIMS SAVE FILE'!$M:$M & 'MIMS SAVE FILE'!V:V, …)
   works as an ordinary formula: the & over two whole columns has to produce an
   array, not implicitly intersect down to one row. arrayArgs says which
   positions get that treatment (pattern P9, 8,564 cells).

   LAZY FUNCTIONS. IF must not evaluate the branch it does not take, or
       =IF(A1=0,"",1/A1)
   would divide by zero anyway. IF, IFERROR and IFNA are therefore lazy.

   UNSUPPORTED, declared. Functions that appear in the source estate but are
   out of scope (catalog §8) are still registered, so a pasted formula parses
   and reports honestly instead of silently showing #NAME?.
   ============================================================================= */
(function (XT) {
  'use strict';

  var FN = Object.create(null);
  XT.FN = FN;

  XT.defineFunction = function (spec) {
    if (!spec || !spec.name) throw new Error('defineFunction needs a name');
    spec.min = spec.min === undefined ? 0 : spec.min;
    spec.max = spec.max === undefined ? 255 : spec.max;
    spec.arrayArgs = spec.arrayArgs || [];
    spec.pats = spec.pats || [];
    spec.group = spec.group || 'other';
    FN[spec.name] = spec;
    (spec.aliases || []).forEach(function (a) { FN[a.toUpperCase()] = spec; });
    return spec;
  };

  // Register a function we knowingly do not implement.
  XT.declareUnsupported = function (name, why, pats) {
    XT.defineFunction({
      name: name, min: 0, max: 255, group: 'unsupported',
      sig: name + '(…)',
      doc: 'Out of scope for this trainer.',
      estate: why || '',
      pats: pats || [],
      unsupported: true,
      impl: function () {
        return XT.V.name(name + ' is deliberately not implemented in the trainer. ' + (why || ''));
      }
    });
  };

  XT.functionList = function () {
    var seen = {}, out = [];
    Object.keys(FN).forEach(function (k) {
      var f = FN[k];
      if (seen[f.name]) return;
      seen[f.name] = 1;
      out.push(f);
    });
    out.sort(function (a, b) { return a.name < b.name ? -1 : 1; });
    return out;
  };

  // Names offered by the formula-bar autocomplete, most useful first.
  XT.functionNames = function (prefix) {
    var p = String(prefix || '').toUpperCase();
    return XT.functionList()
      .filter(function (f) { return !f.unsupported && f.name.indexOf(p) === 0; })
      .map(function (f) { return f.name; });
  };

})(window.XT = window.XT || {});
