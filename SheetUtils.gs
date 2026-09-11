/**
 * SheetUtils — the helpers every Apps Script project ends up rewriting.
 * Fast, batched, and safe on large sheets.
 *
 *   SheetUtils.readObjects('Orders')                 -> [{Email:'a@b.com', Amount: 12}, ...]
 *   SheetUtils.writeObjects('Clean', rows)            -> writes headers + rows in one call
 *   SheetUtils.dedupe('Leads', ['Email'])             -> removes duplicate rows by key columns
 *   SheetUtils.mergeSheets(['Jan','Feb','Mar'], 'All') -> stacks sheets with the same headers
 *   SheetUtils.upsert('Customers', 'Email', rows)     -> update matching rows, append new ones
 *   SheetUtils.withLock(function(){ ... })            -> prevents concurrent trigger runs
 *   SheetUtils.retry(function(){ return UrlFetchApp.fetch(url); }, 3)
 *
 * Built by SheetFix Studio — ishubham8.gumroad.com/l/wfxzs
 */
var SheetUtils = (function () {
  function ss() { return SpreadsheetApp.getActiveSpreadsheet(); }
  function sheet(name, create) {
    var sh = ss().getSheetByName(name);
    if (!sh && create) sh = ss().insertSheet(name);
    if (!sh) throw new Error('Sheet not found: ' + name);
    return sh;
  }

  function readObjects(name, opts) {
    opts = opts || {};
    var sh = sheet(name);
    var values = sh.getDataRange().getValues();
    if (values.length < 2) return [];
    var headers = values[0].map(function (h) { return String(h).trim(); });
    var out = [];
    for (var r = 1; r < values.length; r++) {
      var row = values[r];
      if (opts.skipEmpty !== false && row.every(function (c) { return c === '' || c == null; })) continue;
      var o = { _row: r + 1 };
      headers.forEach(function (h, i) { if (h) o[h] = row[i]; });
      out.push(o);
    }
    return out;
  }

  function writeObjects(name, rows, opts) {
    opts = opts || {};
    var sh = sheet(name, true);
    var headers = opts.headers || uniqueKeys_(rows);
    var data = rows.map(function (o) { return headers.map(function (h) { return o[h] == null ? '' : o[h]; }); });
    if (opts.append && sh.getLastRow() > 0) {
      if (data.length) sh.getRange(sh.getLastRow() + 1, 1, data.length, headers.length).setValues(data);
    } else {
      sh.clearContents();
      sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
      sh.setFrozenRows(1);
      if (data.length) sh.getRange(2, 1, data.length, headers.length).setValues(data);
    }
    return data.length;
  }

  function dedupe(name, keyColumns, opts) {
    opts = opts || {};
    var sh = sheet(name);
    var values = sh.getDataRange().getValues();
    var headers = values[0];
    var idx = keyColumns.map(function (k) {
      var i = headers.indexOf(k);
      if (i < 0) throw new Error('Column not found: ' + k);
      return i;
    });
    var seen = {}, keep = [values[0]], removed = 0;
    for (var r = 1; r < values.length; r++) {
      var key = idx.map(function (i) {
        var v = values[r][i];
        return opts.caseSensitive ? String(v) : String(v).trim().toLowerCase();
      }).join('\u0001');
      if (seen[key]) { removed++; continue; }
      seen[key] = true;
      keep.push(values[r]);
    }
    sh.clearContents();
    sh.getRange(1, 1, keep.length, headers.length).setValues(keep);
    return removed;
  }

  function mergeSheets(names, targetName, opts) {
    opts = opts || {};
    var all = [], headers = null;
    names.forEach(function (n) {
      var v = sheet(n).getDataRange().getValues();
      if (v.length < 2) return;
      if (!headers) headers = v[0].concat(opts.sourceColumn ? [opts.sourceColumn] : []);
      for (var r = 1; r < v.length; r++) {
        var row = v[r].slice(0, headers.length - (opts.sourceColumn ? 1 : 0));
        while (row.length < headers.length - (opts.sourceColumn ? 1 : 0)) row.push('');
        if (opts.sourceColumn) row.push(n);
        all.push(row);
      }
    });
    var t = sheet(targetName, true);
    t.clearContents();
    if (!headers) return 0;
    t.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    if (all.length) t.getRange(2, 1, all.length, headers.length).setValues(all);
    return all.length;
  }

  function upsert(name, keyColumn, rows) {
    var sh = sheet(name, true);
    var existing = readObjects(name);
    var headers = sh.getLastRow() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] : uniqueKeys_(rows);
    var byKey = {};
    existing.forEach(function (o) { byKey[String(o[keyColumn]).toLowerCase()] = o._row; });
    var appends = [], updated = 0;
    rows.forEach(function (o) {
      var k = String(o[keyColumn]).toLowerCase();
      var vals = headers.map(function (h) { return o[h] == null ? '' : o[h]; });
      if (byKey[k]) { sh.getRange(byKey[k], 1, 1, headers.length).setValues([vals]); updated++; }
      else appends.push(vals);
    });
    if (!sh.getLastRow()) sh.getRange(1, 1, 1, headers.length).setValues([headers]);
    if (appends.length) sh.getRange(sh.getLastRow() + 1, 1, appends.length, headers.length).setValues(appends);
    return { updated: updated, inserted: appends.length };
  }

  function withLock(fn, timeoutMs) {
    var lock = LockService.getScriptLock();
    if (!lock.tryLock(timeoutMs || 30000)) { Logger.log('Another run holds the lock; skipping.'); return null; }
    try { return fn(); } finally { lock.releaseLock(); }
  }

  function retry(fn, times, baseMs) {
    var err;
    for (var i = 0; i < (times || 3); i++) {
      try { return fn(); } catch (e) { err = e; Utilities.sleep((baseMs || 500) * Math.pow(2, i)); }
    }
    throw err;
  }

  function uniqueKeys_(rows) {
    var seen = {}, keys = [];
    rows.forEach(function (o) { Object.keys(o).forEach(function (k) { if (k !== '_row' && !seen[k]) { seen[k] = 1; keys.push(k); } }); });
    return keys;
  }

  return { readObjects: readObjects, writeObjects: writeObjects, dedupe: dedupe, mergeSheets: mergeSheets, upsert: upsert, withLock: withLock, retry: retry };
})();



That is all six files. Questions or a custom adaptation: message me through the Fiverr link in the Read me section.