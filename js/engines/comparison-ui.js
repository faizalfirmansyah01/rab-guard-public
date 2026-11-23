/* Thin UI adapter for independent comparison ingestion and result rendering. */
(function (window, document) {
  'use strict';

  var state = { canonicalA: null, canonicalB: null, fileA: null, fileB: null, readyA: false, readyB: false, result: null, loadTokenA: 0, loadTokenB: 0 };
  var elements = {};
  var initialized = false;

  function byId(id) { return document.getElementById(id); }
  function setText(element, value) { if (element) element.textContent = value === null || value === undefined ? '' : String(value); }
  function formatMoney(value) {
    if (typeof value !== 'number' || !Number.isFinite(value)) return 'Tidak tersedia';
    return 'Rp ' + value.toLocaleString('id-ID', { maximumFractionDigits: 2 });
  }
  function pathLabel(path) {
    if (!Array.isArray(path) || !path.length) return 'Hierarki tidak tersedia';
    var ordered = path.slice().sort(function (a, b) { return Number(a.levelIndex || 0) - Number(b.levelIndex || 0); });
    var labels = [], next = 0;
    ordered.forEach(function (level) {
      while (next < Number(level.levelIndex || 0)) { labels.push('[level ' + (next + 1) + ' kosong]'); next++; }
      labels.push(level.value === null ? '[belum terselesaikan]' : String(level.value));
      next = Number(level.levelIndex || 0) + 1;
    });
    return labels.join(' > ');
  }
  function sourceLabel(ref) {
    if (!ref) return 'Tidak tersedia';
    return [ref.workbook, ref.sheet, ref.row ? 'baris ' + ref.row : null].filter(Boolean).join(' • ');
  }
  function detailsNode(title, value) {
    var details = document.createElement('details');
    var summary = document.createElement('summary'); summary.textContent = title; details.appendChild(summary);
    var pre = document.createElement('pre'); pre.style.whiteSpace = 'pre-wrap'; pre.style.overflowWrap = 'anywhere';
    pre.textContent = JSON.stringify(value, null, 2); details.appendChild(pre); return details;
  }
  function statCard(label, value) {
    var box = document.createElement('div'); box.className = 'card';
    var title = document.createElement('strong'); title.textContent = label; box.appendChild(title);
    var number = document.createElement('p'); number.textContent = value; box.appendChild(number);
    return box;
  }
  function flattenResult(result) {
    if (!result) return [];
    return (result.changes || []).map(function (entry) { return { kind: 'change', value: entry }; })
      .concat((result.candidates || []).map(function (entry) { return { kind: 'uncertain', value: entry }; }))
      .concat((result.relationships || []).map(function (entry) { return { kind: 'split_merge', value: entry }; }));
  }
  function getVisibleEntries(result, filter) {
    var rows = flattenResult(result);
    filter = filter || 'all';
    return rows.filter(function (row) {
      var value = row.value;
      if (filter === 'matched') return row.kind === 'change' && (value.status === 'EXACT_MATCH' || value.status === 'HIGH_CONFIDENCE_MATCH');
      if (filter === 'added') return value.status === 'ADDED';
      if (filter === 'removed') return value.status === 'REMOVED';
      if (filter === 'uncertain') return row.kind === 'uncertain';
      if (filter === 'split_merge') return row.kind === 'split_merge';
      return true;
    });
  }
  function filterRows(rows) {
    return getVisibleEntries({ changes: rows.filter(function (row) { return row.kind === 'change'; }).map(function (row) { return row.value; }),
      candidates: rows.filter(function (row) { return row.kind === 'uncertain'; }).map(function (row) { return row.value; }),
      relationships: rows.filter(function (row) { return row.kind === 'split_merge'; }).map(function (row) { return row.value; }) }, elements.filter ? elements.filter.value : 'all');
  }
  function renderRows() {
    if (!elements.rows || !state.result) return;
    elements.rows.textContent = '';
    var all = flattenResult(state.result), visible = filterRows(all);
    setText(elements.visibleCount, 'Menampilkan ' + visible.length + ' dari ' + all.length + ' baris komparasi. Ringkasan total tetap menggunakan seluruh workbook.');
    if (!visible.length) {
      var empty = document.createElement('tr'), cell = document.createElement('td'); cell.colSpan = 5; cell.textContent = 'Tidak ada hasil untuk filter ini.'; empty.appendChild(cell); elements.rows.appendChild(empty); return;
    }
    visible.forEach(function (row) {
      var value = row.value, tr = document.createElement('tr');
      var status = document.createElement('td'); status.textContent = value.status || row.kind; tr.appendChild(status);
      var changes = document.createElement('td'); changes.textContent = (value.changeTypes || value.reasons || [value.reason] || []).filter(Boolean).join(', '); tr.appendChild(changes);
      var a = document.createElement('td'); a.textContent = sourceLabel(value.sourceA || (value.itemsA && value.itemsA.map(sourceLabel).join(' | '))); tr.appendChild(a);
      var b = document.createElement('td'); b.textContent = sourceLabel(value.sourceB || (value.itemsB && value.itemsB.map(sourceLabel).join(' | '))); tr.appendChild(b);
      var evidence = document.createElement('td');
      if (value.changes) {
        var desc = document.createElement('p'); desc.textContent = (value.sourceA && value.sourceA.description || '—') + ' → ' + (value.sourceB && value.sourceB.description || '—'); evidence.appendChild(desc);
        var amount = value.changes.amount || {};
        var amountText = document.createElement('p'); amountText.textContent = 'Jumlah impor: ' + formatMoney(amount.before) + ' → ' + formatMoney(amount.after) + '; delta ' + formatMoney(amount.delta) + (amount.percentageChange === null ? ' (persentase N/A)' : ' (' + amount.percentageChange.toFixed(2) + '%)'); evidence.appendChild(amountText);
        evidence.appendChild(detailsNode('Nilai, category path, dan provenance kedua sumber', { matchEvidence: value.matchEvidence, changes: value.changes, sourceA: value.sourceA, sourceB: value.sourceB }));
      } else if (value.sourceA || value.sourceB) {
        evidence.appendChild(detailsNode('Kandidat dan provenance', value));
      } else {
        evidence.appendChild(detailsNode('Item dalam kelompok yang perlu diverifikasi', value));
      }
      tr.appendChild(evidence); elements.rows.appendChild(tr);
    });
  }
  function renderResult() {
    var result = state.result, summary = result && result.summary;
    if (!result || !summary) return;
    elements.results.style.display = 'block';
    if (result.status === 'failed') {
      setText(elements.resultStatus, 'Komparasi gagal: ' + (result.errors || []).join(' '));
      elements.summary.textContent = '';
      elements.rows.textContent = '';
      return;
    }
    setText(elements.resultStatus, result.status === 'partial' ? 'Komparasi selesai dengan item/kondisi yang perlu verifikasi.' : 'Komparasi selesai.');
    elements.summary.textContent = '';
    var cards = [
      ['Workbook A — nilai impor', formatMoney(summary.workbookA.importedAmount) + ' (' + summary.workbookA.importedAmountCoverage + '/' + summary.workbookA.workItems + ' nilai tersedia' + (summary.workbookA.missingImportedAmountItems ? '; parsial' : '') + ')'],
      ['Workbook A — hitung ulang', formatMoney(summary.workbookA.recalculatedAmount) + ' (' + summary.workbookA.calculableItems + '/' + summary.workbookA.workItems + ' calculable' + (summary.workbookA.incompleteItems ? '; parsial' : '') + ')'],
      ['Workbook B — nilai impor', formatMoney(summary.workbookB.importedAmount) + ' (' + summary.workbookB.importedAmountCoverage + '/' + summary.workbookB.workItems + ' nilai tersedia' + (summary.workbookB.missingImportedAmountItems ? '; parsial' : '') + ')'],
      ['Workbook B — hitung ulang', formatMoney(summary.workbookB.recalculatedAmount) + ' (' + summary.workbookB.calculableItems + '/' + summary.workbookB.workItems + ' calculable' + (summary.workbookB.incompleteItems ? '; parsial' : '') + ')'],
      ['Pasangan / baru / dihapus', summary.matchedItems + ' / ' + summary.addedItems + ' / ' + summary.removedItems],
      ['Delta nilai impor B − A', formatMoney(summary.importedNetChange)],
      ['Perubahan item terklasifikasi', formatMoney(summary.classifiedImportedAmountChange)],
      ['Kandidat / split-merge / ambigu', summary.uncertainItems + ' / ' + summary.possibleSplitMergeGroups + ' / ' + summary.ambiguousCodeGroups],
      ['Delta belum terjelaskan', formatMoney(summary.unreconciledImportedAmountChange)]
    ];
    cards.forEach(function (entry) { elements.summary.appendChild(statCard(entry[0], entry[1])); });
    renderRows();
  }
  function updateCompareButton() { if (elements.run) elements.run.disabled = !(state.canonicalA && state.canonicalB && state.readyA && state.readyB); }
  function ingestFile(side, file) {
    if (!file) return Promise.resolve(null);
    var tokenKey = side === 'A' ? 'loadTokenA' : 'loadTokenB';
    var token = ++state[tokenKey];
    var status = side === 'A' ? elements.statusA : elements.statusB;
    if (side === 'A') state.fileA = file; else state.fileB = file;
    if (side === 'A') state.readyA = false; else state.readyB = false;
    state.result = null;
    if (elements.results) elements.results.style.display = 'none';
    updateCompareButton();
    setText(status, 'Membaca, menginspeksi, dan menormalisasi seluruh workbook…');
    setText(elements.error, '');
    return window.XlsxReader.readFileAsync(file).then(function (workbook) {
      if (state[tokenKey] !== token) return null;
      var inspection = window.WorkbookInspector.inspect(workbook);
      var canonical = window.DataNormalizer.normalizeWorkbook(workbook, inspection);
      var validation = window.IngestionValidator.validate(canonical);
      if (!canonical.workItems.length || !canonical.statistics.processedBoqSheets) {
        var unresolved = (canonical.unresolvedSheets || []).map(function (sheet) { return sheet.name + ': ' + (sheet.unresolvedReasons || []).join(' '); }).join(' | ');
        throw new Error('Tidak ada worksheet BOQ yang berhasil diproses. ' + unresolved);
      }
      if (!validation.valid) throw new Error(validation.errors.join(' '));
      canonical.project.sourceFilename = file.name;
      canonical.project.projectName = canonical.project.projectName || file.name;
      var unresolvedCount = (canonical.unresolvedSheets || []).length;
      setText(status, 'Siap: ' + canonical.workItems.length + ' item dari ' + canonical.statistics.processedBoqSheets + ' sheet BOQ; ' + unresolvedCount + ' sheet unresolved.' + (validation.warnings.length ? ' Ada catatan validasi.' : ''));
      if (state[tokenKey] !== token) return null;
      if (side === 'A') { state.canonicalA = canonical; state.readyA = true; } else { state.canonicalB = canonical; state.readyB = true; }
      state.result = null;
      elements.results.style.display = 'none';
      updateCompareButton();
      return canonical;
    }).catch(function (error) {
      if (state[tokenKey] !== token) return null;
      setText(status, 'File baru gagal diproses; hasil komparasi sebelumnya dibatalkan. Unggah ulang sumber yang valid.');
      setText(elements.error, (side === 'A' ? 'Workbook A: ' : 'Workbook B: ') + (error.message || String(error)));
      updateCompareButton();
      return null;
    });
  }
  function runComparison() {
    if (!state.canonicalA || !state.canonicalB || !state.readyA || !state.readyB) return;
    setText(elements.error, '');
    state.result = window.AnalysisEngineAdapter.runComparison(window.ComparisonEngine, state.canonicalA, state.canonicalB);
    renderResult();
  }
  function init() {
    if (initialized) return;
    elements.fileA = byId('comparison-file-a'); elements.fileB = byId('comparison-file-b');
    elements.statusA = byId('comparison-status-a'); elements.statusB = byId('comparison-status-b');
    elements.error = byId('comparison-error'); elements.run = byId('btn-run-rab-comparison');
    elements.results = byId('comparison-results'); elements.resultStatus = byId('comparison-result-status');
    elements.summary = byId('comparison-summary'); elements.filter = byId('comparison-filter');
    elements.visibleCount = byId('comparison-visible-count'); elements.rows = byId('comparison-rows');
    if (!elements.fileA || !elements.fileB || !elements.run) return;
    initialized = true;
    elements.fileA.addEventListener('change', function () { ingestFile('A', elements.fileA.files && elements.fileA.files[0]); });
    elements.fileB.addEventListener('change', function () { ingestFile('B', elements.fileB.files && elements.fileB.files[0]); });
    elements.run.addEventListener('click', runComparison);
    elements.filter.addEventListener('change', renderRows);
  }
  var ComparisonUI = {
    init: init,
    ingestFile: ingestFile,
    run: runComparison,
    getVisibleEntries: getVisibleEntries,
    getState: function () { return { canonicalA: state.canonicalA, canonicalB: state.canonicalB, fileA: state.fileA, fileB: state.fileB, readyA: state.readyA, readyB: state.readyB, result: state.result }; },
    restoreState: function (snapshot) {
      snapshot = snapshot || {};
      state.canonicalA = snapshot.canonicalA || null;
      state.canonicalB = snapshot.canonicalB || null;
      state.fileA = snapshot.fileA || null;
      state.fileB = snapshot.fileB || null;
      state.readyA = snapshot.readyA !== false && !!state.canonicalA;
      state.readyB = snapshot.readyB !== false && !!state.canonicalB;
      state.result = snapshot.result || null;
      updateCompareButton();
      if (state.result) renderResult();
      else if (elements.results) elements.results.style.display = 'none';
      return { restored: !!(state.canonicalA || state.canonicalB || state.result) };
    },
    render: renderResult
  };
  window.ComparisonUI = ComparisonUI;
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init, { once: true }); else init();
})(typeof window !== 'undefined' ? window : this, typeof document !== 'undefined' ? document : null);
