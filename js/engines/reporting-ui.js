/* UI adapter: renders and exports stored engine results without rerunning analysis. */
(function (window, document) {
  'use strict';

  var state = { document: null, bound: false };
  function el(id) { return document.getElementById(id); }
  function setText(id, value) { var node = el(id); if (node) node.textContent = value; }
  function currentFilename(canonical) {
    return canonical && canonical.project && canonical.project.sourceFilename || null;
  }
  function sameName(a, b) { return !!a && !!b && String(a).toLowerCase() === String(b).toLowerCase(); }
  function storedResults() {
    var ui = window.UniversalUI && window.UniversalUI.getCurrentState ? window.UniversalUI.getCurrentState() : {};
    var compare = window.ComparisonUI && window.ComparisonUI.getState ? window.ComparisonUI.getState() : {};
    var canonical = ui.currentCanonicalData || null;
    var filename = currentFilename(canonical);
    var contextId = filename ? 'workbook:' + filename.toLowerCase() : null;
    var engines = {};
    if (ui.currentAuditResult) engines.audit = { engineName: 'UniversalChecker', status: ui.currentAuditResult.status || 'complete', result: ui.currentAuditResult, sourceFilename: filename, sourceWorkbooks: filename ? [{ filename: filename }] : [], contextId: contextId };
    if (ui.currentUnitConsistencyResult) engines.unit = { engineName: 'UnitConsistencyEngine', result: ui.currentUnitConsistencyResult, sourceFilename: filename, sourceWorkbooks: filename ? [{ filename: filename }] : [], contextId: contextId };
    if (compare.result) {
      var names = [];
      var summary = compare.result.summary || {};
      if (summary.workbookA && summary.workbookA.filename) names.push(summary.workbookA.filename);
      if (summary.workbookB && summary.workbookB.filename) names.push(summary.workbookB.filename);
      if (!names.length) (compare.result.sourceReferences || []).forEach(function (ref) { var n = ref.workbook || ref.filename; if (n && names.indexOf(n) < 0) names.push(n); });
      engines.comparison = { engineName: 'RABComparisonEngine', result: compare.result, sourceWorkbooks: names.map(function (name, index) { return { filename: name, role: index === 0 ? 'A' : 'B' }; }), contextId: contextId && names.some(function (name) { return sameName(name, filename); }) ? contextId : null };
    }
    return { engines: engines, canonical: canonical, filename: filename, contextId: contextId };
  }
  function availableType(type, engines) {
    if (type === 'combined') {
      var entries = ['audit', 'unit', 'comparison'].map(function (key) { return engines[key]; }).filter(function (entry) { return !!(entry && entry.result); });
      return entries.length >= 2 && !!entries[0].contextId && entries.every(function (entry) { return entry.contextId === entries[0].contextId; });
    }
    return !!(engines[type] && engines[type].result && engines[type].result.status !== 'failed');
  }
  function build(type) {
    var stored = storedResults();
    if (!availableType(type, stored.engines)) throw new Error('Hasil untuk jenis laporan ini belum tersedia. Jalankan analisis yang diperlukan terlebih dahulu.');
    var engines = Object.assign({}, stored.engines);
    var warnings = [];
    if (type === 'combined') {
      // Only include comparison when it explicitly contains the current workbook;
      // otherwise it is a separate project/version context and must remain separate.
      if (engines.comparison && !engines.comparison.contextId) {
        warnings.push('Hasil komparasi tersedia, tetapi tidak dimasukkan karena pasangan workbook tidak memiliki konteks sumber yang sama dengan audit aktif. Buat laporan Komparasi secara terpisah.');
        engines.comparison = { engineName: 'RABComparisonEngine', status: 'excluded', result: null, sourceWorkbooks: [] };
      }
      var contexts = Object.keys(engines).filter(function (key) { return engines[key].result && engines[key].result.status !== 'failed'; }).map(function (key) { return engines[key].contextId; }).filter(Boolean);
      if (contexts.length && contexts.some(function (value) { return value !== stored.contextId; })) throw new Error('Hasil engine tidak berasal dari konteks workbook yang sama.');
    }
    var projectName = stored.canonical && stored.canonical.project && stored.canonical.project.projectName || null;
    var sourceWorkbooks = [];
    var sourceKeys = type === 'combined' ? ['audit', 'unit', 'comparison'] : [type];
    sourceKeys.forEach(function (key) { (engines[key] && engines[key].sourceWorkbooks || []).forEach(function (book) { sourceWorkbooks.push(book); }); });
    var canonicalSummary = stored.canonical && stored.canonical.statistics ? stored.canonical.statistics : {};
    return window.ReportingEngine.createDocument({
      reportType: type,
      engines: engines,
      warnings: warnings,
      sourceWorkbooks: sourceWorkbooks,
      canonicalSummary: canonicalSummary,
      projectMetadata: projectName ? { projectName: projectName, sourceFilename: stored.filename } : (stored.filename ? { sourceFilename: stored.filename } : null),
      assumptions: ['Laporan menyalin hasil analisis yang telah tersimpan; ekspor tidak menjalankan ulang audit atau komparasi.'],
      limitations: []
    });
  }
  function refresh() {
    var stored = storedResults(), selector = el('report-type');
    if (!selector) return;
    Array.prototype.forEach.call(selector.options, function (option) { option.disabled = !availableType(option.value, stored.engines); });
    if (selector.selectedOptions.length && selector.selectedOptions[0].disabled) {
      var first = Array.prototype.find.call(selector.options, function (option) { return !option.disabled; });
      if (first) selector.value = first.value;
    }
    setText('report-source', stored.filename ? 'Workbook: ' + stored.filename : (stored.engines.comparison ? 'Workbook A/B: ' + (stored.engines.comparison.sourceWorkbooks || []).map(function (x) { return x.filename; }).join(' • ') : 'Belum ada hasil engine tersimpan.'));
    setText('report-status', Object.keys(stored.engines).length ? 'Mencakup hasil analisis yang sudah tersedia. Ekspor selalu memuat seluruh hasil laporan, bukan filter tampilan audit.' : 'Belum tersedia hasil analisis untuk dilaporkan. Unggah workbook dan jalankan audit, atau selesaikan komparasi.');
    updateSummary();
  }
  function updateSummary() {
    var box = el('report-summary');
    if (!box) return;
    box.textContent = '';
    var type = el('report-type').value;
    try {
      state.document = build(type);
      var doc = state.document;
      var lines = [window.ReportingEngine.reportTypeTitle(type), 'Temuan: ' + doc.summary.findings,
        'ERROR ' + doc.summary.severity.ERROR + ' • WARNING ' + doc.summary.severity.WARNING + ' • INFO ' + doc.summary.severity.INFO,
        'Sumber: ' + (doc.sourceWorkbooks.map(function (book) { return book.filename; }).filter(Boolean).join(', ') || 'identitas tidak tersedia')];
      if (doc.summary.workItems !== null) lines.push('Pekerjaan: ' + doc.summary.workItems);
      if (doc.summary.comparison) lines.push('Rekonsiliasi impor lengkap: ' + (doc.summary.comparison.importedTotalsComplete === true ? 'Ya' : doc.summary.comparison.importedTotalsComplete === false ? 'Tidak' : 'Tidak diketahui'));
      box.textContent = lines.join('\n');
      setText('report-status', 'Dokumen laporan siap. Hasil sumber tidak diubah; filter audit tidak membatasi isi ekspor.');
    } catch (error) {
      state.document = null;
      box.textContent = '';
      setText('report-status', error.message);
    }
  }
  function preview() {
    updateSummary();
    if (!state.document) return;
    var popup = window.open('', '_blank');
    if (!popup) { setText('report-status', 'Pratinjau diblokir browser. Izinkan pop-up untuk membuka pratinjau laporan.'); return; }
    popup.document.open(); popup.document.write(window.ReportingEngine.renderPrintHtml(state.document)); popup.document.close();
    setText('report-status', 'Pratinjau dibuka. Untuk membuat PDF, pilih “Cetak / Simpan sebagai PDF” lalu pilih destinasi PDF.');
  }
  function exportXlsx() {
    try {
      updateSummary();
      if (!state.document) return;
      if (!window.XLSX || !window.Blob || !window.URL || !window.URL.createObjectURL) throw new Error('Ekspor XLSX tidak tersedia pada browser ini.');
      var blob = new Blob([window.ReportingEngine.toXlsx(state.document)], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
      var url = window.URL.createObjectURL(blob), link = document.createElement('a');
      link.href = url; link.download = window.ReportingEngine.fileName(state.document, 'xlsx');
      link.style.display = 'none'; document.body.appendChild(link); link.click(); link.remove();
      window.setTimeout(function () { window.URL.revokeObjectURL(url); }, 1000);
      setText('report-status', 'Unduhan XLSX dimulai: ' + link.download);
    } catch (error) { setText('report-status', 'Ekspor XLSX gagal: ' + error.message); }
  }
  function init() {
    if (state.bound) return;
    var selector = el('report-type');
    if (!selector) return;
    state.bound = true;
    selector.addEventListener('change', updateSummary);
    el('btn-report-preview').addEventListener('click', preview);
    el('btn-report-xlsx').addEventListener('click', exportXlsx);
    var nav = el('nav-btn-reporting');
    if (nav) nav.addEventListener('click', refresh);
    window.addEventListener('rabguard:analysis-updated', refresh);
    refresh();
  }
  window.ReportingUI = { init: init, refresh: refresh, buildDocument: build, getState: function () { return { document: state.document }; },
    restoreDocument: function (reportDocument) { state.document = reportDocument || null; if (state.document) { setText('report-status', 'Report snapshot dipulihkan; tidak menjalankan ulang engine.'); setText('report-summary', state.document.summary ? JSON.stringify(state.document.summary, null, 2) : ''); } else refresh(); return { restored: !!state.document }; } };
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init, { once: true }); else init();
})(window, document);
