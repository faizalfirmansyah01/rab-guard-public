/* Thin presentation adapter; all material arithmetic stays in its engine. */
(function (window, document) {
  'use strict';
  var result = null, initialized = false, activeCanonical = null, mappingDecisions = Object.create(null);
  function el(id) { return document.getElementById(id); }
  function text(id, value) { var node = el(id); if (node) node.textContent = value === null || value === undefined ? '—' : String(value); }
  function cell(row, value) { var td = document.createElement('td'); td.textContent = value === null || value === undefined || value === '' ? '—' : String(value); row.appendChild(td); }
  function sourceText(references) { return (references || []).map(function (ref) { return [ref.workbook, ref.sheet, ref.row, ref.cell].filter(Boolean).join(' · '); }).filter(Boolean).join('\n') || '—'; }
  function refresh() {
    var ui = window.UniversalUI, state = ui && ui.getCurrentState ? ui.getCurrentState() : {};
    var canonical = state.currentCanonicalData;
    if (canonical !== activeCanonical) { activeCanonical = canonical || null; mappingDecisions = Object.create(null); result = null; }
    if (state.currentMaterialLogisticsResult !== result) {
      result = state.currentMaterialLogisticsResult || null;
      if (!result) {
        text('material-logistics-coverage', '');
        ['tbody-material-logistics', 'tbody-material-logistics-details', 'material-logistics-findings'].forEach(function (id) { var node = el(id); if (node) node.textContent = ''; });
      }
    }
    var analyses = canonical && Array.isArray(canonical.ahspAnalyses) ? canonical.ahspAnalyses : [];
    var button = el('btn-run-material-logistics');
    if (button) button.disabled = !canonical || !canonical.workItems || !canonical.workItems.length || !analyses.length;
    text('material-logistics-source', canonical && canonical.project ? canonical.project.sourceFilename : 'Belum ada workbook aktif.');
    text('material-logistics-status', !canonical ? 'Unggah workbook dan selesaikan audit terlebih dahulu.' : (!analyses.length ?
      ((canonical.ahspWarnings || []).join(' ') || 'Belum ditemukan tabel AHSP dengan kode, uraian, koefisien, dan satuan yang dapat dinormalisasi.') :
      analyses.length + ' analisis AHSP berhasil dinormalisasi. Jalankan perhitungan untuk meninjau kecocokan dan coverage.'));
    if (result) render(result);
  }
  function render(output) {
    result = output;
    var coverage = output.coverage || {};
    text('material-logistics-status', 'Status: ' + output.status + '. ' + (output.warnings || []).join(' '));
    text('material-logistics-coverage', 'Item BOQ: ' + (coverage.totalWorkItems || 0) +
      ' • volume/satuan valid: ' + (coverage.validVolumeItems || 0) +
      ' • terhubung AHSP: ' + (coverage.linkedWorkItems || 0) +
      ' • berkontribusi: ' + (coverage.calculableWorkItems || 0) +
      ' • tidak ditemukan: ' + (coverage.notFound || 0) +
      ' • ambigu/perlu tinjau: ' + (coverage.ambiguous || 0) +
      ' • satuan tidak kompatibel: ' + (coverage.incompatibleUnit || 0));
    var materials = el('tbody-material-logistics');
    if (materials) {
      materials.textContent = '';
      (output.materials || []).forEach(function (material) {
        var row = document.createElement('tr');
        cell(row, material.resourceCode); cell(row, material.description); cell(row, material.specification);
        cell(row, material.unit); cell(row, material.totalQuantity); cell(row, material.contributingWorkItems);
        materials.appendChild(row);
      });
      if (!output.materials || !output.materials.length) { var empty = document.createElement('tr'); var emptyCell = document.createElement('td'); emptyCell.colSpan = 6; emptyCell.textContent = 'Belum ada material terverifikasi yang dapat dihitung.'; empty.appendChild(emptyCell); materials.appendChild(empty); }
    }
    var details = el('tbody-material-logistics-details');
    if (details) {
      details.textContent = '';
      (output.workItemDetails || []).forEach(function (item) {
        if (!item.resourceDetails || !item.resourceDetails.length) {
          var unmatched = document.createElement('tr'); cell(unmatched, item.code); cell(unmatched, item.description); cell(unmatched, item.matchStatus); cell(unmatched, item.reason); cell(unmatched, '—'); cell(unmatched, '—'); cell(unmatched, sourceText(item.sourceReferences)); details.appendChild(unmatched); return;
        }
        item.resourceDetails.forEach(function (resource) {
          var row = document.createElement('tr'); cell(row, item.code); cell(row, item.description); cell(row, item.matchStatus);
          cell(row, resource.description + (resource.specification ? ' · ' + resource.specification : ''));
          cell(row, resource.coefficient === null ? '—' : resource.coefficient + ' ' + (resource.unit || ''));
          cell(row, resource.quantity === null ? resource.reason : resource.quantity + ' ' + resource.unit);
          cell(row, sourceText((item.sourceReferences || []).concat(resource.sourceReferences || [])));
          details.appendChild(row);
        });
      });
      if (!output.workItemDetails || !output.workItemDetails.length) { var noRows = document.createElement('tr'); var noCell = document.createElement('td'); noCell.colSpan = 7; noCell.textContent = 'Belum ada item pekerjaan untuk dianalisis.'; noRows.appendChild(noCell); details.appendChild(noRows); }
    }
    var findings = el('material-logistics-findings');
    if (findings) { findings.textContent = ''; (output.findings || []).forEach(function (finding) { var li = document.createElement('li'); li.textContent = finding.title + ': ' + finding.explanation; findings.appendChild(li); }); }
    var mappings = el('tbody-material-logistics-mappings');
    if (mappings) {
      mappings.textContent = '';
      (output.mappings || []).forEach(function (mapping) {
        var row = document.createElement('tr'); cell(row, mapping.workCode);
        cell(row, mapping.status + (mapping.userSelected ? ' · dipilih pengguna' : ''));
        var candidateCell = document.createElement('td'), select = document.createElement('select'), label = document.createElement('label');
        label.textContent = 'Pilih kandidat AHSP untuk ' + (mapping.workCode || mapping.workItemId); select.setAttribute('aria-label', label.textContent);
        var blank = document.createElement('option'); blank.value = ''; blank.textContent = 'Lihat kandidat'; select.appendChild(blank);
        (mapping.candidates || []).forEach(function (candidate) {
          var option = document.createElement('option'); option.value = candidate.analysisId;
          var source = candidate.sourceReferences && candidate.sourceReferences[0];
          option.textContent = (candidate.workCode || 'Tanpa kode') + ' · ' + (candidate.workDescription || 'Uraian tidak tersedia') +
            (source ? ' · ' + [source.sheet, source.row, source.cell].filter(Boolean).join('!') : '');
          if (mapping.userSelectedAnalysisId === candidate.analysisId) option.selected = true;
          select.appendChild(option);
        });
        candidateCell.appendChild(select); row.appendChild(candidateCell); cell(row, mapping.reason || '—');
        var action = document.createElement('td');
        if ((mapping.candidates || []).length) {
          var button = document.createElement('button'); button.type = 'button'; button.textContent = 'Gunakan pilihan';
          button.setAttribute('data-apply-ahsp', mapping.workItemId || ''); button.disabled = !mappingDecisions[mapping.workItemId]; action.appendChild(button);
          select.setAttribute('data-ahsp-select', mapping.workItemId || '');
        } else action.textContent = '—';
        row.appendChild(action); mappings.appendChild(row);
      });
      if (!output.mappings || !output.mappings.length) { var emptyMapping = document.createElement('tr'); var emptyMappingCell = document.createElement('td'); emptyMappingCell.colSpan = 5; emptyMappingCell.textContent = 'Tidak ada pemetaan pekerjaan.'; emptyMapping.appendChild(emptyMappingCell); mappings.appendChild(emptyMapping); }
    }
    if (window.UniversalUI && window.UniversalUI.setMaterialLogisticsResult) window.UniversalUI.setMaterialLogisticsResult(output);
  }
  function analyze() {
    var ui = window.UniversalUI, canonical = ui && ui.getCurrentState && ui.getCurrentState().currentCanonicalData;
    if (!canonical) { text('material-logistics-status', 'Data canonical belum tersedia. Selesaikan audit workbook terlebih dahulu.'); return null; }
    if (!canonical.ahspAnalyses || !canonical.ahspAnalyses.length) { text('material-logistics-status', 'Analisis tidak dijalankan: belum ada AHSP canonical yang dapat digunakan.'); return null; }
    try {
      var next = window.AnalysisEngineAdapter.run(window.MaterialLogisticsEngine, canonical, { mappingDecisions: mappingDecisions });
      render(next);
      return next;
    } catch (error) { text('material-logistics-status', 'Analisis kebutuhan material gagal: ' + error.message); return null; }
  }
  function init() {
    if (initialized) return; initialized = true;
    var button = el('btn-run-material-logistics'); if (button) button.addEventListener('click', analyze);
    var mappingBody = el('tbody-material-logistics-mappings');
    if (mappingBody) {
      mappingBody.addEventListener('change', function (event) {
        var select = event.target && event.target.getAttribute('data-ahsp-select');
        if (!select) return;
        if (event.target.value) mappingDecisions[select] = event.target.value; else delete mappingDecisions[select];
        result = null;
        if (window.UniversalUI && window.UniversalUI.setMaterialLogisticsResult) window.UniversalUI.setMaterialLogisticsResult(null);
        text('material-logistics-status', 'Keputusan pemetaan berubah. Hasil sebelumnya tidak lagi current; terapkan pilihan dan hitung ulang.');
        if (window.UniversalUI && window.UniversalUI.persistProjectState) window.UniversalUI.persistProjectState();
        var buttons = mappingBody.querySelectorAll('[data-apply-ahsp]');
        Array.prototype.forEach.call(buttons, function (action) { if (action.getAttribute('data-apply-ahsp') === select) action.disabled = !event.target.value; });
      });
      mappingBody.addEventListener('click', function (event) {
        var itemId = event.target && event.target.getAttribute('data-apply-ahsp');
        if (!itemId || !mappingDecisions[itemId]) return;
        analyze();
      });
    }
    refresh();
  }
  function clear() { result = null; mappingDecisions = Object.create(null); activeCanonical = null; text('material-logistics-coverage', ''); ['tbody-material-logistics', 'tbody-material-logistics-details', 'tbody-material-logistics-mappings', 'material-logistics-findings'].forEach(function (id) { var node = el(id); if (node) node.textContent = ''; }); refresh(); }
  window.MaterialLogisticsUI = { init: init, refresh: refresh, analyze: analyze, clear: clear, getResult: function () { return result; }, render: render,
    getState: function () { return { result: result, mappingDecisions: Object.assign({}, mappingDecisions) }; },
    restoreState: function (snapshot, canonical) { snapshot = snapshot || {}; activeCanonical = canonical || null; mappingDecisions = Object.assign(Object.create(null), snapshot.mappingDecisions || {}); result = snapshot.result || null; if (result) render(result); else refresh(); return { restored: !!result, decisions: Object.keys(mappingDecisions).length }; } };
  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init, { once: true }); else init();
})(window, document);
