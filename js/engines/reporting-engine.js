/* Snapshot builder and local PDF/Excel export adapters for RAB Guard reports. */
(function (window) {
  'use strict';

  var VERSION = '1.0.0';
  var SCHEMA = 'rab-guard-report/1';
  var idSequence = 0;
  var ALLOWED_TYPES = { audit: true, unit: true, comparison: true, combined: true };

  function clone(value, seen) {
    if (value === null || typeof value !== 'object') return value;
    if (value instanceof Date) return new Date(value.getTime());
    seen = seen || new WeakMap();
    if (seen.has(value)) return seen.get(value);
    var copy = Array.isArray(value) ? [] : {};
    seen.set(value, copy);
    Object.keys(value).forEach(function (key) { copy[key] = clone(value[key], seen); });
    return copy;
  }
  function freeze(value, seen) {
    if (!value || typeof value !== 'object' || Object.isFrozen(value)) return value;
    seen = seen || new WeakSet();
    if (seen.has(value)) return value;
    seen.add(value);
    Object.keys(value).forEach(function (key) { freeze(value[key], seen); });
    return Object.freeze(value);
  }
  function own(object, key) { return Object.prototype.hasOwnProperty.call(object || {}, key); }
  function nonEmpty(value) { return value !== null && value !== undefined && String(value).trim() !== ''; }
  function stableRefKey(ref) {
    return [ref.workbook || ref.filename || '', ref.sheet || '', ref.row || '', ref.cell || ref.cellRef || '', ref.field || '', ref.itemId || ''].join('|');
  }
  function sourceReferences(entry) {
    var refs = [];
    if (entry.result && Array.isArray(entry.result.sourceReferences)) refs = refs.concat(entry.result.sourceReferences);
    (entry.result && Array.isArray(entry.result.findings) ? entry.result.findings : []).forEach(function (finding) {
      if (Array.isArray(finding.sourceReferences)) refs = refs.concat(finding.sourceReferences);
      // UniversalChecker predates the common engine-result contract and stores
      // its primary provenance in sourceSheet/sourceRow/sourceCell fields.
      if (finding.sourceSheet || finding.sourceRow || finding.sourceCell) {
        refs.push({
          workbook: entry.sourceFilename || null,
          sheet: finding.sourceSheet || null,
          row: Number.isFinite(finding.sourceRow) && finding.sourceRow > 0 ? finding.sourceRow : null,
          cell: finding.sourceCell && !/^baris\s+\d+$/i.test(String(finding.sourceCell)) ? finding.sourceCell : null,
          code: finding.itemCode || null,
          description: finding.itemDescription || null
        });
      }
      if (Array.isArray(finding.relatedItems)) finding.relatedItems.forEach(function (item) {
        if (item.source) refs.push(item.source);
        else if (item.sheet || item.row || item.sourceCell) refs.push({ sheet: item.sheet, row: item.row, cell: item.sourceCell, code: item.code, description: item.description, categoryPath: item.categoryPath });
      });
    });
    if (entry.result && entry.result.changes) entry.result.changes.forEach(function (change) { if (change.sourceA) refs.push(change.sourceA); if (change.sourceB) refs.push(change.sourceB); });
    if (entry.result && entry.result.candidates) entry.result.candidates.forEach(function (candidate) { if (candidate.sourceA) refs.push(candidate.sourceA); if (candidate.sourceB) refs.push(candidate.sourceB); });
    if (entry.result && entry.result.relationships) entry.result.relationships.forEach(function (relation) { refs = refs.concat(relation.itemsA || [], relation.itemsB || []); });
    return refs.filter(Boolean);
  }
  function inferWorkbooks(entry) {
    if (Array.isArray(entry.sourceWorkbooks) && entry.sourceWorkbooks.length) return entry.sourceWorkbooks;
    var summary = entry.result && entry.result.summary || {};
    if (summary.workbookA || summary.workbookB) return [summary.workbookA, summary.workbookB].filter(Boolean).map(function (book) {
      return { filename: book.filename || null, role: book === summary.workbookA ? 'A' : 'B' };
    });
    var names = Object.create(null);
    sourceReferences(entry).forEach(function (ref) { var name = ref.workbook || ref.filename; if (name) names[name] = true; });
    return Object.keys(names).map(function (filename) { return { filename: filename }; });
  }
  function selectedEntries(reportType, entries) {
    var map = { audit: ['audit'], unit: ['unit'], comparison: ['comparison'], combined: ['audit', 'unit', 'comparison'] };
    return map[reportType].map(function (key) {
      var entry = entries[key];
      if (entry) return Object.assign({ key: key }, entry);
      return reportType === 'combined' ? { key: key, engineName: ({ audit: 'UniversalChecker', unit: 'UnitConsistencyEngine', comparison: 'RABComparisonEngine' })[key], status: 'not_run', result: null, sourceWorkbooks: [] } : null;
    }).filter(Boolean);
  }
  function uniqueReferences(entries) {
    var index = Object.create(null), output = [];
    entries.forEach(function (entry) { sourceReferences(entry).forEach(function (ref) {
      var copy = clone(ref), key = stableRefKey(copy);
      if (!index[key]) { index[key] = true; output.push(copy); }
    }); });
    return output;
  }
  function normalizeFinding(finding, entry) {
    var copy = clone(finding || {});
    copy.originEngine = entry.result && (entry.result.engineName || entry.engineName) || entry.engineName || 'UnknownEngine';
    copy.sourceReferences = Array.isArray(copy.sourceReferences) ? copy.sourceReferences : [];
    if (!copy.sourceReferences.length && (copy.sourceSheet || copy.sourceRow || copy.sourceCell)) {
      copy.sourceReferences.push({
        workbook: entry.sourceFilename || null,
        sheet: copy.sourceSheet || null,
        row: Number.isFinite(copy.sourceRow) && copy.sourceRow > 0 ? copy.sourceRow : null,
        cell: copy.sourceCell && !/^baris\s+\d+$/i.test(String(copy.sourceCell)) ? copy.sourceCell : null,
        code: copy.itemCode || null,
        description: copy.itemDescription || null
      });
    }
    return copy;
  }
  function engineSummaries(entries) {
    return entries.map(function (entry) {
      var result = entry.result || {};
      return {
        engineName: result.engineName || entry.engineName || 'UnknownEngine',
        engineVersion: result.engineVersion || null,
        status: result.status || entry.status || 'not_run',
        summary: clone(result.summary || {}),
        findingsCount: Array.isArray(result.findings) ? result.findings.length : null,
        warnings: clone(result.warnings || []),
        errors: clone(result.errors || [])
      };
    });
  }
  function collectWarnings(entries, input) {
    var warnings = clone(input.warnings || []);
    entries.forEach(function (entry) {
      var result = entry.result || {};
      (result.warnings || []).forEach(function (message) { warnings.push((result.engineName || entry.engineName) + ': ' + message); });
      (result.errors || []).forEach(function (message) { warnings.push((result.engineName || entry.engineName) + ' gagal: ' + message); });
      if (entry.status === 'excluded') warnings.push((entry.engineName || entry.key) + ' tersedia tetapi tidak disertakan karena konteks sumber tidak cocok.');
      else if (!entry.result || result.status === 'not_run') warnings.push((entry.engineName || entry.key) + ' tidak dijalankan; laporan hanya mencakup hasil engine yang tersedia.');
    });
    return Array.from(new Set(warnings));
  }
  function validateDocument(document) {
    var errors = [], warnings = [];
    if (!document || document.schemaVersion !== SCHEMA) errors.push('Versi kontrak Report Document tidak valid.');
    if (!document || !document.reportId || !ALLOWED_TYPES[document.reportType] || !document.generatedAt) errors.push('Identitas dasar laporan tidak lengkap.');
    if (!document || !Array.isArray(document.analysisEngines) || !Array.isArray(document.sections) || !Array.isArray(document.findings) || !Array.isArray(document.sourceReferences)) errors.push('Koleksi inti laporan tidak valid.');
    if (document && document.findings) document.findings.forEach(function (finding, index) {
      if (!finding.originEngine) errors.push('Temuan ' + index + ' tidak memiliki identitas engine asal.');
      if (!Array.isArray(finding.sourceReferences) || !finding.sourceReferences.length) warnings.push('Temuan ' + (finding.id || index + 1) + ' tidak memiliki provenance sel yang tersedia.');
    });
    if (document && !(document.sourceWorkbooks || []).length) warnings.push('Identitas workbook tidak tersedia dalam metadata laporan.');
    return { valid: errors.length === 0, errors: errors, warnings: warnings };
  }
  function severityCounts(findings) {
    var counts = { ERROR: 0, WARNING: 0, INFO: 0, OTHER: 0 };
    findings.forEach(function (finding) { var key = String(finding.severity || '').toUpperCase(); if (counts[key] === undefined) counts.OTHER++; else counts[key]++; });
    return counts;
  }
  function reportTypeTitle(type) {
    return ({ audit: 'Laporan Audit RAB', unit: 'Laporan Pengecekan Satuan', comparison: 'Laporan Komparasi RAB (CCO)', combined: 'Laporan Gabungan RAB Guard' })[type] || 'Laporan RAB Guard';
  }

  function createDocument(input) {
    input = input || {};
    var type = input.reportType;
    if (!ALLOWED_TYPES[type]) throw new Error('Jenis laporan tidak didukung.');
    var entries = selectedEntries(type, input.engines || {});
    if (type !== 'combined' && (!entries.length || !entries[0].result)) throw new Error('Hasil engine untuk jenis laporan ini belum tersedia.');
    var resultEntries = entries.filter(function (entry) { return entry.result && entry.result.status !== 'not_run'; });
    if (!resultEntries.length) throw new Error('Belum ada hasil analisis yang dapat dimasukkan ke laporan.');
    if (type === 'combined') {
      var contexts = resultEntries.map(function (entry) { return entry.contextId || null; });
      var knownContexts = contexts.filter(Boolean);
      if (knownContexts.length !== contexts.length || new Set(knownContexts).size > 1) {
        throw new Error('Hasil engine gabungan tidak memiliki konteks proyek yang sama dan terverifikasi.');
      }
    }

    var findings = [];
    resultEntries.forEach(function (entry) {
      (entry.result.findings || []).forEach(function (finding) { findings.push(normalizeFinding(finding, entry)); });
    });
    var comparisonEntry = resultEntries.filter(function (entry) { return entry.key === 'comparison'; })[0] || null;
    var stats = input.canonicalSummary || {};
    var sourceWorkbooks = [];
    resultEntries.forEach(function (entry) { sourceWorkbooks = sourceWorkbooks.concat(inferWorkbooks(entry)); });
    if (Array.isArray(input.sourceWorkbooks)) sourceWorkbooks = sourceWorkbooks.concat(input.sourceWorkbooks);
    var workbookSeen = Object.create(null);
    sourceWorkbooks = sourceWorkbooks.filter(function (book) {
      if (!book || !book.filename) return false;
      var key = [book.filename, book.sheet || '', book.role || ''].join('|');
      if (workbookSeen[key]) return false;
      workbookSeen[key] = true; return true;
    }).map(function (book) { return clone(book); });
    var allLimits = clone(input.limitations || []);
    resultEntries.forEach(function (entry) { (entry.result.limitations || []).forEach(function (value) { allLimits.push({ engine: entry.result.engineName || entry.engineName, text: value }); }); });
    var sections = [];
    resultEntries.forEach(function (entry) {
      var name = entry.result.engineName || entry.engineName;
      sections.push({ id: entry.key, title: reportTypeTitle(entry.key), engineName: name,
        status: entry.result.status || entry.status || 'complete', findingCount: (entry.result.findings || []).length,
        detailCount: entry.key === 'comparison' ? ((entry.result.changes || []).length + (entry.result.candidates || []).length + (entry.result.relationships || []).length) : null });
    });
    var comparison = comparisonEntry ? {
      summary: clone(comparisonEntry.result.summary || {}),
      changes: clone(comparisonEntry.result.changes || []),
      candidates: clone(comparisonEntry.result.candidates || []),
      relationships: clone(comparisonEntry.result.relationships || [])
    } : null;
    var unitEntry = resultEntries.filter(function (entry) { return entry.key === 'unit'; })[0];
    var auditEntry = resultEntries.filter(function (entry) { return entry.key === 'audit'; })[0];
    var document = {
      schemaVersion: SCHEMA,
      reportId: input.reportId || ('RGR-' + Date.now().toString(36) + '-' + (++idSequence).toString(36)),
      reportType: type,
      reportVersion: VERSION,
      generatedAt: input.generatedAt || new Date().toISOString(),
      applicationVersion: input.applicationVersion || null,
      projectMetadata: clone(input.projectMetadata || null),
      sourceWorkbooks: sourceWorkbooks,
      analysisEngines: engineSummaries(entries),
      summary: {
        workItems: Number.isFinite(stats.totalWorkItems) ? stats.totalWorkItems : (Number.isFinite(stats.workItems) ? stats.workItems : null),
        findings: findings.length,
        severity: severityCounts(findings),
        audit: auditEntry ? clone(auditEntry.result.summary || {}) : null,
        unit: unitEntry ? clone(unitEntry.result.summary || {}) : null,
        comparison: comparison ? clone(comparison.summary) : null
      },
      sections: sections,
      findings: findings,
      comparisonDetails: comparison,
      sourceReferences: uniqueReferences(resultEntries),
      assumptions: clone(input.assumptions || []),
      limitations: allLimits,
      warnings: collectWarnings(entries, input),
      validationStatus: null
    };
    var validation = validateDocument(document);
    document.validationStatus = clone(validation);
    if (!validation.valid) throw new Error('Report Document tidak valid: ' + validation.errors.join(' '));
    return freeze(document);
  }

  function safeString(value) {
    if (value === null || value === undefined) return '';
    if (typeof value === 'object') return JSON.stringify(value);
    return String(value);
  }
  function numberOrBlank(value) { return typeof value === 'number' && Number.isFinite(value) ? value : ''; }
  function cellRef(ref) { return ref.cell || ref.cellRef || (ref.cells && Object.keys(ref.cells).map(function (key) { return key + ':' + (ref.cells[key].cellRef || ''); }).join('; ')) || ''; }
  function workbookName(ref) { return ref.workbook || ref.filename || ''; }
  function findingRows(document) {
    var rows = [['Finding ID', 'Engine Asal', 'Rule ID', 'Severity', 'Interpretasi', 'Judul', 'Penjelasan', 'Rekomendasi', 'Workbook', 'Sheet', 'Baris', 'Sel', 'Item/Kode', 'Deskripsi', 'Category Path', 'Nilai Aktual', 'Nilai Pembanding', 'Selisih']];
    document.findings.forEach(function (finding) {
      var refs = finding.sourceReferences || [], ref = refs[0] || {};
      rows.push([finding.id || '', finding.originEngine || '', finding.ruleId || '', finding.severity || '', finding.interpretation || '',
        finding.title || finding.message || '', finding.explanation || finding.reason || '', finding.recommendedAction || '',
        workbookName(ref), ref.sheet || '', ref.row || '', cellRef(ref), finding.itemCode || ref.code || '', finding.description || ref.description || '',
        finding.categoryPath || ref.categoryPath || '', own(finding, 'actualValue') ? finding.actualValue : (finding.evidence && finding.evidence.actualValue),
        own(finding, 'comparisonValue') ? finding.comparisonValue : (finding.evidence && finding.evidence.comparisonValue),
        own(finding, 'difference') ? finding.difference : (finding.evidence && finding.evidence.difference)]);
    });
    return rows;
  }
  function unitRows(document) {
    var rows = [['Rule ID', 'Severity', 'Interpretasi', 'Temuan', 'Alasan', 'Item', 'Kode', 'Deskripsi', 'Satuan', 'Sheet', 'Baris', 'Sel', 'Category Path', 'Tindakan']];
    document.findings.filter(function (finding) { return /UnitConsistency/i.test(finding.originEngine || ''); }).forEach(function (finding) {
      var related = finding.relatedItems && finding.relatedItems.length ? finding.relatedItems : [null];
      related.forEach(function (item) {
        var ref = item && (item.source || item.sourceReference) || finding.sourceReferences && finding.sourceReferences[0] || {};
        rows.push([finding.ruleId || '', finding.severity || '', finding.interpretation || '', finding.title || '', finding.explanation || '',
          item && item.itemId || '', item && item.code || '', item && item.description || '', item && item.unit || '', ref.sheet || item && item.sheet || '',
          ref.row || item && item.row || '', cellRef(ref) || (item && item.sourceCell) || '', item && item.categoryPath || '', finding.recommendedAction || '']);
      });
    });
    return rows;
  }
  function comparisonRows(document) {
    var rows = [['Status', 'Perubahan', 'Kode A', 'Kode B', 'Deskripsi A', 'Deskripsi B', 'Satuan A', 'Satuan B', 'Volume A', 'Volume B', 'Delta Volume', 'Delta Volume %', 'Harga Satuan A', 'Harga Satuan B', 'Delta Harga Satuan', 'Jumlah A (Impor)', 'Jumlah B (Impor)', 'Delta Jumlah', 'Delta Jumlah %', 'Hierarchy A', 'Hierarchy B', 'Workbook A', 'Sheet A', 'Baris A', 'Sel A', 'Workbook B', 'Sheet B', 'Baris B', 'Sel B', 'Alasan/Kandidat']];
    var details = document.comparisonDetails;
    if (!details) return rows;
    (details.changes || []).forEach(function (change) {
      var c = change.changes || {}, a = change.sourceA || {}, b = change.sourceB || {};
      var volume = c.volume || {}, price = c.unitPrice || {}, amount = c.amount || {};
      rows.push([change.status || '', (change.changeTypes || []).join(', '), c.code && c.code.before || a.code || '', c.code && c.code.after || b.code || '',
        c.description && c.description.before || a.description || '', c.description && c.description.after || b.description || '',
        c.unit && c.unit.before || a.unit || '', c.unit && c.unit.after || b.unit || '', numberOrBlank(volume.before), numberOrBlank(volume.after), numberOrBlank(volume.delta), numberOrBlank(volume.percentageChange),
        numberOrBlank(price.before), numberOrBlank(price.after), numberOrBlank(price.delta), numberOrBlank(amount.before), numberOrBlank(amount.after), numberOrBlank(amount.delta), numberOrBlank(amount.percentageChange),
        pathText(a.categoryPath), pathText(b.categoryPath), workbookName(a), a.sheet || '', a.row || '', cellRef(a), workbookName(b), b.sheet || '', b.row || '', cellRef(b), safeString(change.matchEvidence || '')]);
    });
    (details.candidates || []).forEach(function (candidate) {
      var a = candidate.sourceA || {}, b = candidate.sourceB || {};
      rows.push([candidate.status || 'AMBIGUOUS', 'Kandidat perlu verifikasi', a.code || '', b.code || '', a.description || '', b.description || '', a.unit || '', b.unit || '', '', '', '', '', '', '', '', numberOrBlank(a.originalValues && a.originalValues.amount), numberOrBlank(b.originalValues && b.originalValues.amount), '', '', pathText(a.categoryPath), pathText(b.categoryPath), workbookName(a), a.sheet || '', a.row || '', cellRef(a), workbookName(b), b.sheet || '', b.row || '', cellRef(b), (candidate.reasons || []).join(' ')]);
    });
    (details.relationships || []).forEach(function (group) {
      rows.push([group.status || 'AMBIGUOUS', 'Kelompok split/merge atau kode berulang', group.code || '', '',
        safeString(group.itemsA || group.itemIndicesA || ''), safeString(group.itemsB || group.itemIndicesB || ''), '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', group.reason || '']);
    });
    return rows;
  }
  function pathText(path) {
    if (!Array.isArray(path)) return safeString(path);
    return path.slice().sort(function (a, b) { return Number(a.levelIndex || 0) - Number(b.levelIndex || 0); }).map(function (level) { return (Number(level.levelIndex) + 1) + ':' + (level.value === null ? '[kosong]' : safeString(level.value)); }).join(' > ');
  }
  function sourceRows(document) {
    var rows = [['Workbook', 'Sheet', 'Baris', 'Sel/Range', 'Field', 'Item ID', 'Kode', 'Deskripsi', 'Category Path', 'Nilai Asli', 'Formula', 'Cell Status', 'Engine Ref']];
    document.sourceReferences.forEach(function (ref) {
      var cells = ref.cells || {};
      var entries = Object.keys(cells);
      if (!entries.length) entries = [''];
      entries.forEach(function (field) {
        var cell = cells[field] || {};
        rows.push([workbookName(ref), ref.sheet || '', ref.row || '', cell.cellRef || cell.cell || ref.cell || ref.cellRef || '', field, ref.itemId || ref.id || '', ref.code || '', ref.description || '', ref.categoryPath || '', own(cell, 'rawValue') ? cell.rawValue : (own(cell, 'val') ? cell.val : ref.originalValue), cell.formula || ref.formula || '', cell.valueStatus || ref.valueStatus || '', ref.originEngine || '']);
      });
    });
    return rows;
  }
  function assumptionsRows(document) {
    var rows = [['Jenis', 'Sumber', 'Keterangan']];
    document.assumptions.forEach(function (value) { rows.push(['Asumsi', '', safeString(value)]); });
    document.limitations.forEach(function (value) { rows.push(['Keterbatasan', value.engine || '', value.text || safeString(value)]); });
    document.warnings.forEach(function (value) { rows.push(['Peringatan', '', safeString(value)]); });
    if (rows.length === 1) rows.push(['Informasi', '', 'Tidak ada asumsi atau keterbatasan tambahan yang dicatat.']);
    return rows;
  }
  function makeSheet(rows) {
    var sheet = window.XLSX.utils.aoa_to_sheet(rows, { cellDates: false });
    var width = Math.max.apply(Math, rows.map(function (row) { return row.length; }));
    sheet['!cols'] = Array.from({ length: width }, function (_, col) {
      var size = Math.min(48, Math.max(12, rows.slice(0, 20).reduce(function (best, row) { return Math.max(best, safeString(row[col]).length); }, 0) + 2));
      return { wch: size };
    });
    if (rows.length > 1) sheet['!autofilter'] = { ref: 'A1:' + window.XLSX.utils.encode_col(width - 1) + rows.length };
    Object.keys(sheet).forEach(function (key) {
      if (key.charAt(0) === '!') return;
      var cell = sheet[key];
      if (typeof cell.v === 'string') { cell.t = 's'; delete cell.f; delete cell.F; }
    });
    return sheet;
  }
  function buildExcelWorkbook(document) {
    if (!window.XLSX || !window.XLSX.utils || typeof window.XLSX.write !== 'function') throw new Error('SheetJS lokal belum mendukung ekspor workbook.');
    var wb = window.XLSX.utils.book_new();
    var summaryRows = [['RAB Guard - ' + reportTypeTitle(document.reportType)], ['Report ID', document.reportId], ['Versi dokumen', document.reportVersion], ['Dibuat', document.generatedAt],
      ['Status validasi', document.validationStatus.valid ? 'Valid' : 'Tidak valid'], ['Proyek', document.projectMetadata && document.projectMetadata.projectName || 'Tidak tersedia'],
      ['Temuan', document.summary.findings], ['ERROR', document.summary.severity.ERROR], ['WARNING', document.summary.severity.WARNING], ['INFO', document.summary.severity.INFO],
      ['Jumlah pekerjaan', numberOrBlank(document.summary.workItems)]];
    document.sourceWorkbooks.forEach(function (book, index) { summaryRows.push(['Workbook Sumber ' + (index + 1), book.filename]); if (book.sheet) summaryRows.push(['Sheet sumber ' + (index + 1), book.sheet]); });
    if (document.summary.unit) Object.keys(document.summary.unit).forEach(function (key) { summaryRows.push(['Unit - ' + key, document.summary.unit[key]]); });
    if (document.summary.comparison) Object.keys(document.summary.comparison).forEach(function (key) { summaryRows.push(['Komparasi - ' + key, document.summary.comparison[key]]); });
    window.XLSX.utils.book_append_sheet(wb, makeSheet(summaryRows), 'Ringkasan');
    if (document.analysisEngines.some(function (engine) { return engine.engineName === 'UniversalChecker'; })) window.XLSX.utils.book_append_sheet(wb, makeSheet(findingRows(document)), 'Temuan Audit');
    if (document.analysisEngines.some(function (engine) { return engine.engineName === 'UnitConsistencyEngine'; })) window.XLSX.utils.book_append_sheet(wb, makeSheet(unitRows(document)), 'Pengecekan Satuan');
    if (document.comparisonDetails) window.XLSX.utils.book_append_sheet(wb, makeSheet(comparisonRows(document)), 'Detail Komparasi');
    if (document.sourceReferences.length) window.XLSX.utils.book_append_sheet(wb, makeSheet(sourceRows(document)), 'Referensi Sumber');
    if (document.assumptions.length || document.limitations.length || document.warnings.length) window.XLSX.utils.book_append_sheet(wb, makeSheet(assumptionsRows(document)), 'Asumsi & Batasan');
    wb.Props = { Title: reportTypeTitle(document.reportType), Subject: 'RAB Guard Report Document ' + document.reportId, Author: 'RAB Guard' };
    return wb;
  }
  function xlsxBuffer(document) {
    var output = window.XLSX.write(buildExcelWorkbook(document), { bookType: 'xlsx', type: 'array', compression: true });
    return output instanceof Uint8Array ? output : new Uint8Array(output);
  }

  function escapeHtml(value) {
    return safeString(value).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }
  function displayValue(value) { return value === null || value === undefined || value === '' ? 'Tidak tersedia' : safeString(value); }
  function sourceDisplay(ref) {
    var loc = [workbookName(ref), ref.sheet || '', ref.row ? 'baris ' + ref.row : '', cellRef(ref)].filter(Boolean).join(' • ');
    return loc || 'Referensi sumber tidak tersedia';
  }
  function printFindingRows(document) {
    if (!document.findings.length) return '<tr><td colspan="5">Tidak ada temuan untuk engine yang tercantum.</td></tr>';
    return document.findings.map(function (finding) {
      var refs = finding.sourceReferences || [];
      return '<tr><td>' + escapeHtml(finding.severity || 'OTHER') + '</td><td><strong>' + escapeHtml(finding.title || finding.message || finding.ruleId || 'Temuan') + '</strong><br><small>' + escapeHtml(finding.originEngine + ' / ' + (finding.ruleId || '')) + '</small></td><td>' + escapeHtml(finding.explanation || finding.reason || 'Penjelasan tidak tersedia.') + '</td><td>' + escapeHtml(refs.map(sourceDisplay).join('\n') || 'Referensi tidak tersedia') + '</td><td>' + escapeHtml(finding.recommendedAction || 'Tidak tersedia') + '</td></tr>';
    }).join('');
  }
  function printComparisonRows(details) {
    if (!details) return '';
    var rows = [];
    (details.changes || []).forEach(function (change) {
      var c = change.changes || {}, a = change.sourceA || {}, b = change.sourceB || {}, amt = c.amount || {}, vol = c.volume || {}, price = c.unitPrice || {};
      rows.push('<tr><td>' + escapeHtml(change.status || '') + '</td><td>' + escapeHtml((change.changeTypes || []).join(', ')) + '<br>' + escapeHtml((c.description && c.description.before) || a.description || '') + ' → ' + escapeHtml((c.description && c.description.after) || b.description || '') + '</td><td>' + escapeHtml(sourceDisplay(a)) + '<br>Kode/satuan: ' + escapeHtml(a.code || '') + ' / ' + escapeHtml(a.unit || '') + '<br>Volume: ' + escapeHtml(displayValue(vol.before)) + '<br>Harga: ' + escapeHtml(displayValue(price.before)) + '<br>Jumlah impor: ' + escapeHtml(displayValue(amt.before)) + '<br>Hierarki: ' + escapeHtml(pathText(a.categoryPath)) + '</td><td>' + escapeHtml(sourceDisplay(b)) + '<br>Kode/satuan: ' + escapeHtml(b.code || '') + ' / ' + escapeHtml(b.unit || '') + '<br>Volume: ' + escapeHtml(displayValue(vol.after)) + '<br>Harga: ' + escapeHtml(displayValue(price.after)) + '<br>Jumlah impor: ' + escapeHtml(displayValue(amt.after)) + '<br>Hierarki: ' + escapeHtml(pathText(b.categoryPath)) + '</td><td>Delta nilai: ' + escapeHtml(displayValue(amt.delta)) + (amt.percentageChange === null ? ' (N/A)' : ' (' + escapeHtml(amt.percentageChange) + '%)') + '<br>Delta volume: ' + escapeHtml(displayValue(vol.delta)) + (vol.percentageChange === null ? ' (N/A)' : ' (' + escapeHtml(vol.percentageChange) + '%)') + '</td></tr>');
    });
    (details.candidates || []).forEach(function (candidate) { rows.push('<tr><td>' + escapeHtml(candidate.status || 'AMBIGUOUS') + '</td><td>Kandidat perlu verifikasi<br>' + escapeHtml((candidate.reasons || []).join(' ')) + '</td><td>' + escapeHtml(candidate.sourceA ? sourceDisplay(candidate.sourceA) : 'Tidak tersedia') + '<br>' + escapeHtml(candidate.sourceA && candidate.sourceA.description || '') + '</td><td>' + escapeHtml(candidate.sourceB ? sourceDisplay(candidate.sourceB) : 'Tidak tersedia') + '<br>' + escapeHtml(candidate.sourceB && candidate.sourceB.description || '') + '</td><td>Skor heuristik: ' + escapeHtml(displayValue(candidate.score)) + '</td></tr>'); });
    (details.relationships || []).forEach(function (group) { rows.push('<tr><td>' + escapeHtml(group.status || 'AMBIGUOUS') + '</td><td>' + escapeHtml(group.reason || 'Relasi perlu verifikasi') + '</td><td>' + escapeHtml((group.itemsA || []).map(sourceDisplay).join('\n')) + '</td><td>' + escapeHtml((group.itemsB || []).map(sourceDisplay).join('\n')) + '</td><td>Kode: ' + escapeHtml(group.code || '') + '</td></tr>'); });
    return rows.join('') || '<tr><td colspan="5">Tidak ada detail perubahan.</td></tr>';
  }
  function renderPrintHtml(document) {
    var isWide = document.reportType === 'comparison' || !!document.comparisonDetails;
    var title = reportTypeTitle(document.reportType);
    var workbookList = (document.sourceWorkbooks || []).map(function (book) { return '<li>' + escapeHtml(book.filename || 'Tidak tersedia') + (book.role ? ' (' + escapeHtml(book.role) + ')' : '') + (book.sheet ? ' — ' + escapeHtml(book.sheet) : '') + '</li>'; }).join('') || '<li>Identitas workbook tidak tersedia.</li>';
    var engines = document.analysisEngines.map(function (engine) { return '<li><strong>' + escapeHtml(engine.engineName) + '</strong>: ' + escapeHtml(engine.status) + (engine.findingsCount === null ? '' : ', ' + engine.findingsCount + ' temuan') + '</li>'; }).join('');
    var assumptions = (document.assumptions || []).map(function (value) { return '<li>' + escapeHtml(value) + '</li>'; }).join('') || '<li>Tidak ada asumsi tambahan yang dicatat.</li>';
    var limitations = (document.limitations || []).map(function (value) { return '<li>' + escapeHtml(value.text || value) + (value.engine ? ' <small>(' + escapeHtml(value.engine) + ')</small>' : '') + '</li>'; }).join('') || '<li>Tidak ada keterbatasan tambahan yang dicatat.</li>';
    var warnings = (document.warnings || []).map(function (value) { return '<li>' + escapeHtml(value) + '</li>'; }).join('');
    var summary = document.summary;
    var summaryRows = [['Jumlah pekerjaan', summary.workItems], ['Jumlah temuan', summary.findings], ['ERROR', summary.severity.ERROR], ['WARNING', summary.severity.WARNING], ['INFO', summary.severity.INFO]];
    if (summary.comparison) {
      summaryRows.push(['Item cocok', summary.comparison.matchedItems], ['Item baru', summary.comparison.addedItems], ['Item dihapus', summary.comparison.removedItems], ['Nilai impor B − A', summary.comparison.importedNetChange], ['Belum direkonsiliasi', summary.comparison.unreconciledImportedAmountChange]);
    }
    if (summary.unit) summaryRows.push(['Item dengan unit dikenal', summary.unit.recognizedUnits], ['Indikasi konflik dimensi', summary.unit.conflictGroups], ['Varian unit perlu tinjau', summary.unit.variantGroups]);
    var summaryHtml = summaryRows.map(function (row) { return '<tr><th>' + escapeHtml(row[0]) + '</th><td>' + escapeHtml(displayValue(row[1])) + '</td></tr>'; }).join('');
    var refRows = document.sourceReferences.map(function (ref) { return '<tr><td>' + escapeHtml(workbookName(ref)) + '</td><td>' + escapeHtml(ref.sheet || '') + '</td><td>' + escapeHtml(ref.row || '') + '</td><td>' + escapeHtml(cellRef(ref)) + '</td><td>' + escapeHtml(ref.code || '') + '</td><td>' + escapeHtml(ref.description || '') + '</td><td>' + escapeHtml(pathText(ref.categoryPath)) + '</td></tr>'; }).join('');
    var comparisonSection = document.comparisonDetails ? '<section><h2>Detail Komparasi</h2><table><thead><tr><th>Status</th><th>Jenis perubahan / item</th><th>Sumber Workbook A</th><th>Sumber Workbook B</th><th>Selisih</th></tr></thead><tbody>' + printComparisonRows(document.comparisonDetails) + '</tbody></table></section>' : '';
    return '<!DOCTYPE html><html lang="id"><head><meta charset="utf-8"><meta name="viewport" content="width=device-width, initial-scale=1"><title>' + escapeHtml(title + ' - ' + document.reportId) + '</title><style>' +
      '@page{size:A4 ' + (isWide ? 'landscape' : 'portrait') + ';margin:14mm 12mm 16mm}*{box-sizing:border-box}body{font:9pt/1.4 Arial,Helvetica,sans-serif;color:#172033;margin:0}.toolbar{position:sticky;top:0;background:#eef3fb;padding:12px 16px;border-bottom:1px solid #cbd5e1;display:flex;justify-content:space-between;gap:12px;align-items:center}.toolbar button{font:inherit;padding:8px 14px}.toolbar p{margin:0}main{max-width:1100px;margin:18px auto;padding:0 10px}h1{font-size:20pt;margin:0 0 5px;color:#12315b}h2{font-size:13pt;margin:20px 0 7px;border-bottom:1px solid #cad5e3;padding-bottom:4px;break-after:avoid}h3{font-size:10pt;margin:12px 0 5px}p{margin:5px 0}.muted{color:#526176}.identity{margin:10px 0 14px}.summary{width:100%;max-width:500px;border-collapse:collapse}.summary th,.summary td{border:1px solid #d6deea;padding:5px 7px;text-align:left}.summary th{width:62%;background:#f3f6fa}table{width:100%;border-collapse:collapse;table-layout:fixed;margin:7px 0 14px}thead{display:table-header-group}th,td{border:1px solid #cbd5e1;padding:5px 6px;text-align:left;vertical-align:top;overflow-wrap:anywhere;word-break:break-word}th{background:#eef2f7;font-weight:700}tr{break-inside:avoid}small{font-size:7.5pt;color:#48566b}ul{margin:5px 0;padding-left:20px}li{margin:2px 0;overflow-wrap:anywhere}.empty{color:#64748b}.source-list{columns:2;column-gap:28px}.report-note{margin:8px 0 0;color:#58677d;font-size:8pt;overflow-wrap:anywhere}.block{white-space:pre-wrap}@media print{.toolbar{display:none}main{max-width:none;margin:0;padding:0}a{color:inherit;text-decoration:none}}' +
      '</style></head><body><div class="toolbar"><p><strong>Pratinjau PDF:</strong> pilih “Cetak / Simpan sebagai PDF”, lalu pilih destinasi PDF pada dialog browser. Laporan tidak dikirim ke server.</p><button type="button" onclick="window.print()">Cetak / Simpan sebagai PDF</button></div><main>' +
      '<header><h1>' + escapeHtml(title) + '</h1><p><strong>Report ID:</strong> ' + escapeHtml(document.reportId) + ' &nbsp; <strong>Dibuat:</strong> ' + escapeHtml(document.generatedAt) + '</p><p class="muted"><strong>Proyek:</strong> ' + escapeHtml(document.projectMetadata && document.projectMetadata.projectName || 'Tidak tersedia') + ' &nbsp; <strong>Kontrak:</strong> ' + escapeHtml(document.schemaVersion) + '</p><p class="report-note">Dokumen dibuat secara lokal oleh RAB Guard. Hasil laporan menyalin hasil engine yang tersedia dan tidak menjalankan ulang audit atau komparasi.</p></header>' +
      '<section><h2>Workbook Sumber</h2><ul>' + workbookList + '</ul></section><section><h2>Ringkasan</h2><table class="summary"><tbody>' + summaryHtml + '</tbody></table></section>' +
      '<section><h2>Engine yang Dicakup</h2><ul>' + (engines || '<li>Tidak ada informasi engine.</li>') + '</ul>' + (warnings ? '<h3>Peringatan dan status parsial</h3><ul>' + warnings + '</ul>' : '') + '</section>' +
      (document.findings.length || document.analysisEngines.some(function (e) { return e.engineName === 'UniversalChecker' || e.engineName === 'UnitConsistencyEngine'; }) ? '<section><h2>Temuan Analisis</h2><table><thead><tr><th>Severity</th><th>Judul / Engine / Rule</th><th>Penjelasan</th><th>Referensi Sumber</th><th>Tindakan</th></tr></thead><tbody>' + printFindingRows(document) + '</tbody></table></section>' : '') +
      comparisonSection +
      (document.sourceReferences.length ? '<section><h2>Referensi Sumber</h2><table><thead><tr><th>Workbook</th><th>Sheet</th><th>Baris</th><th>Sel/Range</th><th>Kode</th><th>Uraian</th><th>Category Path</th></tr></thead><tbody>' + (refRows || '<tr><td colspan="7">Tidak tersedia.</td></tr>') + '</tbody></table></section>' : '<section><h2>Referensi Sumber</h2><p class="empty">Engine tidak menyediakan referensi sumber untuk hasil ini.</p></section>') +
      '<section><h2>Asumsi</h2><ul>' + assumptions + '</ul></section><section><h2>Keterbatasan Analisis</h2><ul>' + limitations + '</ul></section>' +
      '</main></body></html>';
  }

  function safeFilePart(value) {
    var part = safeString(value).normalize ? safeString(value).normalize('NFKD').replace(/[\u0300-\u036f]/g, '') : safeString(value);
    return part.replace(/[<>:"/\\|?*\x00-\x1F]/g, '_').replace(/\s+/g, '_').replace(/[. ]+$/g, '').slice(0, 70) || 'Tanpa_Nama';
  }
  function fileName(document, extension) {
    var project = document.projectMetadata && (document.projectMetadata.projectName || document.projectMetadata.sourceFilename) || document.sourceWorkbooks[0] && document.sourceWorkbooks[0].filename || 'RAB';
    var date = String(document.generatedAt).slice(0, 10).replace(/-/g, '-');
    return 'RABGuard_' + safeFilePart(document.reportType) + '_' + safeFilePart(project) + '_' + date + '.' + extension;
  }

  var ReportingEngine = {
    engineName: 'ReportingEngine',
    engineVersion: VERSION,
    schemaVersion: SCHEMA,
    createDocument: createDocument,
    validateDocument: validateDocument,
    buildExcelWorkbook: buildExcelWorkbook,
    toXlsx: xlsxBuffer,
    renderPrintHtml: renderPrintHtml,
    fileName: fileName,
    reportTypeTitle: reportTypeTitle,
    escapeHtml: escapeHtml
  };
  window.ReportingEngine = ReportingEngine;
})(typeof window !== 'undefined' ? window : this);
