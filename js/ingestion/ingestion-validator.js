/**
 * RAB GUARD — Universal XLSX Ingestion
 * Module: IngestionValidator (js/ingestion/ingestion-validator.js)
 * 
 * Tanggung Jawab:
 * - Memvalidasi integritas Model Data Kanonikal pasca-normalisasi.
 * - Menghitung metrik kelengkapan data (kelengkapan volume, harga, rumus).
 * - Menghasilkan laporan validasi awal sebelum eksekusi mesin audit dan simulasi.
 */

(function (window) {
  'use strict';

  var IngestionValidator = {
    /**
     * Memvalidasi data kanonikal.
     * @param {Object} canonicalData
     * @returns {Object}
     */
    validate: function (canonicalData) {
      var errors = [];
      var warnings = [];

      if (!canonicalData) {
        return {
          valid: false,
          errors: ['Data kanonikal kosong atau null.'],
          warnings: [],
          stats: {}
        };
      }

      if (!canonicalData.project || !canonicalData.project.sourceFilename) {
        errors.push('Metadata proyek tidak lengkap.');
      }

      var items = canonicalData.workItems || [];
      if (items.length === 0) {
        errors.push('Tidak ditemukan baris item pekerjaan yang valid dari lembar yang dipilih.');
      }

      var missingDescCount = 0;
      var missingVolCount = 0;
      var missingPriceCount = 0;
      var missingAmountCount = 0;
      var formulaAmountCount = 0;
      var hierarchyIssues = [];
      var hierarchyAmbiguousCount = 0;
      var hierarchyPartialCount = 0;

      for (var i = 0; i < items.length; i++) {
        var it = items[i];
        if (!it.description || !it.description.trim()) missingDescCount++;
        if (it.volume === null || it.volume === undefined) missingVolCount++;
        if (it.unitPrice === null || it.unitPrice === undefined) missingPriceCount++;
        if (it.amount === null || it.amount === undefined) missingAmountCount++;
        if (it.source && it.source.amountFormula) formulaAmountCount++;
        this.validateHierarchyRecord(it, 'workItems[' + i + ']', hierarchyIssues);
        if (it.hierarchy && it.hierarchy.status === 'ambiguous') hierarchyAmbiguousCount++;
        if (it.hierarchy && it.hierarchy.status === 'partial') hierarchyPartialCount++;
      }

      (canonicalData.subtotals || []).forEach(function (subtotal, index) {
        this.validateHierarchyRecord(subtotal, 'subtotals[' + index + ']', hierarchyIssues);
      }, this);

      if (hierarchyIssues.length) warnings.push('Terdapat ' + hierarchyIssues.length + ' masalah interpretasi/provenance hierarki; validitas finansial item dinilai terpisah.');

      if (missingDescCount > 0) {
        warnings.push('Terdapat ' + missingDescCount + ' item pekerjaan tanpa uraian deskripsi.');
      }
      if (missingVolCount > 0) {
        warnings.push('Terdapat ' + missingVolCount + ' item pekerjaan tanpa volume kuantitas numerik.');
      }
      if (missingPriceCount > 0) {
        warnings.push('Terdapat ' + missingPriceCount + ' item pekerjaan tanpa harga satuan.');
      }
      if (missingAmountCount > 0) {
        warnings.push('Terdapat ' + missingAmountCount + ' item pekerjaan tanpa jumlah harga.');
      }

      var completenessScore = 100;
      if (items.length > 0) {
        var penalty = (missingDescCount * 25 + missingVolCount * 25 + missingPriceCount * 25 + missingAmountCount * 25) / items.length;
        completenessScore = Math.max(0, Math.round(100 - penalty));
      }

      return {
        valid: errors.length === 0,
        errors: errors,
        warnings: warnings,
        stats: {
          totalItems: items.length,
          totalHeadings: (canonicalData.headings || []).length,
          totalSubtotals: (canonicalData.subtotals || []).length,
          itemsWithFormula: formulaAmountCount,
          completenessScore: completenessScore,
          hierarchyAmbiguousItems: hierarchyAmbiguousCount,
          hierarchyPartialItems: hierarchyPartialCount,
          hierarchyIssues: hierarchyIssues
        }
      };
    },

    validateHierarchyRecord: function (record, label, issues) {
      var path = record && record.categoryPath;
      var hierarchy = record && record.hierarchy;
      var allowedStatus = ['resolved', 'partial', 'ambiguous', 'none'];
      var allowedResolution = ['resolved', 'inherited', 'missing', 'ambiguous'];
      var allowedOrigins = ['column', 'vertical_heading', 'merged_cell'];
      var cellRefPattern = /^[A-Z]{1,3}[1-9]\d*$/;
      if (!Array.isArray(path) || !hierarchy || allowedStatus.indexOf(hierarchy.status) === -1 || !hierarchy.structure || !Array.isArray(hierarchy.unresolved)) {
        issues.push(label + ': struktur hierarchy tidak lengkap.');
        return;
      }
      var seen = {};
      var previous = -1;
      path.forEach(function (level, index) {
        if (!Number.isInteger(level.levelIndex) || level.levelIndex < 0 || seen[level.levelIndex] || level.levelIndex <= previous) {
          issues.push(label + ': indeks level tidak valid, duplikat, atau tidak terurut (' + index + ').');
        }
        seen[level.levelIndex] = true;
        previous = level.levelIndex;
        if (allowedResolution.indexOf(level.resolution) === -1 || (level.value === null && level.resolution !== 'missing' && level.resolution !== 'ambiguous')) {
          issues.push(label + ': nilai/status level tidak konsisten (' + index + ').');
        }
        var p = level.provenance || {};
        var sourceRef = p.originalValueCell && p.originalValueCell.ref;
        if (allowedOrigins.indexOf(p.origin) === -1 || typeof p.sheet !== 'string' || !p.sheet.trim() ||
            !p.originalValueCell || !sourceRef || !cellRefPattern.test(sourceRef) ||
            !p.appliedTo || !Number.isInteger(p.appliedTo.row) || p.appliedTo.row < 1 ||
            (p.appliedTo.cellRef !== null && p.appliedTo.cellRef !== undefined && !cellRefPattern.test(p.appliedTo.cellRef))) {
          issues.push(label + ': provenance sumber atau tujuan tidak valid (' + index + ').');
        }
        if (p.headerSourceCell && p.headerSourceCell.ref && !cellRefPattern.test(p.headerSourceCell.ref)) {
          issues.push(label + ': referensi header hierarchy tidak valid (' + index + ').');
        }
        if (p.origin === 'vertical_heading' && (!p.verticalHeadingSourceCell || !p.verticalHeadingSourceCell.ref || !cellRefPattern.test(p.verticalHeadingSourceCell.ref))) {
          issues.push(label + ': provenance heading vertikal tidak lengkap (' + index + ').');
        }
        if (p.origin === 'merged_cell') {
          var merge = p.mergedAnchor;
          if (!merge || !merge.ref || !cellRefPattern.test(merge.ref) || !merge.range || !p.inherited && p.appliedTo.row !== Number(String(merge.ref).match(/\d+$/)[0])) {
            issues.push(label + ': provenance merged cell tidak memiliki anchor/range yang konsisten (' + index + ').');
          }
          if (merge && merge.range && window.XLSX && window.XLSX.utils && window.XLSX.utils.decode_range) {
            try {
              var bounds = window.XLSX.utils.decode_range(merge.range);
              var appliedRef = p.appliedTo.cellRef;
              var applied = appliedRef ? window.XLSX.utils.decode_cell(appliedRef) : null;
              if (!applied || applied.c !== bounds.s.c || applied.r < bounds.s.r || applied.r > bounds.e.r || bounds.s.c !== bounds.e.c) {
                issues.push(label + ': tujuan inheritance berada di luar merged range (' + index + ').');
              }
            } catch (ignoreInvalidMerge) {
              issues.push(label + ': rentang merged cell tidak dapat dibaca (' + index + ').');
            }
          }
        }
        if (level.resolution === 'inherited' && (!p.inherited || !(p.mergedAnchor || p.verticalHeadingSourceCell))) {
          issues.push(label + ': inheritance tidak memiliki bukti struktural (' + index + ').');
        }
        if (level.resolution === 'resolved' && p.inherited) {
          issues.push(label + ': level resolved tidak boleh ditandai inherited (' + index + ').');
        }
      });
      if (hierarchy.status === 'none' && path.length) issues.push(label + ': status none harus memiliki path kosong.');
      if (hierarchy.status === 'none' && hierarchy.unresolved.length) issues.push(label + ': status none tidak boleh memiliki unresolved evidence.');
      if ((hierarchy.status === 'ambiguous' || hierarchy.status === 'partial') && !hierarchy.unresolved.length) issues.push(label + ': status hierarchy belum terselesaikan tanpa evidence.');
      if ((hierarchy.status === 'resolved' || hierarchy.status === 'none') && (hierarchy.unresolved.length || path.some(function (x) { return x.resolution === 'missing' || x.resolution === 'ambiguous'; }))) {
        issues.push(label + ': status resolved bertentangan dengan level unresolved.');
      }
      hierarchy.unresolved.forEach(function (evidence, index) {
        if (!evidence || typeof evidence.reason !== 'string' || !evidence.reason.trim() || !Array.isArray(evidence.candidates) || !Array.isArray(evidence.evidenceSources)) {
          issues.push(label + ': unresolved evidence tidak lengkap (' + index + ').');
        }
      });
      if (path.length && record.category === undefined) issues.push(label + ': compatibility category tidak tersedia.');
    }
  };

  window.IngestionValidator = IngestionValidator;

})(typeof window !== 'undefined' ? window : this);
