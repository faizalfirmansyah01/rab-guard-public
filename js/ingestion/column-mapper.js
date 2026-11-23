/**
 * RAB GUARD — Universal XLSX Ingestion
 * Module: ColumnMapper (js/ingestion/column-mapper.js)
 * 
 * Tanggung Jawab:
 * - Mengelola konfigurasi pemetaan kolom dari nama fisik Excel ke skema kanonikal.
 * - Memfasilitasi koreksi manual oleh pengguna bila deteksi otomatis tidak yakin.
 * - Memvalidasi kelengkapan kolom wajib (description, volume, unit_price, amount).
 */

(function (window) {
  'use strict';

  var CANONICAL_FIELDS = [
    { key: 'item_number', label: 'Nomor Urut (NO)', required: false },
    { key: 'item_code', label: 'Kode Pekerjaan (KODE)', required: false },
    { key: 'analysis_code', label: 'Kode Analisa / AHSP', required: false },
    { key: 'description', label: 'Uraian Pekerjaan', required: true },
    { key: 'volume', label: 'Volume / Kuantitas', required: false },
    { key: 'unit', label: 'Satuan Pengukuran', required: false },
    { key: 'unit_price', label: 'Harga Satuan (Rp)', required: false },
    { key: 'amount', label: 'Jumlah Harga (Rp)', required: false },
    { key: 'category', label: 'Divisi / Kategori', required: false }
  ];

  var ColumnMapper = {
    getCanonicalFields: function () {
      return CANONICAL_FIELDS;
    },

    /**
     * Membuat mapping konfigurasi awal dari hasil HeaderDetector.
     * @param {Object} detectionResult Hasil dari HeaderDetector.detect()
     * @returns {Object} Konfigurasi pemetaan kolom siap pakai/diedit
     */
    createInitialMapping: function (detectionResult) {
      var mapping = {};
      var rawHeaders = detectionResult ? detectionResult.rawHeaders || [] : [];
      var detected = detectionResult ? detectionResult.mappings || {} : {};

      for (var i = 0; i < CANONICAL_FIELDS.length; i++) {
        var f = CANONICAL_FIELDS[i];
        if (detected[f.key]) {
          mapping[f.key] = {
            colIndex: detected[f.key].colIndex,
            colLetter: detected[f.key].colLetter,
            originalHeader: detected[f.key].originalHeader,
            isAuto: true,
            confidence: 100
          };
        } else {
          mapping[f.key] = {
            colIndex: -1,
            colLetter: '',
            originalHeader: '',
            isAuto: false,
            confidence: 0
          };
        }
      }

      return {
        headerRowIndex: detectionResult ? detectionResult.headerRowIndex : -1,
        headerRow1Based: detectionResult ? detectionResult.headerRow1Based : -1,
        headerRowEndIndex: detectionResult && Number.isInteger(detectionResult.headerRowEndIndex)
          ? detectionResult.headerRowEndIndex : (detectionResult ? detectionResult.headerRowIndex : -1),
        headerRowEnd1Based: detectionResult && Number.isInteger(detectionResult.headerRowEnd1Based)
          ? detectionResult.headerRowEnd1Based : (detectionResult ? detectionResult.headerRow1Based : -1),
        fields: mapping,
        categoryLevels: (detected.categoryLevels || (detected.category ? [detected.category] : [])).map(function (level, index) {
          var header = String(level.originalHeader || '').toLowerCase();
          var explicit = header.match(/(?:level|tier|tingkat)\s*(\d+)/);
          var levelIndex = explicit ? Number(explicit[1]) - 1 : (Number.isInteger(level.levelIndex) ? level.levelIndex : index);
          return { levelIndex: levelIndex,
            colIndex: level.colIndex, colLetter: level.colLetter, originalHeader: level.originalHeader,
            semanticLevel: level.semanticLevel || null, semanticEvidence: level.semanticEvidence || null,
            isAuto: true, confidence: 100 };
        }),
        rawHeaders: rawHeaders,
        unresolvedFields: detectionResult && Array.isArray(detectionResult.conflicts)
          ? detectionResult.conflicts.map(function (conflict) { return conflict.field; })
            .filter(function (field, index, all) { return all.indexOf(field) === index; })
          : []
      };
    },

    /**
     * Memperbarui pemetaan kolom secara manual.
     * @param {Object} currentConfig
     * @param {string} fieldKey
     * @param {number|string} colSpec Indeks angka atau huruf kolom (misal 'C' atau 2)
     * @param {string} [headerText]
     * @returns {Object} Updated config
     */
    updateMappingField: function (currentConfig, fieldKey, colSpec, headerText) {
      var config = JSON.parse(JSON.stringify(currentConfig));
      var colIndex = -1;
      var colLetter = '';

      if (typeof colSpec === 'number') {
        colIndex = colSpec;
        if (colIndex >= 0 && window.XLSX) {
          colLetter = window.XLSX.utils.encode_col(colIndex);
        }
      } else if (typeof colSpec === 'string' && colSpec.trim()) {
        colLetter = colSpec.trim().toUpperCase();
        if (window.XLSX) {
          try {
            colIndex = window.XLSX.utils.decode_col(colLetter);
          } catch (e) {
            colIndex = -1;
          }
        }
      }

      config.fields[fieldKey] = {
        colIndex: colIndex,
        colLetter: colLetter,
        originalHeader: headerText || (colLetter ? 'Kolom ' + colLetter : ''),
        isAuto: false,
        confidence: 100
      };
      config.unresolvedFields = (config.unresolvedFields || []).filter(function (field) {
        return field !== fieldKey;
      });

      return config;
    },

    /**
     * Memvalidasi apakah mapping sudah memenuhi syarat minimum untuk normalisasi.
     * @param {Object} mappingConfig
     * @returns {{ valid: boolean, errors: string[], warnings: string[] }}
     */
    validateMapping: function (mappingConfig) {
      var errors = [];
      var warnings = [];

      if (!mappingConfig || !mappingConfig.fields) {
        return { valid: false, errors: ['Konfigurasi pemetaan kolom kosong.'], warnings: [] };
      }

      if (Array.isArray(mappingConfig.unresolvedFields) && mappingConfig.unresolvedFields.length) {
        errors.push('Pilih kolom secara manual untuk pemetaan yang memiliki beberapa kandidat: ' +
          mappingConfig.unresolvedFields.join(', ') + '.');
      }

      var fields = mappingConfig.fields;
      var categoryLevels = Array.isArray(mappingConfig.categoryLevels) ? mappingConfig.categoryLevels : [];
      var levelIndexes = {};
      categoryLevels.forEach(function (level) {
        if (!Number.isInteger(level.levelIndex) || level.levelIndex < 0 || levelIndexes[level.levelIndex]) {
          errors.push('Indeks level kategori tidak valid atau duplikat.');
        }
        levelIndexes[level.levelIndex] = true;
        if (!Number.isInteger(level.colIndex) || level.colIndex < 0 || !level.colLetter || !level.originalHeader) {
          errors.push('Mapping kategori bertingkat tidak memiliki identitas kolom/header yang valid.');
        }
      });

      // Cek field wajib
      if (!fields.description || fields.description.colIndex < 0) {
        errors.push('Kolom "Uraian Pekerjaan" (description) wajib dipetakan.');
      }
      if (!fields.amount || fields.amount.colIndex < 0) {
        warnings.push('Kolom "Jumlah Harga" (amount) belum dipetakan. Mesin hitung akan bergantung pada perkalian volume × harga satuan.');
      }
      if (!fields.volume || fields.volume.colIndex < 0) {
        warnings.push('Kolom "Volume" belum dipetakan. Beberapa aturan matematis tidak dapat diverifikasi.');
      }
      if (!fields.unit_price || fields.unit_price.colIndex < 0) {
        warnings.push('Kolom "Harga Satuan" belum dipetakan.');
      }

      // Cek duplikasi kolom
      var colUsage = {};
      for (var k in fields) {
        if (k === 'category' && categoryLevels.some(function (level) { return level.colIndex === fields[k].colIndex; })) continue;
        var idx = fields[k].colIndex;
        if (idx >= 0) {
          if (colUsage[idx]) {
            errors.push('Kolom ' + fields[k].colLetter + ' dipetakan ke lebih dari satu field kanonikal: "' + colUsage[idx] + '" dan "' + k + '".');
          } else {
            colUsage[idx] = k;
          }
        }
      }
      categoryLevels.forEach(function (level) {
        if (colUsage[level.colIndex]) errors.push('Kolom ' + level.colLetter + ' dipetakan sebagai kategori dan field ' + colUsage[level.colIndex] + '.');
        else colUsage[level.colIndex] = 'category level ' + level.levelIndex;
      });

      return {
        valid: errors.length === 0,
        errors: errors,
        warnings: warnings
      };
    }
  };

  ColumnMapper.CANONICAL_FIELDS = CANONICAL_FIELDS;
  window.ColumnMapper = ColumnMapper;

})(typeof window !== 'undefined' ? window : this);
