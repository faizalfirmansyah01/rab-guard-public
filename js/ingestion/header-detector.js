/**
 * RAB GUARD — Universal XLSX Ingestion
 * Module: HeaderDetector (js/ingestion/header-detector.js)
 * 
 * Tanggung Jawab:
 * - Memindai 25 baris awal lembar kerja untuk menemukan baris header sejati.
 * - Mengenali sinonim istilah RAB dalam Bahasa Indonesia dan Bahasa Inggris.
 * - Mengutamakan frasa spesifik lebih panjang (e.g. 'harga satuan' sebelum 'satuan').
 * - Mengembalikan indeks baris header, pemetaan kolom otomatis, dan skor keyakinan.
 */

(function (window) {
  'use strict';

  var DICTIONARY = {
    // Separate serial numbers from real work-item codes. In many BOQ sheets,
    // the leftmost NO column is only a sequence number; it must not consume
    // the canonical item_code mapping before a later KODE column is inspected.
    item_number: [
      'nomor urut', 'no urut', 'no item', 'item no', 'nomor', 'number', 'no'
    ],
    item_code: [
      'kode pekerjaan', 'kode item', 'item code', 'work code', 'kode', 'code', 'pos'
    ],
    analysis_code: [
      'kode analisa', 'kode analisis', 'kode ahsp', 'analysis code', 'analisa code', 'ahsp code'
    ],
    description: [
      'uraian pekerjaan', 'nama pekerjaan', 'item pekerjaan', 'work description',
      'item description', 'jenis pekerjaan', 'mata pembayaran', 'lingkup pekerjaan',
      'description', 'deskripsi', 'pekerjaan', 'uraian', 'keterangan'
    ],
    volume: [
      'perkiraan kuantitas', 'jumlah volume', 'kuantitas', 'quantity', 'volume', 'qty', 'vol'
    ],
    unit_price: [
      'harga satuan pekerjaan', 'harga satuan rp', 'harga satuan', 'unit price',
      'unit rate', 'h satuan', 'h sat', 'tarif', 'harga'
    ],
    amount: [
      'jumlah harga', 'total harga', 'total biaya', 'jumlah rp', 'total rp',
      'subtotal', 'amount', 'jumlah', 'total', 'biaya'
    ],
    unit: [
      'satuan ukuran', 'satuan', 'unit', 'uom', 'sat'
    ],
    category: [
      'sub divisi', 'subdivisi', 'sub division', 'subdivision', 'sub kelompok', 'subkelompok',
      'sub group', 'subgroup', 'sub kategori', 'subkategori', 'subcategory', 'sub category',
      'kategori', 'category', 'kelompok', 'group', 'divisi', 'division', 'bagian', 'section',
      'trade', 'discipline', 'subdiscipline', 'sub discipline', 'paket', 'package',
      'klasifikasi', 'classification', 'jenis pekerjaan', 'work group', 'work subgroup'
      , 'level 1', 'level 2', 'level 3', 'level 4', 'tingkat 1', 'tingkat 2', 'tingkat 3', 'tingkat 4',
      'tier 1', 'tier 2', 'tier 3', 'tier 4'
    ]
  };

  var HeaderDetector = {
    /**
     * Memindai sheet untuk mendeteksi baris header dan memetakan kolom.
     * @param {Object} sheet Sheet dari SheetJS
     * @param {Object} [options]
     * @returns {Object}
     */
    detect: function (sheet, options) {
      if (!sheet || !sheet['!ref']) {
        return {
          found: false,
          headerRowIndex: -1,
          headerRow1Based: -1,
          confidence: 0,
          mappings: {},
          rawHeaders: [],
          candidates: [],
          conflicts: [],
          ambiguous: false
        };
      }

      options = options || {};
      var maxScanRows = options.maxScanRows || 25;

      var range = window.XLSX.utils.decode_range(sheet['!ref']);
      var endScanRow = Math.min(range.e.r, range.s.r + maxScanRows);

      var bestCandidate = null;
      var highestScore = -1;
      var candidates = [];

      for (var r = range.s.r; r <= endScanRow; r++) {
        var rowCandidate = this.evaluateRow(sheet, r, range.s.c, range.e.c);
        if (rowCandidate.score > 0) candidates.push(rowCandidate);
        if (rowCandidate.score > highestScore) {
          highestScore = rowCandidate.score;
          bestCandidate = rowCandidate;
        }
      }

      // Evaluate adjacent header bands only when combining their labels adds
      // recognized fields (for example, a merged "Harga" group over "Satuan").
      for (var bandRow = range.s.r; bandRow < endScanRow; bandRow++) {
        var bandCandidate = this.evaluateHeaderBlock(sheet, bandRow, range.s.c, range.e.c);
        var firstRow = this.evaluateRow(sheet, bandRow, range.s.c, range.e.c);
        var secondRow = this.evaluateRow(sheet, bandRow + 1, range.s.c, range.e.c);
        if (bandCandidate && bandCandidate.score > Math.max(firstRow.score, secondRow.score)) {
          candidates.push(bandCandidate);
        }
      }

      candidates.sort(function (a, b) { return b.score - a.score; });
      if (candidates.length && (!bestCandidate || candidates[0].score > bestCandidate.score)) {
        bestCandidate = candidates[0];
      }
      if (Number.isInteger(options.headerRowIndex) &&
          options.headerRowIndex >= range.s.r && options.headerRowIndex <= range.e.r) {
        bestCandidate = this.evaluateRow(sheet, options.headerRowIndex, range.s.c, range.e.c);
      }

      if (!bestCandidate || (bestCandidate.score < 2 && options.headerRowIndex === undefined)) {
        return {
          found: false,
          headerRowIndex: -1,
          headerRow1Based: -1,
          confidence: 0,
          mappings: {},
          rawHeaders: [],
          candidates: candidates.slice(0, 5).map(function (candidate) {
            return this.describeCandidate(candidate);
          }, this),
          conflicts: bestCandidate && bestCandidate.conflicts ? bestCandidate.conflicts : [],
          ambiguous: false,
          reason: 'Tidak ditemukan baris yang memenuhi kriteria header RAB standar.'
        };
      }

      var topCandidates = candidates.filter(function (candidate) {
        return candidate.score === candidates[0].score;
      });
      var ambiguous = false;
      if (options.headerRowIndex === undefined && topCandidates.length > 1) {
        var ranked = topCandidates.map(function (candidate) {
          candidate.dataEvidence = this.evaluateDataEvidence(sheet, candidate, range, candidates);
          return candidate;
        }, this);
        ranked.sort(function (a, b) {
          return b.dataEvidence.score - a.dataEvidence.score || a.rowIndex - b.rowIndex;
        });
        var bestEvidence = ranked[0].dataEvidence.score;
        var tiedEvidence = ranked.filter(function (candidate) {
          return candidate.dataEvidence.score === bestEvidence;
        });
        var sameMapping = tiedEvidence.every(function (candidate) {
          return this.mappingSignature(candidate.mappings) === this.mappingSignature(tiedEvidence[0].mappings);
        }, this);
        if (tiedEvidence.length === 1 || sameMapping) {
          bestCandidate = ranked[0];
        } else {
          ambiguous = true;
        }
      }
      if (!ambiguous && bestCandidate && !bestCandidate.dataEvidence) {
        bestCandidate.dataEvidence = this.evaluateDataEvidence(sheet, bestCandidate, range, candidates);
      }
      var confidence = this.getConfidence(bestCandidate);
      var ambiguityReason = null;
      if (ambiguous || (bestCandidate.conflicts && bestCandidate.conflicts.length)) {
        ambiguityReason = (bestCandidate.conflicts && bestCandidate.conflicts.length)
          ? 'Kandidat header memiliki konflik kolom: ' + bestCandidate.conflicts.map(function (conflict) {
            return conflict.field + ' (' + (conflict.firstColumn || '?') + '/' + (conflict.competingColumn || '?') + ')';
          }).join(', ') + '.'
          : 'Kandidat header dengan skor ' + bestCandidate.score + ' memiliki mapping berbeda dan bukti data setara pada baris ' +
            topCandidates.filter(function (candidate) {
              return candidate.dataEvidence && candidate.dataEvidence.score === bestCandidate.dataEvidence.score;
            }).map(function (candidate) { return candidate.rowIndex + 1; }).join(', ') + '.';
      }

      return {
        found: true,
        headerRowIndex: bestCandidate.rowIndex,
        headerRow1Based: bestCandidate.rowIndex + 1,
        headerRowEndIndex: bestCandidate.endRowIndex !== undefined ? bestCandidate.endRowIndex : bestCandidate.rowIndex,
        headerRowEnd1Based: (bestCandidate.endRowIndex !== undefined ? bestCandidate.endRowIndex : bestCandidate.rowIndex) + 1,
        confidence: confidence,
        score: bestCandidate.score,
        mappings: bestCandidate.mappings,
        rawHeaders: bestCandidate.rawHeaders,
        candidates: candidates.slice(0, 5).map(function (candidate) {
          return this.describeCandidate(candidate);
        }, this),
        conflicts: bestCandidate.conflicts || [],
        dataEvidence: bestCandidate.dataEvidence || { score: 0, dataRows: 0 },
        ambiguityReason: ambiguityReason,
        ambiguous: ambiguous || !!(bestCandidate.conflicts && bestCandidate.conflicts.length)
      };
    },

    mappingSignature: function (mappings) {
      var source = mappings || {};
      var keys = Object.keys(source).filter(function (key) { return key !== 'categoryLevels'; }).sort();
      var signature = keys.map(function (key) { return key + ':' + source[key].colIndex; });
      if (Array.isArray(source.categoryLevels)) {
        signature.push('categoryLevels:' + source.categoryLevels.map(function (level) { return level.colIndex; }).join(','));
      }
      return signature.join('|');
    },

    isHeaderLabel: function (value) {
      var clean = String(value || '').toLowerCase().replace(/[_\-.():\[\]]/g, ' ').replace(/\s+/g, ' ').trim();
      for (var field in DICTIONARY) {
        if (DICTIONARY[field].indexOf(clean) !== -1) return true;
      }
      return false;
    },

    evaluateHeaderBlock: function (sheet, rowIndex, startCol, endCol) {
      var mappings = {};
      var rawHeaders = [];
      var conflicts = [];
      var usedFields = {};
      var categoryCandidates = [];
      var score = 0;
      for (var c = startCol; c <= endCol; c++) {
        var first = sheet[window.XLSX.utils.encode_cell({ r: rowIndex, c: c })];
        var second = sheet[window.XLSX.utils.encode_cell({ r: rowIndex + 1, c: c })];
        var firstText = first && first.v !== undefined && first.v !== null ? String(first.v).trim() : '';
        var secondText = second && second.v !== undefined && second.v !== null ? String(second.v).trim() : '';
        var text = [firstText, secondText].filter(Boolean).join(' ').trim();
        rawHeaders.push({ colIndex: c, colLetter: window.XLSX.utils.encode_col(c), text: text });
        if (!text) continue;
        var cleanText = text.toLowerCase().replace(/[_\-.():\[\]]/g, ' ').replace(/\s+/g, ' ').trim();
        var matchedField = this.matchDictionary(cleanText);
        if (matchedField === 'category') {
          var duplicateCategory = categoryCandidates.filter(function (candidate) { return candidate.normalizedHeader === cleanText; })[0];
          if (duplicateCategory) {
            conflicts.push({ field: 'category_hierarchy', firstColumn: duplicateCategory.colLetter,
              competingColumn: window.XLSX.utils.encode_col(c), firstHeader: duplicateCategory.originalHeader,
              competingHeader: text });
          }
          categoryCandidates.push({ colIndex: c, colLetter: window.XLSX.utils.encode_col(c),
            originalHeader: text, normalizedHeader: cleanText });
          score += 1;
          continue;
        }
        if (matchedField && usedFields[matchedField]) {
          conflicts.push({ field: matchedField, competingColumn: window.XLSX.utils.encode_col(c), competingHeader: text });
          continue;
        }
        if (!matchedField) continue;
        usedFields[matchedField] = true;
        mappings[matchedField] = {
          colIndex: c,
          colLetter: window.XLSX.utils.encode_col(c),
          originalHeader: text
        };
        if (matchedField === 'description') score += 4;
        else if (matchedField === 'volume' || matchedField === 'unit_price' || matchedField === 'amount') score += 3;
        else if (matchedField === 'unit' || matchedField === 'item_code' || matchedField === 'analysis_code') score += 2;
        else score += 1;
      }
      categoryCandidates.sort(function (a, b) { return a.colIndex - b.colIndex; });
      if (categoryCandidates.length) {
        mappings.category = categoryCandidates[0];
        mappings.categoryLevels = categoryCandidates;
      }
      if (!mappings.description || !mappings.volume || !mappings.unit_price || conflicts.length) return null;
      return { rowIndex: rowIndex, endRowIndex: rowIndex + 1, score: score, mappings: mappings, rawHeaders: rawHeaders, conflicts: conflicts, multirow: true };
    },

    evaluateDataEvidence: function (sheet, candidate, range, candidates) {
      var signature = this.mappingSignature(candidate.mappings);
      var endRow = candidate.endRowIndex !== undefined ? candidate.endRowIndex : candidate.rowIndex;
      var stopRow = Math.min(range.e.r, endRow + 8);
      candidates.forEach(function (other) {
        var otherEnd = other.endRowIndex !== undefined ? other.endRowIndex : other.rowIndex;
        if (other.rowIndex > candidate.rowIndex && this.mappingSignature(other.mappings) === signature) {
          stopRow = Math.min(stopRow, other.rowIndex - 1);
        }
      }, this);
      var score = 0;
      var dataRows = 0;
      var inspectedRows = 0;
      for (var r = endRow + 1; r <= stopRow && inspectedRows < 6; r++) {
        var hasAny = false;
        var descriptionEvidence = false;
        var numericEvidence = 0;
        var rowScore = 0;
        Object.keys(candidate.mappings).filter(function (field) { return field !== 'categoryLevels'; }).forEach(function (field) {
          var meta = candidate.mappings[field];
          var cell = sheet[window.XLSX.utils.encode_cell({ r: r, c: meta.colIndex })];
          if (!cell || cell.v === undefined || cell.v === null || String(cell.v).trim() === '') return;
          hasAny = true;
          var value = String(cell.v).trim();
          if (field === 'description' && /[a-z\u00c0-\u024f]/i.test(value) && !this.isHeaderLabel(value) && !/^\s*(sub\s*total|subtotal|grand\s*total|total)\b/i.test(value)) {
            descriptionEvidence = true;
            rowScore += 2;
          } else if (field === 'volume' || field === 'unit_price' || field === 'amount') {
            if (cell.t === 'n' || cell.f || /^\(?\s*-?[\d.,]+\s*\)?-?$/.test(value)) {
              numericEvidence++;
              rowScore++;
            }
          } else if (field === 'unit' && /[a-zA-Z]/.test(value) && value.length <= 16) {
            rowScore++;
          }
        }, this);
        if (!hasAny) continue;
        inspectedRows++;
        if (descriptionEvidence && numericEvidence > 0) {
          dataRows++;
          score += rowScore;
        }
      }
      return { score: score, dataRows: dataRows };
    },

    isRepeatedHeaderRow: function (sheet, rowIndex, fields) {
      return this.getRepeatedHeaderRowEnd(sheet, rowIndex, fields) >= rowIndex;
    },

    getRepeatedHeaderRowEnd: function (sheet, rowIndex, fields) {
      var range = window.XLSX.utils.decode_range(sheet['!ref'] || 'A1:A1');
      var candidate = this.evaluateRow(sheet, rowIndex, range.s.c, range.e.c);
      var expected = {};
      Object.keys(fields || {}).forEach(function (key) {
        if (fields[key] && fields[key].colIndex >= 0) expected[key] = fields[key].colIndex;
      });
      var expectedSignature = Object.keys(expected).sort().map(function (key) {
        return key + ':' + expected[key];
      }).join('|');
      if (candidate.score > 0 && this.mappingSignature(candidate.mappings) === expectedSignature) return rowIndex;
      if (rowIndex < range.e.r) {
        var block = this.evaluateHeaderBlock(sheet, rowIndex, range.s.c, range.e.c);
        if (block && this.mappingSignature(block.mappings) === expectedSignature) return rowIndex + 1;
      }
      return -1;
    },

    getConfidence: function (candidate) {
      var matchedRequired = 0;
      if (candidate.mappings.description) matchedRequired += 35;
      if (candidate.mappings.volume) matchedRequired += 25;
      if (candidate.mappings.unit_price) matchedRequired += 20;
      if (candidate.mappings.amount) matchedRequired += 20;
      return Math.min(100, matchedRequired);
    },

    describeCandidate: function (candidate) {
      return {
        rowIndex: candidate.rowIndex,
        row1Based: candidate.rowIndex + 1,
        rowEndIndex: candidate.endRowIndex !== undefined ? candidate.endRowIndex : candidate.rowIndex,
        rowEnd1Based: (candidate.endRowIndex !== undefined ? candidate.endRowIndex : candidate.rowIndex) + 1,
        score: candidate.score,
        confidence: this.getConfidence(candidate),
        mappings: candidate.mappings,
        rawHeaders: candidate.rawHeaders,
        conflicts: candidate.conflicts || [],
        dataEvidence: candidate.dataEvidence || null,
        multirow: !!candidate.multirow
      };
    },

    /**
     * Mengevaluasi satu baris sebagai kandidat header.
     */
    evaluateRow: function (sheet, rowIndex, startCol, endCol) {
      var rawHeaders = [];
      var mappings = {};
      var score = 0;
      var usedFields = {};
      var conflicts = [];
      var categoryCandidates = [];

      for (var c = startCol; c <= endCol; c++) {
        var cellRef = window.XLSX.utils.encode_cell({ r: rowIndex, c: c });
        var cell = sheet[cellRef];
        var text = (cell && cell.v !== undefined && cell.v !== null) ? String(cell.v).trim() : '';

        var colLetter = window.XLSX.utils.encode_col(c);
        rawHeaders.push({
          colIndex: c,
          colLetter: colLetter,
          text: text
        });

        if (!text) continue;

        var cleanText = text.toLowerCase()
          .replace(/[_\-\.\:\(\)\[\]]/g, ' ')
          .replace(/\s+/g, ' ')
          .trim();

        // Cari kecocokan ke kamus istilah
        var matchedField = this.matchDictionary(cleanText);
        if (matchedField === 'category') {
          var duplicateCategory = categoryCandidates.filter(function (candidate) { return candidate.normalizedHeader === cleanText; })[0];
          if (duplicateCategory) {
            conflicts.push({ field: 'category_hierarchy', firstColumn: duplicateCategory.colLetter,
              competingColumn: colLetter, firstHeader: duplicateCategory.originalHeader,
              competingHeader: text });
          }
          categoryCandidates.push({ colIndex: c, colLetter: colLetter, originalHeader: text,
            normalizedHeader: cleanText });
          score += 1;
          continue;
        }

        if (matchedField && usedFields[matchedField]) {
          conflicts.push({
            field: matchedField,
            firstColumn: mappings[matchedField].colLetter,
            competingColumn: colLetter,
            firstHeader: mappings[matchedField].originalHeader,
            competingHeader: text
          });
          continue;
        }

        if (matchedField) {
          usedFields[matchedField] = true;
          mappings[matchedField] = {
            colIndex: c,
            colLetter: colLetter,
            originalHeader: text
          };

          // Bobot skor
          if (matchedField === 'description') score += 4;
          else if (matchedField === 'volume' || matchedField === 'quantity') score += 3;
          else if (matchedField === 'unit_price' || matchedField === 'amount') score += 3;
          else if (matchedField === 'unit' || matchedField === 'item_code' || matchedField === 'analysis_code') score += 2;
          else if (matchedField === 'item_number') score += 1;
          else score += 1;
        }
      }

      categoryCandidates.sort(function (a, b) { return a.colIndex - b.colIndex; });
      if (categoryCandidates.length) {
        mappings.category = categoryCandidates[0];
        mappings.categoryLevels = categoryCandidates;
      }

      return {
        rowIndex: rowIndex,
        score: score,
        mappings: mappings,
        rawHeaders: rawHeaders,
        conflicts: conflicts
      };
    },

    /**
     * Mencocokkan string sel dengan kamus kanonikal.
     * Mengutamakan kecocokan eksak, kemudian kecocokan frasa terpanjang.
     */
    matchDictionary: function (cleanText) {
      // 1. Cek kecocokan eksak
      for (var field in DICTIONARY) {
        var words = DICTIONARY[field];
        for (var i = 0; i < words.length; i++) {
          if (cleanText === words[i]) {
            return field;
          }
        }
      }

      // Kumpulkan semua kata kunci dengan panjangnya untuk mencocokkan frasa terpanjang lebih dahulu
      var candidates = [];
      for (var f in DICTIONARY) {
        var wList = DICTIONARY[f];
        for (var k = 0; k < wList.length; k++) {
          candidates.push({
            field: f,
            phrase: wList[k],
            length: wList[k].length
          });
        }
      }

      // Urutkan dari frasa terpanjang ke terpendek
      candidates.sort(function (a, b) {
        return b.length - a.length;
      });

      // 2. Cek kecocokan parsial berbasis frasa terpanjang
      for (var m = 0; m < candidates.length; m++) {
        var cand = candidates[m];
        if (cand.length >= 3) {
          // Cocokkan batas kata atau substring
          var regex = new RegExp('\\b' + cand.phrase + '\\b', 'i');
          if (regex.test(cleanText) || cleanText.indexOf(cand.phrase) !== -1) {
            return cand.field;
          }
        }
      }

      return null;
    }
  };

  window.HeaderDetector = HeaderDetector;

})(typeof window !== 'undefined' ? window : this);
