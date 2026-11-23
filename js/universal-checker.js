/**
 * RAB GUARD — Universal XLSX Audit Engine
 * Module: UniversalChecker (js/universal-checker.js)
 * 
 * Tanggung Jawab:
 * - Menjalankan aturan audit universal terhadap Model Data Kanonikal.
 * - Mencakup 18 aturan pemeriksaan struktural, matematis, dan konsistensi data.
 * - Setiap temuan memiliki identitas jelas, tingkat keparahan (ERROR/WARNING/INFO),
 *   rujukan sel fisik Excel (provenance), selisih nilai, dan saran investigasi.
 * - Sepenuhnya independen dan tidak mengasumsikan format proyek tertentu.
 */

(function (window) {
  'use strict';

  var UniversalChecker = {
    /**
     * Menjalankan audit menyeluruh terhadap data kanonikal.
     * @param {Object} canonicalData
     * @param {Object} [options]
     * @returns {Object} Hasil audit lengkap
     */
    runAudit: function (canonicalData, options) {
      if (!canonicalData || !canonicalData.workItems) {
        return {
          totalFindings: 0,
          summary: { error: 0, warning: 0, info: 0 },
          findings: [],
          auditedAt: new Date().toISOString()
        };
      }

      options = options || {};
      var tolerance = options.tolerance !== undefined ? options.tolerance : 0.01;

      var items = canonicalData.workItems || [];
      var subtotals = canonicalData.subtotals || [];
      var findings = [];
      var findingCounter = 1;

      function addFinding(ruleId, severity, explanation, origVal, expVal, diff, sourceSheet, sourceCell, sourceRow, action, certainty, title, itemDescription, relatedItems) {
        var cert = certainty;
        if (!cert) {
          if (ruleId === 'amount_differs_qty_price' || ruleId === 'broken_formula_reference' || ruleId === 'missing_amount') {
            cert = 'MATHEMATICAL_PROOF';
          } else if (ruleId === 'negative_values' || ruleId === 'duplicate_item_codes' || ruleId === 'inconsistent_unit_for_code' || ruleId === 'inconsistent_unit_price' || ruleId === 'invalid_numeric_values') {
            cert = 'PATTERN_ANOMALY';
          } else {
            cert = 'CONTEXT_REQUIRED';
          }
        }

        var certLabel = 'Bukti Matematis';
        if (cert === 'PATTERN_ANOMALY') certLabel = 'Anomali Pola';
        else if (cert === 'CONTEXT_REQUIRED') certLabel = 'Perlu Konfirmasi';
        else if (cert === 'EXTERNAL_COMPARISON') certLabel = 'Pembanding Luar';

        var finding = {
          id: 'UF-' + (findingCounter++),
          ruleId: ruleId,
          severity: severity, // 'ERROR' | 'WARNING' | 'INFO'
          certainty: cert,
          certaintyLabel: certLabel,
          title: title || null,
          itemDescription: itemDescription || null,
          explanation: explanation,
          originalValue: origVal !== undefined ? origVal : null,
          expectedValue: expVal !== undefined ? expVal : null,
          difference: diff !== undefined ? diff : null,
          sourceSheet: sourceSheet || canonicalData.project.activeSheet || 'RAB',
          sourceCell: sourceCell || ('Baris ' + sourceRow),
          sourceRow: sourceRow || 0,
          recommendedAction: action || 'Periksa kembali sel Excel terkait.'
        };
        if (Array.isArray(relatedItems)) {
          finding.relatedItems = relatedItems;
          // Retain an easy-to-consume path collection for clients that do not
          // need the rest of each source item's context.
          finding.categoryPaths = relatedItems.map(function (item) { return item.categoryPath; });
        }
        findings.push(finding);
      }

      function buildRelatedItemContexts(group) {
        return group.map(function (item) {
          var source = item.source || {};
          var cells = source.cells || {};
          var codeCell = cells.item_code && cells.item_code.cellRef
            ? source.sheet + '!' + cells.item_code.cellRef
            : (source.itemCodeCell || null);
          var path = Array.isArray(item.categoryPath) ? item.categoryPath : [];
          var pathCopy;
          try { pathCopy = JSON.parse(JSON.stringify(path)); } catch (ignoreNonSerializablePath) { pathCopy = []; }
          return {
            itemCode: item.code || item.analysisCode || null,
            description: item.description || null,
            sheet: source.sheet || null,
            sourceRow: source.row || null,
            sourceCell: codeCell,
            categoryPath: pathCopy,
            hierarchyStatus: item.hierarchy && item.hierarchy.status || item.hierarchyStatus || (pathCopy.length ? 'unresolved' : 'none'),
            categoryProvenance: pathCopy.map(function (level) {
              return { levelIndex: level.levelIndex, provenance: level.provenance || null };
            })
          };
        });
      }

      var codeMap = {};
      var descMap = {};

      // -------------------------------------------------------------
      // 1. Audit Per-Item Pekerjaan (WorkItem)
      // -------------------------------------------------------------
      for (var i = 0; i < items.length; i++) {
        var it = items[i];
        var s = it.source || {};
        var row = s.row || (i + 1);
        var sheetName = s.sheet || canonicalData.project.activeSheet;
        var amountCell = s.amountCell || (sheetName + '!KolomJumlah' + row);
        var cells = s.cells || {};
        function hasExcelError(field) {
          var cellInfo = cells[field];
          return !!(cellInfo && (cellInfo.cellType === 'e' || cellInfo.valueStatus === 'error'));
        }

        // R-01: missing_description
        if (!it.description || !it.description.trim()) {
          var descCell = cells.description ? (sheetName + '!' + cells.description.cellRef) : ('Baris ' + row);
          addFinding(
            'missing_description',
            'ERROR',
            'Item pekerjaan tidak memiliki uraian/deskripsi nama pekerjaan.',
            '',
            'Teks Uraian Pekerjaan',
            null,
            sheetName,
            descCell,
            row,
            'Isi deskripsi pekerjaan agar ruang lingkup biaya terdefinisi dengan jelas.'
          );
        }

        // R-02: missing_unit
        if ((!it.unit || !it.unit.trim()) && it.volume !== null && it.volume > 0) {
          var unitCell = cells.unit ? (sheetName + '!' + cells.unit.cellRef) : ('Baris ' + row);
          addFinding(
            'missing_unit',
            'WARNING',
            'Volume pekerjaan terisi (' + it.volume + ') namun satuan pengukuran kosong.',
            '',
            'Satuan (m1, m2, m3, kg, dll)',
            null,
            sheetName,
            unitCell,
            row,
            'Lengkapi satuan pekerjaan untuk menghindari multitafsir dasar perhitungan harga.'
          );
        }

        // R-03: missing_quantity
        if (!hasExcelError('volume') && (it.volume === null || it.volume === undefined) && it.unitPrice !== null && it.unitPrice > 0) {
          var volCell = cells.volume ? (sheetName + '!' + cells.volume.cellRef) : ('Baris ' + row);
          addFinding(
            'missing_quantity',
            'WARNING',
            'Harga satuan terisi (' + it.unitPrice + ') namun volume kuantitas kosong.',
            null,
            'Volume > 0',
            null,
            sheetName,
            volCell,
            row,
            'Tentukan volume rencana pekerjaan untuk item ini.'
          );
        }

        // R-04: missing_unit_price
        if (!hasExcelError('unit_price') && (it.unitPrice === null || it.unitPrice === undefined) && it.volume !== null && it.volume > 0) {
          var priceCell = cells.unit_price ? (sheetName + '!' + cells.unit_price.cellRef) : ('Baris ' + row);
          addFinding(
            'missing_unit_price',
            'WARNING',
            'Volume pekerjaan terisi (' + it.volume + ') namun harga satuan kosong.',
            null,
            'Harga Satuan (Rp)',
            null,
            sheetName,
            priceCell,
            row,
            'Lengkapi harga satuan pekerjaan melalui referensi AHSP atau input manual.'
          );
        }

        // R-05: missing_amount
        if (!hasExcelError('amount') && (it.amount === null || it.amount === undefined) && it.volume !== null && it.unitPrice !== null) {
          var expAmt = it.volume * it.unitPrice;
          addFinding(
            'missing_amount',
            'ERROR',
            'Volume (' + it.volume + ') dan Harga Satuan (' + it.unitPrice + ') terisi, namun kolom Jumlah kosong.',
            null,
            expAmt,
            expAmt,
            sheetName,
            amountCell,
            row,
            'Terapkan formula perkalian volume x harga satuan pada sel jumlah harga.'
          );
        }

        // R-08: invalid_numeric_values & R-09: negative_values
        ['volume', 'unit_price', 'amount'].forEach(function (fKey) {
          var val = it[fKey === 'unit_price' ? 'unitPrice' : fKey];
          var cInfo = cells[fKey];
          var cRef = cInfo ? (sheetName + '!' + cInfo.cellRef) : ('Baris ' + row);

          if (hasExcelError(fKey)) {
            var errorText = cInfo.errorText || '#ERROR!';
            var errorCodeText = cInfo.errorCode !== null && cInfo.errorCode !== undefined
              ? ' Kode error Excel: ' + String(cInfo.errorCode) + '.'
              : '';
            var formulaText = cInfo.formula ? ' Formula sumber: ' + String(cInfo.formula) + '.' : '';
            var fieldLabel = fKey === 'unit_price' ? 'unit_price (harga satuan)' : fKey;
            var errorTitle = 'Error Excel pada ' + fieldLabel;
            var errorAction = 'Periksa dan perbaiki formula atau nilai sumber pada sel ' + cRef + ', lalu impor ulang workbook.';
            addFinding(
              'excel_error_' + fKey,
              'ERROR',
              'Sel ' + cRef + ' pada kolom ' + fieldLabel + ' mengandung ' + errorText +
                '.' + errorCodeText + formulaText +
                ' Nilai error Excel tidak dapat dipakai sebagai nilai numerik. Item: ' + (it.description || '(tanpa uraian)') + '.',
              errorText,
              null,
              null,
              sheetName,
              cRef,
              row,
              errorAction,
              'CONTEXT_REQUIRED',
              errorTitle,
              it.description || null
            );
          } else if (cInfo && typeof cInfo.val === 'string' && /#value!|#div\/0!|#ref!|#name\?/i.test(cInfo.val)) {
            addFinding(
              'invalid_numeric_values',
              'ERROR',
              'Sel memuat teks kesalahan formula Excel: ' + cInfo.val,
              cInfo.val,
              'Angka Valid',
              null,
              sheetName,
              cRef,
              row,
              'Perbaiki rumus formula sel Excel yang menghasilkan error.'
            );
          } else if (val !== null && val < 0) {
            addFinding(
              'negative_values',
              'ERROR',
              'Nilai kolom ' + fKey + ' bernilai negatif (' + val + '). Nilai RAB tidak boleh negatif.',
              val,
              '>= 0',
              val,
              sheetName,
              cRef,
              row,
              'Periksa apakah ada kesalahan tanda minus atau formula pengurangan yang keliru.'
            );
          }
        });

        // R-11: amount_differs_qty_price (Perkalian Matematis)
        if (typeof it.volume === 'number' && typeof it.unitPrice === 'number' && typeof it.amount === 'number') {
          var expectedAmount = it.volume * it.unitPrice;
          var diff = Math.abs(it.amount - expectedAmount);
          if (diff > tolerance) {
            addFinding(
              'amount_differs_qty_price',
              'ERROR',
              'Jumlah harga Excel (' + it.amount + ') tidak sama dengan Volume (' + it.volume + ') x Harga Satuan (' + it.unitPrice + '). Selisih: ' + (it.amount - expectedAmount).toFixed(2),
              it.amount,
              expectedAmount,
              it.amount - expectedAmount,
              sheetName,
              amountCell,
              row,
              'Sinkronkan formula sel Jumlah agar mengalikan kolom Volume dan Harga Satuan secara konsisten.'
            );
          }
        }

        // R-14: broken_formula_reference
        if (s.amountFormula && /#ref!|#name\?/i.test(s.amountFormula)) {
          addFinding(
            'broken_formula_reference',
            'ERROR',
            'Rumus sel Jumlah memuat referensi rusak: ' + s.amountFormula,
            s.amountFormula,
            'Formula Rujukan Valid',
            null,
            sheetName,
            amountCell,
            row,
            'Hubungkan kembali rujukan sel formula yang terputus atau terhapus.'
          );
        }

        // R-17: suspicious_extreme_quantity
        if (typeof it.volume === 'number' && (it.volume > 1000000 || (it.volume > 0 && it.volume < 0.0001))) {
          var vRef = cells.volume ? (sheetName + '!' + cells.volume.cellRef) : ('Baris ' + row);
          addFinding(
            'suspicious_extreme_quantity',
            'INFO',
            'Volume kuantitas bernilai ekstrem (' + it.volume + ').',
            it.volume,
            'Pemeriksaan Kewajaran',
            null,
            sheetName,
            vRef,
            row,
            'Verifikasi apakah satuan pekerjaan atau angka volume telah sesuai dengan gambar rencana.'
          );
        }

        // R-18: suspicious_zero_price
        if (typeof it.volume === 'number' && it.volume > 0 && it.unitPrice === 0) {
          var pRef = cells.unit_price ? (sheetName + '!' + cells.unit_price.cellRef) : ('Baris ' + row);
          addFinding(
            'suspicious_zero_price',
            'WARNING',
            'Pekerjaan memiliki volume (' + it.volume + ') namun harga satuan bernilai Rp0.',
            0,
            'Harga > Rp0',
            null,
            sheetName,
            pRef,
            row,
            'Pastikan apakah pekerjaan ini sengaja digratiskan / include dalam item lain atau belum ditentukan harganya.'
          );
        }

        // Kumpulkan data untuk cek konsistensi kode dan uraian
        if (it.code) {
          if (!codeMap[it.code]) codeMap[it.code] = [];
          codeMap[it.code].push(it);
        }
        if (it.description && it.description.length > 5) {
          var normDesc = it.description.toLowerCase().trim();
          if (!descMap[normDesc]) descMap[normDesc] = [];
          descMap[normDesc].push(it);
        }
      }

      // -------------------------------------------------------------
      // 2. Pemeriksaan Duplikasi & Konsistensi Lintas Baris
      // -------------------------------------------------------------
      // R-06 & R-15: Kode item duplikat atau satuan tidak konsisten
      for (var code in codeMap) {
        var group = codeMap[code];
        if (group.length > 1) {
          var firstUnit = group[0].unit;
          var unitMismatch = false;
          for (var g = 1; g < group.length; g++) {
            if (group[g].unit && firstUnit && group[g].unit.toLowerCase() !== firstUnit.toLowerCase()) {
              unitMismatch = true;
              break;
            }
          }

          if (unitMismatch) {
            addFinding(
              'inconsistent_unit_for_code',
              'WARNING',
              'Kode item "' + code + '" digunakan pada beberapa baris dengan satuan yang berbeda.',
              group.map(function(x){ return x.unit; }).join(', '),
              firstUnit,
              null,
              group[0].source.sheet,
              group.map(function (x) { return (x.source && x.source.cells && x.source.cells.item_code ? x.source.sheet + '!' + x.source.cells.item_code.cellRef : (x.source.sheet + '!Baris' + x.source.row)); }).join(', '),
              group[0].source.row,
              'Samakan satuan pengukuran untuk kode item pekerjaan yang sama.',
              null,
              null,
              null,
              buildRelatedItemContexts(group)
            );
          } else {
            addFinding(
              'duplicate_item_codes',
              'INFO',
              'Kode item "' + code + '" muncul lebih dari 1 kali (' + group.length + ' baris).',
              group.length + ' kali',
              'Kode Unik',
              null,
              group[0].source.sheet,
              group.map(function (x) { return (x.source && x.source.cells && x.source.cells.item_code ? x.source.sheet + '!' + x.source.cells.item_code.cellRef : (x.source.sheet + '!Baris' + x.source.row)); }).join(', '),
              group[0].source.row,
              'Pastikan apakah pengulangan kode disengaja untuk sub-pekerjaan yang berbeda.',
              null,
              null,
              null,
              buildRelatedItemContexts(group)
            );
          }
        }
      }

      // R-16: inconsistent_unit_price (Uraian sama persis namun harga beda signifikan)
      for (var d in descMap) {
        var dGroup = descMap[d];
        if (dGroup.length > 1) {
          var prices = dGroup.map(function (x) { return x.unitPrice; }).filter(function (p) { return typeof p === 'number' && p > 0; });
          if (prices.length > 1) {
            var minP = Math.min.apply(null, prices);
            var maxP = Math.max.apply(null, prices);
            if (minP > 0 && ((maxP - minP) / minP) > 0.50) {
              addFinding(
                'inconsistent_unit_price',
                'WARNING',
                'Uraian pekerjaan serupa memiliki perbedaan harga satuan yang signifikan (>50%): Min Rp' + minP.toLocaleString('id-ID') + ' vs Max Rp' + maxP.toLocaleString('id-ID'),
                'Rp' + maxP.toLocaleString('id-ID'),
                'Rp' + minP.toLocaleString('id-ID'),
                maxP - minP,
                dGroup[0].source.sheet,
                dGroup.map(function (x) { return 'Baris ' + x.source.row; }).join(', '),
                dGroup[0].source.row,
                'Evaluasi apakah disparitas harga satuan pada uraian pekerjaan yang sama ini wajar.'
              );
            }
          }
        }
      }

      // -------------------------------------------------------------
      // 3. Ringkasan Tingkat Keparahan
      // -------------------------------------------------------------
      var summary = { error: 0, warning: 0, info: 0 };
      for (var k = 0; k < findings.length; k++) {
        var sev = findings[k].severity.toLowerCase();
        if (summary[sev] !== undefined) summary[sev]++;
      }

      return {
        totalFindings: findings.length,
        summary: summary,
        findings: findings,
        auditedAt: new Date().toISOString()
      };
    }
  };

  window.UniversalChecker = UniversalChecker;

})(typeof window !== 'undefined' ? window : this);
