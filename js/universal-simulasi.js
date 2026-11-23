/**
 * RAB GUARD — Universal XLSX Simulation & Calculation Engine
 * Module: UniversalSimulasi (js/universal-simulasi.js)
 * 
 * Tanggung Jawab:
 * - Menghitung ulang nilai proyek secara independen dari model data kanonikal.
 * - Menjalankan simulasi eskalasi biaya secara fungsional murni (tanpa memutasikan data asli).
 * - Menghasilkan analisis Pareto kontributor biaya terbesar (kumulatif 80%).
 * - Mengekspor temuan audit dan data kanonikal ke berkas CSV bertanda tangan UTF-8 BOM.
 */

(function (window) {
  'use strict';

  var UniversalSimulasi = {
    /**
     * Menghitung ulang proyek secara independen dari data kanonikal.
     * @param {Object} canonicalData
     * @returns {Object}
     */
    recalculateBaseline: function (canonicalData) {
      var invalid = {
        totalExcel: null, totalImported: null, totalCalculated: null, totalRecalculated: null,
        totalBaseline: null, totalImportedComparable: null, totalRecalculatedComparable: null,
        difference: null, comparableDifference: null, itemsCount: 0,
        calculableItemsCount: 0, uncalculableItemsCount: 0, comparableItemsCount: 0,
        importedAmountItemsCount: 0, coveragePercent: 0,
        isComplete: false, coverageComplete: false, isValid: false, items: [], discrepancies: []
      };
      if (!canonicalData || !Array.isArray(canonicalData.workItems)) return invalid;

      var items = canonicalData.workItems;
      var totalImported = 0;
      var totalRecalculated = 0;
      var totalImportedComparable = 0;
      var totalRecalculatedComparable = 0;
      var calculableItemsCount = 0;
      var comparableItemsCount = 0;
      var importedAmountItemsCount = 0;
      var uncalculableItemsCount = 0;
      var itemResults = [];
      var discrepancies = [];
      function validNumber(value) { return typeof value === 'number' && isFinite(value); }

      for (var i = 0; i < items.length; i++) {
        var it = items[i];
        var hasImported = validNumber(it.amount);
        var product = validNumber(it.volume) && validNumber(it.unitPrice) ? it.volume * it.unitPrice : null;
        var calculable = validNumber(product);
        var recalculated = calculable ? product : null;
        if (hasImported) {
          totalImported += it.amount;
          importedAmountItemsCount++;
        }
        if (calculable) {
          calculableItemsCount++;
          totalRecalculated += recalculated;
          if (hasImported) {
            totalImportedComparable += it.amount;
            totalRecalculatedComparable += recalculated;
            comparableItemsCount++;
          }
        } else {
          uncalculableItemsCount++;
        }
        var result = {
          id: it.id, itemNumber: it.itemNumber, code: it.code, analysisCode: it.analysisCode,
          description: it.description, source: it.source || null,
          importedAmount: hasImported ? it.amount : null, recalculatedAmount: recalculated,
          isCalculable: calculable,
          discrepancy: calculable && hasImported ? recalculated - it.amount : null
        };
        itemResults.push(result);
        if (result.discrepancy !== null && Math.abs(result.discrepancy) > 0.01) discrepancies.push(result);
      }

      var count = items.length;
      var complete = uncalculableItemsCount === 0;
      var hasAggregateComparison = count === 0 ||
        (calculableItemsCount > 0 && importedAmountItemsCount > 0);
      return {
        totalExcel: totalImported, totalImported: totalImported,
        totalCalculated: totalRecalculated, totalRecalculated: totalRecalculated,
        totalBaseline: totalRecalculated, totalImportedComparable: totalImportedComparable,
        totalRecalculatedComparable: totalRecalculatedComparable,
        difference: hasAggregateComparison ? totalRecalculated - totalImported : null,
        comparableDifference: comparableItemsCount > 0
          ? totalRecalculatedComparable - totalImportedComparable : null,
        itemsCount: count, calculableItemsCount: calculableItemsCount,
        uncalculableItemsCount: uncalculableItemsCount,
        comparableItemsCount: comparableItemsCount, importedAmountItemsCount: importedAmountItemsCount,
        coveragePercent: count === 0 ? 100 : (calculableItemsCount / count) * 100,
        isComplete: complete, coverageComplete: complete, isValid: true,
        items: itemResults, discrepancies: discrepancies
      };
    },

    /**
     * Menjalankan simulasi eskalasi harga.
     * @param {Object} canonicalData Model data kanonikal (tidak akan dimutasi)
     * @param {Object} params { percentAll: number, categoryFactors?: Object }
     * @returns {Object} Hasil simulasi lengkap
     */
    simulate: function (canonicalData, params) {
      if (!canonicalData || !Array.isArray(canonicalData.workItems)) {
        throw new Error('Data kanonikal tidak tersedia untuk disimulasikan.');
      }
      params = params || {};
      var percentAll = parseFloat(params.percentAll) || 0;
      var catFactors = params.categoryFactors || {};
      var items = canonicalData.workItems;
      var simulatedItems = [];
      var totalBaseline = 0;
      var totalSimulated = 0;
      var totalRecalculatedBaseline = 0;
      var totalRecalculatedScenario = 0;
      var totalUnadjustedImported = 0;
      var calculableItemsCount = 0;
      var uncalculableItemsCount = 0;
      var categoryTotals = {};
      function validNumber(value) { return typeof value === 'number' && isFinite(value); }

      for (var i = 0; i < items.length; i++) {
        var orig = items[i];
        var cat = orig.category || 'UMUM';
        var factor = 1 + (percentAll / 100);
        if (catFactors[cat] !== undefined) factor = 1 + (parseFloat(catFactors[cat]) / 100);
        var product = validNumber(orig.volume) && validNumber(orig.unitPrice) ? orig.volume * orig.unitPrice : null;
        var calculable = validNumber(product);
        var recalculated = calculable ? product : null;
        var scenarioAmount = calculable ? recalculated * factor : null;
        var hasImported = validNumber(orig.amount);
        var preserved = !calculable && hasImported ? orig.amount : null;
        var baseAmount = calculable ? recalculated : preserved;
        var simAmount = calculable ? scenarioAmount : preserved;
        if (calculable) {
          calculableItemsCount++;
          totalRecalculatedBaseline += recalculated;
          totalRecalculatedScenario += scenarioAmount;
        } else {
          uncalculableItemsCount++;
          if (hasImported) totalUnadjustedImported += orig.amount;
        }
        if (baseAmount !== null) totalBaseline += baseAmount;
        if (simAmount !== null) totalSimulated += simAmount;
        if (!categoryTotals[cat]) {
          categoryTotals[cat] = { category: cat, baselineAmount: 0, simulatedAmount: 0 };
        }
        if (baseAmount !== null) categoryTotals[cat].baselineAmount += baseAmount;
        if (simAmount !== null) categoryTotals[cat].simulatedAmount += simAmount;
        simulatedItems.push({
          id: orig.id, itemNumber: orig.itemNumber, code: orig.code,
          analysisCode: orig.analysisCode, description: orig.description, category: cat,
          volume: orig.volume, unit: orig.unit, basePrice: orig.unitPrice,
          simPrice: calculable ? orig.unitPrice * factor : null,
          importedAmount: hasImported ? orig.amount : null,
          recalculatedAmount: recalculated, baseAmount: baseAmount, simAmount: simAmount,
          deltaAmount: calculable ? scenarioAmount - recalculated : 0,
          isCalculable: calculable, isRecalculated: calculable,
          isScenarioAdjusted: calculable, isUnadjusted: !calculable,
          factorApplied: calculable ? factor : null, source: orig.source || null
        });
      }

      var deltaTotal = totalSimulated - totalBaseline;
      var deltaPercent = totalBaseline !== 0 ? (deltaTotal / Math.abs(totalBaseline)) * 100 : 0;
      var sortedByCost = simulatedItems.filter(function (item) {
        return validNumber(item.simAmount);
      }).sort(function (a, b) {
        return b.simAmount - a.simAmount;
      });
      var runningParetoSum = 0;
      var paretoTop10 = [];
      var limit = Math.min(10, sortedByCost.length);
      for (var p = 0; p < limit; p++) {
        var pItem = sortedByCost[p];
        runningParetoSum += pItem.simAmount;
        paretoTop10.push({
          rank: p + 1, code: pItem.code, description: pItem.description,
          category: pItem.category, simAmount: pItem.simAmount,
          isScenarioAdjusted: pItem.isScenarioAdjusted,
          sharePercent: totalSimulated !== 0 ? (pItem.simAmount / totalSimulated) * 100 : 0,
          cumulativePercent: totalSimulated !== 0 ? (runningParetoSum / totalSimulated) * 100 : 0
        });
      }
      var categorySummary = [];
      for (var c in categoryTotals) {
        var cTot = categoryTotals[c];
        categorySummary.push({
          category: cTot.category, baselineAmount: cTot.baselineAmount, simulatedAmount: cTot.simulatedAmount,
          deltaAmount: cTot.simulatedAmount - cTot.baselineAmount,
          deltaPercent: cTot.baselineAmount !== 0
            ? ((cTot.simulatedAmount - cTot.baselineAmount) / Math.abs(cTot.baselineAmount)) * 100 : 0
        });
      }
      var isPartial = uncalculableItemsCount > 0;
      return {
        totalBaseline: totalBaseline, totalSimulated: totalSimulated,
        totalRecalculatedBaseline: totalRecalculatedBaseline,
        totalRecalculatedScenario: totalRecalculatedScenario,
        totalUnadjustedImported: totalUnadjustedImported,
        deltaTotal: deltaTotal, deltaPercent: deltaPercent,
        calculableItemsCount: calculableItemsCount, uncalculableItemsCount: uncalculableItemsCount,
        itemsCount: items.length, coveragePercent: items.length === 0 ? 100 : calculableItemsCount / items.length * 100,
        isComplete: !isPartial, isPartial: isPartial, percentAll: percentAll,
        categories: categorySummary, items: simulatedItems, paretoTop10: paretoTop10,
        simulatedAt: new Date().toISOString()
      };
    },

    /**
     * Mengekspor temuan audit universal ke CSV.
     */
    buildAuditCSV: function (auditResult, canonicalData, projectName) {
      var findings = auditResult && Array.isArray(auditResult.findings) ? auditResult.findings : [];
      var project = canonicalData && canonicalData.project ? canonicalData.project : {};
      var workbookName = project.sourceFilename || projectName || '';
      var selectedSheet = project.activeSheet || '';
      var rows = [
        ['Workbook', workbookName],
        ['Worksheet', selectedSheet],
        [],
        ['Finding ID', 'Rule ID', 'Severity', 'Title', 'Item description', 'Worksheet', 'Source cell',
          'Explanation', 'Original source value', 'Calculated value', 'Expected / guidance value',
          'Difference', 'Difference convention', 'Formula', 'Excel error code', 'Excel error text', 'Recommended action']
      ];

      function cell(value) {
        if (value === null || value === undefined) return '""';
        var text = String(value);
        if (typeof value !== 'number' && /^[\u0000-\u0020]*[=+\-@]/.test(text)) text = "'" + text;
        return '"' + text.replace(/"/g, '""') + '"';
      }

      for (var i = 0; i < findings.length; i++) {
        var f = findings[i] || {};
        var expectedIsNumeric = typeof f.expectedValue === 'number' && isFinite(f.expectedValue);
        var differenceIsNumeric = typeof f.difference === 'number' && isFinite(f.difference);
        var differenceConvention = f.ruleId === 'amount_differs_qty_price'
          ? 'Sumber − hitung ulang' : '';
        rows.push([
          f.id, f.ruleId, f.severity, f.title, f.itemDescription, f.sourceSheet, f.sourceCell,
          f.explanation, f.originalValue,
          expectedIsNumeric ? f.expectedValue : null,
          f.expectedValue,
          differenceIsNumeric ? f.difference : null,
          differenceConvention,
          f.formula, f.errorCode, f.errorText, f.recommendedAction
        ]);
      }

      return '\uFEFF' + rows.map(function (row) {
        return row.map(cell).join(',');
      }).join('\r\n') + '\r\n';
    },

    /** Build one CSV row for every finding-to-item relationship. */
    buildRelatedItemsCSV: function (auditResult) {
      var findings = auditResult && Array.isArray(auditResult.findings) ? auditResult.findings : [];
      var rows = [[
        'Finding reference', 'Finding index', 'Finding rule', 'Finding severity', 'Finding message',
        'Item code', 'Description', 'Sheet', 'Source row', 'Source cell', 'Category path (JSON)',
        'Hierarchy status', 'Category provenance (JSON)'
      ]];

      function serialize(value) {
        if (value === null || value === undefined) return '';
        if (typeof value === 'string') return value;
        try { return JSON.stringify(value); } catch (ignoreSerializationError) { return ''; }
      }
      function csvCell(value) {
        if (value === null || value === undefined) return '""';
        if (typeof value === 'number' && isFinite(value)) return String(value);
        var text = serialize(value);
        if (/^[\u0000-\u0020\uFEFF]*[=+\-@]/.test(text)) text = "'" + text;
        return '"' + text.replace(/"/g, '""') + '"';
      }

      findings.forEach(function (finding, findingIndex) {
        finding = finding || {};
        var related = Array.isArray(finding.relatedItems) ? finding.relatedItems : [];
        related.forEach(function (item) {
          item = item || {};
          var path = Array.isArray(item.categoryPath) ? item.categoryPath : [];
          var provenance = Array.isArray(item.categoryProvenance) ? item.categoryProvenance :
            path.map(function (level) { return { levelIndex: level.levelIndex, provenance: level.provenance || null }; });
          rows.push([
            finding.id || 'Finding ' + (findingIndex + 1), findingIndex + 1, finding.ruleId, finding.severity,
            finding.title ? finding.title + (finding.explanation ? ' — ' + finding.explanation : '') : finding.explanation,
            item.itemCode, item.description, item.sheet, item.sourceRow, item.sourceCell,
            JSON.stringify(path), item.hierarchyStatus || (path.length ? 'unresolved' : 'none'), JSON.stringify(provenance)
          ]);
        });
      });

      return '\uFEFF' + rows.map(function (row) { return row.map(csvCell).join(','); }).join('\r\n') + '\r\n';
    },

    exportRelatedItemsCSV: function (auditResult, projectName, canonicalData) {
      if (projectName && typeof projectName === 'object' && !canonicalData) {
        canonicalData = projectName;
        projectName = null;
      }
      var csvContent = this.buildRelatedItemsCSV(auditResult);
      var workbookName = canonicalData && canonicalData.project && canonicalData.project.sourceFilename;
      var filenameBase = (workbookName || projectName || 'Universal').replace(/\.[^/.]+$/, '');
      this.downloadCSVContent(csvContent, 'Audit_RAB_Item_Terkait_' + filenameBase + '.csv');
      return csvContent;
    },

    exportAuditCSV: function (auditResult, projectName, canonicalData) {
      if (projectName && typeof projectName === 'object' && !canonicalData) {
        canonicalData = projectName;
        projectName = null;
      }
      var csvContent = this.buildAuditCSV(auditResult, canonicalData, projectName);
      var workbookName = canonicalData && canonicalData.project && canonicalData.project.sourceFilename;
      var filenameBase = (workbookName || projectName || 'Universal').replace(/\.[^/.]+$/, '');
      this.downloadCSVContent(csvContent, 'Audit_RAB_' + filenameBase + '.csv');
      return csvContent;
    },

    /**
     * Mengekspor data kanonikal hasil normalisasi ke CSV.
     */
    exportCanonicalCSV: function (canonicalData) {
      if (!canonicalData || !canonicalData.workItems || canonicalData.workItems.length === 0) {
        alert('Tidak ada data kanonikal untuk diekspor.');
        return;
      }

      var items = canonicalData.workItems;
      var rows = [
        ['ID', 'No. Item', 'Kode Pekerjaan', 'Kode Analisa / AHSP', 'Divisi / Kategori', 'Uraian Pekerjaan', 'Satuan', 'Volume', 'Harga Satuan (Rp)', 'Jumlah Harga (Rp)', 'Rujukan Sel Excel']
      ];

      for (var i = 0; i < items.length; i++) {
        var it = items[i];
        rows.push([
          it.id,
          it.itemNumber || '',
          it.code || '',
          it.analysisCode || '',
          it.category || '',
          it.description || '',
          it.unit || '',
          it.volume !== null ? String(it.volume) : '',
          it.unitPrice !== null ? String(it.unitPrice) : '',
          it.amount !== null ? String(it.amount) : '',
          it.source ? (it.source.amountCell || (it.source.sheet + '!Baris' + it.source.row)) : ''
        ]);
      }

      this.downloadCSV(rows, 'Kanonikal_RAB_' + (canonicalData.project.projectName || 'Proyek') + '.csv');
    },

    /**
     * Membentuk Blob UTF-8 BOM dan memicu download browser secara offline.
     */
    downloadCSV: function (rows, filename) {
      var csvContent = '\uFEFF'; // UTF-8 BOM
      for (var r = 0; r < rows.length; r++) {
        var row = rows[r];
        var escaped = [];
        for (var c = 0; c < row.length; c++) {
          var val = (row[c] !== undefined && row[c] !== null) ? String(row[c]) : '';
          val = val.replace(/"/g, '""');
          if (val.search(/([";\n\r])/g) >= 0) {
            val = '"' + val + '"';
          }
          escaped.push(val);
        }
        csvContent += escaped.join(';') + '\r\n';
      }

      var blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
      var url = URL.createObjectURL(blob);
      var link = document.createElement('a');
      link.setAttribute('href', url);
      link.setAttribute('download', filename);
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);
    },

    downloadCSVContent: function (csvContent, filename) {
      var blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
      var url = URL.createObjectURL(blob);
      var link = document.createElement('a');
      link.setAttribute('href', url);
      link.setAttribute('download', filename);
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      URL.revokeObjectURL(url);
    }
  };

  window.UniversalSimulasi = UniversalSimulasi;

})(typeof window !== 'undefined' ? window : this);
