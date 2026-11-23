/**
 * RAB GUARD — Universal XLSX Ingestion
 * Module: DataNormalizer (js/ingestion/data-normalizer.js)
 * 
 * Tanggung Jawab:
 * - Mengubah lembar kerja fisik Excel menjadi Model Data Kanonikal internal yang modular.
 * - Memisahkan baris menjadi WorkItem, Heading (judul divisi), dan Subtotal secara cerdas.
 * - Menjaga provenance asal sel (sheet, baris 1-based, huruf kolom, formula asli).
 * - Tidak mengarang atau mengasumsikan nilai yang hilang (missing/null/NaN ditandai eksplisit).
 */

(function (window) {
  'use strict';

  var DataNormalizer = {
    /**
     * Build one canonical model from every confidently recognized BOQ sheet.
     * Sheets with uncertain roles, headers, or mappings stay in the manifest
     * with an explicit unresolved status and are excluded from totals.
     */
    normalizeWorkbook: function (rawWb, inspection) {
      if (!rawWb || !rawWb.sheets || !window.WorkbookInspector || !window.HeaderDetector || !window.ColumnMapper) {
        throw new Error('Komponen inspeksi workbook tidak tersedia.');
      }
      inspection = inspection || window.WorkbookInspector.inspect(rawWb);
      var manifest = [];
      var workItems = [];
      var headings = [];
      var subtotals = [];
      var excludedRows = [];
      var totalParsedAmount = 0;
      var processedBoqSheets = [];
      var supportingRoles = { RECAP: true, AHSP_PRICE_REFERENCE: true, SCHEDULE: true, VOLUME_BACKUP: true, DASHBOARD: true, PROJECT_ASSUMPTIONS: true, SUPPORTING: true };
      var sheetMetadata = inspection.sheets || [];

      sheetMetadata.forEach(function (meta, sheetIndex) {
        var entry = {
          name: meta.sheetName,
          index: sheetIndex,
          hidden: !!meta.isHidden,
          usedRange: meta.range || null,
          rowCount: meta.rowCount || 0,
          columnCount: meta.colCount || 0,
          role: meta.semanticRole || meta.detectedRole || 'OTHER',
          detectedRole: meta.detectedRole || 'OTHER',
          confidence: meta.semanticConfidence !== undefined ? meta.semanticConfidence : meta.confidence,
          evidence: (meta.clues || []).slice(),
          headerRow: null,
          headerRowEnd: null,
          mapping: null,
          status: 'supporting',
          unresolvedReasons: []
        };
        var sheet = rawWb.sheets[meta.sheetName];
        if (!sheet) {
          entry.status = 'unresolved';
          entry.unresolvedReasons.push('Sheet tidak tersedia pada data workbook yang dibaca.');
        } else if (entry.role === 'PRIMARY_BOQ' && Number(entry.confidence) >= 70) {
          var detection = window.HeaderDetector.detect(sheet);
          entry.headerRow = detection.found ? detection.headerRow1Based : null;
          entry.headerRowEnd = detection.found ? detection.headerRowEnd1Based : null;
          var mapping = window.ColumnMapper.createInitialMapping(detection);
          entry.mapping = mapping;
          if (!detection.found || detection.ambiguous || detection.confidence < 70) {
            entry.status = 'unresolved';
            entry.unresolvedReasons.push(!detection.found ? 'Baris header tidak terdeteksi.' :
              detection.ambiguous ? (detection.ambiguityReason || 'Beberapa kandidat header sama-sama mungkin.') : 'Keyakinan header di bawah ambang proses otomatis.');
          } else {
            var validation = window.ColumnMapper.validateMapping(mapping);
            if (!validation.valid) {
              entry.status = 'unresolved';
              entry.unresolvedReasons = validation.errors.slice();
            } else {
              try {
                var part = this.normalize(rawWb, meta.sheetName, mapping);
                entry.status = 'processed';
                processedBoqSheets.push(meta.sheetName);
                part.workItems.forEach(function (item) {
                  item.id = meta.sheetName + ':' + item.id;
                  item.source.sheetIndex = sheetIndex;
                  workItems.push(item);
                });
                part.headings.forEach(function (heading) {
                  heading.id = meta.sheetName + ':' + heading.id;
                  heading.source.sheetIndex = sheetIndex;
                  headings.push(heading);
                });
                part.subtotals.forEach(function (subtotal) {
                  subtotal.id = meta.sheetName + ':' + subtotal.id;
                  subtotal.source.sheetIndex = sheetIndex;
                  subtotals.push(subtotal);
                });
                part.excludedRows.forEach(function (row) {
                  row.source.sheetIndex = sheetIndex;
                  excludedRows.push(row);
                });
                totalParsedAmount += part.statistics.totalParsedAmount;
              } catch (error) {
                entry.status = 'unresolved';
                entry.unresolvedReasons.push(error.message);
              }
            }
          }
        } else if (entry.role === 'VOLUME_BACKUP') {
          entry.status = 'supporting';
        } else if (entry.role === 'UNKNOWN' || entry.role === 'PRIMARY_BOQ' || entry.role === 'RAB') {
          entry.status = 'unresolved';
          entry.unresolvedReasons.push(meta.semanticRoleReason || 'Peran sheet belum cukup meyakinkan untuk agregasi BOQ.');
        } else if (entry.role === 'OTHER') {
          entry.status = 'unresolved';
          entry.unresolvedReasons.push('Peran sheet belum dapat ditentukan.');
        }
        manifest.push(entry);
      }, this);

      // Link volume-backup rows only when an explicit code maps unambiguously
      // to one or more BOQ items. Evidence remains outside workbook totals.
      manifest.filter(function (entry) { return entry.role === 'VOLUME_BACKUP'; }).forEach(function (entry) {
        var sheet = rawWb.sheets[entry.name];
        var detection = window.HeaderDetector.detect(sheet);
        if (!detection.found || detection.ambiguous) {
          entry.unresolvedReasons.push('Header backup volume belum dapat dipastikan; sheet disimpan sebagai bukti pendukung.');
          return;
        }
        var mapping = window.ColumnMapper.createInitialMapping(detection);
        entry.headerRow = detection.headerRow1Based;
        entry.mapping = mapping;
        var codeField = mapping.fields.item_code;
        var volumeField = mapping.fields.volume;
        if (!codeField || codeField.colIndex < 0 || !volumeField || volumeField.colIndex < 0) {
          entry.unresolvedReasons.push('Kode pekerjaan atau kolom volume tidak terdeteksi untuk pencocokan; sheet tetap sebagai bukti pendukung.');
          return;
        }
        var range = window.XLSX.utils.decode_range(sheet['!ref'] || 'A1:A1');
        var backupDataStart = Number.isInteger(detection.headerRowEndIndex)
          ? detection.headerRowEndIndex + 1 : detection.headerRowIndex + 1;
        for (var rowIndex = backupDataStart; rowIndex <= range.e.r; rowIndex++) {
          var backupRepeatedEnd = window.HeaderDetector.getRepeatedHeaderRowEnd(sheet, rowIndex, mapping.fields);
          if (backupRepeatedEnd >= rowIndex) {
            rowIndex = backupRepeatedEnd;
            continue;
          }
          var row = this.extractRowData(sheet, rowIndex, mapping.fields, entry.name);
          if (!row.item_code) continue;
          var candidates = workItems.filter(function (item) { return item.code && item.code === row.item_code; });
          if (candidates.length !== 1) continue;
          var target = candidates[0];
          var evidence = {
            sheet: entry.name, row: rowIndex + 1, code: row.item_code,
            codeCell: row.source.itemCodeCell || (entry.name + '!' + row.source.cells.item_code.cellRef),
            volume: row.volume,
            volumeCell: row.source.cells.volume ? entry.name + '!' + row.source.cells.volume.cellRef : null,
            formula: row.source.cells.volume ? row.source.cells.volume.formula : null,
            cachedValue: row.source.cells.volume ? row.source.cells.volume.val : null,
            difference: typeof row.volume === 'number' && typeof target.volume === 'number' ? row.volume - target.volume : null
          };
          if (!Array.isArray(target.supportingEvidence)) target.supportingEvidence = [];
          target.supportingEvidence.push(evidence);
        }
      }, this);

      var unresolvedSheets = manifest.filter(function (entry) { return entry.status === 'unresolved'; });
      var supportingSheets = manifest.filter(function (entry) { return supportingRoles[entry.role] && entry.status === 'supporting'; });
      // AHSP sheets remain supporting sources. Extract their structured records
      // separately so they never enter BOQ work-item or financial aggregation.
      var ahspModel;
      try {
        ahspModel = window.AHSPNormalizer && typeof window.AHSPNormalizer.normalize === 'function'
          ? window.AHSPNormalizer.normalize(rawWb, manifest)
          : { schemaVersion: 'rab-guard-ahsp/1', analyses: [], sourceReferences: [], warnings: ['Ekstraktor AHSP tidak tersedia; kebutuhan material belum dapat dianalisis.'] };
      } catch (ahspError) {
        ahspModel = { schemaVersion: 'rab-guard-ahsp/1', analyses: [], sourceReferences: [],
          warnings: ['Ekstraksi AHSP gagal: ' + (ahspError && ahspError.message ? ahspError.message : String(ahspError))] };
        manifest.forEach(function (entry) {
          if (entry.role === 'AHSP_PRICE_REFERENCE' || entry.detectedRole === 'AHSP') entry.ahspExtractionStatus = 'failed';
        });
      }
      var ahspContractInvalid = !Array.isArray(ahspModel.analyses) || !Array.isArray(ahspModel.sourceReferences) || !Array.isArray(ahspModel.warnings);
      if (!Array.isArray(ahspModel.analyses)) ahspModel.analyses = [];
      if (!Array.isArray(ahspModel.sourceReferences)) ahspModel.sourceReferences = [];
      if (!Array.isArray(ahspModel.warnings)) ahspModel.warnings = [];
      if (ahspContractInvalid) ahspModel.warnings.push('Ekstraktor AHSP mengembalikan struktur yang tidak valid; data dipertahankan sebagai unresolved.');
      if (ahspContractInvalid) manifest.forEach(function (entry) {
        if (entry.role === 'AHSP_PRICE_REFERENCE' || entry.detectedRole === 'AHSP') entry.ahspExtractionStatus = 'failed';
      });
      ahspModel.sourceReferences.forEach(function (ref) { if (!ref.workbook) ref.workbook = rawWb.filename || null; });
      ahspModel.analyses.forEach(function (analysis) {
        (analysis.sourceReferences || []).forEach(function (ref) { if (!ref.workbook) ref.workbook = rawWb.filename || null; });
        (analysis.resources || []).forEach(function (resource) {
          (resource.sourceReferences || []).forEach(function (ref) { if (!ref.workbook) ref.workbook = rawWb.filename || null; });
        });
      });
      manifest.filter(function (entry) { return entry.role === 'AHSP_PRICE_REFERENCE' || entry.detectedRole === 'AHSP'; }).forEach(function (entry) {
        var extracted = ahspModel.analyses.filter(function (analysis) { return analysis.source && analysis.source.sheet === entry.name; });
        entry.ahspAnalysisCount = extracted.length;
        entry.ahspResourceCount = extracted.reduce(function (count, analysis) { return count + (analysis.resources || []).length; }, 0);
        if (!ahspContractInvalid && entry.ahspExtractionStatus !== 'failed') entry.ahspExtractionStatus = extracted.length ? 'processed' : 'unresolved';
        if (!extracted.length) {
          var extractionWarning = 'Sheet ' + entry.name + ' terdeteksi sebagai AHSP, tetapi tidak ada analisis terstruktur yang berhasil dinormalisasi.';
          if ((ahspModel.warnings || []).indexOf(extractionWarning) === -1) ahspModel.warnings.push(extractionWarning);
        }
      });
      return {
        project: {
          projectName: rawWb.filename ? rawWb.filename.replace(/\.[^/.]+$/, '') : 'Proyek Tanpa Judul',
          sourceFilename: rawWb.filename || 'unknown.xlsx', fileSize: rawWb.fileSize || 0,
          currency: 'IDR', activeSheet: null, normalizedAt: new Date().toISOString()
        },
        workbookManifest: manifest,
        workItems: workItems, headings: headings, subtotals: subtotals, excludedRows: excludedRows,
        ahspSchemaVersion: ahspModel.schemaVersion,
        ahspAnalyses: ahspModel.analyses,
        ahspSourceReferences: ahspModel.sourceReferences,
        ahspWarnings: ahspModel.warnings,
        statistics: {
          totalSheets: manifest.length, processedBoqSheets: processedBoqSheets.length,
          supportingSheets: supportingSheets.length, unresolvedSheets: unresolvedSheets.length,
          totalWorkItems: workItems.length, totalHeadings: headings.length,
          totalSubtotals: subtotals.length, totalExcludedRows: excludedRows.length,
          totalAhspAnalyses: ahspModel.analyses.length,
          totalAhspResources: ahspModel.analyses.reduce(function (count, analysis) { return count + (analysis.resources || []).length; }, 0),
          totalParsedAmount: totalParsedAmount, processingComplete: unresolvedSheets.length === 0
        },
        processedBoqSheets: processedBoqSheets,
        unresolvedSheets: unresolvedSheets,
        mappingConfig: null
      };
    },

    /**
     * Upgrade canonical snapshots written before the AHSP extension existed.
     * When the original workbook contains AHSP sheets, re-run the actual AHSP
     * extractor so an old snapshot cannot masquerade as a workbook with no
     * AHSP. Extraction failures remain visible as warnings and manifest state.
     */
    ensureAhspContract: function (canonical, rawWb, inspection) {
      if (!canonical || typeof canonical !== 'object') return canonical;
      if (Array.isArray(canonical.ahspAnalyses)) {
        return Object.assign({}, canonical, {
          ahspSchemaVersion: canonical.ahspSchemaVersion || 'rab-guard-ahsp/1',
          ahspSourceReferences: Array.isArray(canonical.ahspSourceReferences) ? canonical.ahspSourceReferences : [],
          ahspWarnings: Array.isArray(canonical.ahspWarnings) ? canonical.ahspWarnings : []
        });
      }

      var upgraded = Object.assign({}, canonical);
      upgraded.ahspSchemaVersion = 'rab-guard-ahsp/1';
      upgraded.ahspAnalyses = [];
      upgraded.ahspSourceReferences = Array.isArray(canonical.ahspSourceReferences) ? canonical.ahspSourceReferences.slice() : [];
      upgraded.ahspWarnings = Array.isArray(canonical.ahspWarnings) ? canonical.ahspWarnings.slice() : [];
      upgraded.workbookManifest = Array.isArray(canonical.workbookManifest)
        ? canonical.workbookManifest.map(function (entry) { return Object.assign({}, entry, { unresolvedReasons: Array.isArray(entry.unresolvedReasons) ? entry.unresolvedReasons.slice() : [] }); })
        : canonical.workbookManifest;

      var hasAhspSource = (upgraded.workbookManifest || []).some(function (entry) {
        return entry.role === 'AHSP_PRICE_REFERENCE' || entry.detectedRole === 'AHSP';
      });
      if (!hasAhspSource && inspection && Array.isArray(inspection.sheets)) {
        hasAhspSource = inspection.sheets.some(function (sheet) {
          return sheet.semanticRole === 'AHSP_PRICE_REFERENCE' || sheet.detectedRole === 'AHSP';
        });
      }
      if (!hasAhspSource) return upgraded;

      if (!rawWb || !rawWb.sheets || !window.AHSPNormalizer || typeof window.AHSPNormalizer.normalize !== 'function') {
        upgraded.ahspWarnings.push('Snapshot lama mengidentifikasi sheet AHSP, tetapi workbook sumber atau ekstraktor AHSP tidak tersedia untuk memulihkan datanya.');
        (upgraded.workbookManifest || []).forEach(function (entry) {
          if (entry.role === 'AHSP_PRICE_REFERENCE' || entry.detectedRole === 'AHSP') {
            entry.ahspExtractionStatus = 'unavailable';
            if (!entry.unresolvedReasons.some(function (reason) { return /AHSP.*tersedia|ekstraksi AHSP/i.test(reason); })) {
              entry.unresolvedReasons.push('Data AHSP belum tersedia pada snapshot dan tidak dapat diekstrak ulang.');
            }
          }
        });
        return upgraded;
      }

      try {
        inspection = inspection || (window.WorkbookInspector && window.WorkbookInspector.inspect(rawWb));
        var extractionManifest = upgraded.workbookManifest;
        var manifestHasAhsp = Array.isArray(extractionManifest) && extractionManifest.some(function (entry) {
          return entry.role === 'AHSP_PRICE_REFERENCE' || entry.detectedRole === 'AHSP';
        });
        if ((!Array.isArray(extractionManifest) || (!manifestHasAhsp && hasAhspSource)) && inspection && window.WorkbookInspector && window.HeaderDetector && window.ColumnMapper) {
          extractionManifest = this.normalizeWorkbook(rawWb, inspection).workbookManifest;
          var manifestEntries = Array.isArray(upgraded.workbookManifest) ? upgraded.workbookManifest : (upgraded.workbookManifest = []);
          extractionManifest.filter(function (entry) { return entry.role === 'AHSP_PRICE_REFERENCE' || entry.detectedRole === 'AHSP'; })
            .forEach(function (sourceEntry) {
              var targetEntry = manifestEntries.find(function (entry) { return entry.name === sourceEntry.name; });
              if (targetEntry) Object.assign(targetEntry, sourceEntry);
              else manifestEntries.push(Object.assign({}, sourceEntry));
            });
        }
        var model = window.AHSPNormalizer.normalize(rawWb, extractionManifest || []);
        upgraded.ahspSchemaVersion = model.schemaVersion || 'rab-guard-ahsp/1';
        upgraded.ahspAnalyses = Array.isArray(model.analyses) ? model.analyses : [];
        upgraded.ahspSourceReferences = Array.isArray(model.sourceReferences) ? model.sourceReferences : [];
        upgraded.ahspWarnings = Array.isArray(model.warnings) ? model.warnings.slice() : [];
        (upgraded.workbookManifest || []).forEach(function (entry) {
          if (entry.role !== 'AHSP_PRICE_REFERENCE' && entry.detectedRole !== 'AHSP') return;
          var analyses = upgraded.ahspAnalyses.filter(function (analysis) { return analysis.source && analysis.source.sheet === entry.name; });
          entry.ahspAnalysisCount = analyses.length;
          entry.ahspResourceCount = analyses.reduce(function (count, analysis) { return count + (analysis.resources || []).length; }, 0);
          entry.ahspExtractionStatus = analyses.length ? 'processed' : 'unresolved';
          if (!analyses.length) {
            var reason = 'Sheet ' + entry.name + ' dikenali sebagai AHSP, tetapi tidak ada analisis terstruktur yang berhasil diekstrak.';
            if (!entry.unresolvedReasons.some(function (existingReason) { return /AHSP|analisis/i.test(existingReason); })) entry.unresolvedReasons.push(reason);
            if (upgraded.ahspWarnings.indexOf(reason) === -1) upgraded.ahspWarnings.push(reason);
          }
        });
      } catch (error) {
        upgraded.ahspWarnings.push('Ekstraksi ulang AHSP gagal: ' + (error && error.message ? error.message : String(error)));
        (upgraded.workbookManifest || []).forEach(function (entry) {
          if (entry.role === 'AHSP_PRICE_REFERENCE' || entry.detectedRole === 'AHSP') entry.ahspExtractionStatus = 'failed';
        });
      }
      return upgraded;
    },

    /**
     * Menormalisasi sheet Excel terpilih menjadi Model Data Kanonikal.
     * @param {Object} rawWb Objek workbook dari XlsxReader
     * @param {string} sheetName Nama sheet yang diaudit
     * @param {Object} mappingConfig Konfigurasi pemetaan kolom dari ColumnMapper
     * @returns {Object} Canonical Data Model
     */
    normalize: function (rawWb, sheetName, mappingConfig) {
      if (!rawWb || !rawWb.sheets || !rawWb.sheets[sheetName]) {
        throw new Error('Sheet "' + sheetName + '" tidak ditemukan dalam workbook.');
      }

      var sheet = rawWb.sheets[sheetName];
      var categoryLevels = (mappingConfig.categoryLevels || []).map(function (level) {
        var copy = Object.assign({}, level);
        copy.headerCellRef = copy.headerCellRef || this.findCategoryHeaderCell(sheet, copy, mappingConfig.headerRowIndex, mappingConfig.headerRowEndIndex);
        return copy;
      }, this);
      var fields = Object.assign({}, mappingConfig.fields, { categoryLevels: categoryLevels });
      var headerRowIndex = mappingConfig.headerRowIndex;

      var range = window.XLSX.utils.decode_range(sheet['!ref'] || 'A1:A1');
      var headerRowEndIndex = Number.isInteger(mappingConfig.headerRowEndIndex)
        ? mappingConfig.headerRowEndIndex : headerRowIndex;
      var startRow = headerRowEndIndex + 1;
      var endRow = range.e.r;

      var workItems = [];
      var headings = [];
      var subtotals = [];
      var excludedRows = [];
      var currentCategory = 'UMUM';
      var headingPath = [];

      var runningItemId = 1;
      var totalParsedAmount = 0;

      for (var r = startRow; r <= endRow; r++) {
        var rowData = this.extractRowData(sheet, r, fields, sheetName);
        var rowContext = this.extractRowContext(sheet, r, range.s.c, range.e.c);

        // Abaikan baris yang benar-benar kosong pada seluruh field yang dipetakan.
        // Sel liar di kolom yang tidak dipetakan tidak boleh berubah menjadi item.
        if (rowData.isEmpty) {
          continue;
        }

        // Repeated column headers inside a BOQ (often at page/division breaks)
        // are evidence rows, not work items. Keep their physical location in
        // the exclusion manifest instead of silently dropping them.
        var repeatedHeaderEnd = window.HeaderDetector.getRepeatedHeaderRowEnd(sheet, r, fields);
        if (repeatedHeaderEnd >= r) {
          for (var repeatedRow = r; repeatedRow <= repeatedHeaderEnd; repeatedRow++) {
            excludedRows.push({
              rowType: 'REPEATED_HEADER',
              reason: 'Baris mengulang mapping header BOQ yang sudah dipilih.',
              sourceRow: repeatedRow + 1,
              source: { sheet: sheetName, row: repeatedRow + 1 }
            });
          }
          r = repeatedHeaderEnd;
          continue;
        }

        // Baris panduan nomor kolom (mis. 1,2,3,4,5,6) bukan item RAB.
        // Tetap simpan jejaknya sebagai excludedRows agar tidak hilang tanpa alasan.
        if (this.isColumnNumberingGuideRow(rowData, rowContext)) {
          excludedRows.push({
            rowType: 'COLUMN_INDEX_GUIDE',
            reason: 'Baris hanya berisi nomor urut/panduan kolom, bukan data pekerjaan.',
            sourceRow: r + 1,
            source: { sheet: sheetName, row: r + 1 }
          });
          continue;
        }

        // Klasifikasikan subtotal memakai seluruh teks pada baris, bukan hanya
        // kolom Uraian. Banyak workbook menempatkan label JUMLAH di kolom NO.
        var descLower = (rowData.description || '').toLowerCase();
        var subtotalInfo = this.findSubtotalLabel(rowData, rowContext);
        var isSubtotalRow = this.detectSubtotalRow(descLower, rowData, rowContext);

        if (isSubtotalRow) {
          var subtotalLabel = subtotalInfo ? subtotalInfo.text : (rowData.description || 'Subtotal');
          var subtotalLabelCell = subtotalInfo ? subtotalInfo.cellRef : (rowData.source.descriptionCell || null);
          var subtotalCategoryPath = this.createCategoryPath(sheet, r, categoryLevels, headingPath, sheetName, rowData.categoryValues);
          var subtotalHierarchy = this.createHierarchyState(sheet, r, categoryLevels, headingPath, sheetName, subtotalCategoryPath);
          var subtotalSource = {
            sheet: sheetName,
            row: r + 1,
            amountCell: rowData.source.amountCell,
            amountFormula: rowData.source.amountFormula,
            labelCell: subtotalLabelCell
          };
          subtotals.push({
            id: 'subtotal-' + (subtotals.length + 1),
            label: subtotalLabel,
            description: subtotalLabel,
            amount: rowData.amount,
            sourceRow: r + 1,
            sourceCell: rowData.source.amountCell,
            formula: rowData.source.amountFormula,
            category: currentCategory,
            categoryPath: subtotalCategoryPath,
            hierarchy: subtotalHierarchy,
            source: subtotalSource
          });
          continue;
        }

        var isHeadingRow = this.detectHeadingRow(rowData, rowContext);
        if (isHeadingRow) {
          currentCategory = (rowData.description || this.findHeadingLabel(rowContext) || '').trim();
          var depth = this.getHeadingLevel(currentCategory);
          var headingProv = this.createHeadingLevel(sheetName, r + 1, currentCategory, rowData);
          if (depth === null) {
            headingPath = [{ levelIndex: 0, value: currentCategory, sourceLabel: currentCategory,
              resolution: 'ambiguous', provenance: headingProv }];
          } else {
            headingPath = headingPath.filter(function (level) { return level.levelIndex < depth; });
            headingPath.push({ levelIndex: depth, value: this.stripHeadingMarker(currentCategory), sourceLabel: currentCategory,
              resolution: 'resolved', provenance: headingProv });
          }
          headings.push({
            id: 'heading-' + (headings.length + 1),
            code: rowData.item_code,
            analysisCode: rowData.analysis_code || '',
            title: currentCategory,
            sourceRow: r + 1,
            source: { sheet: sheetName, row: r + 1 }, categoryPath: headingPath.slice(),
            hierarchy: { status: depth === null ? 'ambiguous' : 'resolved', structure: 'vertical_headings', unresolved: depth === null ? [{ reason: 'Kedalaman heading tidak memiliki bukti eksplisit.', levelIndex: 0, candidates: [], evidenceSources: [headingProv] }] : [] }
          });
          continue;
        }

        // Baris Item Pekerjaan (termasuk item yang datanya belum lengkap).
        // Nomor urut, kode pekerjaan, dan kode analisa disimpan terpisah.
        var itemCategoryPath = this.createCategoryPath(sheet, r, categoryLevels, headingPath, sheetName, rowData.categoryValues);
        var itemHierarchy = this.createHierarchyState(sheet, r, categoryLevels, headingPath, sheetName, itemCategoryPath);
        var item = {
          id: 'item-' + runningItemId++,
          itemNumber: rowData.item_number || '',
          code: rowData.item_code || '',
          analysisCode: rowData.analysis_code || '',
          description: rowData.description || '',
          unit: rowData.unit || '',
          volume: rowData.volume,
          unitPrice: rowData.unit_price,
          amount: rowData.amount,
          categoryPath: itemCategoryPath,
          hierarchy: itemHierarchy,
          hierarchyStatus: itemHierarchy.status,
          hierarchyUnresolvedEvidence: itemHierarchy.unresolved,
          category: this.projectCompatibilityCategory(currentCategory, itemCategoryPath),
          rowType: 'WORK_ITEM',
          source: rowData.source
        };

        if (typeof item.amount === 'number' && !isNaN(item.amount)) {
          totalParsedAmount += item.amount;
        }

        workItems.push(item);
      }

      var projectModel = {
        projectName: rawWb.filename ? rawWb.filename.replace(/\.[^/.]+$/, '') : 'Proyek Tanpa Judul',
        sourceFilename: rawWb.filename || 'unknown.xlsx',
        fileSize: rawWb.fileSize || 0,
        currency: 'IDR',
        activeSheet: sheetName,
        normalizedAt: new Date().toISOString()
      };

      return {
        project: projectModel,
        workItems: workItems,
        headings: headings,
        subtotals: subtotals,
        excludedRows: excludedRows,
        // Single-sheet normalization does not imply AHSP was scanned. Keep
        // the extension shape stable and distinguish this from an empty scan.
        ahspSchemaVersion: 'rab-guard-ahsp/1',
        ahspAnalyses: [],
        ahspSourceReferences: [],
        ahspWarnings: [],
        ahspExtractionStatus: 'not_scanned',
        statistics: {
          totalWorkItems: workItems.length,
          totalHeadings: headings.length,
          totalSubtotals: subtotals.length,
          totalExcludedRows: excludedRows.length,
          totalParsedAmount: totalParsedAmount
        },
        mappingConfig: mappingConfig
      };
    },

    /**
     * Mengekstrak nilai dan formula untuk sebuah baris sesuai mapping kolom.
     */
    extractRowData: function (sheet, rowIndex, fields, sheetName) {
      var row1Based = rowIndex + 1;
      var data = {
        isEmpty: true,
        item_code: null,
        description: null,
        unit: null,
        volume: null,
        unit_price: null,
        amount: null,
        source: {
          sheet: sheetName,
          row: row1Based,
          cells: {},
          amountFormula: null,
          priceFormula: null,
          amountCell: null
        }
      };

      var anyNonEmpty = false;

      for (var fKey in fields) {
        var fMeta = fields[fKey];
        if (fKey === 'categoryLevels' || fMeta.colIndex < 0) continue;

        var cellRef = window.XLSX.utils.encode_cell({ r: rowIndex, c: fMeta.colIndex });
        var cell = sheet[cellRef];

        var rawVal = cell ? cell.v : undefined;
        var formula = cell ? cell.f : null;
        var isErrorCell = !!(cell && cell.t === 'e');
        var errorText = null;
        if (isErrorCell) {
          errorText = cell.w || (window.XLSX.utils.format_cell ? window.XLSX.utils.format_cell(cell) : null) || String(rawVal);
        }

        data.source.cells[fKey] = {
          cellRef: cellRef,
          colLetter: fMeta.colLetter,
          val: rawVal,
          formula: formula,
          cellType: cell ? cell.t : null,
          valueStatus: isErrorCell ? 'error' : (rawVal === undefined || rawVal === null || String(rawVal).trim() === '' ? 'missing' : 'valid'),
          errorCode: isErrorCell ? rawVal : null,
          errorText: errorText
        };

        if (fKey === 'amount') {
          data.source.amountCell = sheetName + '!' + cellRef;
          data.source.amountFormula = formula;
        }
        if (fKey === 'description') {
          data.source.descriptionCell = sheetName + '!' + cellRef;
        }
        if (fKey === 'item_code') {
          data.source.itemCodeCell = sheetName + '!' + cellRef;
        }
        if (fKey === 'analysis_code') {
          data.source.analysisCodeCell = sheetName + '!' + cellRef;
        }
        if (fKey === 'unit_price') {
          data.source.priceFormula = formula;
        }

        // Error-cell numeric codes are provenance, never canonical field values.
        if (isErrorCell) {
          anyNonEmpty = true;
          continue;
        }

        if (rawVal !== undefined && rawVal !== null && String(rawVal).trim() !== '') {
          anyNonEmpty = true;
          if (fKey === 'item_number' || fKey === 'item_code' || fKey === 'analysis_code' || fKey === 'description' || fKey === 'unit' || fKey === 'category' || /^category_level_/.test(fKey)) {
            data[fKey] = String(rawVal).trim();
          } else if (fKey === 'volume' || fKey === 'unit_price' || fKey === 'amount') {
            data[fKey] = this.parseNumericValue(rawVal);
          }
        }
      }

      data.isEmpty = !anyNonEmpty;
      data.categoryValues = [];
      (fields.categoryLevels || []).forEach(function (level) {
        var extracted = this.extractCategoryValue(sheet, rowIndex, level, sheetName);
        data.categoryValues.push({ level: level, value: extracted.value,
          cellRef: extracted.provenance.originalValueCell.ref, mergedRange: extracted.provenance.mergedAnchor && extracted.provenance.mergedAnchor.range,
          inherited: extracted.provenance.inherited, provenance: extracted.provenance, row: rowIndex + 1 });
      }, this);
      if (data.categoryValues.some(function (x) { return x.value !== null; })) data.isEmpty = false;
      return data;
    },

    findVerticalMerge: function (sheet, rowIndex, colIndex) {
      var merges = sheet && Array.isArray(sheet['!merges']) ? sheet['!merges'] : [];
      for (var i = 0; i < merges.length; i++) {
        var m = merges[i];
        if (m.s.c === colIndex && m.e.c === colIndex && rowIndex >= m.s.r && rowIndex <= m.e.r) return m;
      }
      return null;
    },

    findCategoryHeaderCell: function (sheet, level, headerStart, headerEnd) {
      if (!Number.isInteger(level.colIndex) || level.colIndex < 0) return null;
      var start = Number.isInteger(headerStart) ? headerStart : 0;
      var end = Number.isInteger(headerEnd) ? headerEnd : start;
      var expected = String(level.originalHeader || '').trim().toLowerCase().replace(/\s+/g, ' ');
      for (var row = end; row >= start; row--) {
        var directRef = window.XLSX.utils.encode_cell({ r: row, c: level.colIndex });
        var direct = sheet[directRef];
        if (direct && String(direct.v || '').trim().toLowerCase().replace(/\s+/g, ' ') === expected) return directRef;
        var merges = sheet && Array.isArray(sheet['!merges']) ? sheet['!merges'] : [];
        for (var i = 0; i < merges.length; i++) {
          var merge = merges[i];
          if (row >= merge.s.r && row <= merge.e.r && level.colIndex >= merge.s.c && level.colIndex <= merge.e.c) {
            var anchorRef = window.XLSX.utils.encode_cell({ r: merge.s.r, c: merge.s.c });
            var anchor = sheet[anchorRef];
            if (anchor && String(anchor.v || '').trim().toLowerCase().replace(/\s+/g, ' ') === expected) return anchorRef;
          }
        }
      }
      return window.XLSX.utils.encode_cell({ r: end, c: level.colIndex });
    },

    createHeadingLevel: function (sheetName, row, text, rowData) {
      var ref = rowData && rowData.source && rowData.source.descriptionCell ? rowData.source.descriptionCell.split('!').pop() : null;
      return { origin: 'vertical_heading', sheet: sheetName, originalValueCell: { ref: ref, value: text },
        mergedAnchor: null, appliedTo: { row: row, cellRef: ref }, headerSourceCell: null,
        verticalHeadingSourceCell: { ref: ref, value: text }, inherited: true };
    },

    getHeadingLevel: function (text) {
      var value = String(text || '').trim();
      if (/^(?:[IVXLCDM]+)[.)]?\s+/i.test(value)) return 0;
      if (/^\d+(?:\.\d+)+[.)]?\s+/.test(value)) return value.match(/^\d+(?:\.\d+)+/)[0].split('.').length - 1;
      if (/^[a-z][.)]\s+/i.test(value)) return 1;
      return null;
    },

    stripHeadingMarker: function (text) { return String(text || '').replace(/^(?:(?:[IVXLCDM]+|\d+(?:\.\d+)*|[a-z])[.)]?\s+)/i, '').trim(); },

    createCategoryPath: function (sheet, rowIndex, levels, headingPath, sheetName, rowCategoryValues) {
      var result = (headingPath || []).map(function (level) {
        var snapshot = JSON.parse(JSON.stringify(level));
        snapshot.resolution = 'inherited';
        snapshot.provenance.inherited = true;
        snapshot.provenance.appliedTo = { row: rowIndex + 1, cellRef: null };
        return snapshot;
      });
      (levels || []).forEach(function (level, levelIndex) {
        var cached = rowCategoryValues && rowCategoryValues[levelIndex];
        var valueCell = cached ? { value: cached.value, provenance: cached.provenance } : this.extractCategoryValue(sheet, rowIndex, level, sheetName);
        if (valueCell.value === null) {
          result.push({ levelIndex: level.levelIndex, value: null, sourceLabel: level.originalHeader, resolution: 'missing',
            provenance: valueCell.provenance });
          return;
        }
        var semantic = String(level.originalHeader || '').toLowerCase();
        var levelIndex = level.levelIndex;
        var explicitSub = /\bsub|subdiv|subgroup|subkategori/.test(semantic);
        var explicitLevel = semantic.match(/(?:level|tier|tingkat)\s*(\d+)/);
        if (explicitLevel) levelIndex = Math.max(0, Number(explicitLevel[1]) - 1);
        else if (explicitSub && levelIndex === 0) levelIndex = 1;
        var existing = result.filter(function (item) { return item.levelIndex === levelIndex; })[0];
        if (existing) {
          var sameValue = String(existing.value || '').trim().toLowerCase() === String(valueCell.value || '').trim().toLowerCase();
          existing.provenance.sources = (existing.provenance.sources || [JSON.parse(JSON.stringify(existing.provenance))]).concat([valueCell.provenance]);
          if (!sameValue) { existing.resolution = 'ambiguous'; existing.conflictingValue = valueCell.value; }
          return;
        }
        result.push({ levelIndex: levelIndex, value: valueCell.value, sourceLabel: level.originalHeader,
          resolution: valueCell.provenance.inherited ? 'inherited' : 'resolved', provenance: valueCell.provenance });
      }, this);
      result = result.filter(function (level, index) {
        if (level.resolution !== 'missing') return true;
        return !result.slice(index + 1).some(function (later) { return later.value !== null; });
      });
      result.sort(function (a, b) { return a.levelIndex - b.levelIndex; });
      return result;
    },

    extractCategoryValue: function (sheet, rowIndex, level, sheetName) {
      var ref = window.XLSX.utils.encode_cell({ r: rowIndex, c: level.colIndex });
      var cell = sheet[ref]; var merge = this.findVerticalMerge(sheet, rowIndex, level.colIndex);
      var anchorRef = merge ? window.XLSX.utils.encode_cell({ r: merge.s.r, c: merge.s.c }) : ref;
      var anchor = sheet[anchorRef];
      var original = cell && cell.v !== undefined && cell.v !== null && String(cell.v).trim() !== '' ? cell : (merge ? anchor : null);
      var value = original ? String(original.v).trim() : null;
      return { value: value, provenance: { origin: merge ? 'merged_cell' : 'column', sheet: sheetName || null,
        originalValueCell: { ref: original ? anchorRef : ref, value: value },
        mergedAnchor: merge ? { ref: anchorRef, range: window.XLSX.utils.encode_range(merge) } : null,
        appliedTo: { row: rowIndex + 1, cellRef: ref }, headerSourceCell: { ref: level.headerCellRef || null, value: level.originalHeader },
        verticalHeadingSourceCell: null, inherited: !!(merge && merge.s.r !== rowIndex) } };
    },

    createHierarchyState: function (sheet, rowIndex, levels, headingPath, sheetName, categoryPath) {
      var path = categoryPath || this.createCategoryPath(sheet, rowIndex, levels, headingPath, sheetName);
      var missing = path.filter(function (x) { return x.resolution === 'missing'; });
      var ambiguous = path.filter(function (x) { return x.resolution === 'ambiguous'; });
      var indexed = {};
      path.forEach(function (x) { indexed[x.levelIndex] = true; });
      var hasIndexGap = false;
      if (path.length > 1) for (var levelIndex = path[0].levelIndex; levelIndex < path[path.length - 1].levelIndex; levelIndex++) if (!indexed[levelIndex]) hasIndexGap = true;
      var explicitRelation = (levels || []).some(function (level) { return /\bsub|level\s*\d+|tier\s*\d+|tingkat\s*\d+/.test(String(level.originalHeader || '').toLowerCase()); });
      var genericMultiColumn = levels.length > 1 && !explicitRelation;
      var hasResolved = path.some(function (x) { return x.resolution === 'resolved' || x.resolution === 'inherited'; });
      var status = !path.length ? 'none' : (ambiguous.length || genericMultiColumn) ? 'ambiguous' : (missing.length || hasIndexGap) ? 'partial' : 'resolved';
      var unresolved = ambiguous.map(function (x) { return { reason: 'Kedalaman atau relasi hierarchy belum terbukti.', levelIndex: x.levelIndex, candidates: [{ value: x.value, sourceLabel: x.sourceLabel, provenance: x.provenance }], evidenceSources: [x.provenance] }; });
      missing.forEach(function (x) {
        unresolved.push({ reason: 'Level kategori tidak memiliki nilai pada baris sumber.', levelIndex: x.levelIndex,
          candidates: [], evidenceSources: [x.provenance] });
      });
      if (hasIndexGap) {
        for (var gap = path[0].levelIndex; gap < path[path.length - 1].levelIndex; gap++) {
          if (!indexed[gap]) unresolved.push({ reason: 'Indeks level antara tidak memiliki nilai; indeks level sumber tetap dipertahankan.', levelIndex: gap, candidates: [], evidenceSources: levels });
        }
      }
      if (genericMultiColumn) unresolved.push({ reason: 'Beberapa kolom kategori tidak memiliki bukti parent-child eksplisit.', candidates: levels.map(function (x) { return { sourceLabel: x.originalHeader, headerCellRef: x.headerCellRef, levelIndex: x.levelIndex }; }), evidenceSources: levels });
      return { status: status, structure: levels.length && headingPath.length ? 'mixed' : levels.length ? 'columns' : headingPath.length ? 'vertical_headings' : 'none',
        unresolved: unresolved };
    },

    projectCompatibilityCategory: function (legacyCategory, path) {
      if (legacyCategory && legacyCategory !== 'UMUM') return legacyCategory;
      var resolved = (path || []).filter(function (x) { return x.value !== null && (x.resolution === 'resolved' || x.resolution === 'inherited'); });
      return resolved.length ? resolved[resolved.length - 1].value : 'UMUM';
    },

    /**
     * Mengurai nilai numerik dari string atau number Excel secara aman.
     * Mendukung format angka Indonesia (1.250,50) dan Internasional (1,250.50),
     * simbol mata uang, spasi, notasi akuntansi negatif, dan format integer.
     */
    parseNumericValue: function (val) {
      if (typeof val === 'number') {
        return isNaN(val) ? null : val;
      }
      if (typeof val !== 'string') {
        return null;
      }

      var str = val.trim();
      if (!str || str === '-' || str === '--') {
        return null;
      }

      // Tanda ,-. adalah terminator format Rupiah positif, bukan tanda negatif.
      if (/,\s*-$/.test(str)) {
        str = str.replace(/,\s*-$/, '');
      }

      // Deteksi tanda negatif (accounting parentheses atau tanda minus)
      var isNegative = false;
      if (/^\(.*\)$/.test(str)) {
        isNegative = true;
        str = str.substring(1, str.length - 1).trim();
      } else if (/^-\s*/.test(str)) {
        isNegative = true;
        str = str.replace(/^-\s*/, '').trim();
      } else if (/\s*-$/.test(str)) {
        isNegative = true;
        str = str.replace(/\s*-$/, '').trim();
      }

      // Deteksi dan bersihkan prefiks mata uang dan spasi
      var hadCurrency = /^(rp\.?|idr|\$|eur)\s*/i.test(str);
      str = str.replace(/^(rp\.?|idr|\$|eur)\s*/i, '').replace(/\s+/g, '').trim();
      if (!str) return null;

      var hasDot = str.indexOf('.') !== -1;
      var hasComma = str.indexOf(',') !== -1;

      if (hasDot && hasComma) {
        // Keduanya ada: tentukan urutan posisi
        var firstDot = str.indexOf('.');
        var firstComma = str.indexOf(',');
        if (firstDot < firstComma) {
          // Format Indonesia: 1.250.000,50 -> titik = ribuan, koma = desimal
          str = str.replace(/\./g, '').replace(/,/g, '.');
        } else {
          // Format Internasional: 1,250,000.50 -> koma = ribuan, titik = desimal
          str = str.replace(/,/g, '');
        }
      } else if (hasDot && !hasComma) {
        var dotCount = (str.match(/\./g) || []).length;
        if (dotCount > 1) {
          // Lebih dari 1 titik -> pasti pemisah ribuan Indonesia (e.g. 1.250.000)
          str = str.replace(/\./g, '');
        } else {
          // Tepat 1 titik: e.g. 1.250 atau 1250.50 atau 0.75
          var dotParts = str.split('.');
          if (dotParts[1].length !== 3 || dotParts[0] === '0') {
            // Bukan 3 digit desimal atau diawali 0 (e.g. 0.75, 12.50, 1.5) -> desimal internasional
            // Titik tetap dipertahankan sebagai desimal
          } else if (hadCurrency) {
            // Didahului Rp/IDR dan 3 digit -> ribuan Indonesia (e.g. Rp 1.250)
            str = str.replace(/\./g, '');
          } else {
            // Format 1 titik dengan 3 digit (e.g. 1.250): dalam konteks teks RAB Indonesia, titik adalah ribuan
            str = str.replace(/\./g, '');
          }
        }
      } else if (hasComma && !hasDot) {
        var commaCount = (str.match(/,/g) || []).length;
        if (commaCount > 1) {
          // Lebih dari 1 koma -> pasti pemisah ribuan internasional (e.g. 1,250,000)
          str = str.replace(/,/g, '');
        } else {
          // Tepat 1 koma: e.g. 1,250 atau 1250,50 atau 0,75
          var commaParts = str.split(',');
          if (commaParts[1].length !== 3 || commaParts[0] === '0') {
            // Bukan 3 digit atau diawali 0 (e.g. 0,75 atau 1250,50) -> koma desimal Indonesia
            str = commaParts[0] + '.' + commaParts[1];
          } else {
            // Ambigu 1 koma dengan 3 digit (e.g. 1,250):
            // Dalam format numerik standar, ubah koma menjadi titik desimal
            str = commaParts[0] + '.' + commaParts[1];
          }
        }
      }

      var num = parseFloat(str);
      if (isNaN(num)) return null;
      return isNegative ? -num : num;
    },

    /**
     * Mengambil seluruh sel terisi pada satu baris, termasuk kolom yang tidak
     * dipetakan. Ini diperlukan untuk label subtotal yang sering berada di NO,
     * bukan di kolom Uraian Pekerjaan.
     */
    extractRowContext: function (sheet, rowIndex, startCol, endCol) {
      var cells = [];
      var textCells = [];

      for (var c = startCol; c <= endCol; c++) {
        var cellRef = window.XLSX.utils.encode_cell({ r: rowIndex, c: c });
        var cell = sheet[cellRef];
        if (!cell || cell.v === undefined || cell.v === null || String(cell.v).trim() === '') continue;

        var entry = {
          cellRef: cellRef,
          colIndex: c,
          colLetter: window.XLSX.utils.encode_col(c),
          value: cell.v,
          formula: cell.f || null
        };
        cells.push(entry);
        if (typeof cell.v === 'string') textCells.push(entry);
      }

      return { cells: cells, textCells: textCells };
    },

    /**
     * Mendeteksi baris panduan nomor kolom yang kadang disisipkan tepat di
     * bawah header (contoh: NO=1, KODE ANALISA=2, URAIAN=2, SAT=3, dst.).
     */
    isColumnNumberingGuideRow: function (rowData, rowContext) {
      var description = rowData && rowData.description !== null && rowData.description !== undefined
        ? String(rowData.description).trim() : '';
      if (!/^\d+(?:[.,]\d+)?$/.test(description)) return false;
      if (!rowContext || !Array.isArray(rowContext.cells) || rowContext.cells.length < 4) return false;

      for (var i = 0; i < rowContext.cells.length; i++) {
        var cell = rowContext.cells[i];
        if (cell.formula && !/^\s*=?\s*[A-Z]{1,3}\$?\d+\s*[+-]\s*\d+(?:\.\d+)?\s*$/i.test(cell.formula)) return false;
        if (typeof cell.value === 'number' && isFinite(cell.value)) continue;
        if (typeof cell.value === 'string' && /^\d+(?:[.,]\d+)?$/.test(cell.value.trim())) continue;
        return false;
      }
      return true;
    },

    /** Menguji apakah sebuah teks adalah label agregasi, bukan nama item biasa. */
    isSubtotalLabelText: function (text) {
      if (!text) return false;
      var value = String(text).toLowerCase().replace(/\s+/g, ' ').trim();
      var prefix = /^(?:[\s\-*]*(?:\d+[.)]?\s*)?)(?:sub\s*total|sub-total|grand\s*total|total\b|jumlah\s+(?:harga|biaya|total|bagian|pekerjaan|rekapitulasi)\b|rekapitulasi\b)/i;
      var phrase = /\b(?:sub\s*total|sub-total|grand\s*total|jumlah\s+total|total\s+biaya|jumlah\s+biaya|total\s+harga|jumlah\s+harga|jumlah\s+bagian|total\s+bagian|total\s+pekerjaan|jumlah\s+pekerjaan|total\s+proyek|jumlah\s+rekapitulasi)\b/i;
      return prefix.test(value) || phrase.test(value);
    },

    /** Menemukan teks label subtotal dan lokasi sel asalnya pada seluruh baris. */
    findSubtotalLabel: function (rowData, rowContext) {
      var description = rowData && rowData.description ? String(rowData.description).trim() : '';
      if (description && this.isSubtotalLabelText(description)) {
        return {
          text: description,
          cellRef: rowData.source && rowData.source.descriptionCell
            ? rowData.source.descriptionCell.split('!').pop() : null
        };
      }

      if (rowContext && Array.isArray(rowContext.textCells)) {
        for (var i = 0; i < rowContext.textCells.length; i++) {
          var cell = rowContext.textCells[i];
          if (this.isSubtotalLabelText(cell.value)) {
            return { text: String(cell.value).trim(), cellRef: cell.cellRef };
          }
        }
      }
      return null;
    },

    /** Mengambil teks heading fallback dari baris yang tidak memiliki uraian terpetakan. */
    findHeadingLabel: function (rowContext) {
      if (!rowContext || !Array.isArray(rowContext.textCells)) return '';
      for (var i = 0; i < rowContext.textCells.length; i++) {
        var text = String(rowContext.textCells[i].value || '').trim();
        if (/^(?:divisi\b|[ivxlcdm]+\.|[a-z]\.)/i.test(text)) return text;
      }
      return '';
    },

    /**
     * Mendeteksi apakah baris adalah baris Subtotal atau Total secara ketat.
     * Mencegah false positive pada item pekerjaan seperti "Jumlah Reng" atau "Total Station".
     */
    detectSubtotalRow: function (descLower, rowData, rowContext) {
      rowData = rowData || {};
      var description = rowData.description ? String(rowData.description).trim() : '';
      var hasVolume = typeof rowData.volume === 'number' && !isNaN(rowData.volume);
      var hasPrice = typeof rowData.unit_price === 'number' && !isNaN(rowData.unit_price);
      var hasUnit = !!(rowData.unit && String(rowData.unit).trim().length > 0);
      var hasAmount = typeof rowData.amount === 'number' && !isNaN(rowData.amount);
      var hasAmountFormula = !!(rowData.source && rowData.source.amountFormula);

      // Deskripsi nyata yang disertai sinyal kuantitas/satuan/harga tetap item,
      // termasuk nama pekerjaan yang kebetulan diawali kata "Jumlah" atau "Total".
      var hasMeaningfulDescription = description && !/^\d+(?:[.,]\d+)?$/.test(description);
      var hasWorkItemSignals = hasUnit || hasVolume || hasPrice;
      if (hasMeaningfulDescription && hasWorkItemSignals) return false;

      // Subtotal/total harus memiliki nominal atau formula pada kolom jumlah.
      if (!hasAmount && !hasAmountFormula) return false;

      // Cari label di kolom uraian maupun kolom lain (seringkali kolom NO).
      if (this.findSubtotalLabel(rowData, rowContext)) return true;

      // Formula SUM tanpa label hanya dianggap agregat bila tidak ada sinyal
      // item pekerjaan sama sekali. Jangan mengubah item berformula menjadi subtotal.
      if (hasAmountFormula && /^\s*=?\s*sum\s*\(/i.test(rowData.source.amountFormula) &&
          !hasUnit && !hasVolume && !hasPrice && !hasMeaningfulDescription) {
        return true;
      }

      return false;
    },

    /**
     * Mendeteksi judul bagian/divisi. Baris header tetap dipisahkan dari item,
     * dan tidak memerlukan angka/volume/harga untuk dianggap sebagai heading.
     */
    detectHeadingRow: function (rowData, rowContext) {
      var description = rowData && rowData.description ? String(rowData.description).trim() : '';
      var hasNoVol = rowData.volume === null || rowData.volume === undefined;
      var hasNoPrice = rowData.unit_price === null || rowData.unit_price === undefined;
      var hasNoAmount = rowData.amount === null || rowData.amount === undefined;
      var hasNoUnit = !rowData.unit;
      var codeValues = [rowData.item_code, rowData.analysis_code];
      var hasMeaningfulWorkCode = codeValues.some(function (value) {
        var code = value === null || value === undefined ? '' : String(value).trim();
        if (!code) return false;
        // Section labels can land in code columns in common RAB layouts.
        return !/^(?:divisi\b.*|[ivxlcdm]+\.?|[a-z]\.?)$/i.test(code);
      });

      // Explicit work codes identify incomplete items even when all quantities are absent.
      if (hasMeaningfulWorkCode) return false;

      if (description && !/^\d+(?:[.,]\d+)?$/.test(description) && hasNoVol && hasNoPrice && hasNoAmount && hasNoUnit) {
        return true;
      }

      // Fallback untuk heading yang berada di kolom lain sementara mapping
      // uraian tidak terisi pada baris tersebut.
      if (!description && hasNoVol && hasNoPrice && hasNoAmount && hasNoUnit && rowContext && rowContext.textCells) {
        for (var i = 0; i < rowContext.textCells.length; i++) {
          var text = String(rowContext.textCells[i].value || '').trim();
          if (/^(?:divisi\b|[ivxlcdm]+\.|[a-z]\.)/i.test(text)) return true;
        }
      }
      return false;
    }
  };

  window.DataNormalizer = DataNormalizer;

})(typeof window !== 'undefined' ? window : this);
