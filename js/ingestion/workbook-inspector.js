/**
 * RAB GUARD — Universal XLSX Ingestion
 * Module: WorkbookInspector (js/ingestion/workbook-inspector.js)
 * 
 * Tanggung Jawab:
 * - Memeriksa dimensi, lembar kerja tersembunyi, merged cells, dan sebaran tipe sel.
 * - Mendeteksi peran lembar kerja (RAB, AHSP, Rekap, Kurva S, Pendukung) dengan skor keyakinan.
 * - Menyediakan cuplikan data baris awal untuk pratinjau pengguna.
 */

(function (window) {
  'use strict';

  var WorkbookInspector = {
    /**
     * Memeriksa seluruh lembar kerja pada workbook.
     * @param {Object} parsedWb Objek dari XlsxReader
     * @returns {Object} Hasil inspeksi lengkap
     */
    inspect: function (parsedWb) {
      if (!parsedWb || !parsedWb.rawWorkbook) {
        throw new Error('Workbook tidak valid untuk diinspeksi.');
      }

      var rawWb = parsedWb.rawWorkbook;
      var sheetNames = parsedWb.sheetNames || [];
      var sheetsMeta = [];

      var wbProps = (rawWb.Workbook && rawWb.Workbook.Sheets) ? rawWb.Workbook.Sheets : [];

      for (var i = 0; i < sheetNames.length; i++) {
        var name = sheetNames[i];
        var sheet = rawWb.Sheets[name];
        var isHidden = false;

        if (wbProps[i] && wbProps[i].Hidden) {
          isHidden = true;
        }

        var sheetReport = this.inspectSheet(name, sheet, isHidden);
        var headerDetection = window.HeaderDetector && sheet
          ? window.HeaderDetector.detect(sheet) : { found: false, ambiguous: false, confidence: 0, reason: 'Detektor header tidak tersedia.' };
        var initialMapping = window.ColumnMapper && headerDetection.found
          ? window.ColumnMapper.createInitialMapping(headerDetection) : null;
        var mappingValidation = window.ColumnMapper && initialMapping
          ? window.ColumnMapper.validateMapping(initialMapping) : { valid: false, errors: ['Mapping kolom tidak tersedia.'] };
        if (sheetReport.semanticRole === 'UNKNOWN' && sheetReport.detectedRole === 'RAB' &&
            headerDetection.found && !headerDetection.ambiguous && mappingValidation.valid &&
            headerDetection.mappings.description && headerDetection.mappings.volume &&
            headerDetection.mappings.unit_price) {
          sheetReport.semanticRole = 'PRIMARY_BOQ';
          sheetReport.semanticRoleReason = 'Mapping header dan struktur kolom finansial BOQ terdeteksi otomatis.';
          sheetReport.semanticConfidence = Math.max(sheetReport.semanticConfidence || 0, 80);
        }
        sheetReport.headerDetection = {
          found: !!headerDetection.found,
          row1Based: headerDetection.headerRow1Based || null,
          rowEnd1Based: headerDetection.headerRowEnd1Based || headerDetection.headerRow1Based || null,
          confidence: headerDetection.confidence || 0,
          ambiguous: !!headerDetection.ambiguous,
          mappingValid: !!mappingValidation.valid,
          reason: !headerDetection.found ? (headerDetection.reason || 'Header belum terdeteksi.') :
            headerDetection.ambiguous ? (headerDetection.ambiguityReason || 'Kandidat header setara menghasilkan mapping berbeda.') :
              (!mappingValidation.valid ? mappingValidation.errors.join(' ') : null)
        };
        sheetReport.normalizationReady = sheetReport.semanticRole === 'PRIMARY_BOQ' &&
          sheetReport.semanticConfidence >= 70 && sheetReport.headerDetection.found &&
          !sheetReport.headerDetection.ambiguous && sheetReport.headerDetection.mappingValid;
        sheetsMeta.push(sheetReport);
      }

      // Pilih otomatis hanya jika satu-satunya kandidat RAB memiliki keyakinan tinggi.
      var rabCandidates = sheetsMeta.filter(function (meta) {
        return meta.semanticRole === 'PRIMARY_BOQ';
      });
      var trustedRabCandidates = rabCandidates.filter(function (meta) {
        return meta.semanticConfidence >= 70 && meta.normalizationReady;
      });
      var primaryRabSheet = rabCandidates.length === 1 && trustedRabCandidates.length === 1
        ? trustedRabCandidates[0].sheetName
        : null;

      return {
        filename: parsedWb.filename,
        fileSize: parsedWb.fileSize,
        sheetCount: sheetNames.length,
        primaryRabSheet: primaryRabSheet,
        rabCandidates: rabCandidates.map(function (meta) {
          return {
            sheetName: meta.sheetName,
            confidence: meta.confidence,
            semanticRole: meta.semanticRole,
            headerDetection: meta.headerDetection,
            normalizationReady: meta.normalizationReady
          };
        }),
        requiresSheetSelection: sheetNames.length > 0 && !primaryRabSheet,
        sheets: sheetsMeta
      };
    },

    /**
     * Memeriksa satu lembar kerja secara mendalam.
     * @param {string} sheetName
     * @param {Object} sheet
     * @param {boolean} isHidden
     * @returns {Object}
     */
    inspectSheet: function (sheetName, sheet, isHidden) {
      if (!sheet) {
        return {
          sheetName: sheetName,
          isEmpty: true,
          detectedRole: 'EMPTY',
          confidence: 0,
          clues: ['Lembar kerja kosong atau tidak dapat diakses']
        };
      }

      var rangeStr = sheet['!ref'] || 'A1:A1';
      var range = window.XLSX.utils.decode_range(rangeStr);

      var totalRows = range.e.r - range.s.r + 1;
      var totalCols = range.e.c - range.s.c + 1;
      var merges = sheet['!merges'] || [];

      var cellCount = 0;
      var formulaCount = 0;
      var numberCount = 0;
      var stringCount = 0;
      var errorCount = 0;

      var textSnippets = [];

      // Telusuri sel untuk mengumpulkan statistik
      for (var R = range.s.r; R <= range.e.r; R++) {
        for (var C = range.s.c; C <= range.e.c; C++) {
          var cellRef = window.XLSX.utils.encode_cell({ r: R, c: C });
          var cell = sheet[cellRef];
          if (!cell) continue;

          cellCount++;
          if (cell.f) formulaCount++;
          if (cell.t === 'n') numberCount++;
          if (cell.t === 's') {
            stringCount++;
            var sVal = (cell.v !== undefined ? String(cell.v).trim().toLowerCase() : '');
            if (sVal && textSnippets.length < 150) {
              textSnippets.push(sVal);
            }
          }
          if (cell.t === 'e') errorCount++;
        }
      }

      // Ambil sampel 15 baris pertama untuk visualisasi & pemetaan
      var sampleRows = [];
      var maxSampleRows = Math.min(range.e.r, range.s.r + 20);

      for (var rIdx = range.s.r; rIdx <= maxSampleRows; rIdx++) {
        var rowArr = [];
        var hasContent = false;
        for (var cIdx = range.s.c; cIdx <= range.e.c; cIdx++) {
          var cRef = window.XLSX.utils.encode_cell({ r: rIdx, c: cIdx });
          var cVal = sheet[cRef];
          if (cVal && cVal.v !== undefined && cVal.v !== null && String(cVal.v).trim() !== '') {
            hasContent = true;
            rowArr.push({
              colIndex: cIdx,
              colLetter: window.XLSX.utils.encode_col(cIdx),
              val: cVal.v,
              formatted: cVal.w || String(cVal.v),
              formula: cVal.f || null,
              type: cVal.t
            });
          } else {
            rowArr.push({
              colIndex: cIdx,
              colLetter: window.XLSX.utils.encode_col(cIdx),
              val: '',
              formatted: '',
              formula: null,
              type: 'z'
            });
          }
        }
        if (hasContent) {
          sampleRows.push({
            rowIndex: rIdx,
            row1Based: rIdx + 1,
            cells: rowArr
          });
        }
      }

      // Deteksi peran lembar kerja
      var roleAnalysis = this.analyzeRole(sheetName, textSnippets, formulaCount, numberCount, totalCols);
      var semanticRole = this.classifySemanticRole(sheetName, textSnippets, formulaCount, roleAnalysis.role);

      return {
        sheetName: sheetName,
        isHidden: !!isHidden,
        isEmpty: cellCount === 0,
        range: rangeStr,
        rowCount: totalRows,
        colCount: totalCols,
        cellCount: cellCount,
        formulaCount: formulaCount,
        numberCount: numberCount,
        stringCount: stringCount,
        errorCount: errorCount,
        mergeCount: merges.length,
        detectedRole: roleAnalysis.role,
        semanticRole: semanticRole.role,
        semanticRoleReason: semanticRole.reason,
        semanticConfidence: semanticRole.confidence,
        confidence: roleAnalysis.confidence,
        clues: roleAnalysis.clues,
        sampleRows: sampleRows
      };
    },

    classifySemanticRole: function (sheetName, textList, formulaCount, detectedRole) {
      var name = String(sheetName || '').toLowerCase();
      var text = textList.join(' ');
      var hasQtoFactors = /panjang/.test(text) && /lebar/.test(text) && /tinggi/.test(text) && /pengurang/.test(text);
      var hasVolume = /volume|kuantitas|qty/.test(text);
      if (/backup/.test(name) && hasQtoFactors && hasVolume && formulaCount > 0) {
        return { role: 'VOLUME_BACKUP', confidence: 95, reason: 'Struktur memiliki faktor dimensi, pengurang, volume, dan formula kuantitas.' };
      }
      if (/rekap|rekapitulasi|summary/.test(name) && (/total|jumlah|divisi|anggaran/.test(text))) return { role: 'RECAP', confidence: 90, reason: 'Nama sheet dan isi menunjukkan ringkasan total/divisi.' };
      if (/ahsp|analisa|analisis|harga satuan/.test(name) && detectedRole === 'AHSP') {
        return { role: 'AHSP_PRICE_REFERENCE', confidence: 90, reason: 'Nama sheet dan pola konten menunjukkan referensi analisa/harga.' };
      }
      if (/dashboard/.test(name) && (formulaCount > 0 || /total|progress|grafik|realisasi/.test(text))) return { role: 'DASHBOARD', confidence: 90, reason: 'Nama dan struktur isi menunjukkan ringkasan/dashboard.' };
      if (/basis|asumsi/.test(name) && /asumsi|parameter|basis|koefisien|faktor/.test(text)) return { role: 'PROJECT_ASSUMPTIONS', confidence: 85, reason: 'Nama dan isi menunjukkan parameter/asumsi proyek.' };
      if (/kurva|schedule|jadwal/.test(name) || detectedRole === 'KURVA_S') return { role: 'SCHEDULE', confidence: 85, reason: 'Nama atau pola kolom menunjukkan jadwal.' };
      var hasDescription = /uraian|deskripsi|pekerjaan|description/.test(text);
      var hasUnitPrice = /harga satuan|unit price|tarif/.test(text);
      // Many Indonesian BOQ tables label the monetary extension simply
      // "Jumlah". Require the unit-price and quantity signals as well so a
      // QTO count/factor column named Jumlah is not mistaken for money.
      var hasAmount = /jumlah|total harga|amount/.test(text);
      if (hasVolume && hasDescription && hasUnitPrice && hasAmount && detectedRole === 'RAB') {
        return { role: 'PRIMARY_BOQ', confidence: 90, reason: 'Header uraian, volume, harga satuan, dan jumlah harga menunjukkan BOQ finansial.' };
      }
      if (detectedRole === 'AHSP') return { role: 'AHSP_PRICE_REFERENCE', confidence: 70, reason: 'Pola komponen analisa harga terdeteksi.' };
      if (detectedRole === 'REKAP') return { role: 'RECAP', confidence: 70, reason: 'Pola rekapitulasi terdeteksi.' };
      return { role: 'UNKNOWN', confidence: 0, reason: 'Bukti struktur belum cukup untuk menetapkan peran sheet.' };
    },

    /**
     * Menganalisis peran lembar kerja berdasarkan multi-faktor.
     */
    analyzeRole: function (sheetName, textList, formulaCount, numberCount, colCount) {
      var sName = (sheetName || '').toLowerCase();
      var fullText = textList.join(' ');
      var clues = [];

      var scores = {
        RAB: 0,
        AHSP: 0,
        REKAP: 0,
        KURVA_S: 0,
        SUPPORTING: 0
      };

      // 1. Clues Nama Sheet
      if (/rab|anggaran|biaya|boq|bill|kuantitas/.test(sName)) {
        scores.RAB += 35;
        clues.push('Nama lembar mengindikasikan RAB/BOQ (' + sheetName + ')');
      }
      if (/ahsp|analisa|analisis|hsp|satuan/.test(sName)) {
        scores.AHSP += 35;
        clues.push('Nama lembar mengindikasikan AHSP (' + sheetName + ')');
      }
      if (/rekap|rekapitulasi|summary|ringkasan/.test(sName)) {
        scores.REKAP += 40;
        clues.push('Nama lembar mengindikasikan Rekapitulasi (' + sheetName + ')');
      }
      if (/kurva|jadwal|schedule|s-curve|progress/.test(sName)) {
        scores.KURVA_S += 40;
        clues.push('Nama lembar mengindikasikan Kurva S / Jadwal (' + sheetName + ')');
      }
      if (/upah|bahan|material|alat|master|daftar harga/.test(sName)) {
        scores.SUPPORTING += 30;
        clues.push('Nama lembar mengindikasikan Data Pendukung / Master Harga');
      }

      // 2. Clues Header & Kata Kunci Konten
      var hasUraian = /uraian|deskripsi|pekerjaan|description/.test(fullText);
      var hasVol = /volume|vol|qty|kuantitas/.test(fullText);
      var hasSat = /satuan|sat|unit/.test(fullText);
      var hasHarga = /harga satuan|unit price|h\.satuan|tarif/.test(fullText);
      var hasJumlah = /jumlah|total|amount|jumlah harga/.test(fullText);

      if (hasUraian && hasVol && (hasHarga || hasJumlah)) {
        scores.RAB += 45;
        clues.push('Ditemukan kombinasi header standar RAB: Uraian, Volume, Harga Satuan, Jumlah');
      }

      var hasKomponen = /tenaga|upah|bahan|material|alat|peralatan|koefisien|koef/.test(fullText);
      var hasOverhead = /overhead|profit|keuntungan/.test(fullText);

      if (hasKomponen && (hasOverhead || /kategori|komponen/.test(fullText))) {
        scores.AHSP += 45;
        clues.push('Ditemukan istilah analisis harga: Tenaga/Upah, Bahan, Alat, Koefisien, Overhead');
      }

      var hasDivisi = /divisi|bagian|rekapitulasi total|dibulatkan|grand total/.test(fullText);
      if (hasDivisi && !hasVol) {
        scores.REKAP += 35;
        clues.push('Ditemukan struktur rekapitulasi divisi tanpa kolom volume per item');
      }

      var hasMinggu = /minggu ke|week|bobot|kumulatif|rencana/.test(fullText);
      if (hasMinggu && colCount > 10) {
        scores.KURVA_S += 35;
        clues.push('Ditemukan sebaran kolom mingguan dan kata kunci jadwal/bobot');
      }

      // 3. Clues Formula & Angka
      if (formulaCount > 10) {
        scores.RAB += 10;
        scores.AHSP += 10;
      }

      // Tentukan peran tertinggi
      var bestRole = 'OTHER';
      var maxScore = 0;

      for (var r in scores) {
        if (scores[r] > maxScore) {
          maxScore = scores[r];
          bestRole = r;
        }
      }

      var confidence = Math.min(100, maxScore);

      if (confidence < 25) {
        bestRole = 'OTHER';
        clues.push('Pola tidak cocok dengan struktur standar RAB/AHSP');
      }

      return {
        role: bestRole,
        confidence: confidence,
        clues: clues
      };
    }
  };

  window.WorkbookInspector = WorkbookInspector;

})(typeof window !== 'undefined' ? window : this);
