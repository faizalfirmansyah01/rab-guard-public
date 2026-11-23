/* Conservative AHSP table extraction into the Universal canonical model. */
(function (window) {
  'use strict';

  function text(value) { return value === null || value === undefined ? '' : String(value).trim(); }
  function key(value) { return text(value).toLowerCase().normalize('NFKD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-z0-9]+/g, ' ').trim(); }
  function numeric(value) {
    if (typeof value === 'number' && Number.isFinite(value)) return value;
    if (typeof value === 'string' && window.DataNormalizer && window.DataNormalizer.parseNumericValue) return window.DataNormalizer.parseNumericValue(value);
    return null;
  }
  var ALIASES = {
    analysisCode: ['kode analisa', 'kode analisis', 'kode ahsp', 'analysis code', 'ahsp code', 'kode pekerjaan'],
    analysisDescription: ['uraian pekerjaan', 'nama pekerjaan', 'work description'],
    code: ['kode', 'code', 'no analisa'],
    resourceCode: ['kode bahan', 'kode material', 'kode sumber daya', 'kode resource', 'kode komponen', 'resource code', 'material code'],
    description: ['uraian', 'uraian pekerjaan', 'deskripsi', 'description', 'nama pekerjaan', 'nama bahan', 'komponen'],
    resourceDescription: ['uraian bahan', 'uraian sumber daya', 'uraian komponen', 'nama bahan', 'nama material', 'resource description', 'material description'],
    unit: ['satuan', 'sat', 'unit'],
    resourceUnit: ['satuan bahan', 'satuan resource', 'satuan sumber daya', 'unit bahan', 'resource unit'],
    coefficient: ['koefisien', 'koef', 'koef.','indeks', 'kebutuhan', 'kuantitas koefisien', 'coefficient'],
    basisQuantity: ['volume dasar', 'kuantitas dasar', 'basis quantity', 'jumlah dasar', 'per'],
    basisUnit: ['satuan dasar', 'unit dasar', 'basis unit', 'satuan pekerjaan', 'unit pekerjaan', 'satuan analisa'],
    group: ['kelompok', 'jenis', 'kategori', 'golongan', 'resource type', 'jenis sumber daya'],
    specification: ['spesifikasi', 'spesifikasi bahan', 'mutu', 'ukuran', 'specification']
  };
  function identify(header) {
    var normalized = key(header);
    if (!normalized) return null;
    var names = Object.keys(ALIASES);
    for (var i = 0; i < names.length; i++) {
      if (ALIASES[names[i]].some(function (alias) { return normalized === key(alias); })) return names[i];
    }
    return null;
  }
  function isRepeatedHeaderRow(sheet, row, fields) {
    var hits = 0;
    Object.keys(fields).forEach(function (field) {
      if (identify(cellValue(sheet, row, fields[field])) === field) hits++;
    });
    return hits >= 3;
  }
  function columnRef(sheet, row, col) {
    return window.XLSX && window.XLSX.utils ? window.XLSX.utils.encode_cell({ r: row, c: col }) : String(row + 1) + ':' + String(col + 1);
  }
  function cellValue(sheet, row, col) {
    var cell = sheet[columnRef(sheet, row, col)];
    return cell ? cell.v : null;
  }
  function cellNumeric(sheet, row, col) {
    var cell = sheet[columnRef(sheet, row, col)];
    if (!cell || cell.t === 'e' || cell.v === null || cell.v === undefined) return null;
    return numeric(cell.v);
  }
  function headerValue(sheet, row, col) {
    var value = cellValue(sheet, row, col);
    if (nonEmpty(value)) return value;
    (sheet['!merges'] || []).some(function (merge) {
      if (row < merge.s.r || row > merge.e.r || col < merge.s.c || col > merge.e.c) return false;
      value = cellValue(sheet, merge.s.r, merge.s.c); return true;
    });
    return value;
  }
  function cellRef(sheetName, sheet, row, col, field) {
    var ref = columnRef(sheet, row, col), cell = sheet[ref];
    return { field: field, workbook: null, sheet: sheetName, row: row + 1, cell: sheetName + '!' + ref,
      cellRef: ref, value: cell ? cell.v : null, cachedValue: cell && cell.f ? (cell.v !== undefined ? cell.v : null) : null, formula: cell && cell.f || null,
      cellType: cell && cell.t || null, valueStatus: cell && cell.t === 'e' ? 'error' : (cell && cell.v !== undefined ? 'valid' : 'missing'),
      errorCode: cell && cell.t === 'e' ? cell.v : null, errorText: cell && cell.t === 'e' ? (cell.w || null) : null };
  }
  function classifyResource(group, description, code) {
    var evidence = key(group) + ' ' + key(description) + ' ' + key(code);
    if (/\b(lain lain|lainnya|other)\b/.test(key(group))) return { type: 'OTHER', evidence: group };
    if (/\b(upah|tenaga kerja|tenaga|pekerja|tukang|mandor|operator|oh)\b/.test(evidence)) return { type: 'LABOR', evidence: group ? group : description };
    if (/\b(peralatan|alat|excavator|molen|mixer|vibrator|dump truck|truck|crane)\b/.test(evidence)) return { type: 'EQUIPMENT', evidence: group ? group : description };
    if (/\b(bahan|material|semen|pasir|kerikil|split|agregat|besi|baja|beton|bata|pipa|kabel|cat|aspal|kayu|paku|genteng|air)\b/.test(evidence)) return { type: 'MATERIAL', evidence: group ? group : description };
    return { type: 'UNKNOWN', evidence: null };
  }
  function nonEmpty(value) { return value !== null && value !== undefined && text(value) !== ''; }

  var AHSPNormalizer = {
    schemaVersion: 'rab-guard-ahsp/1',
    normalize: function (rawWb, manifest) {
      var output = [], warnings = [], sourceReferences = [];
      (manifest || []).filter(function (entry) { return entry.role === 'AHSP_PRICE_REFERENCE' || entry.detectedRole === 'AHSP'; }).forEach(function (entry) {
        var sheet = rawWb && rawWb.sheets && rawWb.sheets[entry.name];
        if (!sheet || !window.XLSX || !window.XLSX.utils) return;
        var range;
        try { range = window.XLSX.utils.decode_range(sheet['!ref'] || 'A1:A1'); } catch (_) { return; }
        var scanEnd = Math.min(range.e.r, range.s.r + 59), winner = null;
        for (var r = range.s.r; r <= scanEnd; r++) {
          var fields = {}, score = 0, headerEndRow = r;
          for (var c = range.s.c; c <= range.e.c; c++) {
            var parts = [];
            for (var hr = r; hr <= Math.min(r + 1, scanEnd); hr++) {
              var hv = text(headerValue(sheet, hr, c)); if (hv) parts.push(hv);
            }
            var field = null, fieldRow = r;
            for (var pi = 0; pi < parts.length; pi++) { field = identify(parts[pi]); if (field) { fieldRow = r + pi; break; } }
            if (!field) { field = identify(parts.join(' ')); if (field) fieldRow = Math.min(r + 1, scanEnd); }
            if (field && fields[field] === undefined) {
              fields[field] = c; score += field === 'description' || field === 'analysisDescription' || field === 'resourceDescription' || field === 'coefficient' || field === 'unit' ? 3 : 1;
              headerEndRow = Math.max(headerEndRow, fieldRow);
            }
          }
          if ((fields.description !== undefined || fields.analysisDescription !== undefined || fields.resourceDescription !== undefined) && fields.coefficient !== undefined &&
            (fields.unit !== undefined || fields.resourceUnit !== undefined || fields.basisUnit !== undefined) && score >= 7 && (!winner || score > winner.score)) {
            winner = { row: r, endRow: headerEndRow, fields: fields, score: score };
          }
        }
        if (!winner) { warnings.push('Sheet ' + entry.name + ': tabel AHSP tidak dinormalisasi karena header uraian, koefisien, dan satuan belum teridentifikasi bersama.'); return; }
        entry.ahspHeaderRow = winner.row + 1;
        entry.ahspHeaderRowEnd = winner.endRow + 1;
        var current = null, currentGroup = '';
        for (var row = winner.endRow + 1; row <= range.e.r; row++) {
          var f = winner.fields;
          if (isRepeatedHeaderRow(sheet, row, f)) continue;
          var codeCol = f.analysisCode !== undefined ? f.analysisCode : f.code;
          var analysisCode = codeCol !== undefined ? text(cellValue(sheet, row, codeCol)) : '';
          var analysisDescriptionCol = f.analysisDescription !== undefined ? f.analysisDescription : (f.description !== undefined ? f.description : f.resourceDescription);
          var resourceDescriptionCol = f.resourceDescription !== undefined ? f.resourceDescription : (f.description !== undefined ? f.description : f.analysisDescription);
          var description = analysisDescriptionCol !== undefined ? text(cellValue(sheet, row, analysisDescriptionCol)) : '';
          var resourceDescription = resourceDescriptionCol !== undefined ? text(cellValue(sheet, row, resourceDescriptionCol)) : description;
          var coefficient = cellNumeric(sheet, row, f.coefficient);
          var resourceUnitCol = f.resourceUnit !== undefined ? f.resourceUnit : (f.unit !== undefined ? f.unit : f.basisUnit);
          var analysisUnitCol = f.basisUnit !== undefined ? f.basisUnit : (f.unit !== undefined ? f.unit : f.resourceUnit);
          var unit = resourceUnitCol !== undefined ? text(cellValue(sheet, row, resourceUnitCol)) : '';
          var basisQuantity = f.basisQuantity !== undefined ? cellNumeric(sheet, row, f.basisQuantity) : null;
          var basisUnit = analysisUnitCol !== undefined ? text(cellValue(sheet, row, analysisUnitCol)) : '';
          var rowGroup = f.group !== undefined ? text(cellValue(sheet, row, f.group)) : '';
          var resourceCodeForRow = f.resourceCode !== undefined ? text(cellValue(sheet, row, f.resourceCode)) : '';
          var rowHasResource = nonEmpty(resourceDescription) &&
            (Number.isFinite(coefficient) || nonEmpty(resourceCodeForRow) || /^(bahan|material|tenaga kerja|upah|peralatan|alat)$/i.test(rowGroup));
          var rowHasAnalysis = nonEmpty(analysisCode) && nonEmpty(description) && !rowHasResource && (nonEmpty(unit) || basisQuantity !== null);
          if (!rowHasResource && !analysisCode && description && /^(bahan|material|tenaga kerja|upah|peralatan|alat)$/i.test(description)) { currentGroup = description; continue; }
          if (rowHasAnalysis) {
            var analysisRef = [codeCol, analysisDescriptionCol, analysisUnitCol, f.basisQuantity].filter(function (col) { return col !== undefined; }).map(function (col) { return cellRef(entry.name, sheet, row, col, col === codeCol ? 'workCode' : col === analysisDescriptionCol ? 'workDescription' : col === analysisUnitCol ? 'basisUnit' : 'basisQuantity'); });
            current = { analysisId: entry.name + ':' + (analysisCode || 'row-' + (row + 1)), workCode: analysisCode,
              workDescription: description, basisQuantity: basisQuantity, basisUnit: basisUnit || unit,
              sourceReferences: analysisRef, resources: [], source: { sheet: entry.name, row: row + 1 } };
            output.push(current); sourceReferences = sourceReferences.concat(analysisRef);
            currentGroup = '';
            continue;
          }
          if (!rowHasResource) continue;
          var explicitCodeCol = f.resourceCode;
          var resourceCode = explicitCodeCol !== undefined ? text(cellValue(sheet, row, explicitCodeCol)) : '';
          var group = rowGroup || currentGroup;
          var specification = f.specification !== undefined ? text(cellValue(sheet, row, f.specification)) : '';
          var type = classifyResource(group, resourceDescription, resourceCode);
          if (f.analysisCode !== undefined && analysisCode && (!current || norm(current.workCode) !== norm(analysisCode))) {
            current = { analysisId: entry.name + ':' + analysisCode, workCode: analysisCode, workDescription: '', basisQuantity: basisQuantity, basisUnit: basisUnit,
              sourceReferences: [], resources: [], source: { sheet: entry.name, row: row + 1 } };
            output.push(current);
          } else if (!current && codeCol !== undefined && analysisCode) {
            current = { analysisId: entry.name + ':' + analysisCode, workCode: analysisCode, workDescription: '', basisQuantity: null, basisUnit: '', sourceReferences: [], resources: [], source: { sheet: entry.name, row: row + 1 } };
            output.push(current);
          }
          if (!current) { warnings.push('Sheet ' + entry.name + ', baris ' + (row + 1) + ': komponen AHSP tidak memiliki identitas analisis yang dapat ditautkan.'); continue; }
          var cols = [explicitCodeCol, resourceDescriptionCol, f.coefficient, resourceUnitCol, f.group, f.specification].filter(function (col) { return col !== undefined; });
          var refs = cols.map(function (col) {
            var field = col === explicitCodeCol ? 'resourceCode' : col === resourceDescriptionCol ? 'resourceDescription' : col === f.coefficient ? 'coefficient' : col === resourceUnitCol ? 'resourceUnit' : col === f.group ? 'resourceGroup' : 'specification';
            return cellRef(entry.name, sheet, row, col, field);
          });
          var resource = { resourceId: entry.name + ':' + columnRef(sheet, row, resourceDescriptionCol), resourceType: type.type,
            resourceTypeEvidence: type.evidence, resourceCode: resourceCode || null, resourceDescription: resourceDescription,
            resourceSpecification: specification || null, coefficient: coefficient, resourceUnit: unit,
            sourceReferences: refs, source: { sheet: entry.name, row: row + 1 } };
          current.resources.push(resource); sourceReferences = sourceReferences.concat(refs);
        }
      });
      return { schemaVersion: this.schemaVersion, analyses: output, sourceReferences: sourceReferences, warnings: warnings };
    }
  };
  window.AHSPNormalizer = AHSPNormalizer;
})(typeof window !== 'undefined' ? window : this);
