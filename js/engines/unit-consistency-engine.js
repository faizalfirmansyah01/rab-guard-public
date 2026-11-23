/* Conservative unit interpretation and exact-identity consistency review. */
(function (window) {
  'use strict';

  var UNIT_DEFINITIONS = [
    { canonicalUnit: 'm', dimension: 'length', aliases: ['m', 'm1', 'meter', 'meters', 'meter lari', 'meter linear', 'mtr'] },
    { canonicalUnit: 'cm', dimension: 'length', aliases: ['cm', 'centimeter', 'centimeter'] },
    { canonicalUnit: 'mm', dimension: 'length', aliases: ['mm', 'millimeter', 'milimeter'] },
    { canonicalUnit: 'm2', dimension: 'area', aliases: ['m2', 'm^2', 'm²', 'meter persegi', 'mtr2', 'sqm'] },
    { canonicalUnit: 'cm2', dimension: 'area', aliases: ['cm2', 'cm^2', 'cm²', 'centimeter persegi'] },
    { canonicalUnit: 'm3', dimension: 'volume', aliases: ['m3', 'm^3', 'm³', 'meter kubik', 'mtr3', 'cum'] },
    { canonicalUnit: 'cm3', dimension: 'volume', aliases: ['cm3', 'cm^3', 'cm³', 'centimeter kubik'] },
    { canonicalUnit: 'kg', dimension: 'mass', aliases: ['kg', 'kgs', 'kilogram', 'kilograms'] },
    { canonicalUnit: 'g', dimension: 'mass', aliases: ['g', 'gram', 'grams'] },
    { canonicalUnit: 'ton', dimension: 'mass', aliases: ['ton', 'tonne', 'metric ton'] },
    { canonicalUnit: 'l', dimension: 'volume', aliases: ['l', 'lt', 'liter', 'litre', 'liter(s)'] },
    { canonicalUnit: 'jam', dimension: 'time', aliases: ['jam', 'hour', 'hr', 'hours'] },
    { canonicalUnit: 'hari', dimension: 'time', aliases: ['hari', 'day', 'days'] },
    { canonicalUnit: 'oh', dimension: 'labor_time', aliases: ['oh', 'orang-hari', 'orang hari', 'man-day', 'manday'] },
    { canonicalUnit: 'buah', dimension: 'count', aliases: ['buah', 'bh', 'pcs', 'pc', 'piece', 'pieces'] },
    { canonicalUnit: 'unit', dimension: 'count', aliases: ['unit', 'units'] },
    { canonicalUnit: 'set', dimension: 'set_count', aliases: ['set', 'sets'] },
    { canonicalUnit: 'titik', dimension: 'point_count', aliases: ['titik', 'point', 'points'] },
    { canonicalUnit: 'ls', dimension: 'lump_sum', aliases: ['ls', 'l.s', 'lump sum', 'lump-sum'] }
  ];

  var UNIT_INDEX = Object.create(null);
  UNIT_DEFINITIONS.forEach(function (definition) {
    definition.aliases.forEach(function (alias) {
      UNIT_INDEX[normalizeLabel(alias)] = definition;
    });
  });

  function normalizeLabel(value) {
    return String(value === null || value === undefined ? '' : value)
      .trim().toLowerCase().replace(/²/g, '2').replace(/³/g, '3')
      .replace(/\s+/g, ' ');
  }

  function copyValue(value) {
    if (value === null || typeof value !== 'object') return value;
    if (Array.isArray(value)) return value.map(copyValue);
    var copy = {};
    Object.keys(value).forEach(function (key) { copy[key] = copyValue(value[key]); });
    return copy;
  }

  function sourceReference(item, rawUnit) {
    var source = item.source || {};
    var unitCell = source.cells && source.cells.unit || null;
    var sheet = source.sheet || null;
    var row = Number.isFinite(source.row) ? source.row : null;
    var localCell = unitCell && unitCell.cellRef || null;
    return {
      itemId: item.id || null,
      sheet: sheet,
      row: row,
      field: 'unit',
      cell: localCell ? (sheet ? sheet + '!' + localCell : localCell) : null,
      originalValue: rawUnit === undefined ? null : rawUnit,
      valueStatus: unitCell && unitCell.valueStatus || (rawUnit === null || rawUnit === '' ? 'missing' : 'valid'),
      errorCode: unitCell && unitCell.errorCode !== undefined ? unitCell.errorCode : null,
      errorText: unitCell && unitCell.errorText || null,
      formula: unitCell && unitCell.formula || null
    };
  }

  function classifyUnit(value, sourceStatus) {
    var raw = value === null || value === undefined ? '' : String(value);
    var normalized = normalizeLabel(raw);
    if (sourceStatus === 'error') return { raw: raw, normalized: normalized, status: 'source_error', canonicalUnit: null, dimension: null };
    if (!normalized) return { raw: raw, normalized: '', status: 'missing', canonicalUnit: null, dimension: null };
    var definition = UNIT_INDEX[normalized];
    if (!definition) return { raw: raw, normalized: normalized, status: 'unrecognized', canonicalUnit: null, dimension: null };
    return { raw: raw, normalized: normalized, status: 'recognized', canonicalUnit: definition.canonicalUnit, dimension: definition.dimension };
  }

  function stableId(ruleId, references) {
    var input = ruleId + '|' + references.map(function (ref) { return [ref.sheet, ref.row, ref.cell, ref.originalValue].join(':'); }).join('|');
    var hash = 2166136261;
    for (var i = 0; i < input.length; i++) hash = Math.imul(hash ^ input.charCodeAt(i), 16777619);
    return ruleId + '-' + (hash >>> 0).toString(16);
  }

  function resultBase(status, errors) {
    return {
      engineName: 'UnitConsistencyEngine',
      engineVersion: '1.0.0',
      schemaVersion: 'rab-guard-engine-result/1',
      status: status,
      summary: {},
      findings: [],
      metrics: {},
      sourceReferences: [],
      warnings: [],
      errors: errors || [],
      limitations: [
        'Unit tidak disimpulkan dari kata-kata deskripsi pekerjaan saja.',
        'Tidak ada konversi kuantitas otomatis; unit berbeda tetap ditampilkan dalam unit sumber.',
        'Model AHSP/koefisien material belum tersedia pada kontrak canonical saat ini; validasi unit pekerjaan terhadap unit bahan tidak dilakukan.'
      ]
    };
  }

  function makeRelatedItem(record) {
    return {
      itemId: record.item.id || null,
      code: record.item.code || '',
      description: record.item.description || '',
      unit: record.unit.raw,
      normalizedUnit: record.unit.canonicalUnit,
      dimension: record.unit.dimension,
      sheet: record.source.sheet,
      row: record.source.row,
      sourceCell: record.source.cell,
      categoryPath: copyValue(Array.isArray(record.item.categoryPath) ? record.item.categoryPath : []),
      hierarchyStatus: record.item.hierarchyStatus || record.item.hierarchy && record.item.hierarchy.status || (record.item.categoryPath && record.item.categoryPath.length ? 'unresolved' : 'none')
    };
  }

  function makeFinding(ruleId, severity, interpretation, title, explanation, records, evidence, action) {
    var references = records.map(function (record) { return record.source; });
    return {
      id: stableId(ruleId, references),
      ruleId: ruleId,
      severity: severity,
      interpretation: interpretation,
      title: title,
      explanation: explanation,
      evidence: evidence,
      relatedItems: records.map(makeRelatedItem),
      sourceReferences: references.map(copyValue),
      recommendedAction: action
    };
  }

  var UnitConsistencyEngine = {
    name: 'UnitConsistencyEngine',
    engineName: 'UnitConsistencyEngine',
    engineVersion: '1.0.0',

    classifyUnit: function (unit, sourceStatus) {
      return classifyUnit(unit, sourceStatus);
    },

    analyze: function (canonicalData) {
      if (!canonicalData || !Array.isArray(canonicalData.workItems)) {
        return resultBase('failed', ['Canonical Data Contract v1 memerlukan workItems berupa array.']);
      }

      var result = resultBase('complete');
      var records = canonicalData.workItems.map(function (item) {
        item = item || {};
        var rawUnit = item.unit;
        var unitCell = item.source && item.source.cells && item.source.cells.unit;
        var unit = classifyUnit(rawUnit, unitCell && unitCell.valueStatus);
        var source = sourceReference(item, rawUnit);
        result.sourceReferences.push(copyValue(source));
        return { item: item, unit: unit, source: source };
      });

      var counts = { recognizedUnits: 0, missingUnits: 0, unrecognizedUnits: 0, sourceErrorUnits: 0 };
      var groupIndex = Object.create(null);
      records.forEach(function (record) {
        counts[record.unit.status === 'recognized' ? 'recognizedUnits' :
          (record.unit.status === 'missing' ? 'missingUnits' :
            (record.unit.status === 'source_error' ? 'sourceErrorUnits' : 'unrecognizedUnits'))]++;

        if (record.unit.status === 'unrecognized') {
          result.findings.push(makeFinding(
            'unit_unrecognized', 'INFO', 'inconclusive', 'Penulisan satuan belum dikenal',
            'Satuan “' + record.unit.raw + '” belum ada pada registry engine. Engine tidak menyimpulkan dimensi atau melakukan konversi.',
            [record], [{ unit: record.unit.raw, normalizedLabel: record.unit.normalized }],
            'Verifikasi label satuan pada workbook dan tambahkan alias hanya setelah maknanya dipastikan.'
          ));
        }

        if (record.unit.status === 'recognized' && record.item.code && record.item.description) {
          var code = String(record.item.code).trim().toUpperCase();
          var description = String(record.item.description).trim().toLowerCase().replace(/\s+/g, ' ');
          var key = JSON.stringify([code, description]);
          if (!groupIndex[key]) groupIndex[key] = [];
          groupIndex[key].push(record);
        }
      });

      var comparisonGroups = 0;
      var conflictGroups = 0;
      var variantGroups = 0;
      Object.keys(groupIndex).sort().forEach(function (key) {
        var group = groupIndex[key];
        if (group.length < 2) return;
        comparisonGroups++;
        var dimensions = Array.from(new Set(group.map(function (record) { return record.unit.dimension; })));
        var canonicalUnits = Array.from(new Set(group.map(function (record) { return record.unit.canonicalUnit; })));
        var evidence = group.map(function (record) {
          return { unit: record.unit.raw, canonicalUnit: record.unit.canonicalUnit, dimension: record.unit.dimension };
        });
        if (dimensions.length > 1) {
          conflictGroups++;
          result.findings.push(makeFinding(
            'unit_code_dimension_conflict', 'WARNING', 'needs_review', 'Dimensi satuan berbeda pada pekerjaan berkode dan beruraian sama',
            'Kode dan uraian yang sama memiliki satuan dengan dimensi berbeda. Ini adalah indikasi untuk diperiksa; perbedaan kategori atau penggunaan ulang kode dapat menjelaskan variasi tersebut.',
            group, evidence, 'Periksa baris dan kategori sumber untuk memastikan pekerjaan yang dibandingkan benar-benar sama. Jangan mengubah nilai satuan sebelum verifikasi.'
          ));
        } else if (canonicalUnits.length > 1) {
          variantGroups++;
          result.findings.push(makeFinding(
            'unit_code_unit_variant', 'INFO', 'needs_review', 'Unit sejenis berbeda pada pekerjaan berkode dan beruraian sama',
            'Unit-unit ini tercatat pada dimensi yang sama tetapi bukan alias yang identik. Engine tidak mengonversi nilai; basis kuantitas perlu diverifikasi.',
            group, evidence, 'Verifikasi basis satuan pada sumber. Jika diperlukan, lakukan konversi terpisah dengan faktor yang terdokumentasi.'
          ));
        }
      });

      if (counts.missingUnits) {
        result.status = 'partial';
        result.warnings.push(counts.missingUnits + ' item tidak memiliki label satuan; tidak ada dimensi yang diasumsikan. Universal Checker tetap menjadi pemilik temuan missing_unit.');
      }
      if (counts.unrecognizedUnits) {
        result.status = 'partial';
        result.warnings.push(counts.unrecognizedUnits + ' label satuan belum dikenali oleh registry.');
      }
      if (counts.sourceErrorUnits) {
        result.status = 'partial';
        result.warnings.push(counts.sourceErrorUnits + ' sel satuan memiliki status error sumber dan tidak digunakan sebagai label unit.');
      }

      result.summary = {
        itemsAnalyzed: records.length,
        recognizedUnits: counts.recognizedUnits,
        missingUnits: counts.missingUnits,
        unrecognizedUnits: counts.unrecognizedUnits,
        sourceErrorUnits: counts.sourceErrorUnits,
        comparisonGroups: comparisonGroups,
        conflictGroups: conflictGroups,
        variantGroups: variantGroups,
        findingsCount: result.findings.length
      };
      result.metrics = Object.assign({}, result.summary);
      return result;
    }
  };

  window.UnitConsistencyEngine = UnitConsistencyEngine;
})(typeof window !== 'undefined' ? window : this);
