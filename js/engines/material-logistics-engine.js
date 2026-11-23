/* Calculates theoretical resource demand from explicit BOQ-to-AHSP relations. */
(function (window) {
  'use strict';

  var ENGINE = 'MaterialLogisticsEngine', VERSION = '1.0.0';
  function str(value) { return value === null || value === undefined ? '' : String(value).trim(); }
  function norm(value) { return str(value).toLowerCase().normalize('NFKD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-z0-9]+/g, ' ').trim(); }
  function unit(value) {
    var n = norm(value).replace(/\s+/g, '');
    var aliases = { 'm³': 'm3', 'meterkubik': 'm3', 'm^3': 'm3', 'm²': 'm2', 'meterpersegi': 'm2', 'm^2': 'm2', 'meter': 'm', 'kilogram': 'kg', 'liter': 'l', 'ltr': 'l', 'ton': 't', 'tonne': 't', 'oh': 'oh' };
    return aliases[n] || n;
  }
  function finite(value) { return typeof value === 'number' && Number.isFinite(value); }
  function refs(item, filename) {
    var source = item && item.source || {}, result = [];
    Object.keys(source.cells || {}).forEach(function (field) {
      var cell = source.cells[field]; if (!cell || !cell.cellRef) return;
      result.push({ workbook: filename || null, sheet: source.sheet || null, row: source.row || null,
        cell: (source.sheet ? source.sheet + '!' : '') + cell.cellRef, field: field, itemId: item.id || null,
        value: cell.val !== undefined ? cell.val : null, formula: cell.formula || null });
    });
    if (!result.length && (source.sheet || source.row)) result.push({ workbook: filename || null, sheet: source.sheet || null, row: source.row || null, cell: null, itemId: item.id || null });
    return result;
  }
  function refForCell(item, field, filename) {
    var source = item && item.source || {}, cell = source.cells && (source.cells[field] || source.cells[field === 'volume' ? 'quantity' : field]);
    if (!cell && !source.sheet && !source.row) return [];
    return [{ workbook: filename || null, sheet: source.sheet || null, row: source.row || null,
      cell: cell && cell.cellRef ? (source.sheet ? source.sheet + '!' : '') + cell.cellRef : null,
      field: field, itemId: item.id || null, description: item.description || null, code: item.analysisCode || item.code || null,
      formula: cell && cell.formula || null, value: cell && cell.val !== undefined ? cell.val : (field === 'volume' ? item.volume : item[field]) }];
  }
  function overlap(a, b) {
    var left = norm(a).split(' ').filter(Boolean), right = norm(b).split(' ').filter(Boolean);
    if (!left.length || !right.length) return 0;
    var rightSet = Object.create(null); right.forEach(function (token) { rightSet[token] = true; });
    return left.filter(function (token) { return rightSet[token]; }).length / Math.max(left.length, right.length);
  }
  function pathKey(path) {
    return (Array.isArray(path) ? path : []).slice().sort(function (a, b) { return a.levelIndex - b.levelIndex; })
      .map(function (level) { return level && level.resolution !== 'ambiguous' ? norm(level.value) : '?'; }).join('/');
  }
  function newFinding(ruleId, severity, title, explanation, item, filename, candidates) {
    return { id: ENGINE + ':' + ruleId + ':' + (item.id || item.source && item.source.row || 'item'), ruleId: ruleId,
      severity: severity, interpretation: severity === 'ERROR' ? 'confirmed' : 'needs_review', title: title,
      explanation: explanation, affectedItems: [{ itemId: item.id || null, code: item.analysisCode || item.code || null,
        description: item.description || '', categoryPath: item.categoryPath || [], hierarchyStatus: item.hierarchyStatus || 'none' }],
      sourceReferences: refs(item, filename).concat(refForCell(item, 'volume', filename)), candidates: candidates || [],
      recommendedAction: 'Verifikasi kode analisa, satuan, volume, dan sumber AHSP sebelum memakai rekap kebutuhan.' };
  }
  function materialKey(resource) {
    return [norm(resource.resourceCode), norm(resource.resourceDescription), norm(resource.resourceSpecification), unit(resource.resourceUnit)].join('|');
  }
  function refFrom(item, filename) {
    return (item.sourceReferences || []).map(function (source) {
      var ref = Object.assign({}, source); if (!ref.workbook) ref.workbook = filename || null; return ref;
    });
  }
  function descriptionUnitKey(description, basisUnit) {
    return JSON.stringify([norm(description), unit(basisUnit)]);
  }
  function buildAnalysisIndex(analyses) {
    var byId = Object.create(null), byWorkCode = Object.create(null), byDescriptionUnit = Object.create(null);
    function add(index, key, analysis) {
      if (!key) return;
      if (!index[key]) index[key] = [];
      index[key].push(analysis);
    }
    analyses.forEach(function (analysis) {
      add(byId, analysis.analysisId, analysis);
      add(byWorkCode, norm(analysis.workCode), analysis);
      add(byDescriptionUnit, descriptionUnitKey(analysis.workDescription, analysis.basisUnit), analysis);
    });
    return { byId: byId, byWorkCode: byWorkCode, byDescriptionUnit: byDescriptionUnit };
  }
  function matchItem(item, analyses, mappingDecisions, analysisIndex) {
    analysisIndex = analysisIndex || buildAnalysisIndex(analyses);
    var selectedId = mappingDecisions && mappingDecisions[item.id];
    if (selectedId) {
      var selected = analysisIndex.byId[selectedId] || [];
      if (selected.length === 1) {
        var chosen = selected[0];
        if (item.unit && chosen.basisUnit && unit(item.unit) !== unit(chosen.basisUnit)) return { status: 'INCOMPATIBLE_UNIT', candidates: selected, reason: 'Pilihan manual memiliki satuan basis yang tidak kompatibel.', userSelected: true };
        if (finite(chosen.basisQuantity) && chosen.basisQuantity > 0 && chosen.basisUnit) return { status: 'HIGH_CONFIDENCE_MATCH', candidates: selected, analysis: chosen, reason: 'Pemetaan AHSP dipilih pengguna; canonical source tetap tidak berubah.', userSelected: true };
      }
    }
    var itemCode = norm(item.analysisCode || item.code), itemDescription = norm(item.description), itemUnit = unit(item.unit);
    var codeCandidates = itemCode ? (analysisIndex.byWorkCode[itemCode] || []).slice() : [];
    if (codeCandidates.length) {
      var samePath = codeCandidates.filter(function (analysis) {
        return !analysis.categoryPath || !item.categoryPath || !analysis.categoryPath.length || !item.categoryPath.length || pathKey(analysis.categoryPath) === pathKey(item.categoryPath);
      });
      if (samePath.length) codeCandidates = samePath;
      if (codeCandidates.length > 1) return { status: 'AMBIGUOUS', candidates: codeCandidates, reason: 'Beberapa AHSP memakai kode yang sama dan belum dapat dipilih tanpa keputusan sumber yang eksplisit.' };
      var candidate = codeCandidates[0];
      if (itemDescription && candidate.workDescription && overlap(itemDescription, candidate.workDescription) < 0.5) return { status: 'AMBIGUOUS', candidates: codeCandidates, reason: 'Kode analisa sama, tetapi uraian pekerjaan berbeda secara material.' };
      if (itemUnit && candidate.basisUnit && itemUnit !== unit(candidate.basisUnit)) return { status: 'INCOMPATIBLE_UNIT', candidates: codeCandidates, reason: 'Satuan item BOQ dan satuan dasar AHSP berbeda; tidak ada konversi yang terverifikasi.' };
      if (!candidate.basisUnit || !finite(candidate.basisQuantity) || candidate.basisQuantity <= 0) return { status: 'HIGH_CONFIDENCE_MATCH', candidates: codeCandidates, reason: 'Identitas kode cocok, tetapi basis AHSP tidak lengkap atau tidak valid untuk perhitungan.' };
      return { status: itemDescription && candidate.workDescription && itemDescription === norm(candidate.workDescription) ? 'EXACT_MATCH' : 'HIGH_CONFIDENCE_MATCH', candidates: codeCandidates, analysis: candidate, reason: null };
    }
    var descriptionCandidates = itemDescription && itemUnit
      ? (analysisIndex.byDescriptionUnit[descriptionUnitKey(itemDescription, itemUnit)] || []).slice() : [];
    if (descriptionCandidates.length) return { status: descriptionCandidates.length > 1 ? 'AMBIGUOUS' : 'POSSIBLE_MATCH', candidates: descriptionCandidates,
      reason: 'Uraian dan satuan tampak sesuai, tetapi kode analisa BOQ tidak membuktikan identitas yang sama.' };
    return { status: 'NOT_FOUND', candidates: [], reason: 'Tidak ditemukan AHSP dengan kode analisa yang cocok dan bukti relasi yang memadai.' };
  }

  var MaterialLogisticsEngine = {
    engineName: ENGINE, name: ENGINE, engineVersion: VERSION,
    analyze: function (canonical, options) {
      if (!canonical || !Array.isArray(canonical.workItems)) throw new Error('Canonical Data memerlukan workItems array.');
      options = options || {};
      var analyses = Array.isArray(canonical.ahspAnalyses) ? canonical.ahspAnalyses : [];
      var analysisIndex = buildAnalysisIndex(analyses);
      var workItems = canonical.workItems.filter(function (item) { return item && item.rowType !== 'SUBTOTAL' && item.isSubtotal !== true; });
      var filename = canonical.project && canonical.project.sourceFilename || null;
      var findings = [], mappings = [], workItemDetails = [], aggregates = Object.create(null), totalValid = 0, linked = 0, calculable = 0, noVolume = 0, ambiguous = 0, incompatible = 0, notFound = 0, missingBasis = 0, unknownResources = 0, invalidCoefficients = 0, invalidResourceInputs = 0;
      var totalValue = 0, coveredValue = 0, totalValueKnown = true;
      workItems.forEach(function (item) {
        if (finite(item.volume) && item.volume >= 0 && str(item.unit)) totalValid++;
        if (finite(item.amount) && item.amount >= 0) totalValue += item.amount; else totalValueKnown = false;
        var match = matchItem(item, analyses, options.mappingDecisions, analysisIndex), analysis = match.analysis || (match.candidates.length === 1 ? match.candidates[0] : null);
        var itemDetail = { workItemId: item.id || null, code: item.analysisCode || item.code || null, description: item.description || '',
          categoryPath: item.categoryPath || [], hierarchyStatus: item.hierarchyStatus || 'none', volume: finite(item.volume) ? item.volume : null,
          unit: item.unit || null, matchStatus: match.status, reason: match.reason, analysisId: analysis && analysis.analysisId || null,
          analysisCode: analysis && analysis.workCode || null, basisQuantity: analysis && finite(analysis.basisQuantity) ? analysis.basisQuantity : null,
          basisUnit: analysis && analysis.basisUnit || null, sourceReferences: refs(item, filename), ahspSourceReferences: analysis ? refFrom(analysis, filename) : [],
          resourceDetails: [], calculated: false };
        mappings.push({ workItemId: item.id || null, workCode: itemDetail.code, analysisId: itemDetail.analysisId,
          status: match.status, userSelected: !!match.userSelected, userSelectedAnalysisId: match.userSelected ? (match.analysis && match.analysis.analysisId || null) : null,
          candidates: match.candidates.map(function (candidate) { return { analysisId: candidate.analysisId, workCode: candidate.workCode,
            workDescription: candidate.workDescription, basisUnit: candidate.basisUnit, sourceReferences: refFrom(candidate, filename) }; }), reason: match.reason });
        if (match.status === 'NOT_FOUND') { notFound++; findings.push(newFinding('ahsp_not_found', 'WARNING', 'AHSP tidak ditemukan', match.reason, item, filename)); }
        if (match.status === 'AMBIGUOUS' || match.status === 'POSSIBLE_MATCH') { ambiguous++; findings.push(newFinding('ahsp_match_ambiguous', 'WARNING', 'Hubungan BOQ–AHSP perlu verifikasi', match.reason, item, filename, itemDetail && mappings[mappings.length - 1].candidates)); }
        if (match.status === 'INCOMPATIBLE_UNIT') { incompatible++; findings.push(newFinding('ahsp_unit_incompatible', 'WARNING', 'Satuan basis BOQ dan AHSP tidak kompatibel', match.reason, item, filename)); }
        if (!analysis || ['AMBIGUOUS', 'POSSIBLE_MATCH', 'INCOMPATIBLE_UNIT', 'NOT_FOUND'].indexOf(match.status) >= 0) { workItemDetails.push(itemDetail); return; }
        if (!finite(analysis.basisQuantity) || analysis.basisQuantity <= 0 || !analysis.basisUnit) {
          missingBasis++; findings.push(newFinding('ahsp_basis_invalid', 'WARNING', 'Basis AHSP tidak valid', 'Kuantitas basis harus eksplisit, positif, dan memiliki satuan sebelum koefisien dapat diterapkan.', item, filename));
          workItemDetails.push(itemDetail); return;
        }
        linked++;
        if (!finite(item.volume) || item.volume < 0 || !str(item.unit)) {
          noVolume++; findings.push(newFinding('boq_volume_unavailable', 'WARNING', 'Volume pekerjaan tidak dapat dihitung', 'Volume harus berupa angka tidak negatif dan satuan pekerjaan harus tersedia.', item, filename));
          workItemDetails.push(itemDetail); return;
        }
        var itemCalculable = false;
        (analysis.resources || []).forEach(function (resource) {
          var detail = { resourceId: resource.resourceId, resourceType: resource.resourceType || 'UNKNOWN', resourceCode: resource.resourceCode || null,
            description: resource.resourceDescription || '', specification: resource.resourceSpecification || null, coefficient: finite(resource.coefficient) ? resource.coefficient : null,
            unit: resource.resourceUnit || null, quantity: null, includedInMaterialTotal: false, reason: null,
            sourceReferences: refFrom(resource, filename) };
          if (resource.resourceType !== 'MATERIAL') detail.reason = 'Komponen bukan material terverifikasi; tidak dimasukkan ke total material.';
          else if (!finite(resource.coefficient) || resource.coefficient < 0 || !resource.resourceUnit) {
            detail.reason = 'Koefisien atau satuan material tidak valid.'; invalidResourceInputs++;
            if (!finite(resource.coefficient) || resource.coefficient < 0) invalidCoefficients++;
            findings.push({ id: ENGINE + ':ahsp_coefficient_invalid:' + resource.resourceId + ':' + item.id, ruleId: 'ahsp_coefficient_invalid', severity: 'WARNING',
              interpretation: 'inconclusive', title: 'Koefisien material tidak dapat digunakan', explanation: detail.reason,
              affectedItems: [{ itemId: item.id || null, code: item.analysisCode || item.code || null, description: item.description || '' }],
              sourceReferences: refs(item, filename).concat(detail.sourceReferences), recommendedAction: 'Periksa nilai koefisien dan satuan pada sel sumber AHSP.' });
          }
          else {
            detail.quantity = item.volume / analysis.basisQuantity * resource.coefficient;
            if (!Number.isFinite(detail.quantity)) { detail.quantity = null; detail.reason = 'Hasil aritmetika tidak valid.'; }
            else {
              detail.includedInMaterialTotal = true; itemCalculable = true;
              var key = materialKey(resource);
              if (!aggregates[key]) aggregates[key] = { materialId: key, resourceCode: resource.resourceCode || null,
                description: resource.resourceDescription || '', specification: resource.resourceSpecification || null,
                unit: resource.resourceUnit, totalQuantity: 0, contributingWorkItems: 0, sourceReferences: [], resourceTypes: ['MATERIAL'] };
              aggregates[key].totalQuantity += detail.quantity;
              aggregates[key].contributingWorkItems += 1;
              aggregates[key].sourceReferences = aggregates[key].sourceReferences.concat(detail.sourceReferences);
            }
          }
          if (resource.resourceType === 'UNKNOWN') {
            unknownResources++;
            findings.push({ id: ENGINE + ':resource_unknown:' + resource.resourceId + ':' + item.id, ruleId: 'ahsp_resource_unknown', severity: 'WARNING',
              interpretation: 'inconclusive', title: 'Jenis resource AHSP belum diketahui', explanation: 'Resource dipertahankan sebagai bukti tetapi tidak dimasukkan ke kebutuhan material.',
              affectedItems: [{ itemId: item.id || null, code: item.analysisCode || item.code || null, description: item.description || '' }],
              sourceReferences: refs(item, filename).concat(detail.sourceReferences), recommendedAction: 'Tinjau kategori dan deskripsi resource pada sumber AHSP.' });
          }
          itemDetail.resourceDetails.push(detail);
        });
        if (itemCalculable) { calculable++; itemDetail.calculated = true; coveredValue += finite(item.amount) && item.amount >= 0 ? item.amount : 0; }
        workItemDetails.push(itemDetail);
      });
      var materials = Object.keys(aggregates).map(function (key) { return aggregates[key]; });
      var totalItems = workItems.length, coverage = { totalWorkItems: totalItems, validVolumeItems: totalValid, linkedWorkItems: linked,
        calculableWorkItems: calculable, notFound: notFound, ambiguous: ambiguous, incompatibleUnit: incompatible, volumeUnavailable: noVolume,
        unknownResources: unknownResources, invalidCoefficients: invalidCoefficients, invalidResourceInputs: invalidResourceInputs,
        invalidBasis: missingBasis, itemCoverage: totalItems ? calculable / totalItems : 0,
        totalBoqAmount: totalValueKnown ? totalValue : null, coveredBoqAmount: totalValueKnown ? coveredValue : null,
        valueCoverage: totalValueKnown && totalValue !== 0 ? coveredValue / totalValue : (totalValueKnown && totalValue === 0 ? null : null) };
      var partial = totalItems === 0 || calculable !== totalItems || unknownResources > 0 || invalidResourceInputs > 0 || (canonical.ahspWarnings && canonical.ahspWarnings.length);
      var result = { engineName: ENGINE, engineVersion: VERSION, schemaVersion: 'rab-guard-engine-result/1', status: partial ? 'partial' : 'complete',
        summary: { workItems: totalItems, validVolumeItems: totalValid, matchedItems: linked, calculableItems: calculable, materialCount: materials.length,
          noAhspItems: notFound, ambiguousItems: ambiguous, incompatibleUnitItems: incompatible, totalTheoreticalMaterialContributions: materials.length },
        findings: findings, metrics: { formula: 'boqWorkQuantity / ahspBasisQuantity * resourceCoefficient' }, materials: materials,
        workItemDetails: workItemDetails, mappings: mappings, coverage: coverage,
        sourceReferences: workItemDetails.reduce(function (all, detail) { return all.concat(detail.sourceReferences || [], detail.ahspSourceReferences || []); }, [])
          .concat(materials.reduce(function (all, material) { return all.concat(material.sourceReferences); }, []))
          .concat(canonical.ahspSourceReferences || []),
        assumptions: ['Hasil adalah kebutuhan teoritis berdasarkan koefisien AHSP yang diekstrak; tidak termasuk waste, stok, atau rencana pengadaan.'],
        limitations: ['Hanya AHSP dengan kode, tabel, satuan, dan koefisien yang terdeteksi eksplisit yang digunakan.',
          'Pencocokan tanpa kode tidak dilakukan otomatis; alternatif dan kecocokan lemah tetap memerlukan verifikasi.',
          'Konversi satuan material tidak dilakukan tanpa faktor konversi terverifikasi.',
          'Ekstraksi dan klasifikasi resource berbasis bukti struktur/header dan istilah yang dikenal; variasi workbook belum dijamin tercakup.'],
        warnings: (canonical.ahspWarnings || []).slice(), errors: [] };
      return result;
    }
  };
  window.MaterialLogisticsEngine = MaterialLogisticsEngine;
})(typeof window !== 'undefined' ? window : this);
