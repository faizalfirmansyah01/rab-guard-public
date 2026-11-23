/* Conservative, source-traceable comparison of two Universal canonical workbooks. */
(function (window) {
  'use strict';

  var VERSION = '1.0.0';
  var RESULT_SCHEMA = 'rab-guard-engine-result/1';

  function text(value) { return value === null || value === undefined ? '' : String(value).trim(); }
  function normalizeCode(value) { return text(value).toUpperCase().replace(/\s+/g, ' '); }
  function normalizeDescription(value) {
    return text(value).toLowerCase().normalize ? text(value).toLowerCase().normalize('NFKD').replace(/[\u0300-\u036f]/g, '').replace(/[^a-z0-9]+/g, ' ').trim() : text(value).toLowerCase().replace(/[^a-z0-9]+/g, ' ').trim();
  }
  function normalizeUnit(value) { return text(value).toLowerCase().replace(/²/g, '2').replace(/³/g, '3').replace(/\s+/g, ' '); }
  function pathValues(item) {
    return (Array.isArray(item.categoryPath) ? item.categoryPath : []).slice().sort(function (a, b) {
      return Number(a.levelIndex || 0) - Number(b.levelIndex || 0);
    }).map(function (level) { return { levelIndex: level.levelIndex, value: text(level.value), sourceLabel: level.sourceLabel || null, provenance: level.provenance || null }; });
  }
  function pathSignature(item) { return pathValues(item).map(function (level) { return level.levelIndex + ':' + level.value.toLowerCase(); }).join('|'); }
  function descriptionSimilarity(a, b) {
    var left = normalizeDescription(a), right = normalizeDescription(b);
    if (!left || !right) return 0;
    if (left === right) return 1;
    var stop = { pekerjaan: true, pembangunan: true, umum: true, dan: true, yang: true, untuk: true, dengan: true, pada: true };
    var leftTokens = left.split(' ').filter(function (token) { return token && !stop[token]; });
    var rightTokens = right.split(' ').filter(function (token) { return token && !stop[token]; });
    if (!leftTokens.length || !rightTokens.length) return 0;
    var intersection = leftTokens.filter(function (token) { return rightTokens.indexOf(token) !== -1; }).length;
    var jaccard = intersection / Math.max(1, new Set(leftTokens.concat(rightTokens)).size);
    return jaccard;
  }
  function numeric(value) {
    if (typeof value === 'number' && Number.isFinite(value)) return value;
    if (value === null || value === undefined || text(value) === '') return null;
    if (window.DataNormalizer && typeof window.DataNormalizer.parseNumericValue === 'function') {
      var parsed = window.DataNormalizer.parseNumericValue(value);
      return typeof parsed === 'number' && Number.isFinite(parsed) ? parsed : null;
    }
    if (typeof value === 'string' && /^-?\d+(?:\.\d+)?$/.test(value.trim())) return Number(value.trim());
    return null;
  }
  function sourceRef(item, canonical) {
    var source = item.source || {};
    var filename = canonical.project && canonical.project.sourceFilename || null;
    var fields = source.cells || {};
    var refs = {};
    ['item_code', 'description', 'unit', 'volume', 'unit_price', 'amount'].forEach(function (field) {
      var cell = fields[field];
      if (cell) refs[field] = { cellRef: cell.cellRef || null, rawValue: cell.val !== undefined ? cell.val : cell.rawValue, formula: cell.formula || null, valueStatus: cell.valueStatus || null };
    });
    return {
      id: itemKey(item, 0),
      workbook: filename,
      sheet: source.sheet || null,
      row: Number.isFinite(source.row) ? source.row : null,
      cells: refs,
      code: item.code || item.itemNumber || null,
      description: item.description || null,
      unit: item.unit || null,
      categoryPath: pathValues(item),
      hierarchyStatus: item.hierarchyStatus || 'none',
      originalValues: { volume: item.volume, unitPrice: item.unitPrice, amount: item.amount }
    };
  }
  function itemKey(item, index) { return String(item.id || ((item.source && item.source.sheet || 'sheet') + ':' + (item.source && item.source.row || index + 1))); }
  function categoryRelation(a, b) {
    var pa = pathSignature(a), pb = pathSignature(b);
    if (!pa || !pb) return { equal: !pa && !pb, known: false };
    return { equal: pa === pb, known: true };
  }
  function candidateScore(a, b) {
    var codeA = normalizeCode(a.code || a.analysisCode || a.itemNumber), codeB = normalizeCode(b.code || b.analysisCode || b.itemNumber);
    var desc = descriptionSimilarity(a.description, b.description);
    var unitA = normalizeUnit(a.unit), unitB = normalizeUnit(b.unit);
    var category = categoryRelation(a, b);
    var score = 0;
    var reasons = [];
    if (codeA && codeA === codeB) { score += 0.46; reasons.push('Kode pekerjaan sama.'); }
    if (desc === 1) { score += 0.36; reasons.push('Uraian sama setelah normalisasi teks.'); }
    else if (desc >= 0.60) { score += 0.26 * desc; reasons.push('Uraian memiliki kemiripan tekstual; perlu validasi.'); }
    if (unitA && unitA === unitB) { score += 0.12; reasons.push('Satuan sama.'); }
    else if (unitA && unitB) { score += 0.03; reasons.push('Satuan berbeda dan perlu ditinjau.'); }
    if (category.equal && category.known) { score += 0.10; reasons.push('Category path sama.'); }
    else if (category.known && !category.equal) { score -= 0.18; reasons.push('Category path berbeda.'); }
    if (a.hierarchyStatus === 'ambiguous' || b.hierarchyStatus === 'ambiguous' || a.hierarchyStatus === 'partial' || b.hierarchyStatus === 'partial') {
      score -= 0.04;
      reasons.push('Salah satu hierarki tidak lengkap atau ambigu.');
    }
    return { score: Math.max(0, Math.min(1, score)), reasons: reasons, codeSame: !!codeA && codeA === codeB, descriptionSimilarity: desc, unitSame: !!unitA && unitA === unitB, categoryEqual: category.equal, categoryKnown: category.known };
  }
  function pct(delta, oldValue) { return oldValue === 0 || oldValue === null ? null : delta / oldValue * 100; }
  function compareValue(oldValue, newValue) {
    var before = numeric(oldValue), after = numeric(newValue);
    if (before === null || after === null) return { before: before, after: after, delta: null, percentageChange: null, status: 'unavailable' };
    var delta = after - before;
    return { before: before, after: after, delta: delta, percentageChange: pct(delta, before), status: delta === 0 ? 'unchanged' : 'changed' };
  }
  function comparePair(oldItem, newItem, oldCanonical, newCanonical, matchStatus, evidence) {
    var changes = {
      code: { before: oldItem.code || oldItem.analysisCode || oldItem.itemNumber || null, after: newItem.code || newItem.analysisCode || newItem.itemNumber || null },
      description: { before: oldItem.description || null, after: newItem.description || null },
      unit: { before: oldItem.unit || null, after: newItem.unit || null },
      volume: compareValue(oldItem.volume, newItem.volume),
      unitPrice: compareValue(oldItem.unitPrice, newItem.unitPrice),
      amount: compareValue(oldItem.amount, newItem.amount),
      categoryPath: { before: pathValues(oldItem), after: pathValues(newItem), changed: pathSignature(oldItem) !== pathSignature(newItem) }
    };
    var flags = [];
    if (normalizeCode(changes.code.before) !== normalizeCode(changes.code.after)) flags.push('CODE_CHANGED');
    if (normalizeDescription(changes.description.before) !== normalizeDescription(changes.description.after)) flags.push('DESCRIPTION_CHANGED');
    if (normalizeUnit(changes.unit.before) !== normalizeUnit(changes.unit.after)) flags.push('UNIT_CHANGED');
    ['volume', 'unitPrice', 'amount'].forEach(function (field) { if (changes[field].status === 'changed') flags.push(field.toUpperCase() + '_CHANGED'); });
    if (changes.categoryPath.changed) flags.push('HIERARCHY_CHANGED');
    if (!flags.length) flags.push('UNCHANGED');
    return {
      id: 'comparison:' + itemKey(oldItem, 0) + '::' + itemKey(newItem, 0),
      status: matchStatus,
      changeTypes: flags,
      changes: changes,
      matchEvidence: evidence,
      confidenceScore: evidence.score,
      sourceA: sourceRef(oldItem, oldCanonical),
      sourceB: sourceRef(newItem, newCanonical)
    };
  }
  function aggregate(canonical) {
    var items = Array.isArray(canonical.workItems) ? canonical.workItems : [];
    var imported = 0, importedCount = 0, calculated = 0, calculatedCount = 0, incomplete = 0;
    items.forEach(function (item) {
      var amount = numeric(item.amount), volume = numeric(item.volume), price = numeric(item.unitPrice);
      if (amount !== null) { imported += amount; importedCount++; } else incomplete++;
      if (volume !== null && price !== null) { calculated += volume * price; calculatedCount++; }
    });
    return { itemCount: items.length, importedAmount: importedCount ? imported : null, importedAmountCoverage: importedCount, calculatedAmount: calculatedCount ? calculated : null, calculableItems: calculatedCount, incompleteItems: items.length - calculatedCount, missingImportedAmountItems: items.length - importedCount };
  }
  function ambiguousResult(a, b, score, reasons, canonicalA, canonicalB) {
    return { id: 'candidate:' + itemKey(a, 0) + '::' + itemKey(b, 0), status: score >= 0.52 ? 'POSSIBLE_MATCH' : 'AMBIGUOUS', score: score, reasons: reasons, sourceA: sourceRef(a, canonicalA), sourceB: sourceRef(b, canonicalB) };
  }

  var ComparisonEngine = {
    engineName: 'RABComparisonEngine',
    engineVersion: VERSION,

    compare: function (canonicalA, canonicalB) {
      if (!canonicalA || !Array.isArray(canonicalA.workItems) || !canonicalB || !Array.isArray(canonicalB.workItems)) throw new Error('Dua model canonical dengan workItems diperlukan.');
      var aItems = canonicalA.workItems.filter(function (item) { return item && item.rowType !== 'subtotal' && item.type !== 'subtotal'; });
      var bItems = canonicalB.workItems.filter(function (item) { return item && item.rowType !== 'subtotal' && item.type !== 'subtotal'; });
      var usedA = Object.create(null), usedB = Object.create(null), blockedA = Object.create(null), blockedB = Object.create(null), matches = [], candidates = [], relations = [], earlyGroups = Object.create(null);
      var allScores = [];

      // Unequal repeated-code groups may represent a split/merge; do not consume them as pairs.
      var allCodesA = Object.create(null), allCodesB = Object.create(null);
      aItems.forEach(function (item, i) { var code = normalizeCode(item.code || item.analysisCode || item.itemNumber); if (code) (allCodesA[code] || (allCodesA[code] = [])).push(i); });
      bItems.forEach(function (item, i) { var code = normalizeCode(item.code || item.analysisCode || item.itemNumber); if (code) (allCodesB[code] || (allCodesB[code] = [])).push(i); });
      Object.keys(allCodesA).forEach(function (code) {
        var left = allCodesA[code], right = allCodesB[code] || [];
        if (left.length > 0 && right.length > 0 && (left.length !== right.length || left.length > 64 || right.length > 64)) {
          var status = left.length === right.length ? 'AMBIGUOUS' : (left.length === 1 && right.length > 1 ? 'SPLIT_CANDIDATE' : (left.length > 1 && right.length === 1 ? 'MERGE_CANDIDATE' : 'AMBIGUOUS'));
          relations.push({ status: status, code: code, itemIndicesA: left.slice(), itemIndicesB: right.slice(), reason: left.length === right.length
            ? 'Kode berulang terlalu banyak untuk dipasangkan aman secara otomatis; perlu konteks tambahan.'
            : 'Jumlah penggunaan kode berubah di antara versi; kemungkinan split/merge perlu diverifikasi.' });
          earlyGroups[code] = true;
          left.forEach(function (i) { blockedA[i] = true; }); right.forEach(function (i) { blockedB[i] = true; });
        }
      });

      // Only accept an automatic pair when each item has one clear best counterpart.
      var codeIndex = Object.create(null), descriptionIndex = Object.create(null), tokenIndex = Object.create(null);
      bItems.forEach(function (item, bi) {
        var code = normalizeCode(item.code || item.analysisCode || item.itemNumber), description = normalizeDescription(item.description);
        if (code) (codeIndex[code] || (codeIndex[code] = [])).push(bi);
        if (description) (descriptionIndex[description] || (descriptionIndex[description] = [])).push(bi);
        description.split(' ').forEach(function (token) { if (token.length > 2) (tokenIndex[token] || (tokenIndex[token] = [])).push(bi); });
      });
      var scoreByA = Object.create(null), scoreByB = Object.create(null);
      aItems.forEach(function (a, ai) {
        if (blockedA[ai]) return;
        var candidatesByIndex = Object.create(null);
        var code = normalizeCode(a.code || a.analysisCode || a.itemNumber), description = normalizeDescription(a.description);
        (codeIndex[code] || []).forEach(function (bi) { candidatesByIndex[bi] = true; });
        var exactDescriptionCandidates = descriptionIndex[description] || [];
        if (exactDescriptionCandidates.length <= 64) exactDescriptionCandidates.forEach(function (bi) { candidatesByIndex[bi] = true; });
        var hasFrequentTextToken = false;
        description.split(' ').forEach(function (token) {
          var posting = tokenIndex[token] || [];
          if (posting.length <= 64) posting.forEach(function (bi) { candidatesByIndex[bi] = true; });
          else hasFrequentTextToken = true;
        });
        if (!Object.keys(candidatesByIndex).length && hasFrequentTextToken) {
          blockedA[ai] = true;
          candidates.push({ id: 'candidate:' + itemKey(a, ai), status: 'AMBIGUOUS', score: null,
            reasons: ['Kata uraian terlalu umum pada korpus ini untuk menghasilkan kandidat satu-ke-satu secara aman.'],
            sourceA: sourceRef(a, canonicalA), sourceB: null });
        }
        Object.keys(candidatesByIndex).forEach(function (biKey) {
          var bi = Number(biKey);
          if (blockedB[bi]) return;
          var entry = { a: a, b: bItems[bi], ai: ai, bi: bi, data: candidateScore(a, bItems[bi]) };
          if (entry.data.score >= 0.35) {
            allScores.push(entry);
            (scoreByA[ai] || (scoreByA[ai] = [])).push(entry);
            (scoreByB[bi] || (scoreByB[bi] = [])).push(entry);
          }
        });
      });
      bItems.forEach(function (b, bi) {
        if (blockedB[bi] || usedB[bi]) return;
        var description = normalizeDescription(b.description);
        var frequent = description.split(' ').some(function (token) { return (tokenIndex[token] || []).length > 64; });
        var hasCandidate = (scoreByB[bi] || []).length > 0;
        if (frequent && !hasCandidate) {
          blockedB[bi] = true;
          candidates.push({ id: 'candidate:' + itemKey(b, bi), status: 'AMBIGUOUS', score: null,
            reasons: ['Kata uraian terlalu umum pada korpus ini untuk menghasilkan kandidat satu-ke-satu secara aman.'],
            sourceA: null, sourceB: sourceRef(b, canonicalB) });
        }
      });
      var bestA = Object.create(null), bestB = Object.create(null);
      aItems.forEach(function (a, ai) {
        var ranked = (scoreByA[ai] || []).sort(function (x, y) { return y.data.score - x.data.score; });
        if (ranked.length) bestA[ai] = { entry: ranked[0], margin: ranked.length > 1 ? ranked[0].data.score - ranked[1].data.score : 1 };
      });
      bItems.forEach(function (b, bi) {
        var ranked = (scoreByB[bi] || []).sort(function (x, y) { return y.data.score - x.data.score; });
        if (ranked.length) bestB[bi] = { entry: ranked[0], margin: ranked.length > 1 ? ranked[0].data.score - ranked[1].data.score : 1 };
      });

      Object.keys(bestA).forEach(function (key) {
        var ai = Number(key), choice = bestA[ai], bi = choice.entry.bi, reverse = bestB[bi];
        if (usedA[ai] || usedB[bi]) return;
        var score = choice.entry.data.score;
        var unique = reverse && reverse.entry.ai === ai && choice.margin >= 0.08 && reverse.margin >= 0.08;
        var exact = choice.entry.data.codeSame && choice.entry.data.descriptionSimilarity === 1 && choice.entry.data.unitSame &&
          (!choice.entry.data.categoryKnown || choice.entry.data.categoryEqual);
        if (unique && score >= 0.70) {
          var status = exact ? 'EXACT_MATCH' : 'HIGH_CONFIDENCE_MATCH';
          var match = comparePair(aItems[ai], bItems[bi], canonicalA, canonicalB, status, { score: score, reasons: choice.entry.data.reasons });
          matches.push(match); usedA[ai] = true; usedB[bi] = true;
        } else {
          candidates.push(ambiguousResult(aItems[ai], bItems[bi], score, choice.entry.data.reasons, canonicalA, canonicalB));
        }
      });

      // Repeated code groups with one-to-many or many-to-one shape are explicit candidates, never forced pairs.
      var codesA = Object.create(null), codesB = Object.create(null);
      aItems.forEach(function (item, i) { var code = normalizeCode(item.code || item.analysisCode || item.itemNumber); if (code && !usedA[i]) (codesA[code] || (codesA[code] = [])).push(i); });
      bItems.forEach(function (item, i) { var code = normalizeCode(item.code || item.analysisCode || item.itemNumber); if (code && !usedB[i]) (codesB[code] || (codesB[code] = [])).push(i); });
      Object.keys(codesA).forEach(function (code) {
        if (earlyGroups[code]) return;
        var left = codesA[code], right = codesB[code] || [];
        if (left.length === 1 && right.length > 1) relations.push({ status: 'SPLIT_CANDIDATE', code: code, itemIndicesA: left.slice(), itemIndicesB: right.slice(), reason: 'Satu item sumber berbagi kode dengan beberapa item perubahan; struktur perlu diverifikasi.' });
        else if (left.length > 1 && right.length === 1) relations.push({ status: 'MERGE_CANDIDATE', code: code, itemIndicesA: left.slice(), itemIndicesB: right.slice(), reason: 'Beberapa item sumber berbagi kode dengan satu item perubahan; struktur perlu diverifikasi.' });
        else if (left.length > 1 && right.length > 1) relations.push({ status: 'AMBIGUOUS', code: code, itemIndicesA: left.slice(), itemIndicesB: right.slice(), reason: 'Kode berulang pada kedua workbook dan tidak membentuk pasangan satu-ke-satu yang unik.' });
      });
      relations.forEach(function (relation) {
        relation.itemsA = relation.itemIndicesA.map(function (i) { return sourceRef(aItems[i], canonicalA); });
        relation.itemsB = relation.itemIndicesB.map(function (i) { return sourceRef(bItems[i], canonicalB); });
      });

      // High-overlap candidates remain ambiguous and are excluded from added/removed labels.
      var candidateA = Object.create(null), candidateB = Object.create(null);
      candidates.forEach(function (candidate) {
        if (candidate.sourceA) aItems.forEach(function (item, i) { if (itemKey(item, i) === candidate.sourceA.id) candidateA[i] = true; });
        if (candidate.sourceB) bItems.forEach(function (item, i) { if (itemKey(item, i) === candidate.sourceB.id) candidateB[i] = true; });
      });
      var relatedA = Object.create(null), relatedB = Object.create(null);
      relations.forEach(function (relation) { relation.itemIndicesA.forEach(function (i) { relatedA[i] = true; }); relation.itemIndicesB.forEach(function (i) { relatedB[i] = true; }); });
      candidates.forEach(function (candidate) {
        if (candidate.sourceA) aItems.forEach(function (item, i) { if (itemKey(item, i) === candidate.sourceA.id) { candidateA[i] = true; } });
        if (candidate.sourceB) bItems.forEach(function (item, i) { if (itemKey(item, i) === candidate.sourceB.id) { candidateB[i] = true; } });
      });
      var added = [], removed = [];
      bItems.forEach(function (item, i) { if (!usedB[i] && !candidateB[i] && !relatedB[i]) added.push({ status: 'ADDED', sourceB: sourceRef(item, canonicalB), changeTypes: ['ITEM_ADDED'] }); });
      aItems.forEach(function (item, i) { if (!usedA[i] && !candidateA[i] && !relatedA[i]) removed.push({ status: 'REMOVED', sourceA: sourceRef(item, canonicalA), changeTypes: ['ITEM_REMOVED'] }); });
      var totalA = aggregate(canonicalA), totalB = aggregate(canonicalB);
      var matchedDelta = 0, matchedDeltaCount = 0;
      matches.forEach(function (match) { if (match.changes.amount.delta !== null) { matchedDelta += match.changes.amount.delta; matchedDeltaCount++; } });
      var importedTotalsComplete = totalA.missingImportedAmountItems === 0 && totalB.missingImportedAmountItems === 0;
      var netImportedDelta = importedTotalsComplete && totalA.importedAmount !== null && totalB.importedAmount !== null ? totalB.importedAmount - totalA.importedAmount : null;
      var classifiedAmounts = matches.every(function (match) { return match.changes.amount.delta !== null; }) &&
        added.every(function (item) { return numeric(item.sourceB.originalValues.amount) !== null; }) &&
        removed.every(function (item) { return numeric(item.sourceA.originalValues.amount) !== null; });
      var classifiedDelta = classifiedAmounts ? matchedDelta + added.reduce(function (sum, item) { return sum + numeric(item.sourceB.originalValues.amount); }, 0) - removed.reduce(function (sum, item) { return sum + numeric(item.sourceA.originalValues.amount); }, 0) : null;
      var summary = {
        workbookA: { filename: canonicalA.project && canonicalA.project.sourceFilename || null, workItems: totalA.itemCount, importedAmount: totalA.importedAmount, importedAmountCoverage: totalA.importedAmountCoverage, recalculatedAmount: totalA.calculatedAmount, calculableItems: totalA.calculableItems, incompleteItems: totalA.incompleteItems, missingImportedAmountItems: totalA.missingImportedAmountItems },
        workbookB: { filename: canonicalB.project && canonicalB.project.sourceFilename || null, workItems: totalB.itemCount, importedAmount: totalB.importedAmount, importedAmountCoverage: totalB.importedAmountCoverage, recalculatedAmount: totalB.calculatedAmount, calculableItems: totalB.calculableItems, incompleteItems: totalB.incompleteItems, missingImportedAmountItems: totalB.missingImportedAmountItems },
        matchedItems: matches.length, addedItems: added.length, removedItems: removed.length, uncertainItems: candidates.length,
        possibleSplitMergeGroups: relations.filter(function (x) { return x.status === 'SPLIT_CANDIDATE' || x.status === 'MERGE_CANDIDATE'; }).length,
        ambiguousCodeGroups: relations.filter(function (x) { return x.status === 'AMBIGUOUS'; }).length,
        importedNetChange: netImportedDelta,
        matchedImportedAmountChange: matchedDeltaCount ? matchedDelta : null,
        classifiedImportedAmountChange: classifiedDelta,
        unreconciledImportedAmountChange: netImportedDelta === null || classifiedDelta === null ? null : netImportedDelta - classifiedDelta,
        importedTotalsComplete: importedTotalsComplete
      };
      var warnings = [];
      if (candidates.length || relations.length) warnings.push('Sebagian item memiliki kandidat pencocokan atau kemungkinan pecah/gabung dan tidak dipaksakan sebagai item baru/hilang.');
      if (!summary.importedTotalsComplete) warnings.push('Sebagian item tidak memiliki jumlah impor numerik; total impor merupakan cakupan parsial.');
      return {
        engineName: 'RABComparisonEngine', engineVersion: VERSION, schemaVersion: RESULT_SCHEMA,
        status: candidates.length || relations.length || !summary.importedTotalsComplete ? 'partial' : 'complete',
        summary: summary,
        findings: matches,
        metrics: { itemsA: aItems.length, itemsB: bItems.length, matches: matches.length, added: added.length, removed: removed.length, uncertain: candidates.length, comparisonCandidatesEvaluated: allScores.length },
        sourceReferences: matches.reduce(function (refs, match) { refs.push(match.sourceA, match.sourceB); return refs; }, []).concat(added.map(function (item) { return item.sourceB; }), removed.map(function (item) { return item.sourceA; }), candidates.reduce(function (refs, candidate) { if (candidate.sourceA) refs.push(candidate.sourceA); if (candidate.sourceB) refs.push(candidate.sourceB); return refs; }, []), relations.reduce(function (refs, relation) { return refs.concat(relation.itemsA || [], relation.itemsB || []); }, [])),
        warnings: warnings, errors: [],
        limitations: ['Subtotal dan grand total canonical tidak diagregasi sebagai work item.', 'Tidak ada konversi satuan atau kalkulasi ulang formula Excel; nilai kalkulasi memakai volume dan harga numerik yang tersedia.', 'Pencocokan fuzzy bersifat konservatif; kasus kandidat diserahkan untuk verifikasi manusia.'],
        changes: matches.concat(added, removed), candidates: candidates, relationships: relations
      };
    }
  };

  ComparisonEngine.normalizeDescription = normalizeDescription;
  ComparisonEngine.descriptionSimilarity = descriptionSimilarity;
  window.ComparisonEngine = ComparisonEngine;
})(typeof window !== 'undefined' ? window : this);
