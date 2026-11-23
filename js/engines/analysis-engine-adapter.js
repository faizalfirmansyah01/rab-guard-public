/* Shared read-only invocation boundary for Universal analysis engines. */
(function (window) {
  'use strict';

  var RESULT_SCHEMA = 'rab-guard-engine-result/1';

  function cloneValue(value, seen) {
    if (value === null || typeof value !== 'object') return value;
    seen = seen || new WeakMap();
    if (seen.has(value)) return seen.get(value);
    var copy = Array.isArray(value) ? [] : {};
    seen.set(value, copy);
    Object.keys(value).forEach(function (key) {
      copy[key] = cloneValue(value[key], seen);
    });
    return copy;
  }

  function freezeValue(value, seen) {
    if (value === null || typeof value !== 'object' || Object.isFrozen(value)) return value;
    seen = seen || new WeakSet();
    if (seen.has(value)) return value;
    seen.add(value);
    Object.keys(value).forEach(function (key) { freezeValue(value[key], seen); });
    return Object.freeze(value);
  }

  function emptyResult(engine, status, errors) {
    return {
      engineName: engine && (engine.engineName || engine.name) || 'UnknownAnalysisEngine',
      engineVersion: engine && engine.engineVersion || '0.0.0',
      schemaVersion: RESULT_SCHEMA,
      status: status,
      summary: {},
      findings: [],
      metrics: {},
      sourceReferences: [],
      warnings: [],
      errors: errors || [],
      limitations: []
    };
  }

  function standardizeResult(engine, rawResult) {
    if (!rawResult || typeof rawResult !== 'object') {
      return emptyResult(engine, 'failed', ['Engine tidak mengembalikan objek hasil.']);
    }
    var allowedStatus = { complete: true, partial: true, failed: true };
    var status = allowedStatus[rawResult.status] ? rawResult.status : 'failed';
    var result = {
      engineName: rawResult.engineName || engine.engineName || engine.name || 'UnknownAnalysisEngine',
      engineVersion: rawResult.engineVersion || engine.engineVersion || '0.0.0',
      schemaVersion: RESULT_SCHEMA,
      status: status,
      summary: rawResult.summary && typeof rawResult.summary === 'object' ? rawResult.summary : {},
      findings: Array.isArray(rawResult.findings) ? rawResult.findings : [],
      metrics: rawResult.metrics && typeof rawResult.metrics === 'object' ? rawResult.metrics : {},
      sourceReferences: Array.isArray(rawResult.sourceReferences) ? rawResult.sourceReferences : [],
      warnings: Array.isArray(rawResult.warnings) ? rawResult.warnings : [],
      errors: Array.isArray(rawResult.errors) ? rawResult.errors : [],
      limitations: Array.isArray(rawResult.limitations) ? rawResult.limitations : []
    };
    ['changes', 'candidates', 'relationships'].forEach(function (key) {
      if (Array.isArray(rawResult[key])) result[key] = rawResult[key];
    });
    // Domain-specific payloads remain top-level while sharing the standard
    // engine envelope; consumers can ignore fields they do not understand.
    ['materials', 'workItemDetails', 'mappings', 'coverage', 'assumptions',
      'priceComparisons', 'unbalancedPricingFindings', 'frontLoadingFindings',
      'matchingResults', 'riskIndicators', 'referenceSources'].forEach(function (key) {
      if (rawResult[key] !== undefined) result[key] = rawResult[key];
    });
    if (status === 'failed' && result.errors.length === 0) result.errors.push('Engine melaporkan kegagalan tanpa rincian.');
    return result;
  }

  var AnalysisEngineAdapter = {
    schemaVersion: RESULT_SCHEMA,

    run: function (engine, canonicalData, options) {
      if (!engine || typeof engine.analyze !== 'function') {
        return emptyResult(engine, 'failed', ['Engine tidak memiliki fungsi analyze(canonicalData).']);
      }
      if (!canonicalData || typeof canonicalData !== 'object' || !Array.isArray(canonicalData.workItems)) {
        return emptyResult(engine, 'failed', ['Canonical Data Contract v1 memerlukan workItems berupa array.']);
      }

      try {
        var snapshot = freezeValue(cloneValue(canonicalData));
        return standardizeResult(engine, engine.analyze(snapshot, cloneValue(options || {})));
      } catch (error) {
        return emptyResult(engine, 'failed', [error && error.message ? error.message : String(error)]);
      }
    },

    runComparison: function (engine, canonicalA, canonicalB, options) {
      if (!engine || typeof engine.compare !== 'function') {
        return emptyResult(engine, 'failed', ['Engine tidak memiliki fungsi compare(canonicalA, canonicalB).']);
      }
      if (!canonicalA || !Array.isArray(canonicalA.workItems) || !canonicalB || !Array.isArray(canonicalB.workItems)) {
        return emptyResult(engine, 'failed', ['Perbandingan memerlukan dua Canonical Data Contract v1 yang memiliki workItems array.']);
      }
      try {
        var snapshotA = freezeValue(cloneValue(canonicalA));
        var snapshotB = freezeValue(cloneValue(canonicalB));
        return standardizeResult(engine, engine.compare(snapshotA, snapshotB, freezeValue(cloneValue(options || {}))));
      } catch (error) {
        return emptyResult(engine, 'failed', [error && error.message ? error.message : String(error)]);
      }
    }
  };

  window.AnalysisEngineAdapter = AnalysisEngineAdapter;
})(typeof window !== 'undefined' ? window : this);
