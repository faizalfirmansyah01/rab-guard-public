/* Versioned, non-destructive project snapshot contract and validation helpers. */
(function (window) {
  'use strict';
  var VERSION = 1, ENGINE_SCHEMA = 'rab-guard-engine-result/1';
  function clone(value) { return value == null ? value : JSON.parse(JSON.stringify(value)); }
  function stable(value) {
    if (Array.isArray(value)) return '[' + value.map(stable).join(',') + ']';
    if (!value || typeof value !== 'object') return JSON.stringify(value);
    return '{' + Object.keys(value).sort().map(function (key) { return JSON.stringify(key) + ':' + stable(value[key]); }).join(',') + '}';
  }
  function fingerprint(value) {
    var subtle = window.crypto && window.crypto.subtle;
    if (!subtle || typeof TextEncoder === 'undefined') return Promise.resolve({ status: 'UNAVAILABLE', algorithm: null, value: null });
    return subtle.digest('SHA-256', new TextEncoder().encode(stable(value))).then(function (buffer) {
      return { status: 'COMPLETE', algorithm: 'SHA-256', value: Array.prototype.map.call(new Uint8Array(buffer), function (b) { return b.toString(16).padStart(2, '0'); }).join('') };
    }).catch(function () { return { status: 'UNAVAILABLE', algorithm: null, value: null }; });
  }
  function blobFingerprint(blob) {
    if (!blob || typeof blob.arrayBuffer !== 'function') return Promise.resolve({ status: 'UNAVAILABLE', algorithm: null, value: null });
    return blob.arrayBuffer().then(function (buffer) { return fingerprintBytes(buffer); }).catch(function () { return { status: 'UNAVAILABLE', algorithm: null, value: null }; });
  }
  function fingerprintBytes(buffer) {
    var subtle = window.crypto && window.crypto.subtle;
    if (!subtle) return Promise.resolve({ status: 'UNAVAILABLE', algorithm: null, value: null });
    return subtle.digest('SHA-256', buffer).then(function (digest) { return { status: 'COMPLETE', algorithm: 'SHA-256', value: Array.prototype.map.call(new Uint8Array(digest), function (b) { return b.toString(16).padStart(2, '0'); }).join('') }; })
      .catch(function () { return { status: 'UNAVAILABLE', algorithm: null, value: null }; });
  }
  function engineVersion(name) {
    var map = { UniversalChecker: window.UniversalChecker, UnitConsistencyEngine: window.UnitConsistencyEngine,
      MaterialLogisticsEngine: window.MaterialLogisticsEngine, RABComparisonEngine: window.ComparisonEngine,
      PricingRiskEngine: window.PricingRiskEngine, ReportingEngine: window.ReportingEngine };
    var engine = map[name]; return engine && engine.engineVersion || (name === 'UniversalChecker' ? 'legacy-universal-checker/1' : null);
  }
  function resultStatus(result, legacy) {
    if (!result) return 'NOT_RUN';
    if (result.status === 'failed') return 'FAILED';
    if (result.status === 'partial') return 'PARTIAL';
    if (result.status === 'complete' || legacy && Number.isFinite(result.totalFindings)) return 'COMPLETE';
    return 'FAILED';
  }
  function engineRecord(name, result, inputSourceIds, inputs, configuration) {
    var state = resultStatus(result, name === 'UniversalChecker');
    return fingerprint({ engineName: name, engineVersion: engineVersion(name), inputs: inputs, configuration: configuration || null }).then(function (hash) {
      if ((state === 'COMPLETE' || state === 'PARTIAL') && hash.status !== 'COMPLETE') state = 'UNAVAILABLE';
      return { engineName: name, engineVersion: engineVersion(name), resultSchemaVersion: result && result.schemaVersion || (name === 'UniversalChecker' && result ? 'universal-checker/legacy' : null),
        status: state, inputSourceIds: inputSourceIds || [], inputSnapshotIds: [], inputFingerprint: hash, analyzedAt: result && (result.analyzedAt || result.createdAt) || null,
        result: result ? clone(result) : null, warnings: clone(result && result.warnings || []), errors: clone(result && result.errors || []), limitations: clone(result && result.limitations || []) };
    });
  }
  function migrateDocument(doc) {
    if (doc && doc.projectSnapshot && doc.projectSnapshot.snapshotSchemaVersion === VERSION) return { snapshot: clone(doc.projectSnapshot), migrated: false };
    if (doc && doc.projectSnapshot && Number(doc.projectSnapshot.snapshotSchemaVersion) > VERSION) throw new Error('Snapshot memakai schema yang lebih baru dan tidak didukung. Data original dipertahankan.');
    var canonical = doc && doc.canonicalData || null;
    var sid = canonical ? 'canonical:' + (doc.documentId || 'legacy') + ':universal' : null;
    var migrationTime = doc && (doc.updatedAt || doc.importedAt) || 'unknown';
    var snapshot = { snapshotSchemaVersion: VERSION, snapshotId: 'migrated:' + (doc && doc.documentId || 'unknown') + ':v' + VERSION, projectId: doc && doc.projectId || null,
      projectName: null, createdAt: doc && doc.updatedAt || null, updatedAt: migrationTime, applicationVersion: 'unversioned',
      sourceRegistry: [], canonicalSnapshots: {}, engineResults: {}, engineStates: {}, manualDecisions: [], reportSources: [],
      migrationMetadata: { from: doc && doc.projectSnapshot && doc.projectSnapshot.snapshotSchemaVersion || 0, to: VERSION, migratedAt: migrationTime, originalSnapshotPreserved: true, missingFields: [] }, validationMetadata: { status: 'MIGRATED_PARTIAL', errors: [] } };
    if (canonical) {
      snapshot.canonicalSnapshots[sid] = clone(canonical);
      snapshot.sourceRegistry.push({ sourceId: 'source:' + (doc.documentId || 'legacy') + ':universal', documentId: doc.documentId || null, role: 'universalWorkbook', filename: doc.fileName || null,
        mimeType: doc.fileType || null, size: doc.fileSize || null, contentHash: null, addedAt: doc.importedAt || null, snapshotVersion: 1,
        ingestionStatus: doc.ingestionStatus || 'UNKNOWN', sourceMetadata: { availability: doc.documentId ? 'document-reference-unverified' : 'missing-document-reference' }, canonicalSnapshotId: sid });
    } else snapshot.migrationMetadata.missingFields.push('canonicalData');
    snapshot.engineResults.universalChecker = { engineName: 'UniversalChecker', engineVersion: engineVersion('UniversalChecker'), resultSchemaVersion: 'universal-checker/legacy', status: doc && doc.auditResult ? 'STALE' : 'NOT_RUN',
      inputSourceIds: canonical ? ['source:' + (doc.documentId || 'legacy') + ':universal'] : [], inputSnapshotIds: sid ? [sid] : [], inputFingerprint: null, analyzedAt: null,
      result: clone(doc && doc.auditResult || null), warnings: [], errors: [], limitations: ['Snapshot lama tidak menyimpan fingerprint sumber; hasil lama tidak dinyatakan current.'] };
    ['unitConsistency', 'materialLogistics', 'comparison', 'pricingRisk', 'reporting'].forEach(function (key) {
      snapshot.engineResults[key] = { engineName: key, engineVersion: engineVersion(key), resultSchemaVersion: null, status: 'UNAVAILABLE', inputSourceIds: [], inputSnapshotIds: [], inputFingerprint: null, analyzedAt: null, result: null, warnings: [], errors: [], limitations: ['Snapshot lama tidak menyimpan hasil atau identitas input engine ini.'] };
    });
    snapshot.migrationMetadata.missingFields.push('cross-engine-results', 'source-content-hashes');
    return { snapshot: snapshot, migrated: true };
  }
  function validate(snapshot) {
    var errors = [], engineErrors = Object.create(null);
    function coreError(message) { errors.push(message); }
    function engineError(key, message) { if (!engineErrors[key]) engineErrors[key] = []; engineErrors[key].push(message); errors.push(message); }
    if (!snapshot || snapshot.snapshotSchemaVersion !== VERSION) coreError('Versi project snapshot tidak didukung.');
    if (!snapshot || typeof snapshot.snapshotId !== 'string' || !snapshot.snapshotId) coreError('snapshotId tidak tersedia.');
    if (!snapshot || !snapshot.projectId) coreError('projectId tidak tersedia.');
    if (!snapshot || !snapshot.sourceRegistry || !Array.isArray(snapshot.sourceRegistry)) coreError('sourceRegistry tidak valid.');
    if (!snapshot || !snapshot.canonicalSnapshots || typeof snapshot.canonicalSnapshots !== 'object') coreError('canonicalSnapshots tidak valid.');
    if (!snapshot || !snapshot.engineResults || typeof snapshot.engineResults !== 'object') coreError('engineResults tidak valid.');
    if (snapshot && Array.isArray(snapshot.sourceRegistry)) {
      var sourceIds = Object.create(null);
      snapshot.sourceRegistry.forEach(function (source) {
      if (!source.sourceId || !source.role) coreError('Sumber memiliki sourceId/role kosong.');
      if (source.sourceId && sourceIds[source.sourceId]) coreError('sourceId duplikat: ' + source.sourceId);
      if (source.sourceId) sourceIds[source.sourceId] = true;
      if (source.canonicalSnapshotId && !snapshot.canonicalSnapshots[source.canonicalSnapshotId]) coreError('Canonical snapshot sumber tidak ditemukan: ' + source.sourceId);
      if (!source.documentId && !source.canonicalSnapshotId) coreError('Sumber tidak memiliki referensi penyimpanan: ' + source.sourceId);
      });
    }
    if (snapshot && snapshot.canonicalSnapshots) Object.keys(snapshot.canonicalSnapshots).forEach(function (id) {
      if (!snapshot.canonicalSnapshots[id] || !Array.isArray(snapshot.canonicalSnapshots[id].workItems)) coreError('Canonical snapshot tidak valid: ' + id);
    });
    if (snapshot && snapshot.engineResults) Object.keys(snapshot.engineResults).forEach(function (key) {
      var record = snapshot.engineResults[key];
      if (!record || !['NOT_RUN','COMPLETE','PARTIAL','STALE','FAILED','INSUFFICIENT_DATA','UNAVAILABLE'].includes(record.status)) engineError(key, 'Status engine tidak valid: ' + key);
      if (record && record.status === 'COMPLETE' && !record.result) engineError(key, 'Engine COMPLETE tanpa result: ' + key);
      if (record && ['COMPLETE','PARTIAL'].includes(record.status) && (!record.engineVersion || !record.inputFingerprint || record.inputFingerprint.status !== 'COMPLETE' || !record.inputFingerprint.value)) engineError(key, 'Engine ' + record.status + ' tanpa versi/fingerprint SHA-256: ' + key);
      if (record && record.inputSnapshotIds) Object.keys(record.inputSnapshotIds).forEach(function (role) {
        var id = record.inputSnapshotIds[role];
        if (!snapshot.canonicalSnapshots[id] && !(snapshot.supplementalSnapshots && snapshot.supplementalSnapshots[id])) engineError(key, 'Input snapshot engine tidak ditemukan: ' + key + '/' + role);
      });
    });
    if (snapshot && Array.isArray(snapshot.manualDecisions)) snapshot.manualDecisions.forEach(function (decision) {
      if (!decision.decisionId || !decision.engineName || !decision.targetItemIdentity || !decision.selectedCandidateIdentity) coreError('Manual decision tidak lengkap.');
      if (!decision.provenance || !decision.provenance.sourceId || !sourceIdsFor(snapshot).includes(decision.provenance.sourceId)) coreError('Manual decision tidak terikat ke sumber proyek.');
    });
    return { valid: errors.length === 0, coreValid: errors.length === Object.values(engineErrors).reduce(function (total, list) { return total + list.length; }, 0), errors: errors, engineErrors: engineErrors };
  }
  function sourceIdsFor(snapshot) { return Array.isArray(snapshot.sourceRegistry) ? snapshot.sourceRegistry.map(function (source) { return source.sourceId; }) : []; }
  window.ProjectSnapshot = { schemaVersion: VERSION, engineResultSchemaVersion: ENGINE_SCHEMA, clone: clone, stableStringify: stable,
    fingerprint: fingerprint, blobFingerprint: blobFingerprint, engineVersion: engineVersion, resultStatus: resultStatus, engineRecord: engineRecord, migrateDocument: migrateDocument, validate: validate };
})(typeof window !== 'undefined' ? window : this);
