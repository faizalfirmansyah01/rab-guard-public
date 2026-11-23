/*
 * RAB GUARD – Project Manager
 * Coordinates active project and document handling in the UI.
 * Relies on ProjectStorage (IndexedDB) for persistence.
 */

(function (window) {
  'use strict';

  let activeProjectId = null; // currently opened project
  let activeDocumentId = null; // currently opened document within the project

  /** Generate a simple UUID (fallback if crypto not available) */
  function genId() {
    if (window.crypto && crypto.randomUUID) return crypto.randomUUID();
    // simple fallback based on timestamp and random number
    return 'id-' + Date.now() + '-' + Math.floor(Math.random() * 100000);
  }

  function engineStates() { return window.ProjectSnapshot; }
  function addCanonical(snapshot, roleMap, role, canonical) {
    if (!canonical || !Array.isArray(canonical.workItems)) return null;
    var id = 'cs:' + role + ':' + (canonical.project && canonical.project.sourceFilename || 'unnamed') + ':' + Object.keys(snapshot.canonicalSnapshots).length;
    snapshot.canonicalSnapshots[id] = window.ProjectSnapshot.clone(canonical);
    roleMap[role] = id;
    var sourceId = 'src:' + role + ':' + id;
    snapshot.sourceRegistry.push({ sourceId: sourceId, documentId: null, role: role, filename: canonical.project && canonical.project.sourceFilename || null,
      mimeType: null, size: null, contentHash: null, addedAt: null, snapshotVersion: 1, ingestionStatus: 'CANONICAL_SNAPSHOT',
      sourceMetadata: { availability: 'canonical-only', originalBlobPersisted: false }, canonicalSnapshotId: id });
    return { sourceId: sourceId, snapshotId: id };
  }
  async function persistSourceFiles(doc, comparison, pricing) {
    var helper = engineStates(), roles = [
      ['comparisonWorkbookA', comparison.fileA], ['comparisonWorkbookB', comparison.fileB],
      ['pricingProposal', pricing.files && pricing.files.bid], ['pricingReference', pricing.files && pricing.files.reference],
      ['pricingSchedule', pricing.files && pricing.files.schedule]
    ].filter(function (entry) { return entry[1] && typeof entry[1].slice === 'function'; });
    var known = await window.ProjectStorage.getDocumentsByProject(doc.projectId), refs = {}, staged = [], reusedHash = Object.create(null), primaryBlob = await window.ProjectStorage.getWorkbookFile(doc.documentId);
    var primaryHash = primaryBlob ? await helper.blobFingerprint(primaryBlob) : null;
    for (var i = 0; i < roles.length; i++) {
      var role = roles[i][0], file = roles[i][1], hash = await helper.blobFingerprint(file), key = hash.status === 'COMPLETE' ? hash.value : null;
      if (key && primaryHash && primaryHash.status === 'COMPLETE' && key === primaryHash.value) {
        refs[role] = { documentId: doc.documentId, contentHash: hash, fileName: doc.fileName, fileType: doc.fileType, size: doc.fileSize, availability: 'blob-stored' }; continue;
      }
      if (key && reusedHash[key]) { refs[role] = reusedHash[key]; continue; }
      var existing = key && known.find(function (candidate) { return candidate.documentId !== doc.documentId && candidate.snapshotSourceHash && candidate.snapshotSourceHash.value === key; });
      if (existing && await window.ProjectStorage.getWorkbookFile(existing.documentId)) {
        refs[role] = { documentId: existing.documentId, contentHash: hash, fileName: existing.fileName, fileType: existing.fileType, size: existing.fileSize, availability: 'blob-stored' };
        if (key) reusedHash[key] = refs[role]; continue;
      }
      var sourceDocumentId = genId(), sourceDoc = { documentId: sourceDocumentId, projectId: doc.projectId, fileName: file.name || role + '.xlsx',
        fileType: file.type || 'application/octet-stream', fileSize: file.size, documentRole: role, sourceRole: role, snapshotSourceHash: hash,
        documentVersion: 1, ingestionStatus: 'SOURCE_STORED', mappingStatus: 'NOT_APPLICABLE', auditStatus: 'NOT_APPLICABLE', canonicalData: null };
      staged.push({ document: sourceDoc, file: file, role: role, hash: hash });
      var ref = { documentId: sourceDocumentId, contentHash: hash, fileName: sourceDoc.fileName, fileType: sourceDoc.fileType, size: sourceDoc.fileSize, availability: 'blob-stored' };
      refs[role] = ref; if (key) reusedHash[key] = ref;
    }
    if (staged.length) await window.ProjectStorage.addSourceDocumentsWithFiles(staged);
    return { refs: refs, createdDocumentIds: staged.map(function (entry) { return entry.document.documentId; }) };
  }
  async function createVersionedSnapshot(doc, universalState, sourceRefs) {
    var helper = engineStates(), comparison = window.ComparisonUI && window.ComparisonUI.getState ? window.ComparisonUI.getState() : {};
    var pricing = window.PricingRiskUI && window.PricingRiskUI.getState ? window.PricingRiskUI.getState() : {};
    var material = window.MaterialLogisticsUI && window.MaterialLogisticsUI.getState ? window.MaterialLogisticsUI.getState() : { result: universalState.currentMaterialLogisticsResult };
    var report = window.ReportingUI && window.ReportingUI.getState ? window.ReportingUI.getState().document : null;
    var snapshot = { snapshotSchemaVersion: helper.schemaVersion, snapshotId: genId(), projectId: doc.projectId, projectName: null,
      createdAt: new Date().toISOString(), updatedAt: new Date().toISOString(), applicationVersion: 'RABGuard-local', sourceRegistry: [],
      canonicalSnapshots: {}, supplementalSnapshots: {}, engineResults: {}, engineStates: {}, manualDecisions: [], reportSources: [],
      migrationMetadata: { from: helper.schemaVersion, to: helper.schemaVersion, migratedAt: null, originalSnapshotPreserved: true, missingFields: [] },
      validationMetadata: { status: 'PREPARED', errors: [] }, uiState: { comparison: {}, pricingRisk: {} } };
    var roles = {}, primary = addCanonical(snapshot, roles, 'universalWorkbook', universalState.currentCanonicalData);
    if (primary) {
      var project = (await window.ProjectStorage.getAllProjects()).find(function (item) { return item.projectId === doc.projectId; });
      snapshot.projectName = project && project.projectName || null;
      var blob = await window.ProjectStorage.getWorkbookFile(doc.documentId);
      var contentHash = blob ? await helper.blobFingerprint(blob) : { status: 'UNAVAILABLE', algorithm: null, value: null };
      snapshot.sourceRegistry.push({ sourceId: 'src:' + doc.documentId, documentId: doc.documentId, role: 'universalWorkbook', filename: doc.fileName,
        mimeType: doc.fileType || null, size: doc.fileSize || null, contentHash: contentHash, addedAt: doc.importedAt || null, snapshotVersion: 1,
        ingestionStatus: doc.ingestionStatus || 'DONE', sourceMetadata: { availability: blob ? 'blob-stored' : 'blob-missing' }, canonicalSnapshotId: primary.snapshotId });
      // Remove the canonical-only duplicate source record for the primary role.
      snapshot.sourceRegistry.splice(snapshot.sourceRegistry.indexOf(snapshot.sourceRegistry.filter(function (s) { return s.sourceId === primary.sourceId; })[0]), 1);
    }
    var pairA = addCanonical(snapshot, roles, 'comparisonWorkbookA', comparison.readyA === false ? null : comparison.canonicalA), pairB = addCanonical(snapshot, roles, 'comparisonWorkbookB', comparison.readyB === false ? null : comparison.canonicalB);
    var bid = addCanonical(snapshot, roles, 'pricingProposal', pricing.bid), reference = addCanonical(snapshot, roles, 'pricingReference', pricing.reference);
    if (pricing.schedule) { var scheduleId = 'ss:pricingSchedule:' + Object.keys(snapshot.supplementalSnapshots).length; snapshot.supplementalSnapshots[scheduleId] = helper.clone(pricing.schedule); roles.pricingSchedule = scheduleId; }
    if (report) { var reportId = 'rs:' + snapshot.snapshotId; snapshot.supplementalSnapshots[reportId] = helper.clone(report); roles.reportingReport = reportId; snapshot.reportSources.push({ reportSnapshotId: reportId, reportType: report.reportType || null }); }
    Object.keys(sourceRefs || {}).forEach(function (role) {
      var ref = sourceRefs[role], entry = snapshot.sourceRegistry.find(function (source) { return source.role === role && source.canonicalSnapshotId === roles[role]; });
      if (!entry) { entry = { sourceId: 'src:' + role + ':' + ref.documentId, canonicalSnapshotId: roles[role] || null }; snapshot.sourceRegistry.push(entry); }
      entry.sourceId = 'src:' + role + ':' + ref.documentId; entry.documentId = ref.documentId; entry.role = role; entry.filename = ref.fileName;
      entry.mimeType = ref.fileType; entry.size = ref.size; entry.contentHash = ref.contentHash; entry.addedAt = doc.importedAt || null; entry.snapshotVersion = 1;
      entry.ingestionStatus = roles[role] ? 'NORMALIZED' : 'SOURCE_STORED'; entry.sourceMetadata = { availability: ref.availability, originalBlobPersisted: true };
    });
    var materialSourceId = primary && snapshot.sourceRegistry.find(function (s) { return s.documentId === doc.documentId; }).sourceId;
    var specs = [
      ['universalChecker', 'UniversalChecker', universalState.currentAuditResult, ['universalWorkbook'], {}, ['universalWorkbook']],
      ['unitConsistency', 'UnitConsistencyEngine', universalState.currentUnitConsistencyResult, ['universalWorkbook'], {}, ['universalWorkbook']],
      ['materialLogistics', 'MaterialLogisticsEngine', material.result || universalState.currentMaterialLogisticsResult, ['universalWorkbook'], { mappingDecisions: material.mappingDecisions || {} }, ['universalWorkbook']],
      ['comparison', 'RABComparisonEngine', comparison.result, ['comparisonWorkbookA','comparisonWorkbookB'], {}, ['comparisonWorkbookA','comparisonWorkbookB']],
      ['pricingRisk', 'PricingRiskEngine', pricing.result, ['pricingProposal','pricingReference'].concat(pricing.schedule ? ['pricingSchedule'] : []), pricing.configuration || {}, ['pricingProposal','pricingReference'].concat(pricing.schedule ? ['pricingSchedule'] : [])],
      ['reporting', 'ReportingEngine', report, ['universalWorkbook'], { reportType: report && report.reportType || null }, ['reportingReport']]
    ];
    for (var i = 0; i < specs.length; i++) {
      var spec = specs[i], inputs = {}, ids = {}, sourceIds = [];
      spec[5].forEach(function (role) { var id = roles[role]; if (!id) return; ids[role] = id; inputs[role] = role === 'pricingSchedule' || role === 'reportingReport' ? snapshot.supplementalSnapshots[id] : snapshot.canonicalSnapshots[id]; });
      spec[3].forEach(function (role) { if (roles[role]) { var source = snapshot.sourceRegistry.find(function (entry) { return entry.role === role; }); if (source) sourceIds.push(source.sourceId); } });
      var record = await helper.engineRecord(spec[1], spec[2], sourceIds, inputs, spec[4]);
      record.inputSnapshotIds = ids; record.configuration = helper.clone(spec[4]);
      if (spec[0] === 'reporting' && spec[2]) { record.status = 'COMPLETE'; record.resultSchemaVersion = spec[2].schemaVersion || 'rab-guard-report/1'; }
      if (spec[0] === 'materialLogistics' && universalState.currentCanonicalData && !(universalState.currentCanonicalData.ahspAnalyses || []).length && spec[2]) record.status = 'INSUFFICIENT_DATA';
      snapshot.engineResults[spec[0]] = record; snapshot.engineStates[spec[0]] = record.status;
    }
    snapshot.uiState.comparison = { canonicalAId: roles.comparisonWorkbookA || null, canonicalBId: roles.comparisonWorkbookB || null };
    snapshot.uiState.pricingRisk = { bidId: roles.pricingProposal || null, referenceId: roles.pricingReference || null, scheduleId: roles.pricingSchedule || null,
      configuration: helper.clone(pricing.configuration || {}) };
    var decisions = material.mappingDecisions || {}, mappingHash = primary ? await helper.fingerprint(snapshot.canonicalSnapshots[primary.snapshotId]) : { status: 'UNAVAILABLE', value: null };
    Object.keys(decisions).forEach(function (itemId) { snapshot.manualDecisions.push({ decisionId: 'ahsp:' + itemId, engineName: 'MaterialLogisticsEngine', sourceFingerprint: mappingHash.value,
      targetItemIdentity: itemId, selectedCandidateIdentity: decisions[itemId], createdAt: snapshot.updatedAt, decisionVersion: 1,
      provenance: { sourceId: materialSourceId || null }, status: mappingHash.status === 'COMPLETE' ? 'CURRENT' : 'UNVERIFIED' }); });
    snapshot.sourceRegistry.forEach(function (entry) {
      if (!entry.contentHash) entry.contentHash = { status: 'UNAVAILABLE', algorithm: null, value: null };
    });
    var validation = helper.validate(snapshot);
    if (!validation.valid) throw new Error('Project snapshot tidak valid: ' + validation.errors.join(' '));
    snapshot.validationMetadata = { status: 'VALID', errors: [], validatedAt: new Date().toISOString() };
    return snapshot;
  }

  /** Create a new project and make it active */
  async function createProject(name) {
    const proj = {
      projectId: genId(),
      projectName: name || 'Untitled Project',
      description: '',
      projectStatus: 'ACTIVE',
      createdAt: new Date().toISOString(),
      updatedAt: new Date().toISOString()
    };
    await window.ProjectStorage.addProject(proj);
    activeProjectId = proj.projectId;
    // refresh UI list
    renderProjectList();
    return proj;
  }

  /** Load an existing project (by id) */
  async function loadProject(projectId) {
    const all = await window.ProjectStorage.getAllProjects();
    const proj = all.find(p => p.projectId === projectId);
    if (!proj) throw new Error('Project not found');

    activeProjectId = proj.projectId;
    activeDocumentId = null;

    // Bersihkan state dokumen sebelumnya ketika berganti project.
    if (window.UniversalUI) {
      window.UniversalUI.resetWorkbookState();
    }

    renderProjectDetail();
    renderDocumentRegistry();
  }
  /** Rename current project */
  async function renameProject(newName) {
    if (!activeProjectId) return;
    const all = await window.ProjectStorage.getAllProjects();
    const proj = all.find(p => p.projectId === activeProjectId);
    if (!proj) return;
    proj.projectName = newName;
    await window.ProjectStorage.updateProject(proj);
    renderProjectList();
    renderProjectDetail();
  }

  /** Delete current project */
  async function deleteCurrentProject() {
    if (!activeProjectId) return;
    await window.ProjectStorage.deleteProject(activeProjectId);
    activeProjectId = null;
    activeDocumentId = null;
    // reset UI to empty state
    renderProjectList();
    renderProjectDetail();
    renderDocumentRegistry();
    // also clear universal UI state
    if (window.UniversalUI) window.UniversalUI.resetWorkbookState();
  }

  /** Add a new document to the active project */
  async function addDocument(file, category) {
    if (!activeProjectId) {
      // auto‑create a temporary project if none exists
      await createProject('Untitled Project');
    }
    const doc = {
      documentId: genId(),
      projectId: activeProjectId,
      fileName: file.name,
      fileType: file.type || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      documentRole: category || 'Other Spreadsheet',
      fileSize: file.size,
      importedAt: new Date().toISOString(),
      worksheetList: [],
      ingestionStatus: 'PENDING',
      mappingStatus: 'PENDING',
      auditStatus: 'PENDING',
      sourceData: null,
      canonicalData: null,
      wbsData: null,
      documentVersion: 1
    };
    await window.ProjectStorage.addDocumentWithFile(doc, file);
    activeDocumentId = doc.documentId;
    // proceed with normal ingestion flow using the file
    if (window.UniversalUI) {
      window.UniversalUI.handleFileUpload(file);
    }
    renderDocumentRegistry();
  }

  /** Set the active document (by id) and load its stored data into UniversalUI */
  /** Restore a stored document into UniversalUI's private state */
  async function setActiveDocument(documentId) {
    const doc = await window.ProjectStorage.getDocument(documentId);
    if (!doc) throw new Error('Document not found');
    var helper = engineStates(), migrated = helper.migrateDocument(doc), snapshot = migrated.snapshot;
    var validation = helper.validate(snapshot);
    if (!validation.coreValid) throw new Error('Snapshot proyek tidak dapat dipulihkan: ' + validation.errors.join(' '));
    if (snapshot.projectId !== doc.projectId) throw new Error('Identitas project pada snapshot tidak cocok dengan dokumen.');
    Object.keys(validation.engineErrors || {}).forEach(function (key) {
      var broken = snapshot.engineResults[key];
      if (!broken) broken = snapshot.engineResults[key] = { engineName: key, inputSourceIds: [], inputSnapshotIds: {}, warnings: [], errors: [], limitations: [] };
      broken.status = 'FAILED'; broken.result = null;
      broken.errors = (broken.errors || []).concat(validation.engineErrors[key]);
      broken.limitations = (broken.limitations || []).concat(['Hasil engine rusak/tidak valid dan tidak dipulihkan; engine lain tetap dapat dipulihkan.']);
      snapshot.engineStates[key] = 'FAILED';
    });
    var sourceIntegrityById = {}, filesByRole = {}, sourceOutcomes = [];
    for (var sourceIndex = 0; sourceIndex < snapshot.sourceRegistry.length; sourceIndex++) {
      var sourceEntry = snapshot.sourceRegistry[sourceIndex], sourceDocument = sourceEntry.documentId ? await window.ProjectStorage.getDocument(sourceEntry.documentId) : null;
      var sourceBlob = sourceDocument ? await window.ProjectStorage.getWorkbookFile(sourceEntry.documentId) : null, integrity = 'CANONICAL_ONLY';
      if (sourceEntry.documentId && !sourceDocument) integrity = 'MISSING';
      else if (sourceEntry.documentId && !sourceBlob) integrity = 'MISSING';
      else if (sourceBlob && sourceEntry.contentHash && sourceEntry.contentHash.status === 'COMPLETE') {
        var verifiedHash = await helper.blobFingerprint(sourceBlob); integrity = verifiedHash.value === sourceEntry.contentHash.value ? 'VERIFIED' : 'MISMATCH';
      } else if (sourceBlob) integrity = 'UNVERIFIED';
      sourceIntegrityById[sourceEntry.sourceId] = integrity; sourceOutcomes.push(integrity);
      if (sourceBlob && typeof File !== 'undefined') filesByRole[sourceEntry.role] = new File([sourceBlob], sourceEntry.filename || sourceDocument.fileName, { type: sourceEntry.mimeType || sourceDocument.fileType || 'application/octet-stream' });
    }
    var sourceIntegrity = sourceOutcomes.length && sourceOutcomes.every(function (result) { return result === 'VERIFIED'; }) ? 'VERIFIED' : sourceOutcomes.indexOf('MISMATCH') >= 0 ? 'MISMATCH' : sourceOutcomes.indexOf('MISSING') >= 0 ? 'MISSING' : 'PARTIAL';
    var primarySource = snapshot.sourceRegistry.find(function (source) { return source.role === 'universalWorkbook'; });

    var canonicalEntry = primarySource && primarySource.canonicalSnapshotId && snapshot.canonicalSnapshots[primarySource.canonicalSnapshotId];
    var canonical = canonicalEntry || doc.canonicalData || null;
    if (!canonical && doc.sourceData) {
      var inspection = window.WorkbookInspector.inspect(doc.sourceData);
      canonical = window.DataNormalizer.normalizeWorkbook(doc.sourceData, inspection);
    }
    if (canonical && !Array.isArray(canonical.workItems)) throw new Error('Canonical snapshot rusak: workItems bukan array. State aktif lama dipertahankan.');

    var engineStatus = {};
    Object.keys(validation.engineErrors || {}).forEach(function (key) { engineStatus[key] = 'FAILED'; });
    // Resolve fingerprints before committing any state.
    await Promise.all(Object.keys(snapshot.engineResults || {}).map(async function (key) {
      var record = snapshot.engineResults[key];
      if (!record || !record.result || !['COMPLETE','PARTIAL'].includes(record.status)) return;
      var ids = record.inputSnapshotIds || {}, inputs = {}, missing = false;
      Object.keys(ids).forEach(function (role) { var id = ids[role], value = snapshot.canonicalSnapshots[id] || snapshot.supplementalSnapshots && snapshot.supplementalSnapshots[id]; if (value === undefined) missing = true; else inputs[role] = value; });
      if (missing || !record.inputFingerprint || record.inputFingerprint.status !== 'COMPLETE' || helper.engineVersion(record.engineName) !== record.engineVersion) { engineStatus[key] = 'STALE'; return; }
      var actual = await helper.fingerprint({ engineName: record.engineName, engineVersion: record.engineVersion, inputs: inputs, configuration: record.configuration || null });
      engineStatus[key] = actual.value === record.inputFingerprint.value ? record.status : 'STALE';
      (record.inputSourceIds || []).forEach(function (sourceId) { if (sourceIntegrityById[sourceId] === 'MISSING') engineStatus[key] = 'UNAVAILABLE'; else if (sourceIntegrityById[sourceId] === 'MISMATCH') engineStatus[key] = 'STALE'; });
    }));
    if (Object.keys(engineStatus).some(function (key) { return key !== 'reporting' && ['STALE','UNAVAILABLE','FAILED'].includes(engineStatus[key]); }) && engineStatus.reporting === 'COMPLETE') engineStatus.reporting = 'STALE';
    Object.keys(engineStatus).forEach(function (key) { if (snapshot.engineResults[key]) { snapshot.engineResults[key].status = engineStatus[key]; snapshot.engineStates[key] = engineStatus[key]; } });

    var oldUniversal = window.UniversalUI && window.UniversalUI.getCurrentState ? window.UniversalUI.getCurrentState() : null;
    var oldComparison = window.ComparisonUI && window.ComparisonUI.getState ? window.ComparisonUI.getState() : null;
    var oldPricing = window.PricingRiskUI && window.PricingRiskUI.getState ? window.PricingRiskUI.getState() : null;
    var oldReporting = window.ReportingUI && window.ReportingUI.getState ? window.ReportingUI.getState() : null;
    var selected = snapshot.engineResults || {}, eligible = function (key) { return engineStatus[key] === 'COMPLETE' || engineStatus[key] === 'PARTIAL'; };
    var canonicalFor = function (role) { var id = snapshot.uiState && (snapshot.uiState.comparison || {})[role]; return id && snapshot.canonicalSnapshots[id] || null; };
    var comparisonRecord = selected.comparison, pricingRecord = selected.pricingRisk, materialRecord = selected.materialLogistics;
    var canonicalFingerprint = canonical ? await helper.fingerprint(canonical) : { status: 'UNAVAILABLE', value: null };
    var restoredDecisions = (snapshot.manualDecisions || []).reduce(function (out, decision) {
      if (decision.sourceFingerprint && canonicalFingerprint.status === 'COMPLETE' && decision.sourceFingerprint === canonicalFingerprint.value) out[decision.targetItemIdentity] = decision.selectedCandidateIdentity;
      else decision.status = 'STALE';
      return out;
    }, {});
    var state = { currentParsedWb: doc.sourceData || null, currentInspection: null, currentSelectedSheet: doc.selectedSheet || null,
      currentMappingConfig: doc.mappingConfig || null, currentCanonicalData: canonical, currentAuditResult: eligible('universalChecker') ? selected.universalChecker.result : null,
      currentUnitConsistencyResult: eligible('unitConsistency') ? selected.unitConsistency.result : null,
      currentMaterialLogisticsResult: eligible('materialLogistics') ? materialRecord.result : null,
      materialLogisticsState: { result: eligible('materialLogistics') ? materialRecord.result : null, mappingDecisions: restoredDecisions },
      currentFilter: 'all', isSyntheticDemo: false, activeDemoTitle: '', restoreMetadata: { sourceIntegrity: sourceIntegrity, engineStates: engineStatus, snapshotVersion: snapshot.snapshotSchemaVersion, migrated: migrated.migrated } };
    try {
      if (window.UniversalUI && typeof window.UniversalUI.restoreCurrentState === 'function') {
        window.UniversalUI.restoreCurrentState(state, doc.fileName);
      }
      if (window.ComparisonUI && window.ComparisonUI.restoreState) window.ComparisonUI.restoreState({
        canonicalA: eligible('comparison') ? canonicalFor('canonicalAId') : null,
        canonicalB: eligible('comparison') ? canonicalFor('canonicalBId') : null,
        fileA: filesByRole.comparisonWorkbookA || null, fileB: filesByRole.comparisonWorkbookB || null,
        readyA: eligible('comparison'), readyB: eligible('comparison'),
        result: eligible('comparison') ? comparisonRecord.result : null
      });
      if (window.PricingRiskUI && window.PricingRiskUI.restoreState) {
        var pricingState = snapshot.uiState && snapshot.uiState.pricingRisk || {};
        window.PricingRiskUI.restoreState({ bid: eligible('pricingRisk') && pricingState.bidId ? snapshot.canonicalSnapshots[pricingState.bidId] : null,
          reference: eligible('pricingRisk') && pricingState.referenceId ? snapshot.canonicalSnapshots[pricingState.referenceId] : null,
          schedule: eligible('pricingRisk') && pricingState.scheduleId ? snapshot.supplementalSnapshots[pricingState.scheduleId] : null,
          files: { bid: filesByRole.pricingProposal || null, reference: filesByRole.pricingReference || null, schedule: filesByRole.pricingSchedule || null },
          result: eligible('pricingRisk') ? pricingRecord.result : null, configuration: pricingState.configuration });
      }
      if (window.ReportingUI) {
        if (eligible('reporting') && window.ReportingUI.restoreDocument) window.ReportingUI.restoreDocument(selected.reporting.result);
        else if (window.ReportingUI.refresh) window.ReportingUI.refresh();
      }
      if (migrated.migrated || JSON.stringify(doc.projectSnapshot) !== JSON.stringify(snapshot)) { doc.projectSnapshot = snapshot; await window.ProjectStorage.updateDocument(doc); }
      activeDocumentId = documentId;
      activeProjectId = doc.projectId;
      renderDocumentRegistry();
      return { restored: true, snapshot: snapshot, engineStates: engineStatus, sourceIntegrity: sourceIntegrity, migrated: migrated.migrated, partial: sourceIntegrity !== 'VERIFIED' || Object.keys(engineStatus).some(function (key) { return ['STALE','FAILED','UNAVAILABLE','NOT_RUN','INSUFFICIENT_DATA'].includes(engineStatus[key]); }) };
    } catch (error) {
      if (oldUniversal && window.UniversalUI && window.UniversalUI.restoreCurrentState) { try { window.UniversalUI.restoreCurrentState(oldUniversal); } catch (_) { /* retain original failure */ } }
      if (oldComparison && window.ComparisonUI && window.ComparisonUI.restoreState) window.ComparisonUI.restoreState(oldComparison);
      if (oldPricing && window.PricingRiskUI && window.PricingRiskUI.restoreState) window.PricingRiskUI.restoreState(oldPricing);
      if (oldReporting && window.ReportingUI && window.ReportingUI.restoreDocument) window.ReportingUI.restoreDocument(oldReporting.document);
      throw new Error('Pemulihan proyek dibatalkan; state aktif sebelumnya dipertahankan. ' + error.message);
    }
  }

  /** Save current private UniversalUI state into the active document */
  async function persistDocumentUpdates() {
    if (!activeDocumentId) return false;

    if (!window.UniversalUI ||
        typeof window.UniversalUI.getCurrentState !== 'function') {
      throw new Error('UniversalUI.getCurrentState tidak tersedia.');
    }

    const doc = await window.ProjectStorage.getDocument(activeDocumentId);
    if (!doc) return false;

    const state = window.UniversalUI.getCurrentState();

    // Simpan state aktual, bukan membaca window.current... yang tidak ada.
    doc.sourceData = state.currentParsedWb || null;
    doc.selectedSheet = state.currentSelectedSheet || null;
    doc.mappingConfig = state.currentMappingConfig || null;
    doc.canonicalData = state.currentCanonicalData || null;
    doc.auditResult = state.currentAuditResult || null;

    doc.ingestionStatus = state.currentParsedWb ? 'DONE' : 'PENDING';
    doc.mappingStatus = state.currentMappingConfig ? 'DONE' : 'PENDING';
    doc.auditStatus = state.currentAuditResult ? 'DONE' : 'PENDING';

    var comparison = window.ComparisonUI && window.ComparisonUI.getState ? window.ComparisonUI.getState() : {};
    var pricing = window.PricingRiskUI && window.PricingRiskUI.getState ? window.PricingRiskUI.getState() : {};
    var sourceSave = await persistSourceFiles(doc, comparison, pricing);
    try {
      doc.projectSnapshot = await createVersionedSnapshot(doc, state, sourceSave.refs);
      doc.snapshotSchemaVersion = doc.projectSnapshot.snapshotSchemaVersion;
      doc.snapshotUpdatedAt = doc.projectSnapshot.updatedAt;
      doc.updatedAt = new Date().toISOString();

      await window.ProjectStorage.updateDocument(doc);
      doc.snapshotSaveEvidence = { committed: true, snapshotId: doc.projectSnapshot.snapshotId, schemaVersion: doc.projectSnapshot.snapshotSchemaVersion, sourceCount: doc.projectSnapshot.sourceRegistry.length,
        engineStatuses: Object.assign({}, doc.projectSnapshot.engineStates), committedAt: doc.updatedAt };
    } catch (saveError) {
      await Promise.all(sourceSave.createdDocumentIds.map(function (id) { return window.ProjectStorage.deleteDocument(id).catch(function () {}); }));
      throw saveError;
    }
    renderDocumentRegistry();

    return doc;
  }
  /** UI rendering helpers – simplistic implementations using prompts and alerts */
  function renderProjectList() {
    const container = document.getElementById('project-list-overlay');
    if (!container) return;
    window.ProjectStorage.getAllProjects().then(projects => {
      let html = '<h2>Projects</h2><ul>'; // basic list
      projects.forEach(p => {
        html += `<li>${escapeHtml(p.projectName)} ` +
                `<button data-id="${p.projectId}" class="proj-open">Open</button> ` +
                `<button data-id="${p.projectId}" class="proj-rename">Rename</button> ` +
                `<button data-id="${p.projectId}" class="proj-delete">Delete</button></li>`;
      });
      html += '</ul><button id="proj-create">Create New Project</button> ' +
              '<button id="proj-close">Close</button>';
      container.innerHTML = html;
      // bind actions
      container.querySelectorAll('.proj-open').forEach(btn => {
        btn.addEventListener('click', e => loadProject(e.target.dataset.id));
      });
      container.querySelectorAll('.proj-rename').forEach(btn => {
        btn.addEventListener('click', async e => {
          const newName = prompt('Enter new project name:');
          if (newName) await renameProject(newName);
        });
      });
      container.querySelectorAll('.proj-delete').forEach(btn => {
        btn.addEventListener('click', async e => {
          if (confirm('Delete this project and all its documents?')) {
            activeProjectId = e.target.dataset.id;
            await deleteCurrentProject();
          }
        });
      });
      const createBtn = container.querySelector('#proj-create');
      if (createBtn) createBtn.addEventListener('click', async () => {
        const name = prompt('Project name:');
        if (name) await createProject(name);
      });
      const closeBtn = container.querySelector('#proj-close');
      if (closeBtn) closeBtn.addEventListener('click', () => { container.style.display = 'none'; });
    });
  }

  function renderProjectDetail() {
    const panel = document.getElementById('project-detail-panel');
    if (!panel) return;
    if (!activeProjectId) {
      panel.innerHTML = '<p>No project selected.</p>';
      return;
    }
    window.ProjectStorage.getAllProjects().then(projects => {
      const proj = projects.find(p => p.projectId === activeProjectId);
      if (!proj) return;
      const html = `<h3>${escapeHtml(proj.projectName)}</h3>` +
        `<p>Created: ${new Date(proj.createdAt).toLocaleString()}</p>` +
        `<p>Updated: ${new Date(proj.updatedAt).toLocaleString()}</p>` +
        `<button id="proj-add-doc">Add Document</button> ` +
        `<button id="proj-show-docs">Show Documents</button>`;
      panel.innerHTML = html;
      const addBtn = document.getElementById('proj-add-doc');
      if (addBtn) addBtn.addEventListener('click', () => {
        // trigger hidden file input (reuse existing file input)
        if (window.UniversalUI && window.UniversalUI.fileInput) {
          window.UniversalUI.fileInput.click();
        }
      });
      const showBtn = document.getElementById('proj-show-docs');
      if (showBtn) showBtn.addEventListener('click', () => {
        renderDocumentRegistry();
      });
    });
  }

  function renderDocumentRegistry() {
    const container = document.getElementById('document-registry-overlay');
    if (!container) return;
    if (!activeProjectId) {
      container.innerHTML = '<p>No project selected.</p>';
      return;
    }
    window.ProjectStorage.getDocumentsByProject(activeProjectId).then(docs => {
      let html = '<h2>Documents</h2><ul>';
      docs.forEach(d => {
        const isActive = d.documentId === activeDocumentId;
        html += `<li${isActive ? ' style="font-weight: bold;"' : ''}>${escapeHtml(d.fileName)} ` +
                `(${escapeHtml(d.documentRole)}) ` +
                `<button data-id="${d.documentId}" class="doc-open">Open</button> ` +
                `<button data-id="${d.documentId}" class="doc-delete">Delete</button></li>`;
      });
      html += '</ul><button id="doc-close">Close</button>';
      container.innerHTML = html;
      container.querySelectorAll('.doc-open').forEach(btn => {
        btn.addEventListener('click', e => setActiveDocument(e.target.dataset.id));
      });
      container.querySelectorAll('.doc-delete').forEach(btn => {
        btn.addEventListener('click', async e => {
          if (confirm('Delete this document from project?')) {
            await window.ProjectStorage.deleteDocument(e.target.dataset.id);
            if (activeDocumentId === e.target.dataset.id) activeDocumentId = null;
            renderDocumentRegistry();
          }
        });
      });
      const closeBtn = container.querySelector('#doc-close');
      if (closeBtn) closeBtn.addEventListener('click', () => { container.style.display = 'none'; });
    });
  }

  /** Utility – safe HTML escaping */
  function escapeHtml(str) {
    if (!str) return '';
    return String(str).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
  }

  // expose manager API
  window.ProjectManager = {
    init: function () {
      // create hidden overlay containers if not present
      if (!document.getElementById('project-list-overlay')) {
        const overlay = document.createElement('div');
        overlay.id = 'project-list-overlay';
        overlay.style = 'position:fixed;top:0;left:0;width:100%;height:100%;background:#fff;z-index:10000;padding:20px;overflow:auto;display:none;';
        document.body.appendChild(overlay);
      }
      if (!document.getElementById('project-detail-panel')) {
        const panel = document.createElement('div');
        panel.id = 'project-detail-panel';
        panel.style = 'margin-top:8px;';
        document.getElementById('main-nav-tabs').appendChild(panel);
      }
      if (!document.getElementById('document-registry-overlay')) {
        const overlay = document.createElement('div');
        overlay.id = 'document-registry-overlay';
        overlay.style = 'position:fixed;top:0;right:0;width:300px;height:100%;background:#f9f9f9;z-index:10000;padding:10px;overflow:auto;display:none;box-shadow:-2px 0 5px rgba(0,0,0,0.1);';
        document.body.appendChild(overlay);
      }
      // bind Projects button (add to header if missing)
      const header = document.querySelector('.app-header .brand-area');
      if (header && !document.getElementById('btn-projects')) {
        const btn = document.createElement('button');
        btn.id = 'btn-projects';
        btn.textContent = 'Projects';
        btn.className = 'sample-btn';
        btn.style.marginLeft = '12px';
        btn.addEventListener('click', () => {
          const overlay = document.getElementById('project-list-overlay');
          overlay.style.display = 'block';
          renderProjectList();
        });
        header.appendChild(btn);
      }
    },
    createProject,
    loadProject,
    renameProject,
    deleteCurrentProject,
    addDocument,
    setActiveDocument,
    persistDocumentUpdates
  };
})(typeof window !== 'undefined' ? window : this);
