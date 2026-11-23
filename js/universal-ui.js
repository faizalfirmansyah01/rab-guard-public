/**
 * RAB GUARD — Master Universal UI & Application Controller
 * Module: UniversalUI (js/universal-ui.js)
 * 
 * Tanggung Jawab:
 * - Mengatur alur produk berbasis masalah pengguna:
 *   1. Empty State (Pilih berkas sendiri atau buka salah satu dari 3 Demo Sintetis).
 *   2. Ingestion & Inspeksi (Deteksi sheet, pratinjau 15 baris, pemetaan kolom kanonikal).
 *   3. Konfirmasi Pengguna.
 *   4. Dashboard Dinamis (Ringkasan, Daftar Pekerjaan, Audit Integritas, Simulasi Sensitivitas, Feedback).
 * - Menjamin standar akurasi, transparansi, dan proveniensi sel sumber.
 * - Terintegrasi dengan UserFeedback (localStorage, false positive, rating, laporan issue, ekspor).
 * - Bebas dari ketergantungan dataset PNJ di runtime (PNJ murni fixture pengujian internal).
 * - Bebas dari kerentanan XSS (murni native DOM APIs + textContent).
 */

(function (window) {
  'use strict';

  var currentParsedWb = null;
  var currentInspection = null;
  var currentSelectedSheet = null;
  var currentSheetConfirmed = false;
  var currentMappingConfig = null;
  var currentHeaderConfirmed = false;
  var currentHeaderNeedsConfirmation = false;
  var currentCanonicalData = null;
  var currentAuditResult = null;
  var currentUnitConsistencyResult = null;
  var currentMaterialLogisticsResult = null;
  var auditRunInProgress = false;
  var auditAttempted = false;
  var currentFilter = 'all';
  var isSyntheticDemo = false;
  var activeDemoTitle = '';

  function syncNavDisabledReason(button, reason) {
    if (!button) return;
    var reasonNode = button.querySelector('.rg-nav-reason');
    if (reasonNode) reasonNode.textContent = reason || '';
  }
  var isInitialized = false;

  function formatRupiah(val) {
    if (typeof val !== 'number' || isNaN(val)) return 'Rp0';
    return 'Rp ' + val.toLocaleString('id-ID', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }

  function formatNumber(val) {
    if (typeof val !== 'number' || isNaN(val)) return '0';
    return val.toLocaleString('id-ID');
  }

  function getDivisionFilterValue(item) {
    var category = item && typeof item.category === 'string' ? item.category.trim() : '';
    return category && category !== 'UMUM' ? category : 'Belum diklasifikasikan';
  }

  function escapeHtml(str) {
    if (str === null || str === undefined) return '';
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  var UniversalUI = {
    init: function () {
      if (isInitialized) return;
      this.bindElements();
      this.bindEvents();
      this.initFeedbackUI();
      this.setInitialEmptyState();
      isInitialized = true;
    },

    bindElements: function () {
      // Header elements
      this.activeDocBadge = document.getElementById('active-doc-badge');
      this.activeDocName = document.getElementById('active-doc-name');
      this.btnSwitchDoc = document.getElementById('btn-switch-doc');

      // Navigation tabs
      this.navTabs = document.querySelectorAll('.tab-btn[data-tab]');
      this.tabPanels = document.querySelectorAll('.tab-panel');

      this.navBtnIngestion = document.getElementById('nav-btn-ingestion');
      this.navBtnRingkasan = document.getElementById('nav-btn-ringkasan');
      this.navBtnPekerjaan = document.getElementById('nav-btn-pekerjaan');
      this.navBtnAudit = document.getElementById('nav-btn-audit');
      this.navBtnSimulasi = document.getElementById('nav-btn-simulasi');
      this.navBtnUnitConsistency = document.getElementById('nav-btn-unit-consistency');
      this.navBtnMaterialLogistics = document.getElementById('nav-btn-material-logistics');
      this.navBtnFeedback = document.getElementById('nav-btn-feedback');
      this.badgeNavAudit = document.getElementById('badge-nav-audit');

      // Ingestion elements
      this.dropzone = document.getElementById('universal-dropzone');
      this.fileInput = document.getElementById('universal-file-input');
      this.btnBrowse = document.getElementById('btn-browse-file');

      // Inspection & Mapping containers
      this.stage2Container = document.getElementById('ingestion-stage-2');
      this.metaBox = document.getElementById('universal-meta-box');
      this.btnConfirmWorkbook = document.getElementById('btn-confirm-workbook');
      this.workbookConfirmationStatus = document.getElementById('workbook-confirmation-status');
      this.filenameEl = document.getElementById('universal-filename');
      this.fileInfoEl = document.getElementById('universal-file-info');

      this.sheetsContainer = document.getElementById('universal-sheets-container');
      this.tbodySheets = document.getElementById('tbody-universal-sheets');
      this.sheetSelectionNotice = document.getElementById('sheet-selection-notice');

      this.previewContainer = document.getElementById('universal-preview-container');
      this.previewSheetName = document.getElementById('preview-sheet-name');
      this.theadPreview = document.getElementById('thead-universal-preview');
      this.tbodyPreview = document.getElementById('tbody-universal-preview');

      this.mappingContainer = document.getElementById('universal-mapping-container');
      this.mappingHeaderInfo = document.getElementById('mapping-header-info');
      this.headerCorrectionControls = document.getElementById('header-correction-controls');
      this.headerRowChoice = document.getElementById('header-row-choice');
      this.btnConfirmHeaderRow = document.getElementById('btn-confirm-header-row');
      this.tbodyMapping = document.getElementById('tbody-universal-mapping');
      this.btnRunAudit = document.getElementById('btn-run-universal-audit');
      this.mappingActionStatus = document.getElementById('mapping-action-status');

      // Dashboard: Ringkasan
      this.metaSummaryProject = document.getElementById('meta-summary-project');
      this.metaSummarySheet = document.getElementById('meta-summary-sheet');
      this.metaSummaryDate = document.getElementById('meta-summary-date');
      this.metaSummaryScore = document.getElementById('meta-summary-score');

      this.statKpiTotal = document.getElementById('kpi-total-anggaran');
      this.statKpiItems = document.getElementById('kpi-total-items');
      this.statKpiFindings = document.getElementById('kpi-total-temuan');
      this.statKpiIntegrity = document.getElementById('kpi-integritas-status');

      this.tbodySubtotalSummary = document.getElementById('tbody-subtotal-summary');
      this.containerSubtotals = document.getElementById('container-subtotals-breakdown');

      // Dashboard: Pekerjaan
      this.searchPekerjaan = document.getElementById('search-pekerjaan');
      this.tbodyPekerjaan = document.getElementById('tbody-pekerjaan');
      this.footTotalPekerjaan = document.getElementById('foot-total-pekerjaan');
      this.footTotalImported = document.getElementById('foot-total-imported');

      // Dashboard: Audit
      this.statAuditTotal = document.getElementById('univ-stat-total');
      this.statAuditError = document.getElementById('univ-stat-error');
      this.statAuditWarning = document.getElementById('univ-stat-warning');
      this.statAuditInfo = document.getElementById('univ-stat-info');

      this.pillCountAll = document.getElementById('pill-univ-count-all');
      this.pillCountError = document.getElementById('pill-univ-count-error');
      this.pillCountWarning = document.getElementById('pill-univ-count-warning');
      this.pillCountInfo = document.getElementById('pill-univ-count-info');

      this.auditReportSummary = document.getElementById('univ-audit-report-summary');
      this.searchAudit = document.getElementById('univ-audit-search');
      this.auditSort = document.getElementById('univ-audit-sort');
      this.auditSheetFilter = document.getElementById('univ-audit-sheet-filter');
      this.auditDivisionFilter = document.getElementById('univ-audit-division-filter');
      this.auditFilteredCount = document.getElementById('univ-audit-filtered-count');
      this.tbodyFindings = document.getElementById('tbody-univ-findings');
      this.btnExportAudit = document.getElementById('btn-export-univ-audit');
      this.btnExportRelatedItems = document.getElementById('btn-export-univ-related-items');
      this.btnExportCanonical = document.getElementById('btn-export-univ-canonical');
      this.auditSourceViewerContainer = document.getElementById('audit-source-viewer');

      // Dashboard: Simulasi
      this.simInputPct = document.getElementById('univ-sim-input-pct');
      this.btnApplySim = document.getElementById('btn-apply-univ-sim');
      this.btnResetSim = document.getElementById('btn-reset-univ-sim');
      this.simTotalBase = document.getElementById('univ-sim-total-base');
      this.simTotalSim = document.getElementById('univ-sim-total-sim');
      this.simCoverage = document.getElementById('univ-sim-coverage');
      this.simDeltaRp = document.getElementById('univ-sim-delta-rp');
      this.simDeltaPct = document.getElementById('univ-sim-delta-pct');
      this.tbodyPareto = document.getElementById('tbody-univ-pareto');

      // Dashboard: Feedback
      this.feedbackStars = document.querySelectorAll('.star-rating span');
      this.feedbackComment = document.getElementById('feedback-comment');
      this.btnSubmitRating = document.getElementById('btn-submit-rating');

      this.issueSheet = document.getElementById('issue-sheet');
      this.issueCell = document.getElementById('issue-cell');
      this.issueDesc = document.getElementById('issue-desc');
      this.issueExpected = document.getElementById('issue-expected');
      this.btnSubmitIssue = document.getElementById('btn-submit-issue');

      this.propTitle = document.getElementById('prop-title');
      this.propCategory = document.getElementById('prop-category');
      this.propDesc = document.getElementById('prop-desc');
      this.btnSubmitProp = document.getElementById('btn-submit-prop');

      this.tbodyFalsePositives = document.getElementById('tbody-false-positives');
      this.btnExportFeedbackJSON = document.getElementById('btn-export-feedback-json');
      this.btnExportFeedbackCSV = document.getElementById('btn-export-feedback-csv');

      this.unitConsistencyStatus = document.getElementById('unit-consistency-status');
      this.unitConsistencySummary = document.getElementById('unit-consistency-summary');
      this.unitConsistencyWarnings = document.getElementById('unit-consistency-warnings');
      this.tbodyUnitConsistency = document.getElementById('tbody-unit-consistency-findings');
    },

    bindEvents: function () {
      var self = this;

      // Tab navigation
      for (var t = 0; t < this.navTabs.length; t++) {
        this.navTabs[t].addEventListener('click', function () {
          if (this.disabled) return;
          var targetId = this.getAttribute('data-tab');
          self.switchTab(targetId);
        });
      }

      // Switch / Change Document button
      if (this.btnSwitchDoc) {
        this.btnSwitchDoc.addEventListener('click', function () {
          self.resetWorkbookState();
        });
      }

      // File upload handlers
      if (this.btnBrowse && this.fileInput) {
        this.btnBrowse.addEventListener('click', function () {
          self.fileInput.click();
        });
      }

      if (this.fileInput) {
        this.fileInput.addEventListener('change', function (e) {
          if (e.target.files && e.target.files[0]) {
            self.routeFileUpload(e.target.files[0]);
          }
        });
      }

      // Drag and drop
      if (this.dropzone) {
        this.dropzone.addEventListener('dragover', function (e) {
          e.preventDefault();
          e.stopPropagation();
          self.dropzone.classList.add('drag-over');
        });

        this.dropzone.addEventListener('dragleave', function (e) {
          e.preventDefault();
          e.stopPropagation();
          self.dropzone.classList.remove('drag-over');
        });

        this.dropzone.addEventListener('drop', function (e) {
          e.preventDefault();
          e.stopPropagation();
          self.dropzone.classList.remove('drag-over');
          var dt = e.dataTransfer;
          if (dt && dt.files && dt.files.length > 0) {
            self.routeFileUpload(dt.files[0]);
          }
        });
      }

      // One primary workbook action starts the consolidated audit.
      if (this.btnConfirmWorkbook) {
        this.btnConfirmWorkbook.addEventListener('click', function () {
          self.runUniversalAudit();
        });
      }

      if (this.btnConfirmHeaderRow) {
        this.btnConfirmHeaderRow.addEventListener('click', function () {
          self.confirmHeaderRow();
        });
      }

      // Filter pills on Audit tab
      var pills = document.querySelectorAll('.filter-pill[data-univ-filter]');
      for (var p = 0; p < pills.length; p++) {
        pills[p].addEventListener('click', function () {
          for (var k = 0; k < pills.length; k++) pills[k].classList.remove('active');
          this.classList.add('active');
          currentFilter = this.getAttribute('data-univ-filter');
          self.renderFindings();
        });
      }

      if (this.navBtnUnitConsistency) {
        this.navBtnUnitConsistency.addEventListener('click', function () {
          if (this.disabled) return;
          self.renderUnitConsistency();
        });
      }

      // Live search filters
      if (this.searchAudit) {
        this.searchAudit.addEventListener('input', function () {
          self.renderFindings();
        });
      }

      if (this.auditSort) {
        this.auditSort.addEventListener('change', function () {
          self.renderFindings();
        });
      }
      [this.auditSheetFilter, this.auditDivisionFilter].forEach(function (select) {
        if (select) select.addEventListener('change', function () { self.renderFindings(); });
      });

      if (this.searchPekerjaan) {
        this.searchPekerjaan.addEventListener('input', function () {
          self.renderPekerjaanTable();
        });
      }

      // Simulation buttons
      if (this.btnApplySim) {
        this.btnApplySim.addEventListener('click', function () {
          self.runSimulation();
        });
      }

      if (this.btnResetSim) {
        this.btnResetSim.addEventListener('click', function () {
          if (self.simInputPct) self.simInputPct.value = 0;
          self.runSimulation();
        });
      }

      // Export buttons
      if (this.btnExportAudit) {
        this.btnExportAudit.addEventListener('click', function () {
          if (currentAuditResult && window.UniversalSimulasi) {
            var pName = currentCanonicalData ? currentCanonicalData.project.projectName : 'Audit';
            window.UniversalSimulasi.exportAuditCSV(currentAuditResult, pName, currentCanonicalData);
          }
        });
      }

      if (this.btnExportRelatedItems) {
        this.btnExportRelatedItems.addEventListener('click', function () {
          if (currentAuditResult && window.UniversalSimulasi) {
            var pName = currentCanonicalData ? currentCanonicalData.project.projectName : 'Audit';
            window.UniversalSimulasi.exportRelatedItemsCSV(currentAuditResult, pName, currentCanonicalData);
          }
        });
      }

      if (this.btnExportCanonical) {
        this.btnExportCanonical.addEventListener('click', function () {
          if (currentCanonicalData && window.UniversalSimulasi) {
            window.UniversalSimulasi.exportCanonicalCSV(currentCanonicalData);
          }
        });
      }
    },

    setInitialEmptyState: function () {
      this.resetWorkbookState();
    },

    /**
     * Membersihkan dan mereset status workbook secara terpusat (Zero State Bleed).
     * Memastikan tidak ada sisa artefak, cache objek, atau DOM dari workbook sebelumnya.
     */
    resetWorkbookState: function () {
      // 1. Reset seluruh variabel status internal
      currentParsedWb = null;
      currentInspection = null;
      currentSelectedSheet = null;
      currentSheetConfirmed = false;
      currentMappingConfig = null;
      currentHeaderConfirmed = false;
      currentHeaderNeedsConfirmation = false;
      currentCanonicalData = null;
      currentAuditResult = null;
      currentUnitConsistencyResult = null;
      currentMaterialLogisticsResult = null;
      if (window.MaterialLogisticsUI && window.MaterialLogisticsUI.clear) window.MaterialLogisticsUI.clear();
      auditRunInProgress = false;
      auditAttempted = false;
      currentFilter = 'all';
      isSyntheticDemo = false;
      activeDemoTitle = '';

      // 2. Kosongkan input berkas agar berkas yang sama persis dapat diunggah ulang
      if (this.fileInput) {
        this.fileInput.value = '';
      }

      // 3. Reset Header Dokumen Aktif
      if (this.activeDocBadge) this.activeDocBadge.style.display = 'none';
      if (this.activeDocName) this.activeDocName.textContent = '';
      if (this.btnSwitchDoc) this.btnSwitchDoc.style.display = 'none';
      if (this.badgeNavAudit) {
        this.badgeNavAudit.style.display = 'none';
        this.badgeNavAudit.textContent = '0';
      }

      // 4. Sembunyikan Ingestion Stage 2 dan kontainer inspeksi/pemetaan
      if (this.stage2Container) this.stage2Container.style.display = 'none';
      if (this.metaBox) this.metaBox.style.display = 'none';
      if (this.previewContainer) this.previewContainer.style.display = 'none';
      if (this.mappingContainer) this.mappingContainer.style.display = 'none';
      if (this.containerSubtotals) this.containerSubtotals.style.display = 'none';
      if (this.filenameEl) this.filenameEl.textContent = '';
      if (this.fileInfoEl) this.fileInfoEl.textContent = '';
      if (this.btnConfirmWorkbook) this.btnConfirmWorkbook.disabled = true;
      if (this.workbookConfirmationStatus) this.workbookConfirmationStatus.textContent = 'Pilih atau tinjau sheet sebelum melanjutkan.';
      if (this.previewSheetName) this.previewSheetName.textContent = '-';
      if (this.mappingHeaderInfo) this.mappingHeaderInfo.textContent = '';
      if (this.mappingActionStatus) this.mappingActionStatus.textContent = 'Konfirmasi workbook dan selesaikan pemetaan wajib untuk menjalankan audit.';
      if (this.btnRunAudit) this.btnRunAudit.disabled = true;
      if (this.sheetSelectionNotice) this.sheetSelectionNotice.textContent = '';
      if (this.headerCorrectionControls) this.headerCorrectionControls.style.display = 'none';
      if (this.headerRowChoice) this.headerRowChoice.innerHTML = '';

      // 5. Bersihkan tabel DOM pada Stage Ingestion
      if (this.tbodySheets) this.tbodySheets.innerHTML = '';
      if (this.theadPreview) this.theadPreview.innerHTML = '';
      if (this.tbodyPreview) this.tbodyPreview.innerHTML = '';
      if (this.tbodyMapping) this.tbodyMapping.innerHTML = '';

      // 6. Reset Metadata dan Kartu KPI pada Tab Ringkasan
      if (this.metaSummaryProject) this.metaSummaryProject.textContent = 'Nama Dokumen';
      if (this.metaSummarySheet) this.metaSummarySheet.textContent = '-';
      if (this.metaSummaryDate) this.metaSummaryDate.textContent = '-';
      if (this.metaSummaryScore) this.metaSummaryScore.textContent = '—';

      if (this.statKpiTotal) this.statKpiTotal.textContent = '-';
      var subCard = (this.statKpiTotal && this.statKpiTotal.parentElement) ? this.statKpiTotal.parentElement.querySelector('.card-subtitle') : null;
      if (subCard) subCard.textContent = 'Rekalkulasi independen volume × harga';

      if (this.statKpiItems) this.statKpiItems.textContent = '-';
      if (this.statKpiFindings) this.statKpiFindings.textContent = '-';
      if (this.statKpiIntegrity) {
        this.statKpiIntegrity.textContent = 'MEMERIKSA';
        this.statKpiIntegrity.style.color = '';
      }
      if (this.tbodySubtotalSummary) this.tbodySubtotalSummary.innerHTML = '';

      // 7. Reset Tab Pekerjaan
      if (this.searchPekerjaan) this.searchPekerjaan.value = '';
      if (this.tbodyPekerjaan) this.tbodyPekerjaan.innerHTML = '';
      if (this.footTotalPekerjaan) this.footTotalPekerjaan.textContent = 'Rp0';
      if (this.footTotalImported) this.footTotalImported.textContent = 'Rp0';

      // 8. Reset Tab Audit Integritas
      if (this.statAuditTotal) this.statAuditTotal.textContent = '0';
      if (this.statAuditError) this.statAuditError.textContent = '0';
      if (this.statAuditWarning) this.statAuditWarning.textContent = '0';
      if (this.statAuditInfo) this.statAuditInfo.textContent = '0';
      if (this.pillCountAll) this.pillCountAll.textContent = '0';
      if (this.pillCountError) this.pillCountError.textContent = '0';
      if (this.pillCountWarning) this.pillCountWarning.textContent = '0';
      if (this.pillCountInfo) this.pillCountInfo.textContent = '0';
      if (this.searchAudit) this.searchAudit.value = '';
      if (this.auditSort) this.auditSort.value = 'source';
      if (this.auditFilteredCount) this.auditFilteredCount.textContent = '0 / 0 temuan';
      if (this.auditReportSummary) {
        var reportIds = ['univ-report-filename', 'univ-report-sheet', 'univ-report-items',
          'univ-report-coverage', 'univ-report-imported', 'univ-report-recalculated', 'univ-report-difference'];
        reportIds.forEach(function (id, index) {
          var node = document.getElementById(id);
          if (node) node.textContent = index === 0 ? 'Belum ada workbook' : '-';
        });
      }
      if (this.tbodyFindings) this.tbodyFindings.innerHTML = '';
      this.clearUnitConsistencyResult();

      // 9. Reset Tab Simulasi Sensitivitas
      if (this.simInputPct) this.simInputPct.value = '0';
      if (this.simTotalBase) this.simTotalBase.textContent = 'Rp0';
      if (this.simTotalSim) this.simTotalSim.textContent = 'Rp0';
      if (this.simCoverage) this.simCoverage.textContent = 'Cakupan simulasi belum dihitung.';
      if (this.simDeltaRp) this.simDeltaRp.textContent = '+Rp0';
      if (this.simDeltaPct) this.simDeltaPct.textContent = '+0.00%';
      if (this.tbodyPareto) this.tbodyPareto.innerHTML = '';

      // 10. Nonaktifkan Tab Dashboard dan Kembali ke Tab Ingestion (Empty State Murni)
      this.setDashboardTabsEnabled(false);
      this.switchTab('tab-ingestion');
    },

    getCurrentState: function () {
      return {
        currentParsedWb: currentParsedWb,
        currentInspection: currentInspection,
        currentSelectedSheet: currentSelectedSheet,
        currentSheetConfirmed: currentSheetConfirmed,
        currentMappingConfig: currentMappingConfig,
        currentHeaderConfirmed: currentHeaderConfirmed,
        currentHeaderNeedsConfirmation: currentHeaderNeedsConfirmation,
        currentCanonicalData: currentCanonicalData,
        currentAuditResult: currentAuditResult,
        currentUnitConsistencyResult: currentUnitConsistencyResult,
        currentMaterialLogisticsResult: currentMaterialLogisticsResult,
        currentFilter: currentFilter,
        isSyntheticDemo: isSyntheticDemo,
        activeDemoTitle: activeDemoTitle
      };
    },

    persistProjectState: function () {
      if (!window.ProjectManager ||
          typeof window.ProjectManager.persistDocumentUpdates !== 'function') {
        return Promise.resolve(false);
      }

      return window.ProjectManager.persistDocumentUpdates().catch(function (err) {
        console.error('[RAB Guard] Gagal menyimpan state dokumen:', err);
        var message = 'Perubahan belum tersimpan' + (err && err.name === 'QuotaExceededError' ? ' karena ruang penyimpanan browser penuh.' : ': ' + (err && err.message || 'penyimpanan gagal.')) + ' Coba ulangi penyimpanan setelah masalah diperbaiki.';
        if (this.workbookConfirmationStatus) this.workbookConfirmationStatus.textContent = message;
        return false;
      }.bind(this));
    },

    restoreCurrentState: function (state, displayName) {
      state = state || {};

      // Mulai dari kondisi bersih, lalu masukkan state dokumen terpilih.
      this.resetWorkbookState();

      currentParsedWb = state.currentParsedWb || null;
      currentInspection = currentParsedWb && window.WorkbookInspector
        ? window.WorkbookInspector.inspect(currentParsedWb)
        : (state.currentInspection || null);

      currentSelectedSheet = state.currentSelectedSheet || null;
      currentSheetConfirmed = state.currentSheetConfirmed === true;
      currentMappingConfig = state.currentMappingConfig || null;
      currentHeaderConfirmed = state.currentHeaderConfirmed === true;
      currentHeaderNeedsConfirmation = state.currentHeaderNeedsConfirmation === true;
      currentCanonicalData = state.currentCanonicalData || null;
      if (currentCanonicalData && window.DataNormalizer && typeof window.DataNormalizer.ensureAhspContract === 'function') {
        currentCanonicalData = window.DataNormalizer.ensureAhspContract(
          currentCanonicalData, currentParsedWb, currentInspection
        );
      }
      currentAuditResult = state.currentAuditResult || null;
      currentUnitConsistencyResult = state.currentUnitConsistencyResult || null;
      currentMaterialLogisticsResult = state.currentMaterialLogisticsResult || null;
      auditRunInProgress = false;
      currentFilter = state.currentFilter || 'all';

      // Dokumen yang dipulihkan adalah file nyata, bukan demo sintetis.
      isSyntheticDemo = false;
      activeDemoTitle = '';

      if (currentParsedWb && currentInspection) {
        if (this.stage2Container) this.stage2Container.style.display = 'block';
        if (this.metaBox) this.metaBox.style.display = 'block';

        if (this.filenameEl) {
          this.filenameEl.textContent =
            displayName || currentParsedWb.filename || 'Workbook';
        }

        if (this.fileInfoEl) {
          var restoredSize = Number(currentParsedWb.fileSize) || 0;
          this.fileInfoEl.textContent =
            'Ukuran: ' + (restoredSize / 1024).toFixed(1) +
            ' KB • ' + currentInspection.sheetCount +
            ' Lembar Kerja • Terdeteksi Utama: ' +
            (currentInspection.primaryRabSheet || '-');
        }

        this.renderSheetsTable();

        var availableNames = currentParsedWb.sheetNames ||
          Object.keys(currentParsedWb.sheets || {});

        if (!currentSelectedSheet ||
            !currentParsedWb.sheets ||
            !currentParsedWb.sheets[currentSelectedSheet]) {
          currentSelectedSheet =
            currentInspection.primaryRabSheet ||
            availableNames[0] ||
            null;
        }

        var restoredSheetMeta = currentInspection.sheets
          ? currentInspection.sheets.find(function (s) {
              return s.sheetName === currentSelectedSheet;
            })
          : null;

        if (restoredSheetMeta) {
          this.renderPreview(restoredSheetMeta);

          var restoredSheet =
            currentParsedWb.sheets[currentSelectedSheet];

          if (restoredSheet &&
              window.HeaderDetector &&
              window.ColumnMapper) {
            try {
              var restoredDetection =
                window.HeaderDetector.detect(restoredSheet);

              if (!currentMappingConfig) {
                currentMappingConfig =
                  window.ColumnMapper.createInitialMapping(restoredDetection);
              }

              this.renderMappingInterface(restoredDetection);
            } catch (restoreMappingError) {
              console.error(
                '[RAB Guard] Gagal memulihkan tampilan mapping:',
                restoreMappingError
              );
            }
          }
        }

        if (this.stage2Container) {
          this.stage2Container.style.display = 'block';
        }
      }

      var restoredTitle =
        (currentCanonicalData &&
         currentCanonicalData.project &&
         currentCanonicalData.project.projectName) ||
        displayName ||
        (currentParsedWb && currentParsedWb.filename) ||
        'Dokumen';

      if (this.activeDocBadge &&
          (currentParsedWb || currentCanonicalData)) {
        this.activeDocBadge.style.display = 'inline-flex';
      }

      if (this.activeDocName &&
          (currentParsedWb || currentCanonicalData)) {
        this.activeDocName.textContent =
          restoredTitle +
          (currentSelectedSheet ? ' (' + currentSelectedSheet + ')' : '');
      }

      if (this.btnSwitchDoc &&
          (currentParsedWb || currentCanonicalData)) {
        this.btnSwitchDoc.style.display = 'inline-block';
      }

      // Pulihkan dashboard hanya jika data normalisasi dan hasil audit ada.
      if (currentCanonicalData && currentAuditResult) {
        try {
          var restoredValidation = window.IngestionValidator
            ? window.IngestionValidator.validate(currentCanonicalData)
            : { valid: true, errors: [] };

          var restoredBaseline = window.UniversalSimulasi.recalculateBaseline(currentCanonicalData);
          this.renderRingkasanTab(restoredValidation);
          this.renderPekerjaanTable();
          this.renderFindings();
          if (restoredBaseline.calculableItemsCount > 0) this.runSimulation();
          else this.setSimulationUnavailable('Tidak ada item dengan volume dan harga satuan numerik yang valid.');
          this.renderFalsePositivesTable();
          var restoredUnitStatus = state.restoreMetadata && state.restoreMetadata.engineStates
            ? state.restoreMetadata.engineStates.unitConsistency
            : null;
          if (!currentUnitConsistencyResult && !restoredUnitStatus) {
            this.runUnitConsistencyAnalysis();
          } else if (!currentUnitConsistencyResult && this.unitConsistencyStatus) {
            this.unitConsistencyStatus.textContent = 'Hasil Unit Consistency tidak dipulihkan sebagai hasil terkini (' + restoredUnitStatus + '). Jalankan analisis kembali untuk membuat hasil baru.';
          }
          if (window.MaterialLogisticsUI && typeof window.MaterialLogisticsUI.restoreState === 'function') {
            window.MaterialLogisticsUI.restoreState(state.materialLogisticsState || { result: currentMaterialLogisticsResult }, currentCanonicalData);
          }
          this.setDashboardTabsEnabled(true, restoredBaseline.calculableItemsCount > 0,
            Array.isArray(currentCanonicalData.workItems) && currentCanonicalData.workItems.length > 0);
          this.switchTab('tab-ringkasan');
        } catch (restoreDashboardError) {
          this.setDashboardTabsEnabled(false);
          console.error(
            '[RAB Guard] Gagal memulihkan dashboard:',
            restoreDashboardError
          );
        }
      } else {
        this.setDashboardTabsEnabled(false);
        this.switchTab('tab-ingestion');
      }

      if (state.restoreMetadata && state.restoreMetadata.engineStates) {
        var restoreStatus = Object.keys(state.restoreMetadata.engineStates).map(function (key) { return key + ': ' + state.restoreMetadata.engineStates[key]; }).join(' · ');
        if (this.workbookConfirmationStatus) this.workbookConfirmationStatus.textContent =
          'Pemulihan snapshot v' + state.restoreMetadata.snapshotVersion + ' (' + state.restoreMetadata.sourceIntegrity + '). Status engine: ' + restoreStatus;
      }

      this.updateWorkflowActions();

      return true;
    },
    setDashboardTabsEnabled: function (enabled, simulationAvailable, unitAnalysisAvailable) {
      var tabs = [this.navBtnRingkasan, this.navBtnPekerjaan, this.navBtnAudit];
      tabs.forEach(function (btn) {
        if (!btn) return;
        if (enabled) {
          btn.removeAttribute('disabled');
          btn.removeAttribute('title');
          syncNavDisabledReason(btn, '');
        } else {
          btn.setAttribute('disabled', 'disabled');
          btn.title = 'Unggah workbook dan jalankan audit terlebih dahulu.';
          syncNavDisabledReason(btn, btn.title);
        }
      });
      if (this.navBtnSimulasi) {
        var canOpenSimulation = !!enabled && !!simulationAvailable;
        this.navBtnSimulasi.disabled = !canOpenSimulation;
        this.navBtnSimulasi.title = canOpenSimulation
          ? ''
          : (enabled ? 'Simulasi memerlukan sedikitnya satu item dengan volume dan harga satuan numerik valid.'
            : 'Simulasi tersedia setelah audit dan memerlukan data hitung ulang yang valid.');
        syncNavDisabledReason(this.navBtnSimulasi, this.navBtnSimulasi.title);
      }
      if (this.navBtnUnitConsistency) {
        var canOpenUnitAnalysis = !!enabled && !!unitAnalysisAvailable;
        this.navBtnUnitConsistency.disabled = !canOpenUnitAnalysis;
        this.navBtnUnitConsistency.title = canOpenUnitAnalysis ? '' :
          (enabled ? 'Pengecekan satuan memerlukan setidaknya satu item pekerjaan.' : 'Jalankan audit workbook terlebih dahulu.');
        syncNavDisabledReason(this.navBtnUnitConsistency, this.navBtnUnitConsistency.title);
      }
      if (this.navBtnMaterialLogistics) {
        var canOpenMaterialLogistics = !!enabled && !!currentCanonicalData;
        this.navBtnMaterialLogistics.disabled = !canOpenMaterialLogistics;
        this.navBtnMaterialLogistics.title = canOpenMaterialLogistics ? '' :
          (enabled ? 'Canonical BOQ belum tersedia.' : 'Jalankan audit workbook terlebih dahulu.');
        syncNavDisabledReason(this.navBtnMaterialLogistics, this.navBtnMaterialLogistics.title);
      }
      if (window.MaterialLogisticsUI && window.MaterialLogisticsUI.refresh) window.MaterialLogisticsUI.refresh();
    },

    setMaterialLogisticsResult: function (result) {
      currentMaterialLogisticsResult = result || null;
    },

    setSimulationUnavailable: function (reason) {
      if (this.simCoverage) this.simCoverage.textContent = 'Simulasi tidak tersedia: ' + reason;
      if (this.simTotalBase) this.simTotalBase.textContent = '—';
      if (this.simTotalSim) this.simTotalSim.textContent = '—';
      if (this.simDeltaRp) this.simDeltaRp.textContent = '—';
      if (this.simDeltaPct) this.simDeltaPct.textContent = '—';
      if (this.tbodyPareto) {
        this.tbodyPareto.innerHTML = '';
        var row = document.createElement('tr');
        var cell = document.createElement('td');
        cell.colSpan = 7;
        cell.textContent = reason;
        row.appendChild(cell);
        this.tbodyPareto.appendChild(row);
      }
    },

    switchTab: function (tabId) {
      if (!tabId) return;
      for (var i = 0; i < this.navTabs.length; i++) {
        var btn = this.navTabs[i];
        if (btn.getAttribute('data-tab') === tabId) {
          btn.classList.add('active');
        } else {
          btn.classList.remove('active');
        }
      }

      for (var j = 0; j < this.tabPanels.length; j++) {
        var panel = this.tabPanels[j];
        if (panel.id === tabId) {
          panel.classList.add('active');
        } else {
          panel.classList.remove('active');
        }
      }

      if (tabId === 'tab-audit') {
        this.renderSourceViewer();
      }
    },

    routeFileUpload: function (file) {
      if (!file) return;

      if (window.ProjectManager && typeof window.ProjectManager.addDocument === 'function') {
        window.ProjectManager.addDocument(file).catch(function (err) {
          console.error('[RAB Guard] Gagal menyimpan dokumen:', err);
          alert('Gagal menyimpan dokumen proyek: ' + (err.message || err));
        });
        return;
      }

      // Fallback jika Project Workspace belum tersedia
      this.handleFileUpload(file);
    },
    handleFileUpload: function (file) {
      var self = this;
      if (!window.XlsxReader) {
        alert('Modul XlsxReader belum siap.');
        return;
      }

      isSyntheticDemo = false;
      activeDemoTitle = '';

      return window.XlsxReader.readFileAsync(file)
        .then(function (wb) {
          self.processLoadedWorkbook(wb, file.name);
        })
        .catch(function (err) {
          alert('Gagal memproses berkas Excel: ' + err.message);
        });
    },

    loadSyntheticDemo: function (key, filename, demoTitle) {
      if (!window.TEST_WORKBOOKS_B64 || !window.TEST_WORKBOOKS_B64[key] || !window.b64ToUint8Array) {
        alert('Fixtures demo sintetis belum termuat.');
        return;
      }

      isSyntheticDemo = true;
      activeDemoTitle = demoTitle;

      var bytes = window.b64ToUint8Array(window.TEST_WORKBOOKS_B64[key]);
      var wb = window.XlsxReader.readUint8Array(bytes, filename);
      this.processLoadedWorkbook(wb, filename);
    },

    processLoadedWorkbook: function (wb, displayName) {
      var loadedAsDemo = isSyntheticDemo;
      var loadedDemoTitle = activeDemoTitle;
      this.resetWorkbookState();
      isSyntheticDemo = loadedAsDemo;
      activeDemoTitle = loadedDemoTitle;
      currentParsedWb = wb;
      currentInspection = window.WorkbookInspector.inspect(wb);

      if (this.stage2Container) this.stage2Container.style.display = 'block';
      if (this.metaBox) this.metaBox.style.display = 'block';

      var labelText = isSyntheticDemo ? ('[DATA SINTETIS] ' + activeDemoTitle) : (displayName || wb.filename);
      if (this.filenameEl) this.filenameEl.textContent = labelText;
      if (this.fileInfoEl) {
        var kb = (wb.fileSize / 1024).toFixed(1);
        var readyCandidate = currentInspection.sheets.filter(function (sheetMeta) { return sheetMeta.normalizationReady; })[0];
        var semanticCandidate = currentInspection.sheets.filter(function (sheetMeta) { return sheetMeta.semanticRole === 'PRIMARY_BOQ'; })[0];
        this.fileInfoEl.textContent = 'Ukuran: ' + kb + ' KB • ' + currentInspection.sheetCount + ' Lembar Kerja • ' +
          (readyCandidate
            ? 'Kandidat BOQ siap dinormalisasi: ' + readyCandidate.sheetName
            : (semanticCandidate ? 'Kandidat BOQ terdeteksi; header/mapping belum dipastikan' : 'Kandidat BOQ belum dipastikan'));
      }

      this.renderSheetsTable();

      // Workbook-level workflow: do not require per-sheet choices before audit.
      currentSelectedSheet = null;
      currentSheetConfirmed = false;
      currentMappingConfig = null;
      currentHeaderConfirmed = false;
      currentHeaderNeedsConfirmation = false;
      if (this.mappingContainer) this.mappingContainer.style.display = 'none';
      if (this.sheetSelectionNotice) {
        this.sheetSelectionNotice.textContent = 'Seluruh ' + currentInspection.sheetCount +
          ' sheet telah diperiksa. Sheet BOQ yang yakin akan diproses bersama; sheet yang belum jelas akan dicantumkan sebagai belum terselesaikan.';
      }
      this.updateWorkflowActions();

      if (this.stage2Container) {
        this.stage2Container.scrollIntoView({ behavior: 'smooth' });
      }
    },

    renderSheetsTable: function () {
      var self = this;
      if (!this.tbodySheets || !currentInspection) return;

      if (this.sheetsContainer) this.sheetsContainer.style.display = 'block';
      this.tbodySheets.innerHTML = '';

      currentInspection.sheets.forEach(function (s) {
        var tr = document.createElement('tr');
        if (s.sheetName === currentSelectedSheet) {
          tr.style.backgroundColor = 'var(--primary-light)';
        }

        var shownRole = s.semanticRole === 'UNKNOWN' || !s.semanticRole && s.confidence < 70
          ? 'UNRESOLVED' : (s.semanticRole || s.detectedRole || 'UNRESOLVED');
        var roleClass = 'role-other';
        if (shownRole === 'PRIMARY_BOQ') roleClass = 'role-rab';
        else if (shownRole === 'AHSP_PRICE_REFERENCE') roleClass = 'role-ahsp';
        else if (shownRole === 'RECAP') roleClass = 'role-rekap';
        else if (shownRole === 'SCHEDULE') roleClass = 'role-kurva';

        var td1 = document.createElement('td');
        var strong1 = document.createElement('strong');
        strong1.textContent = s.sheetName;
        td1.appendChild(strong1);
        if (s.isHidden) {
          var hiddenSpan = document.createElement('span');
          hiddenSpan.style.fontSize = '11px';
          hiddenSpan.style.color = 'var(--text-muted)';
          hiddenSpan.textContent = ' (Tersembunyi)';
          td1.appendChild(hiddenSpan);
        }
        tr.appendChild(td1);

        var td2 = document.createElement('td');
        td2.textContent = (s.range || '-') + ' (' + s.rowCount + ' baris × ' + s.colCount + ' kolom)';
        tr.appendChild(td2);

        var td3 = document.createElement('td');
        td3.className = 'num-col';
        td3.textContent = s.cellCount + ' sel (' + s.formulaCount + ' rumus)';
        tr.appendChild(td3);

        var td4 = document.createElement('td');
        var roleBadge = document.createElement('span');
        roleBadge.className = 'role-badge ' + roleClass;
        roleBadge.textContent = shownRole;
        td4.appendChild(roleBadge);
        tr.appendChild(td4);

        var td5 = document.createElement('td');
        td5.className = 'center-col';
        var strong5 = document.createElement('strong');
        strong5.textContent = (s.semanticConfidence !== undefined ? s.semanticConfidence : s.confidence) + '%';
        td5.appendChild(strong5);
        tr.appendChild(td5);

        var td6 = document.createElement('td');
        td6.className = 'center-col';
        var normalizedSheet = currentCanonicalData && currentCanonicalData.workbookManifest
          ? currentCanonicalData.workbookManifest.filter(function (entry) { return entry.name === s.sheetName; })[0]
          : null;
        if (normalizedSheet && normalizedSheet.status === 'processed') {
          td6.textContent = 'Berhasil dinormalisasi';
        } else if (normalizedSheet && normalizedSheet.status === 'unresolved') {
          td6.textContent = 'Belum diproses: ' + (normalizedSheet.unresolvedReasons || []).join(' ');
        } else if (shownRole === 'PRIMARY_BOQ' && s.normalizationReady) {
          td6.textContent = 'Header terdeteksi baris ' + s.headerDetection.row1Based + ' — siap dinormalisasi';
        } else if (shownRole === 'PRIMARY_BOQ') {
          td6.textContent = 'Kandidat BOQ; ' + ((s.headerDetection && s.headerDetection.reason) || 'header/mapping belum dipastikan');
        } else {
          td6.textContent = shownRole === 'UNRESOLVED' ? 'Perlu ditinjau setelah audit' : 'Referensi / tidak diagregasi';
        }
        tr.appendChild(td6);

        self.tbodySheets.appendChild(tr);
      });

    },

    selectSheet: function (sheetName) {
      if (!currentParsedWb || !currentParsedWb.sheets || !currentParsedWb.sheets[sheetName]) return false;
      this.invalidateAuditResults();
      currentSelectedSheet = sheetName;
      currentSheetConfirmed = false;
      this.renderSheetsTable();

      var sheetMeta = currentInspection.sheets.find(function (s) { return s.sheetName === sheetName; });
      this.renderPreview(sheetMeta);

      var sheet = currentParsedWb.sheets[sheetName];
      var detection = window.HeaderDetector.detect(sheet);
      currentMappingConfig = window.ColumnMapper.createInitialMapping(detection);
      currentHeaderNeedsConfirmation = !detection.found || detection.confidence < 70 || detection.ambiguous;
      currentHeaderConfirmed = !currentHeaderNeedsConfirmation;

      this.renderMappingInterface(detection);
      if (this.sheetSelectionNotice) {
        this.sheetSelectionNotice.textContent = 'Lembar "' + sheetName + '" dipilih. Periksa pratinjau dan konfirmasi baris header bila diminta.';
      }
      this.updateWorkflowActions();
      this.persistProjectState();
      return true;
    },

    invalidateAuditResults: function () {
      currentCanonicalData = null;
      currentAuditResult = null;
      currentUnitConsistencyResult = null;
      currentMaterialLogisticsResult = null;
      this.clearUnitConsistencyResult();
      if (this.badgeNavAudit) {
        this.badgeNavAudit.style.display = 'none';
        this.badgeNavAudit.textContent = '0';
      }
      this.setDashboardTabsEnabled(false);
      if (this.navBtnIngestion && !this.navBtnIngestion.classList.contains('active')) {
        this.switchTab('tab-ingestion');
      }
    },

    confirmWorkbook: function () {
      if (!currentParsedWb || !currentSelectedSheet || !currentMappingConfig) {
        if (this.workbookConfirmationStatus) {
          this.workbookConfirmationStatus.textContent = currentParsedWb
            ? 'Pilih lembar kerja RAB sebelum melanjutkan.'
            : 'Workbook belum berhasil dibaca.';
        }
        this.updateWorkflowActions();
        return false;
      }

      currentSheetConfirmed = true;
      if (this.workbookConfirmationStatus) {
        this.workbookConfirmationStatus.textContent = 'Workbook dan lembar "' + currentSelectedSheet +
          '" dikonfirmasi. Tinjau header serta pemetaan sebelum menjalankan audit.';
      }
      this.updateWorkflowActions();
      if (this.mappingContainer && this.mappingContainer.scrollIntoView) {
        this.mappingContainer.scrollIntoView({ behavior: 'smooth', block: 'start' });
      }
      this.persistProjectState();
      return true;
    },

    updateWorkflowActions: function () {
      var hasWorkbook = !!currentParsedWb;
      var ready = hasWorkbook;
      if (this.btnRunAudit) {
        this.btnRunAudit.disabled = true;
      }
      if (this.btnConfirmWorkbook) this.btnConfirmWorkbook.disabled = !ready || !!currentAuditResult || auditRunInProgress;
      if (this.workbookConfirmationStatus && hasWorkbook && !currentAuditResult && !auditAttempted) {
        this.workbookConfirmationStatus.textContent = 'Workbook siap diperiksa. Semua worksheet akan diproses otomatis sejauh dapat dikenali.';
      }

      if (!this.mappingActionStatus) return;
      if (currentAuditResult) {
        this.mappingActionStatus.textContent = 'Audit selesai. Hasil telah ditampilkan; ubah sheet atau pemetaan untuk menjalankan ulang.';
      } else if (!hasWorkbook) {
        this.mappingActionStatus.textContent = 'Unggah workbook untuk memulai pemeriksaan seluruh sheet.';
      } else {
        this.mappingActionStatus.textContent = 'Pemeriksaan akan memproses semua sheet BOQ yang teridentifikasi yakin dan melaporkan sheet yang belum jelas.';
      }
    },

    confirmHeaderRow: function () {
      if (!currentParsedWb || !currentSelectedSheet || !this.headerRowChoice) return false;
      var row1Based = parseInt(this.headerRowChoice.value, 10);
      if (!Number.isInteger(row1Based) || row1Based < 1) return false;

      var sheet = currentParsedWb.sheets[currentSelectedSheet];
      var detection = window.HeaderDetector.detect(sheet, { headerRowIndex: row1Based - 1 });
      if (!detection.found) {
        if (this.mappingHeaderInfo) {
          this.mappingHeaderInfo.textContent = 'Baris ' + row1Based + ' belum dikenali sebagai header. Pilih baris lain atau koreksi pemetaan kolom.';
        }
        currentHeaderConfirmed = false;
        currentHeaderNeedsConfirmation = true;
        this.updateWorkflowActions();
        return false;
      }

      this.invalidateAuditResults();
      currentMappingConfig = window.ColumnMapper.createInitialMapping(detection);
      currentHeaderConfirmed = true;
      currentHeaderNeedsConfirmation = false;
      this.renderMappingInterface(detection);
      this.updateWorkflowActions();
      this.persistProjectState();
      return true;
    },

    renderPreview: function (sheetMeta) {
      if (!sheetMeta || !this.previewContainer) return;
      this.previewContainer.style.display = 'block';
      if (this.previewSheetName) this.previewSheetName.textContent = sheetMeta.sheetName;

      this.theadPreview.innerHTML = '';
      this.tbodyPreview.innerHTML = '';

      var sampleRows = sheetMeta.sampleRows || [];
      if (sampleRows.length === 0) {
        var emptyTr = document.createElement('tr');
        var emptyTd = document.createElement('td');
        emptyTd.colSpan = 5;
        emptyTd.className = 'center-col';
        emptyTd.textContent = 'Tidak ada data untuk dipratinjau.';
        emptyTr.appendChild(emptyTd);
        this.tbodyPreview.appendChild(emptyTr);
        return;
      }

      var firstCells = sampleRows[0].cells;
      var trHead = document.createElement('tr');
      var thRow = document.createElement('th');
      thRow.style.width = '50px';
      thRow.className = 'center-col';
      thRow.textContent = 'Baris';
      trHead.appendChild(thRow);

      firstCells.forEach(function (c) {
        var thCol = document.createElement('th');
        thCol.textContent = 'Kolom ' + c.colLetter;
        trHead.appendChild(thCol);
      });
      this.theadPreview.appendChild(trHead);

      sampleRows.forEach(function (r) {
        var tr = document.createElement('tr');
        var tdRow = document.createElement('td');
        tdRow.className = 'center-col';
        tdRow.style.color = 'var(--text-muted)';
        tdRow.style.fontSize = '11px';
        tdRow.textContent = r.row1Based;
        tr.appendChild(tdRow);

        r.cells.forEach(function (c) {
          var td = document.createElement('td');
          if (c.formula) {
            var formulaSpan = document.createElement('span');
            formulaSpan.className = 'source-cell-tag';
            formulaSpan.style.marginRight = '4px';
            formulaSpan.textContent = '=' + c.formula;
            td.appendChild(formulaSpan);
          }

          var valText = '';
          if (c.val !== null && c.val !== undefined) {
            valText = (typeof c.val === 'number') ? formatNumber(c.val) : String(c.val);
          }
          var textNode = document.createTextNode(valText);
          td.appendChild(textNode);
          tr.appendChild(td);
        });

        this.tbodyPreview.appendChild(tr);
      }, this);
    },

    renderMappingInterface: function (detection) {
      var self = this;
      if (!this.mappingContainer || !currentMappingConfig) return;

      this.mappingContainer.style.display = 'block';
      if (this.mappingHeaderInfo) {
        var infoText = detection.found
          ? ('Header ' + (currentHeaderConfirmed ? 'dikonfirmasi' : 'terdeteksi') + ' di baris ke-' + detection.headerRow1Based + ' (Keyakinan: ' + detection.confidence + '%).')
          : 'Header belum terdeteksi otomatis dengan yakin. Pilih baris header atau petakan kolom secara manual.';
        if (detection.ambiguous && !currentHeaderConfirmed) {
          infoText += ' Beberapa kandidat baris header memiliki skor setara; konfirmasi baris yang benar.';
        }
        if (currentMappingConfig.unresolvedFields && currentMappingConfig.unresolvedFields.length) {
          infoText += ' Pilih kolom secara manual untuk bidang dengan beberapa kandidat: ' +
            currentMappingConfig.unresolvedFields.join(', ') + '.';
        }
        this.mappingHeaderInfo.textContent = infoText;
      }

      if (this.headerCorrectionControls && this.headerRowChoice) {
        this.headerCorrectionControls.style.display = currentHeaderNeedsConfirmation ? 'block' : 'none';
        this.headerRowChoice.innerHTML = '';
        if (currentHeaderNeedsConfirmation) {
          var sheet = currentParsedWb && currentParsedWb.sheets[currentSelectedSheet];
          var range = sheet && sheet['!ref'] ? window.XLSX.utils.decode_range(sheet['!ref']) : { s: { r: 0 }, e: { r: 0 } };
          var candidateByRow = {};
          (detection.candidates || []).forEach(function (candidate) {
            candidateByRow[candidate.row1Based] = candidate;
          });
          for (var row = range.s.r; row <= range.e.r; row++) {
            var row1 = row + 1;
            var candidate = candidateByRow[row1];
            var option = document.createElement('option');
            option.value = row1;
            option.textContent = 'Baris ' + row1 + (candidate ? ' — kandidat header (' + candidate.confidence + '%)' : '');
            option.selected = detection.found && detection.headerRow1Based === row1;
            this.headerRowChoice.appendChild(option);
          }
        }
      }

      this.tbodyMapping.innerHTML = '';
      var fieldsMeta = window.ColumnMapper.CANONICAL_FIELDS || (window.ColumnMapper.getCanonicalFields ? window.ColumnMapper.getCanonicalFields() : []);
      var headerCols = detection.columns || detection.rawHeaders || [];

      fieldsMeta.forEach(function (f) {
        var mapped = currentMappingConfig.fields[f.key] || {};
        var tr = document.createElement('tr');

        var td1 = document.createElement('td');
        var strongField = document.createElement('strong');
        strongField.textContent = f.label;
        td1.appendChild(strongField);
        tr.appendChild(td1);

        var td2 = document.createElement('td');
        var select = document.createElement('select');
        select.className = 'mapping-select';
        select.setAttribute('data-field', f.key);

        var optNone = document.createElement('option');
        optNone.value = '-1';
        optNone.textContent = '-- Tidak Dipetakan / Kosong --';
        select.appendChild(optNone);

        headerCols.forEach(function (c) {
          var opt = document.createElement('option');
          opt.value = c.colIndex;
          var isChosen = (mapped.colIndex === c.colIndex);
          if (isChosen) opt.selected = true;
          var headerText = c.headerText || c.text || '';
          var headerName = headerText ? (' - "' + headerText + '"') : '';
          opt.textContent = 'Kolom ' + c.colLetter + headerName;
          select.appendChild(opt);
        });
        td2.appendChild(select);
        tr.appendChild(td2);

        var td3 = document.createElement('td');
        var headerSpan = document.createElement('span');
        headerSpan.style.fontFamily = 'var(--font-mono)';
        headerSpan.style.fontSize = '12px';
        headerSpan.textContent = mapped.originalHeader || '-';
        td3.appendChild(headerSpan);
        tr.appendChild(td3);

        var td4 = document.createElement('td');
        var reqBadge = document.createElement('span');
        if (f.required) {
          reqBadge.style.color = 'var(--color-error)';
          reqBadge.style.fontWeight = '700';
          reqBadge.textContent = 'Wajib';
        } else {
          reqBadge.style.color = 'var(--text-muted)';
          reqBadge.textContent = 'Opsional';
        }
        td4.appendChild(reqBadge);
        tr.appendChild(td4);

        self.tbodyMapping.appendChild(tr);
      });

      var selects = this.tbodyMapping.querySelectorAll('.mapping-select');
      for (var s = 0; s < selects.length; s++) {
        selects[s].addEventListener('change', function () {
          var fKey = this.getAttribute('data-field');
          var chosenColIdx = parseInt(this.value, 10);
          var headerCell = (currentMappingConfig.rawHeaders || []).filter(function (header) {
            return header.colIndex === chosenColIdx;
          })[0];
          currentMappingConfig = window.ColumnMapper.updateMappingField(
            currentMappingConfig, fKey, chosenColIdx, headerCell ? headerCell.text : ''
          );
          self.invalidateAuditResults();
          self.renderMappingInterface(detection);
          self.updateWorkflowActions();
          self.persistProjectState();
        });
      }
      this.updateWorkflowActions();
    },

    runUniversalAudit: function () {
      if (auditRunInProgress) return false;
      if (currentAuditResult && currentCanonicalData) {
        if (this.mappingActionStatus) this.mappingActionStatus.textContent = 'Audit sudah dijalankan; hasil tersedia pada Ringkasan Proyek.';
        return false;
      }
      if (!currentParsedWb) {
        if (this.workbookConfirmationStatus) this.workbookConfirmationStatus.textContent = 'Unggah workbook terlebih dahulu.';
        return false;
      }

      auditAttempted = true;
      auditRunInProgress = true;
      this.updateWorkflowActions();
      try {
        currentCanonicalData = window.DataNormalizer.normalizeWorkbook(currentParsedWb, currentInspection);
        currentMaterialLogisticsResult = null;
        if (!currentCanonicalData.statistics.processedBoqSheets) {
          currentAuditResult = null;
          this.renderSheetsTable();
          var unresolvedSummary = (currentCanonicalData.unresolvedSheets || []).map(function (entry) {
            return entry.name + ': ' + entry.unresolvedReasons.join(' ');
          }).join(' | ');
          if (this.workbookConfirmationStatus) {
            this.workbookConfirmationStatus.textContent = 'Tidak ditemukan sheet BOQ yang cukup yakin untuk diaudit. Tidak ada total nol yang ditampilkan sebagai hasil audit.';
          }
          if (this.mappingActionStatus) this.mappingActionStatus.textContent =
            'Audit belum dijalankan. Alasan: ' + (unresolvedSummary || 'struktur BOQ tidak terdeteksi.');
          if (this.sheetSelectionNotice) this.sheetSelectionNotice.textContent =
            'Belum ada BOQ yang dapat diaudit. ' + (unresolvedSummary || 'Struktur worksheet tidak memenuhi pola BOQ yang didukung.');
          return false;
        }
        currentSelectedSheet = 'Seluruh Workbook';
        this.renderSheetsTable();
        var firstProcessed = currentCanonicalData.workbookManifest.find(function (entry) { return entry.status === 'processed'; });
        currentMappingConfig = firstProcessed ? firstProcessed.mapping : null;
        var valIngest = window.IngestionValidator.validate(currentCanonicalData);

        var hasManifestForPartialReport = currentCanonicalData.workbookManifest &&
          currentCanonicalData.workbookManifest.length > 0 && currentCanonicalData.unresolvedSheets &&
          currentCanonicalData.unresolvedSheets.length > 0;
        if (!valIngest.valid && !hasManifestForPartialReport) {
          currentCanonicalData = null;
          if (this.mappingActionStatus) this.mappingActionStatus.textContent = 'Normalisasi gagal: ' + valIngest.errors.join(' ');
          alert('Normalisasi Gagal:\n- ' + valIngest.errors.join('\n- '));
          return false;
        }

        if (!valIngest.valid && this.mappingActionStatus) {
          this.mappingActionStatus.textContent = 'Tidak ada item BOQ yang dapat diproses. Manifest dan alasan sheet unresolved tetap ditampilkan untuk ditinjau.';
        }

        currentAuditResult = window.UniversalChecker.runAudit(currentCanonicalData);
        this.runUnitConsistencyAnalysis();
        this.populateWorkbookFindingFilters();

        // Update Top Header Info
        var docTitle = isSyntheticDemo ? activeDemoTitle : (currentCanonicalData.project.projectName || currentParsedWb.filename);
        if (this.activeDocBadge) this.activeDocBadge.style.display = 'inline-flex';
        if (this.activeDocName) this.activeDocName.textContent = docTitle + ' (Seluruh Workbook)';
        if (this.btnSwitchDoc) this.btnSwitchDoc.style.display = 'inline-block';

        // Update Nav Badge
        if (this.badgeNavAudit) {
          this.badgeNavAudit.style.display = 'inline-block';
          this.badgeNavAudit.textContent = currentAuditResult.totalFindings;
        }

        // Render All Dashboard Panels
        this.renderRingkasanTab(valIngest);
        this.renderAuditReportSummary();
        this.renderPekerjaanTable();
        this.renderFindings();
        this.renderSourceViewer();
        var baselineForModules = window.UniversalSimulasi.recalculateBaseline(currentCanonicalData);
        var simulationAvailable = baselineForModules.calculableItemsCount > 0;
        if (simulationAvailable) this.runSimulation();
        else this.setSimulationUnavailable('Tidak ada item dengan volume dan harga satuan numerik yang valid.');
        this.renderFalsePositivesTable();

        // Unlock modules only after canonical data, audit state, and reports render successfully.
        this.setDashboardTabsEnabled(true, simulationAvailable,
          Array.isArray(currentCanonicalData.workItems) && currentCanonicalData.workItems.length > 0);

        if (this.mappingActionStatus) {
          this.mappingActionStatus.textContent = 'Audit selesai: ' + currentAuditResult.totalFindings +
            ' temuan. Ringkasan dan hasil audit siap ditinjau.';
        }
        if (this.workbookConfirmationStatus) this.workbookConfirmationStatus.textContent =
          'Audit workbook selesai. Ringkasan, Audit, dan modul yang memenuhi prasyarat sudah tersedia.';

        // Switch to Ringkasan Tab
        this.switchTab('tab-ringkasan');
        this.persistProjectState();
        return true;

      } catch (err) {
        currentCanonicalData = null;
        currentAuditResult = null;
        currentUnitConsistencyResult = null;
        currentMaterialLogisticsResult = null;
        this.clearUnitConsistencyResult();
        this.setDashboardTabsEnabled(false);
        if (this.navBtnIngestion && !this.navBtnIngestion.classList.contains('active')) {
          this.switchTab('tab-ingestion');
        }
        if (this.workbookConfirmationStatus) this.workbookConfirmationStatus.textContent = 'Audit gagal. Workbook tetap dimuat; periksa pesan kesalahan dan coba lagi.';
        if (this.mappingActionStatus) this.mappingActionStatus.textContent = 'Audit gagal: ' + err.message;
        alert('Terjadi kesalahan saat memproses audit: ' + err.message);
        console.error(err);
        return false;
      } finally {
        auditRunInProgress = false;
        this.updateWorkflowActions();
      }
    },

    clearUnitConsistencyResult: function () {
      if (this.unitConsistencyStatus) this.unitConsistencyStatus.textContent = 'Jalankan audit workbook untuk menganalisis satuan.';
      if (this.unitConsistencySummary) this.unitConsistencySummary.textContent = '';
      if (this.unitConsistencyWarnings) this.unitConsistencyWarnings.textContent = '';
      if (this.tbodyUnitConsistency) this.tbodyUnitConsistency.textContent = '';
    },

    runUnitConsistencyAnalysis: function () {
      if (!currentCanonicalData) {
        currentUnitConsistencyResult = null;
        this.clearUnitConsistencyResult();
        return null;
      }
      if (!window.AnalysisEngineAdapter) {
        currentUnitConsistencyResult = {
          engineName: 'UnitConsistencyEngine', engineVersion: '1.0.0',
          schemaVersion: 'rab-guard-engine-result/1', status: 'failed', summary: {},
          findings: [], metrics: {}, sourceReferences: [], warnings: [],
          errors: ['AnalysisEngineAdapter tidak tersedia.'], limitations: []
        };
      } else {
        try {
          currentUnitConsistencyResult = window.AnalysisEngineAdapter.run(
            window.UnitConsistencyEngine, currentCanonicalData
          );
        } catch (error) {
          currentUnitConsistencyResult = {
            engineName: 'UnitConsistencyEngine', engineVersion: '1.0.0',
            schemaVersion: 'rab-guard-engine-result/1', status: 'failed', summary: {},
            findings: [], metrics: {}, sourceReferences: [], warnings: [],
            errors: [error && error.message ? error.message : String(error)], limitations: []
          };
        }
      }
      try {
        this.renderUnitConsistency();
      } catch (renderError) {
        currentUnitConsistencyResult = {
          engineName: 'UnitConsistencyEngine', engineVersion: '1.0.0',
          schemaVersion: 'rab-guard-engine-result/1', status: 'failed', summary: {},
          findings: [], metrics: {}, sourceReferences: [], warnings: [],
          errors: ['Gagal menampilkan hasil pengecekan satuan: ' + (renderError && renderError.message ? renderError.message : String(renderError))],
          limitations: []
        };
        if (this.unitConsistencyStatus) this.unitConsistencyStatus.textContent = currentUnitConsistencyResult.errors[0];
      }
      return currentUnitConsistencyResult;
    },

    renderUnitConsistency: function () {
      var result = currentUnitConsistencyResult;
      if (!this.unitConsistencyStatus || !this.unitConsistencySummary ||
          !this.unitConsistencyWarnings || !this.tbodyUnitConsistency) return;

      this.unitConsistencyWarnings.textContent = '';
      this.tbodyUnitConsistency.textContent = '';
      if (!result) {
        this.unitConsistencyStatus.textContent = 'Hasil pengecekan satuan belum tersedia.';
        this.unitConsistencySummary.textContent = '';
        return;
      }

      var statusLabel = result.status === 'complete' ? 'Selesai' :
        (result.status === 'partial' ? 'Selesai sebagian' : 'Analisis gagal');
      this.unitConsistencyStatus.textContent = statusLabel + ' — ' + result.findings.length + ' temuan untuk ditinjau.';
      var summary = result.summary || {};
      this.unitConsistencySummary.textContent = (summary.itemsAnalyzed || 0) + ' item dianalisis; ' +
        (summary.recognizedUnits || 0) + ' satuan dikenali; ' + (summary.missingUnits || 0) +
        ' satuan kosong; ' + (summary.unrecognizedUnits || 0) + ' label belum dikenal; ' +
        (summary.sourceErrorUnits || 0) + ' sel error sumber.';

      (result.warnings || []).concat(result.errors || []).forEach(function (message) {
        var li = document.createElement('li');
        li.textContent = message;
        this.unitConsistencyWarnings.appendChild(li);
      }, this);

      var self = this;
      function addTextCell(row, text) {
        var cell = document.createElement('td');
        cell.textContent = text === null || text === undefined ? '—' : String(text);
        row.appendChild(cell);
        return cell;
      }
      (result.findings || []).forEach(function (finding) {
        var tr = document.createElement('tr');
        addTextCell(tr, finding.severity || 'INFO');
        var findingCell = addTextCell(tr, '');
        var title = document.createElement('strong');
        title.textContent = finding.title || finding.ruleId || 'Temuan satuan';
        findingCell.appendChild(title);
        var explanation = document.createElement('div');
        explanation.textContent = finding.explanation || '';
        findingCell.appendChild(explanation);
        var contextCell = addTextCell(tr, '');
        (finding.relatedItems || []).forEach(function (item) {
          var itemLine = document.createElement('p');
          var path = Array.isArray(item.categoryPath) ? item.categoryPath.slice().sort(function (a, b) {
            return Number(a.levelIndex) - Number(b.levelIndex);
          }) : [];
          var pathParts = [];
          if (path.length) {
            var maxLevel = path.reduce(function (max, level) { return Math.max(max, Number(level.levelIndex) || 0); }, 0);
            for (var levelIndex = 0; levelIndex <= maxLevel; levelIndex++) {
              var level = path.find(function (candidate) { return Number(candidate.levelIndex) === levelIndex; });
              pathParts.push(level && level.value ? String(level.value) : '[tingkat ' + (levelIndex + 1) + ' tidak tersedia]');
            }
          }
          itemLine.textContent = (item.code ? item.code + ' — ' : '') + (item.description || 'Uraian tidak tersedia') +
            ' | ' + (item.unit || 'satuan kosong') + (item.dimension ? ' (' + item.dimension + ')' : '') +
            ' | hierarki ' + (item.hierarchyStatus || 'none') + (pathParts.length ? ' | ' + pathParts.join(' > ') : ' | kategori tidak tersedia');
          contextCell.appendChild(itemLine);
        });
        addTextCell(tr, (finding.sourceReferences || []).map(function (reference) {
          return (reference.cell || reference.sheet || 'lokasi tidak tersedia') +
            (reference.row ? ' (baris ' + reference.row + ')' : '');
        }).join('; '));
        addTextCell(tr, finding.recommendedAction || 'Tinjau sumber satuan.');
        self.tbodyUnitConsistency.appendChild(tr);
      });
      if (!result.findings.length) {
        var emptyRow = document.createElement('tr');
        var emptyCell = document.createElement('td');
        emptyCell.colSpan = 5;
        emptyCell.textContent = result.status === 'failed'
          ? 'Engine tidak menghasilkan analisis yang dapat digunakan.'
          : 'Tidak ada variasi satuan yang perlu ditinjau berdasarkan pemeriksaan yang tersedia.';
        emptyRow.appendChild(emptyCell);
        this.tbodyUnitConsistency.appendChild(emptyRow);
      }
    },

    renderRingkasanTab: function (valIngest) {
      if (!currentCanonicalData || !currentAuditResult) return;

      var baseCalc = window.UniversalSimulasi.recalculateBaseline(currentCanonicalData);
      var items = currentCanonicalData.workItems || [];
      var subtotals = currentCanonicalData.subtotals || [];

      if (this.metaSummaryProject) {
        this.metaSummaryProject.textContent = isSyntheticDemo ? activeDemoTitle : (currentCanonicalData.project.projectName || currentParsedWb.filename);
      }
      if (this.metaSummarySheet) {
        this.metaSummarySheet.textContent = currentSelectedSheet + ' (Baris Header ke-' + currentMappingConfig.headerRow1Based + ')';
      }
      if (this.metaSummaryDate) {
        this.metaSummaryDate.textContent = new Date().toLocaleString('id-ID');
      }
      if (this.metaSummaryScore) {
        var score = valIngest ? valIngest.stats.completenessScore : 100;
        this.metaSummaryScore.textContent = score + ' / 100';
      }

      if (this.statKpiTotal) {
        if (baseCalc && baseCalc.isValid && typeof baseCalc.totalRecalculated === 'number') {
          this.statKpiTotal.textContent = formatRupiah(baseCalc.totalRecalculated);
        } else {
          this.statKpiTotal.textContent = '-';
        }
      }

      var subCard = (this.statKpiTotal && this.statKpiTotal.parentElement) ? this.statKpiTotal.parentElement.querySelector('.card-subtitle') : null;
      if (subCard) {
        if (!baseCalc || !baseCalc.isValid) {
          subCard.textContent = 'Kalkulasi tidak tersedia';
        } else if (!baseCalc.isComplete) {
          subCard.textContent = 'Hitung ulang: ' + baseCalc.calculableItemsCount + '/' + baseCalc.itemsCount +
            ' item; sumber: ' + formatRupiah(baseCalc.totalImported) + '; selisih total (hitung ulang − sumber): ' +
            (baseCalc.difference === null ? 'tidak tersedia' : formatRupiah(baseCalc.difference));
        } else if (baseCalc.difference === null) {
          subCard.textContent = 'Sumber: ' + formatRupiah(baseCalc.totalImported) + '; pembanding tersedia untuk ' + baseCalc.comparableItemsCount + '/' + baseCalc.itemsCount + ' item';
        } else if (Math.abs(baseCalc.difference) > 0.01) {
          subCard.textContent = 'Sumber: ' + formatRupiah(baseCalc.totalImported) +
            ' (Selisih total, hitung ulang − sumber: ' + formatRupiah(baseCalc.difference) + ')';
        } else {
          subCard.textContent = 'Sumber: ' + formatRupiah(baseCalc.totalImported) + ' (selisih total: Rp0)';
        }
      }

      if (this.statKpiItems) this.statKpiItems.textContent = items.length + ' Item (' + subtotals.length + ' Subtotal)';
      if (this.statKpiFindings) {
        this.statKpiFindings.textContent = currentAuditResult.totalFindings + ' Temuan (' + currentAuditResult.summary.error + ' Error, ' + currentAuditResult.summary.warning + ' Warning)';
      }
      if (this.statKpiIntegrity) {
        if (currentAuditResult.summary.error > 0) {
          this.statKpiIntegrity.textContent = 'DEVIASI DITEMUKAN';
          this.statKpiIntegrity.style.color = 'var(--color-error)';
        } else if (currentAuditResult.summary.warning > 0) {
          this.statKpiIntegrity.textContent = 'PERLU DITINJAU';
          this.statKpiIntegrity.style.color = 'var(--color-tinjau)';
        } else {
          this.statKpiIntegrity.textContent = 'INTEGRITAS BAIK';
          this.statKpiIntegrity.style.color = 'var(--color-success)';
        }
      }

      // Render Subtotals Breakdown if present
      if (this.tbodySubtotalSummary && this.containerSubtotals) {
        this.tbodySubtotalSummary.innerHTML = '';
        if (subtotals.length > 0) {
          this.containerSubtotals.style.display = 'block';
          subtotals.forEach(function (sub) {
            var tr = document.createElement('tr');
            var tdDesc = document.createElement('td');
            var strongDesc = document.createElement('strong');
            strongDesc.textContent = sub.description || 'Subtotal';
            tdDesc.appendChild(strongDesc);
            tr.appendChild(tdDesc);

            var tdCell = document.createElement('td');
            tdCell.className = 'center-col';
            var cellTag = document.createElement('span');
            cellTag.className = 'source-cell-tag';
            cellTag.textContent = sub.source ? (sub.source.amountCell || (sub.source.sheet + '!Baris' + sub.source.row)) : '-';
            tdCell.appendChild(cellTag);
            tr.appendChild(tdCell);

            var tdVal = document.createElement('td');
            tdVal.className = 'num-col';
            var strongVal = document.createElement('strong');
            strongVal.textContent = formatRupiah(sub.amount);
            tdVal.appendChild(strongVal);
            tr.appendChild(tdVal);

            this.tbodySubtotalSummary.appendChild(tr);
          }, this);
        } else {
          this.containerSubtotals.style.display = 'none';
        }
      }
    },

    renderPekerjaanTable: function () {
      if (!currentCanonicalData || !this.tbodyPekerjaan) return;

      var items = currentCanonicalData.workItems || [];
      var query = (this.searchPekerjaan ? this.searchPekerjaan.value : '').toLowerCase().trim();

      var filtered = items.filter(function (it) {
        if (!query) return true;
        var searchStr = ((it.itemNumber || '') + ' ' + (it.code || '') + ' ' + (it.analysisCode || '') + ' ' + (it.description || '') + ' ' + (it.unit || '')).toLowerCase();
        return searchStr.indexOf(query) !== -1;
      });

      this.tbodyPekerjaan.innerHTML = '';

      if (filtered.length === 0) {
        if (this.footTotalImported) this.footTotalImported.textContent = formatRupiah(0);
        if (this.footTotalPekerjaan) this.footTotalPekerjaan.textContent = formatRupiah(0);
        var emptyTr = document.createElement('tr');
        var emptyTd = document.createElement('td');
        emptyTd.colSpan = 8;
        emptyTd.className = 'center-col';
        emptyTd.style.padding = '24px';
        emptyTd.textContent = 'Tidak ada item pekerjaan yang cocok dengan pencarian.';
        emptyTr.appendChild(emptyTd);
        this.tbodyPekerjaan.appendChild(emptyTr);
        return;
      }

      var importedGrandSum = 0;
      var recalculatedGrandSum = 0;

      filtered.forEach(function (it, idx) {
        var tr = document.createElement('tr');
        var hasImportedAmount = typeof it.amount === 'number' && isFinite(it.amount);
        var isCalculable = typeof it.volume === 'number' && isFinite(it.volume) &&
          typeof it.unitPrice === 'number' && isFinite(it.unitPrice);
        var recalculatedAmount = isCalculable ? it.volume * it.unitPrice : null;
        if (hasImportedAmount) importedGrandSum += it.amount;
        if (isCalculable) recalculatedGrandSum += recalculatedAmount;

        // 1. No
        var tdNo = document.createElement('td');
        tdNo.className = 'center-col';
        tdNo.textContent = it.itemNumber || (idx + 1);
        tr.appendChild(tdNo);

        // 2. Kode
        var tdCode = document.createElement('td');
        tdCode.className = 'center-col';
        var codeTag = document.createElement('code');
        codeTag.textContent = it.code || it.analysisCode || '-';
        tdCode.appendChild(codeTag);
        tr.appendChild(tdCode);

        // 3. Uraian
        var tdDesc = document.createElement('td');
        var strongDesc = document.createElement('strong');
        strongDesc.textContent = it.description || '-';
        tdDesc.appendChild(strongDesc);
        tr.appendChild(tdDesc);

        // 4. Volume
        var tdVol = document.createElement('td');
        tdVol.className = 'num-col';
        tdVol.textContent = it.volume !== null ? formatNumber(it.volume) : '-';
        tr.appendChild(tdVol);

        // 5. Satuan
        var tdUnit = document.createElement('td');
        tdUnit.className = 'center-col';
        tdUnit.textContent = it.unit || '-';
        tr.appendChild(tdUnit);

        // 6. Harga Satuan
        var tdPrice = document.createElement('td');
        tdPrice.className = 'num-col';
        tdPrice.textContent = it.unitPrice !== null ? formatRupiah(it.unitPrice) : '-';
        tr.appendChild(tdPrice);

        // 7. Jumlah Biaya
        var tdAmt = document.createElement('td');
        tdAmt.className = 'num-col';
        var strongAmt = document.createElement('strong');
        strongAmt.textContent = 'Sumber: ' + (hasImportedAmount ? formatRupiah(it.amount) : '—') +
          ' | Hitung ulang: ' + (isCalculable ? formatRupiah(recalculatedAmount) : '—');
        tdAmt.appendChild(strongAmt);
        tr.appendChild(tdAmt);

        // 8. Sel Rujukan / Formula
        var tdCell = document.createElement('td');
        tdCell.className = 'center-col';
        var cellTag = document.createElement('span');
        cellTag.className = 'source-cell-tag';
        var sInfo = it.source || {};
        var refText = sInfo.amountCell || ('B' + sInfo.row);
        if (sInfo.sheet && refText.indexOf('!') === -1) refText = sInfo.sheet + '!' + refText;
        if (sInfo.amountFormula) refText += ' (=' + sInfo.amountFormula + ')';
        cellTag.textContent = refText;
        tdCell.appendChild(cellTag);
        tr.appendChild(tdCell);

        this.tbodyPekerjaan.appendChild(tr);
      }, this);

      if (this.footTotalPekerjaan) {
        this.footTotalPekerjaan.textContent = formatRupiah(recalculatedGrandSum);
      }
      if (this.footTotalImported) this.footTotalImported.textContent = formatRupiah(importedGrandSum);
    },

    formatRelatedCategoryPath: function (item) {
      var path = Array.isArray(item && item.categoryPath) ? item.categoryPath.slice() : [];
      path.sort(function (a, b) { return Number(a.levelIndex) - Number(b.levelIndex); });
      if (!path.length) return 'Tidak ada kategori sumber';
      var byIndex = {};
      path.forEach(function (level) { byIndex[level.levelIndex] = level; });
      var segments = [];
      for (var levelIndex = 0; levelIndex <= path[path.length - 1].levelIndex; levelIndex++) {
        var level = byIndex[levelIndex];
        if (!level) segments.push('Level ' + (levelIndex + 1) + ' belum tersedia');
        else if (level.value === null || level.value === undefined || level.value === '') segments.push('Level ' + (levelIndex + 1) + ' kosong');
        else segments.push(String(level.value));
      }
      return segments.join(' > ');
    },

    appendRelatedItemsDetails: function (container, finding) {
      var relatedItems = finding && Array.isArray(finding.relatedItems) ? finding.relatedItems : [];
      if (!relatedItems.length) return false;
      var group = document.createElement('details');
      group.style.marginTop = '8px';
      var summary = document.createElement('summary');
      summary.textContent = 'Item terkait (' + relatedItems.length + ')';
      group.appendChild(summary);
      var list = document.createElement('ol');
      list.style.paddingLeft = '22px';
      list.style.maxHeight = '320px';
      list.style.overflowY = 'auto';
      list.style.overflowWrap = 'anywhere';
      relatedItems.forEach(function (item) {
        item = item || {};
        var entry = document.createElement('li');
        entry.style.marginTop = '8px';
        entry.style.overflowWrap = 'anywhere';
        function appendLine(label, value) {
          var line = document.createElement('div');
          var labelNode = document.createElement('strong');
          labelNode.textContent = label + ': ';
          line.appendChild(labelNode);
          var valueNode = document.createElement('span');
          valueNode.textContent = value === null || value === undefined || value === '' ? 'Tidak tersedia' : String(value);
          line.appendChild(valueNode);
          entry.appendChild(line);
        }
        appendLine('Kode item', item.itemCode);
        appendLine('Uraian', item.description);
        appendLine('Lokasi sumber', [item.sheet || 'Sheet tidak tersedia', item.sourceRow ? 'baris ' + item.sourceRow : null, item.sourceCell || null].filter(Boolean).join(' · '));
        appendLine('Category path', this.formatRelatedCategoryPath(item));
        appendLine('Status hierarchy', item.hierarchyStatus || (item.categoryPath && item.categoryPath.length ? 'tidak tersedia' : 'none'));
        var provenance = Array.isArray(item.categoryProvenance) ? item.categoryProvenance :
          (Array.isArray(item.categoryPath) ? item.categoryPath.map(function (level) { return { levelIndex: level.levelIndex, provenance: level.provenance || null }; }) : []);
        var provenanceDetails = document.createElement('details');
        var provenanceSummary = document.createElement('summary');
        provenanceSummary.textContent = 'Provenance kategori';
        provenanceDetails.appendChild(provenanceSummary);
        var provenanceText = document.createElement('pre');
        provenanceText.style.whiteSpace = 'pre-wrap';
        provenanceText.style.overflowWrap = 'anywhere';
        provenanceText.textContent = provenance.length ? JSON.stringify(provenance, null, 2) : 'Tidak ada provenance kategori sumber.';
        provenanceDetails.appendChild(provenanceText);
        entry.appendChild(provenanceDetails);
        list.appendChild(entry);
      }, this);
      group.appendChild(list);
      container.appendChild(group);
      return true;
    },

    createFindingRow: function (f) {
      var self = this;
      var tr = document.createElement('tr');
      tr.style.cursor = 'pointer';
      tr.title = 'Klik untuk melihat lokasi sumber di spreadsheet viewer';
      tr.addEventListener('click', function (e) {
        if (e.target.tagName === 'BUTTON' || e.target.closest('button') || e.target.tagName === 'SUMMARY' || e.target.closest('summary')) return;
        var prev = self.tbodyFindings.querySelectorAll('tr.finding-row-selected');
        prev.forEach(function (r) { r.classList.remove('finding-row-selected'); });
        tr.classList.add('finding-row-selected');
        self.focusFindingInSourceViewer(f);
      });

      var isFP = window.UserFeedback && window.UserFeedback.isFalsePositive(f.id);
      if (isFP) {
        tr.style.backgroundColor = '#fefce8';
      }

      var badgeClass = 'badge-info';
      if (f.severity === 'ERROR') badgeClass = 'badge-fail';
      else if (f.severity === 'WARNING') badgeClass = 'badge-tinjau';

      var origStr = (f.originalValue !== null && f.originalValue !== undefined)
        ? (typeof f.originalValue === 'number' ? formatNumber(f.originalValue) : String(f.originalValue))
        : '-';
      var expStr = (f.expectedValue !== null && f.expectedValue !== undefined)
        ? (typeof f.expectedValue === 'number' ? formatNumber(f.expectedValue) : String(f.expectedValue))
        : '-';

      // 1. ID
      var tdId = document.createElement('td');
      var strongId = document.createElement('strong');
      strongId.textContent = f.id || '-';
      tdId.appendChild(strongId);
      if (isFP) {
        var fpBadge = document.createElement('span');
        fpBadge.className = 'badge-fp';
        fpBadge.style.display = 'block';
        fpBadge.style.marginTop = '4px';
        fpBadge.textContent = 'False Positive';
        tdId.appendChild(fpBadge);
      }
      tr.appendChild(tdId);

      // 2. Severity Badge
      var tdSev = document.createElement('td');
      tdSev.className = 'center-col';
      var badge = document.createElement('span');
      badge.className = 'test-badge ' + badgeClass;
      badge.textContent = f.severity || '-';
      tdSev.appendChild(badge);
      tr.appendChild(tdSev);

      // 3. Rule ID & Certainty Badge
      var tdRule = document.createElement('td');
      var codeRule = document.createElement('code');
      codeRule.style.fontSize = '12px';
      codeRule.textContent = f.ruleId || '-';
      tdRule.appendChild(codeRule);

      if (f.certaintyLabel) {
        var cBadge = document.createElement('span');
        var cClass = 'certainty-context';
        if (f.certainty === 'MATHEMATICAL_PROOF') cClass = 'certainty-math';
        else if (f.certainty === 'PATTERN_ANOMALY') cClass = 'certainty-pattern';
        cBadge.className = 'badge-certainty ' + cClass;
        cBadge.style.display = 'block';
        cBadge.style.marginTop = '4px';
        cBadge.textContent = f.certaintyLabel;
        tdRule.appendChild(cBadge);
      }
      tr.appendChild(tdRule);

      // 4. Explanation
      var tdExp = document.createElement('td');
      var presentation = this.getFindingPresentation(f);
      var findingTitle = document.createElement('strong');
      findingTitle.textContent = presentation.title;
      tdExp.appendChild(findingTitle);
      var categoryText = document.createElement('small');
      categoryText.textContent = presentation.category;
      categoryText.style.display = 'block';
      tdExp.appendChild(categoryText);
      var explanationText = document.createElement('div');
      explanationText.textContent = presentation.explanation;
      tdExp.appendChild(explanationText);
      var itemContext = document.createElement('small');
      itemContext.textContent = 'Item: ' + presentation.itemDescription;
      itemContext.style.display = 'block';
      itemContext.style.marginTop = '4px';
      tdExp.appendChild(itemContext);
      var impactText = document.createElement('small');
      impactText.style.display = 'block';
      impactText.style.marginTop = '4px';
      impactText.textContent = presentation.financialImpact;
      tdExp.appendChild(impactText);
      tr.appendChild(tdExp);

      // 5. Original Value
      var tdOrig = document.createElement('td');
      tdOrig.className = 'num-col';
      tdOrig.textContent = origStr;
      tr.appendChild(tdOrig);

      // 6. Expected Value
      var tdExpected = document.createElement('td');
      tdExpected.className = 'num-col';
      tdExpected.textContent = expStr;
      tr.appendChild(tdExpected);

      // 7. Source Location (Provenance)
      var tdCell = document.createElement('td');
      tdCell.className = 'center-col';
      var locInfo = this.formatFindingLocation(f);
      f.formattedLocation = locInfo.primary;

      if (locInfo.isMissing) {
        var missingSpan = document.createElement('span');
        missingSpan.className = 'source-cell-missing';
        missingSpan.textContent = 'Lokasi sumber tidak tersedia';
        tdCell.appendChild(missingSpan);
      } else {
        var cellTag = document.createElement('button');
        cellTag.type = 'button';
        cellTag.className = 'source-cell-tag' + (f.severity === 'ERROR' ? ' has-error' : ' has-warning');
        cellTag.title = 'Klik untuk membuka ' + locInfo.primary + ' di Spreadsheet Viewer';
        cellTag.textContent = locInfo.primary;

        cellTag.addEventListener('click', function (e) {
          e.stopPropagation();
          var prev = self.tbodyFindings.querySelectorAll('tr.finding-row-selected');
          prev.forEach(function (r) { r.classList.remove('finding-row-selected'); });
          tr.classList.add('finding-row-selected');
          self.focusFindingInSourceViewer(f, locInfo.primary);
          // Also populate the search box with the target address
          var searchInput = document.getElementById('sv-search-input');
          if (searchInput) searchInput.value = locInfo.primary;
        });
        tdCell.appendChild(cellTag);

        if (locInfo.additionalCount > 0) {
          var multiPill = document.createElement('span');
          multiPill.className = 'source-cell-multi-pill';
          multiPill.title = 'Total ' + locInfo.all.length + ' lokasi: ' + locInfo.all.join(', ');
          multiPill.textContent = '+' + locInfo.additionalCount + ' lokasi';
          multiPill.addEventListener('click', function (e) {
            e.stopPropagation();
            var prev = self.tbodyFindings.querySelectorAll('tr.finding-row-selected');
            prev.forEach(function (r) { r.classList.remove('finding-row-selected'); });
            tr.classList.add('finding-row-selected');
            self.focusFindingInSourceViewer(f, locInfo.primary);
          });
          tdCell.appendChild(multiPill);
        }
      }
      tr.appendChild(tdCell);

      // 8. Recommended Action & FP Button
      var tdAction = document.createElement('td');
      tdAction.style.fontSize = '12px';
      tdAction.style.color = 'var(--text-muted)';
      var actionText = document.createElement('div');
      actionText.textContent = f.recommendedAction || '-';
      tdAction.appendChild(actionText);

      var detail = document.createElement('details');
      detail.style.marginTop = '6px';
      var detailSummary = document.createElement('summary');
      detailSummary.textContent = 'Rincian bukti';
      detail.appendChild(detailSummary);
      function appendDetail(label, value) {
        var entry = document.createElement('div');
        entry.style.marginTop = '4px';
        var labelNode = document.createElement('strong');
        labelNode.textContent = label + ': ';
        entry.appendChild(labelNode);
        var valueNode = document.createElement('span');
        valueNode.textContent = value === null || value === undefined || value === '' ? 'Tidak tersedia' : String(value);
        entry.appendChild(valueNode);
        detail.appendChild(entry);
      }
      appendDetail('Penjelasan lengkap', f.explanation);
      appendDetail('Nilai sumber', origStr);
      appendDetail('Nilai hitung / ekspektasi', expStr);
      appendDetail('Selisih (sesuai aturan)', f.difference);
      if (f.formula) appendDetail('Formula Excel', f.formula);
      if (f.errorCode !== null && f.errorCode !== undefined) appendDetail('Kode error Excel', f.errorCode);
      if (f.errorText) appendDetail('Teks error Excel', f.errorText);
      appendDetail('Saran tindakan', f.recommendedAction);
      this.appendRelatedItemsDetails(detail, f);
      tdAction.appendChild(detail);

      var btnFp = document.createElement('button');
      btnFp.type = 'button';
      btnFp.className = 'sample-btn';
      btnFp.style.marginTop = '6px';
      btnFp.style.fontSize = '11px';
      btnFp.textContent = isFP ? 'Batal False Positive' : '⚠️ Tandai False Positive';
      btnFp.addEventListener('click', function () {
        if (isFP) {
          window.UserFeedback.unmarkFalsePositive(f.id);
          self.renderFindings();
          self.renderFalsePositivesTable();
        } else {
          var reason = prompt('Masukkan alasan penandaan false positive untuk temuan ' + f.id + ' (misal: "Disengaja dalam kontrak" / "Pembulatan wajar"):');
          if (reason !== null && reason.trim()) {
            window.UserFeedback.markFalsePositive(f.id, f.ruleId, f.explanation, reason);
            self.renderFindings();
            self.renderFalsePositivesTable();
          }
        }
      });
      tdAction.appendChild(btnFp);

      tr.appendChild(tdAction);

      return tr;
    },

    buildAuditReportSummary: function (canonicalData, auditResult) {
      var baseline = canonicalData && window.UniversalSimulasi
        ? window.UniversalSimulasi.recalculateBaseline(canonicalData) : null;
      var findings = auditResult && Array.isArray(auditResult.findings) ? auditResult.findings : [];
      var counts = { ERROR: 0, WARNING: 0, INFO: 0 };
      findings.forEach(function (finding) {
        var severity = finding && String(finding.severity || '').toUpperCase();
        if (Object.prototype.hasOwnProperty.call(counts, severity)) counts[severity]++;
      });
      var itemsCount = baseline ? baseline.itemsCount : 0;
      var importedCount = baseline ? baseline.importedAmountItemsCount : 0;
      var calculableCount = baseline ? baseline.calculableItemsCount : 0;
      var comparableCount = baseline ? baseline.comparableItemsCount : 0;
      return {
        workbook: canonicalData && canonicalData.project ? canonicalData.project.sourceFilename || '' : '',
        worksheet: canonicalData && canonicalData.project ? canonicalData.project.activeSheet || '' : '',
        itemsCount: itemsCount,
        coverageText: baseline
          ? calculableCount + '/' + itemsCount + ' dapat dihitung; ' + baseline.uncalculableItemsCount + ' belum lengkap; nilai sumber tersedia untuk ' + importedCount + '/' + itemsCount + ' item'
          : 'Cakupan tidak tersedia',
        importedAmount: baseline && importedCount > 0 ? formatRupiah(baseline.totalImported) : 'Tidak tersedia',
        recalculatedAmount: baseline && calculableCount > 0 ? formatRupiah(baseline.totalRecalculated) : 'Tidak tersedia',
        difference: baseline && comparableCount > 0
          ? formatRupiah(baseline.comparableDifference) + ' (hitung ulang − sumber; ' + comparableCount + ' item pembanding)'
          : 'Tidak tersedia',
        totalFindings: findings.length,
        severityCounts: counts,
        baseline: baseline
      };
    },

    getFilteredFindings: function (findings, severity, query, sortBy, filters) {
      var all = Array.isArray(findings) ? findings : [];
      filters = filters || {};
      var normalizedSeverity = String(severity || 'all').toUpperCase();
      var normalizedQuery = String(query || '').toLowerCase().trim();
      var canonicalItems = filters.canonicalData && Array.isArray(filters.canonicalData.workItems)
        ? filters.canonicalData.workItems : [];
      var canonicalItemBySource = null;
      if (filters.division && filters.division !== 'all') {
        canonicalItemBySource = new Map();
        canonicalItems.forEach(function (candidate) {
          if (!candidate || !candidate.source) return;
          var sheet = candidate.source.sheet;
          var byRow = canonicalItemBySource.get(sheet);
          if (!byRow) { byRow = new Map(); canonicalItemBySource.set(sheet, byRow); }
          // Keep the first matching record, as Array.find did before indexing.
          if (!byRow.has(candidate.source.row)) byRow.set(candidate.source.row, candidate);
        });
      }
      var filtered = all.filter(function (finding) {
        finding = finding || {};
        if (normalizedSeverity !== 'ALL' && String(finding.severity || '').toUpperCase() !== normalizedSeverity) return false;
        if (filters.sheet && filters.sheet !== 'all' && String(finding.sourceSheet || '') !== filters.sheet) return false;
        if (filters.division && filters.division !== 'all') {
          var byRow = canonicalItemBySource.get(finding.sourceSheet);
          var item = byRow && byRow.get(finding.sourceRow);
          var division = getDivisionFilterValue(item);
          if (division !== filters.division) return false;
        }
        if (!normalizedQuery) return true;
        var searchable = [finding.itemDescription, finding.ruleId, finding.sourceCell,
          finding.sourceSheet, finding.title, finding.explanation].join(' ').toLowerCase();
        return searchable.indexOf(normalizedQuery) !== -1;
      });
      if (sortBy === 'severity') {
        var rank = { ERROR: 0, WARNING: 1, INFO: 2 };
        filtered.sort(function (a, b) {
          var ar = Object.prototype.hasOwnProperty.call(rank, String(a.severity || '').toUpperCase()) ? rank[String(a.severity).toUpperCase()] : 3;
          var br = Object.prototype.hasOwnProperty.call(rank, String(b.severity || '').toUpperCase()) ? rank[String(b.severity).toUpperCase()] : 3;
          return ar - br;
        });
      } else if (sortBy === 'source') {
        filtered.sort(function (a, b) {
          var aLoc = String(a.sourceSheet || '') + '!' + String(a.sourceCell || '');
          var bLoc = String(b.sourceSheet || '') + '!' + String(b.sourceCell || '');
          return aLoc.localeCompare(bLoc);
        });
      }
      return { findings: filtered, totalCount: all.length, filteredCount: filtered.length };
    },

    getFindingPresentation: function (finding) {
      var f = finding || {};
      var isArithmeticDifference = f.ruleId === 'amount_differs_qty_price' &&
        typeof f.difference === 'number' && isFinite(f.difference);
      var category = 'Perlu tinjauan; dampak belum dipastikan';
      if (f.ruleId === 'amount_differs_qty_price') category = 'Selisih aritmetika terukur';
      else if (/^missing_/.test(f.ruleId || '')) category = 'Data sumber belum lengkap';
      else if (/^(?:excel_error_|broken_formula_reference|invalid_numeric_values)/.test(f.ruleId || '')) category = 'Error nilai atau formula sumber';
      else if (String(f.severity || '').toUpperCase() === 'INFO') category = 'Informasi';
      return {
        id: f.id || 'ID tidak tersedia',
        severity: f.severity || 'Tingkat tidak tersedia',
        ruleId: f.ruleId || 'Aturan tidak tersedia',
        title: f.title || f.ruleId || 'Temuan audit',
        itemDescription: f.itemDescription || 'Uraian item tidak tersedia',
        sourceSheet: f.sourceSheet || 'Lembar tidak tersedia',
        sourceCell: f.sourceCell || 'Sel tidak tersedia',
        explanation: f.explanation || 'Penjelasan tidak tersedia',
        originalValue: f.originalValue,
        expectedValue: f.expectedValue,
        difference: f.difference,
        financialImpact: isArithmeticDifference
          ? 'Selisih item (sumber − hitung ulang): ' + formatRupiah(f.difference)
          : 'Dampak finansial tidak dihitung dari temuan ini',
        category: category,
        recommendedAction: f.recommendedAction || 'Saran tindakan tidak tersedia',
        formula: f.formula || null,
        errorCode: f.errorCode,
        errorText: f.errorText || null
      };
    },

    getFindingsEmptyMessage: function (totalCount, filteredCount) {
      return totalCount === 0
        ? 'Tidak ada temuan audit untuk workbook ini.'
        : (filteredCount === 0 ? 'Tidak ada temuan yang sesuai dengan filter.' : '');
    },

    renderAuditReportSummary: function () {
      if (!this.auditReportSummary || !currentCanonicalData) return;
      var summary = this.buildAuditReportSummary(currentCanonicalData, currentAuditResult);
      var values = {
        'univ-report-filename': summary.workbook || 'Tidak tersedia',
        'univ-report-sheet': summary.worksheet || 'Tidak tersedia',
        'univ-report-items': String(summary.itemsCount),
        'univ-report-coverage': summary.coverageText,
        'univ-report-imported': summary.importedAmount,
        'univ-report-recalculated': summary.recalculatedAmount,
        'univ-report-difference': summary.difference
      };
      Object.keys(values).forEach(function (id) {
        var node = document.getElementById(id);
        if (node) node.textContent = values[id];
      });
      var stats = currentCanonicalData.statistics || {};
      var coverage = document.getElementById('univ-report-workbook-coverage');
      if (coverage) coverage.textContent = stats.processedBoqSheets + ' sheet BOQ diproses; ' +
        stats.supportingSheets + ' sheet pendukung; ' + stats.unresolvedSheets + ' sheet belum terselesaikan dari ' + stats.totalSheets + ' sheet diperiksa.';
      var unresolvedList = document.getElementById('univ-report-unresolved-sheets');
      if (unresolvedList) {
        unresolvedList.innerHTML = '';
        (currentCanonicalData.unresolvedSheets || []).forEach(function (entry) {
          var li = document.createElement('li');
          li.textContent = entry.name + ': ' + entry.unresolvedReasons.join(' ');
          unresolvedList.appendChild(li);
        });
      }
    },

    renderFindings: function () {
      if (!currentAuditResult || !this.tbodyFindings) return;

      var findings = currentAuditResult.findings || [];
      var query = (this.searchAudit ? this.searchAudit.value : '').toLowerCase().trim();
      var sortBy = this.auditSort ? this.auditSort.value : 'source';
      var result = this.getFilteredFindings(findings, currentFilter, query, sortBy, {
        sheet: this.auditSheetFilter ? this.auditSheetFilter.value : 'all',
        division: this.auditDivisionFilter ? this.auditDivisionFilter.value : 'all',
        canonicalData: currentCanonicalData
      });

      var summary = this.buildAuditReportSummary(currentCanonicalData, currentAuditResult);
      if (this.statAuditTotal) this.statAuditTotal.textContent = summary.totalFindings;
      if (this.statAuditError) this.statAuditError.textContent = summary.severityCounts.ERROR;
      if (this.statAuditWarning) this.statAuditWarning.textContent = summary.severityCounts.WARNING;
      if (this.statAuditInfo) this.statAuditInfo.textContent = summary.severityCounts.INFO;

      if (this.pillCountAll) this.pillCountAll.textContent = summary.totalFindings;
      if (this.pillCountError) this.pillCountError.textContent = summary.severityCounts.ERROR;
      if (this.pillCountWarning) this.pillCountWarning.textContent = summary.severityCounts.WARNING;
      if (this.pillCountInfo) this.pillCountInfo.textContent = summary.severityCounts.INFO;
      if (this.auditFilteredCount) this.auditFilteredCount.textContent = result.filteredCount + ' / ' + result.totalCount + ' temuan';

      this.tbodyFindings.innerHTML = '';

      if (result.filteredCount === 0) {
        var emptyTr = document.createElement('tr');
        var emptyTd = document.createElement('td');
        emptyTd.colSpan = 8;
        emptyTd.className = 'center-col';
        emptyTd.style.padding = '24px';
        emptyTd.style.color = 'var(--color-success)';
        emptyTd.style.fontWeight = '600';
        emptyTd.textContent = this.getFindingsEmptyMessage(result.totalCount, result.filteredCount);
        emptyTr.appendChild(emptyTd);
        this.tbodyFindings.appendChild(emptyTr);
        return;
      }

      var self = this;
      result.findings.forEach(function (f) {
        var tr = self.createFindingRow(f);
        self.tbodyFindings.appendChild(tr);
      });
    },

    populateWorkbookFindingFilters: function () {
      var manifest = currentCanonicalData && currentCanonicalData.workbookManifest || [];
      var sheetSelect = this.auditSheetFilter;
      var divisionSelect = this.auditDivisionFilter;
      if (!sheetSelect || !divisionSelect) return;
      sheetSelect.innerHTML = '<option value="all">Semua worksheet</option>';
      divisionSelect.innerHTML = '<option value="all">Semua divisi</option>';
      var sheets = {};
      manifest.filter(function (entry) { return entry.status === 'processed'; }).forEach(function (entry) {
        sheets[entry.name] = true;
      });
      Object.keys(sheets).sort().forEach(function (name) {
        var option = document.createElement('option'); option.value = name; option.textContent = name; sheetSelect.appendChild(option);
      });
      var divisions = {};
      (currentCanonicalData.workItems || []).forEach(function (item) {
        var division = getDivisionFilterValue(item);
        divisions[division] = true;
      });
      Object.keys(divisions).sort().forEach(function (name) {
        var option = document.createElement('option'); option.value = name; option.textContent = name; divisionSelect.appendChild(option);
      });
      sheetSelect.disabled = false;
      divisionSelect.disabled = false;
      sheetSelect.value = 'all';
      divisionSelect.value = 'all';
    },

    runSimulation: function () {
      if (!currentCanonicalData || !this.tbodyPareto) return;

      var pct = this.simInputPct ? (parseFloat(this.simInputPct.value) || 0) : 0;
      var simResult = window.UniversalSimulasi.simulate(currentCanonicalData, { percentAll: pct });

      if (this.simTotalBase) this.simTotalBase.textContent = formatRupiah(simResult.totalBaseline);
      if (this.simTotalSim) this.simTotalSim.textContent = formatRupiah(simResult.totalSimulated);
      if (this.simCoverage) {
        this.simCoverage.textContent = simResult.isPartial
          ? 'Parsial: ' + simResult.calculableItemsCount + '/' + simResult.itemsCount + ' item dihitung; ' +
            simResult.uncalculableItemsCount + ' item tidak disesuaikan (nilai sumber dipertahankan).'
          : 'Lengkap: seluruh ' + simResult.itemsCount + ' item dihitung ulang.';
      }
      if (this.simDeltaRp) {
        var sign = simResult.deltaTotal >= 0 ? '+' : '';
        this.simDeltaRp.textContent = sign + formatRupiah(simResult.deltaTotal);
      }
      if (this.simDeltaPct) {
        var signPct = simResult.deltaPercent >= 0 ? '+' : '';
        this.simDeltaPct.textContent = signPct + simResult.deltaPercent.toFixed(2) + '%';
      }

      this.tbodyPareto.innerHTML = '';
      var pareto = simResult.paretoTop10 || [];

      if (pareto.length === 0) {
        var emptyTr = document.createElement('tr');
        var emptyTd = document.createElement('td');
        emptyTd.colSpan = 7;
        emptyTd.className = 'center-col';
        emptyTd.textContent = 'Tidak ada item biaya yang tersedia.';
        emptyTr.appendChild(emptyTd);
        this.tbodyPareto.appendChild(emptyTr);
        return;
      }

      pareto.forEach(function (p) {
        var tr = document.createElement('tr');

        var tdRank = document.createElement('td');
        tdRank.className = 'center-col';
        var strongRank = document.createElement('strong');
        strongRank.textContent = p.rank;
        tdRank.appendChild(strongRank);
        tr.appendChild(tdRank);

        var tdCode = document.createElement('td');
        tdCode.textContent = p.code || '-';
        tr.appendChild(tdCode);

        var tdDesc = document.createElement('td');
        var strongDesc = document.createElement('strong');
        strongDesc.textContent = p.description || '-';
        tdDesc.appendChild(strongDesc);
        tr.appendChild(tdDesc);

        var tdCat = document.createElement('td');
        tdCat.textContent = p.category || '-';
        tr.appendChild(tdCat);

        var tdAmt = document.createElement('td');
        tdAmt.className = 'num-col';
        tdAmt.style.color = 'var(--primary)';
        tdAmt.style.fontWeight = '700';
        tdAmt.textContent = formatRupiah(p.simAmount) + (p.isScenarioAdjusted ? '' : ' (sumber, tidak disesuaikan)');
        tr.appendChild(tdAmt);

        var tdShare = document.createElement('td');
        tdShare.className = 'num-col';
        tdShare.textContent = (typeof p.sharePercent === 'number' ? p.sharePercent.toFixed(2) : '0.00') + '%';
        tr.appendChild(tdShare);

        var tdCum = document.createElement('td');
        tdCum.className = 'num-col';
        var strongCum = document.createElement('strong');
        strongCum.textContent = (typeof p.cumulativePercent === 'number' ? p.cumulativePercent.toFixed(2) : '0.00') + '%';
        tdCum.appendChild(strongCum);
        tr.appendChild(tdCum);

        this.tbodyPareto.appendChild(tr);
      }, this);
    },

    initFeedbackUI: function () {
      var self = this;
      var selectedStars = 5;

      if (this.feedbackStars) {
        this.feedbackStars.forEach(function (star, index) {
          star.addEventListener('click', function () {
            selectedStars = index + 1;
            self.feedbackStars.forEach(function (s, i) {
              if (i < selectedStars) s.classList.add('active');
              else s.classList.remove('active');
            });
          });
        });
      }

      if (this.btnSubmitRating) {
        this.btnSubmitRating.addEventListener('click', function () {
          var comment = self.feedbackComment ? self.feedbackComment.value : '';
          window.UserFeedback.saveRating(selectedStars, comment);
          alert('Terima kasih! Penilaian Anda (' + selectedStars + ' Bintang) telah disimpan di browser lokal.');
          if (self.feedbackComment) self.feedbackComment.value = '';
        });
      }

      if (this.btnSubmitIssue) {
        this.btnSubmitIssue.addEventListener('click', function () {
          var sName = self.issueSheet ? self.issueSheet.value : '';
          var cRef = self.issueCell ? self.issueCell.value : '';
          var desc = self.issueDesc ? self.issueDesc.value : '';
          var exp = self.issueExpected ? self.issueExpected.value : '';

          if (!desc.trim()) {
            alert('Mohon isi deskripsi permasalahan yang ditemukan.');
            return;
          }

          window.UserFeedback.reportParsingIssue(sName, cRef, desc, exp);
          alert('Laporan permasalahan pembacaan Excel berhasil disimpan di browser lokal.');
          if (self.issueSheet) self.issueSheet.value = '';
          if (self.issueCell) self.issueCell.value = '';
          if (self.issueDesc) self.issueDesc.value = '';
          if (self.issueExpected) self.issueExpected.value = '';
        });
      }

      if (this.btnSubmitProp) {
        this.btnSubmitProp.addEventListener('click', function () {
          var title = self.propTitle ? self.propTitle.value : '';
          var cat = self.propCategory ? self.propCategory.value : 'RULE';
          var desc = self.propDesc ? self.propDesc.value : '';

          if (!title.trim() || !desc.trim()) {
            alert('Mohon lengkapi judul dan deskripsi usulan.');
            return;
          }

          window.UserFeedback.submitProposal(title, desc, cat);
          alert('Usulan fitur / aturan baru berhasil disimpan secara lokal.');
          if (self.propTitle) self.propTitle.value = '';
          if (self.propDesc) self.propDesc.value = '';
        });
      }

      if (this.btnExportFeedbackJSON) {
        this.btnExportFeedbackJSON.addEventListener('click', function () {
          window.UserFeedback.exportJSON();
        });
      }

      if (this.btnExportFeedbackCSV) {
        this.btnExportFeedbackCSV.addEventListener('click', function () {
          window.UserFeedback.exportCSV();
        });
      }

      this.renderFalsePositivesTable();
    },

    renderFalsePositivesTable: function () {
      if (!this.tbodyFalsePositives || !window.UserFeedback) return;
      var list = window.UserFeedback.getFalsePositivesList();
      this.tbodyFalsePositives.innerHTML = '';

      if (list.length === 0) {
        var emptyTr = document.createElement('tr');
        var emptyTd = document.createElement('td');
        emptyTd.colSpan = 5;
        emptyTd.className = 'center-col';
        emptyTd.textContent = 'Belum ada temuan yang ditandai sebagai false positive.';
        emptyTr.appendChild(emptyTd);
        this.tbodyFalsePositives.appendChild(emptyTr);
        return;
      }

      var self = this;
      list.forEach(function (fp) {
        var tr = document.createElement('tr');

        var tdId = document.createElement('td');
        var strongId = document.createElement('strong');
        strongId.textContent = fp.findingId;
        tdId.appendChild(strongId);
        tr.appendChild(tdId);

        var tdRule = document.createElement('td');
        var codeRule = document.createElement('code');
        codeRule.textContent = fp.ruleId;
        tdRule.appendChild(codeRule);
        tr.appendChild(tdRule);

        var tdExp = document.createElement('td');
        tdExp.textContent = fp.explanation;
        tr.appendChild(tdExp);

        var tdReason = document.createElement('td');
        tdReason.style.color = '#92400e';
        tdReason.style.fontWeight = '600';
        tdReason.textContent = fp.userReason;
        tr.appendChild(tdReason);

        var tdAction = document.createElement('td');
        tdAction.className = 'center-col';
        var btnCancel = document.createElement('button');
        btnCancel.type = 'button';
        btnCancel.className = 'sample-btn';
        btnCancel.textContent = 'Batalkan';
        btnCancel.addEventListener('click', function () {
          window.UserFeedback.unmarkFalsePositive(fp.findingId);
          self.renderFindings();
          self.renderFalsePositivesTable();
        });
        tdAction.appendChild(btnCancel);
        tr.appendChild(tdAction);

        self.tbodyFalsePositives.appendChild(tr);
      });
    },

    renderSourceViewer: function () {
      if (!window.SourceViewer || !currentParsedWb) return;
      var container = this.auditSourceViewerContainer || document.getElementById('audit-source-viewer');
      if (!container) return;
      if (!window.SourceViewer.isInitialized || window.SourceViewer.container !== container) {
        window.SourceViewer.init(container);
      }

      if (currentAuditResult && currentAuditResult.findings) {
        var counts = this.calculateSheetAnomalyCounts(currentAuditResult.findings);
        window.SourceViewer.setAnomalyCounts(counts);
      }

      var primarySheet = (currentInspection && currentInspection.primaryRabSheet) ||
                         currentParsedWb.primaryRabSheet ||
                         (currentParsedWb.sheetNames && currentParsedWb.sheetNames[0]);
      window.SourceViewer.open(currentParsedWb, primarySheet);
    },

    calculateSheetAnomalyCounts: function (findings) {
      var counts = {};
      if (!Array.isArray(findings)) return counts;
      var self = this;
      findings.forEach(function (f) {
        var locInfo = self.formatFindingLocation(f);
        var sheetSet = {};

        if (locInfo.sheet) sheetSet[locInfo.sheet] = true;
        if (locInfo.all && locInfo.all.length > 0) {
          locInfo.all.forEach(function (loc) {
            if (loc.indexOf('!') !== -1) {
              sheetSet[loc.split('!')[0].trim()] = true;
            }
          });
        }
        if (f.sourceSheet) sheetSet[f.sourceSheet] = true;

        var isError = String(f.severity || '').toUpperCase() === 'ERROR';
        Object.keys(sheetSet).forEach(function (sh) {
          if (!counts[sh]) counts[sh] = { error: 0, warning: 0, total: 0 };
          counts[sh].total++;
          if (isError) counts[sh].error++;
          else counts[sh].warning++;
        });
      });
      return counts;
    },

    formatFindingLocation: function (f) {
      if (!f) return { primary: null, all: [], sheet: '', isMissing: true };
      var sheet = f.sourceSheet || '';
      var rawCell = f.sourceCell || f.cellReference || f.cell || '';

      if (rawCell && rawCell.indexOf('!') !== -1) {
        var p = rawCell.split('!');
        sheet = p[0].trim();
        rawCell = p[1].trim();
      }

      if (!sheet && currentInspection && currentInspection.primaryRabSheet) {
        sheet = currentInspection.primaryRabSheet;
      }
      if (!sheet && currentParsedWb && currentParsedWb.sheetNames) {
        sheet = currentParsedWb.sheetNames[0];
      }

      var locations = [];
      if (Array.isArray(f.sourceReferences) && f.sourceReferences.length > 0) {
        locations = f.sourceReferences.map(function (r) {
          var s = r.sheet || sheet;
          var c = r.cell || r.range || '';
          return s ? (s + '!' + c) : c;
        });
      } else if (rawCell && rawCell.indexOf(',') !== -1) {
        var parts = rawCell.split(',').map(function (s) { return s.trim(); });
        locations = parts.map(function (part) {
          var rowM = part.match(/^Baris\s*(\d+)$/i);
          if (rowM) {
            var colLetter = (f.ruleId === 'inconsistent_unit_price') ? 'G' : 'F';
            return sheet ? (sheet + '!' + colLetter + rowM[1]) : (colLetter + rowM[1]);
          }
          return sheet ? (sheet + '!' + part) : part;
        });
      } else if (rawCell) {
        var rowMatch = rawCell.match(/^Baris\s*(\d+)$/i);
        if (rowMatch) {
          var col = (f.ruleId === 'inconsistent_unit_price') ? 'G' : 'F';
          locations = [sheet ? (sheet + '!' + col + rowMatch[1]) : (col + rowMatch[1])];
        } else {
          locations = [sheet ? (sheet + '!' + rawCell) : rawCell];
        }
      }

      if (locations.length === 0) {
        return {
          primary: null,
          all: [],
          sheet: sheet,
          isMissing: true
        };
      }

      return {
        primary: locations[0],
        additionalCount: locations.length - 1,
        all: locations,
        sheet: sheet,
        isMissing: false
      };
    },

    focusFindingInSourceViewer: function (f, specificTarget) {
      if (!window.SourceViewer || !currentParsedWb) return;
      this.renderSourceViewer();

      var targetRef = specificTarget || f.formattedLocation || f.sourceCell || f.cellReference || '';
      var sheetName = null;
      var cellRef = null;

      if (targetRef && targetRef.indexOf('!') !== -1) {
        var bangParts = targetRef.split('!');
        sheetName = bangParts[0].trim();
        cellRef = bangParts[1].trim();
      } else {
        cellRef = targetRef;
      }

      if (!sheetName) {
        sheetName = f.sourceSheet || f.sheetName || f.sheet ||
                    (currentInspection && currentInspection.primaryRabSheet) ||
                    (currentParsedWb.sheetNames && currentParsedWb.sheetNames[0]);
      }

      window.SourceViewer.open(currentParsedWb, sheetName, cellRef, f);
    }
  };

  UniversalUI.escapeHtml = escapeHtml;
  UniversalUI.formatRupiah = formatRupiah;
  UniversalUI.formatNumber = formatNumber;
  window.UniversalUI = UniversalUI;

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', function () {
      UniversalUI.init();
    });
  } else {
    UniversalUI.init();
  }

})(typeof window !== 'undefined' ? window : this);
