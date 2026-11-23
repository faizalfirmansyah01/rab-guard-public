// js/ui/source-viewer.js
// Embedded Spreadsheet Source Viewer for RAB Guard
// Read-only inspection of original XLSX workbook data with visual anomaly highlighting

(function (window) {
  'use strict';

  function colIndexToLetter(idx) {
    var letter = '';
    while (idx >= 0) {
      letter = String.fromCharCode((idx % 26) + 65) + letter;
      idx = Math.floor(idx / 26) - 1;
    }
    return letter;
  }

  // eslint-disable-next-line no-unused-vars
  function letterToColIndex(letter) {
    var idx = 0;
    var upper = String(letter || '').toUpperCase().replace(/[^A-Z]/g, '');
    for (var i = 0; i < upper.length; i++) {
      idx = idx * 26 + (upper.charCodeAt(i) - 64);
    }
    return idx - 1;
  }

  function cleanCoordinate(coord) {
    if (!coord) return '';
    return String(coord).trim().replace(/\$/g, '').toUpperCase();
  }

  var SourceViewer = {
    container: null,
    gridWrapper: null,
    gridTable: null,
    formulaBar: null,
    nameBox: null,
    tabBar: null,
    filenameSpan: null,
    sheetnameSpan: null,
    metaSpan: null,
    findingBanner: null,
    currentSheet: null,
    currentParsedWb: null,
    anomalyCounts: {},
    isInitialized: false,

    init: function (container) {
      if (typeof container === 'string') {
        container = document.getElementById(container);
      }
      if (!container) return false;
      this.container = container;
      this.container.innerHTML = '';
      this.container.className = 'sv-container';

      // 1. Header Bar
      var header = document.createElement('div');
      header.className = 'sv-header';
      header.innerHTML = [
        '<div class="sv-header-title">',
        '  <span class="sv-header-badge">XLSX SOURCE VIEWER</span>',
        '  <span id="sv-filename" style="font-weight: 700;">-</span>',
        '  <span style="color: #64748b;">&bull;</span>',
        '  <span id="sv-sheetname" style="color: #38bdf8;">-</span>',
        '</div>',
        '<div class="sv-header-meta" id="sv-meta">Mode Baca Asli (Read-Only)</div>'
      ].join('');
      this.container.appendChild(header);

      // 2. Selected Finding Context Banner
      var banner = document.createElement('div');
      banner.className = 'sv-finding-banner';
      banner.id = 'sv-finding-banner';
      banner.innerHTML = [
        '<div class="sv-finding-banner-top">',
        '  <div class="sv-finding-banner-tags">',
        '    <span class="sv-finding-banner-id" id="sv-banner-id">UF-?</span>',
        '    <span id="sv-banner-severity" class="badge-fail">ERROR</span>',
        '    <strong id="sv-banner-rule">-</strong>',
        '    <span class="sv-finding-banner-loc" id="sv-banner-loc">-</span>',
        '  </div>',
        '  <button type="button" class="sv-finding-banner-close" id="sv-banner-close" title="Tutup konteks">&times;</button>',
        '</div>',
        '<div class="sv-finding-banner-desc" id="sv-banner-desc">-</div>',
        '<div class="sv-finding-banner-values" id="sv-banner-values"></div>'
      ].join('');
      this.container.appendChild(banner);

      // 3. Anomaly Legend Bar
      var legend = document.createElement('div');
      legend.className = 'sv-legend';
      legend.innerHTML = [
        '<span class="sv-legend-title">Petunjuk Visual:</span>',
        '<div class="sv-legend-item"><span class="sv-legend-swatch swatch-primary"></span><span>Lokasi anomali utama (Merah)</span></div>',
        '<div class="sv-legend-item"><span class="sv-legend-swatch swatch-supporting"></span><span>Sel pendukung perhitungan (Kuning)</span></div>',
        '<div class="sv-legend-item"><span class="sv-legend-swatch swatch-selected"></span><span>Sel aktif dipilih (Biru)</span></div>'
      ].join('');
      this.container.appendChild(legend);

      // 4. Formula Bar, Name Box & Search Bar
      var formulaContainer = document.createElement('div');
      formulaContainer.className = 'sv-formula-bar-container';
      formulaContainer.innerHTML = [
        '<div class="sv-name-box" id="sv-cell-address" title="Alamat Sel Aktif">-</div>',
        '<span class="sv-formula-fx" aria-hidden="true">fx</span>',
        '<input type="text" id="sv-formula" class="sv-formula-input" readonly placeholder="Formula atau nilai asli sel" title="Formula / Nilai Sel">',
        '<div class="sv-search-bar" id="sv-search-bar">',
        '<input type="text" id="sv-search-input" class="sv-search-input" placeholder="Contoh: RAB!F353 atau F353" title="Cari alamat sel">',
        '<button type="button" id="sv-search-btn" class="sv-search-btn">Cari</button>',
        '<button type="button" id="sv-clear-btn" class="sv-clear-btn">&#x2716;</button>',
        '<span id="sv-search-error" class="sv-search-error"></span>',
        '</div>'
      ].join('');
      this.container.appendChild(formulaContainer);

      // 5. Grid Wrapper & Table
      var gridWrapper = document.createElement('div');
      gridWrapper.className = 'sv-grid-wrapper';
      gridWrapper.id = 'sv-grid-wrapper';
      var gridTable = document.createElement('table');
      gridTable.className = 'sv-grid';
      gridTable.id = 'sv-grid';
      gridWrapper.appendChild(gridTable);
      this.container.appendChild(gridWrapper);

      // 6. Worksheet Tab Bar
      var tabBar = document.createElement('div');
      tabBar.className = 'sv-tab-bar';
      tabBar.id = 'sv-tab-bar';
      this.container.appendChild(tabBar);

      // Cache DOM references
      this.gridWrapper = gridWrapper;
      this.gridTable = gridTable;
      this.formulaBar = formulaContainer.querySelector('#sv-formula');
      this.nameBox = formulaContainer.querySelector('#sv-cell-address');
      this.tabBar = tabBar;
      this.filenameSpan = header.querySelector('#sv-filename');
      this.sheetnameSpan = header.querySelector('#sv-sheetname');
      this.metaSpan = header.querySelector('#sv-meta');
      this.findingBanner = banner;

      // Banner close event
      var bannerCloseBtn = banner.querySelector('#sv-banner-close');
      if (bannerCloseBtn) {
        bannerCloseBtn.addEventListener('click', function () {
          banner.classList.remove('active');
        });
      }

      // Attach search event handlers
      var self = this;
      var searchInput = formulaContainer.querySelector('#sv-search-input');
      var searchBtn = formulaContainer.querySelector('#sv-search-btn');
      var clearBtn = formulaContainer.querySelector('#sv-clear-btn');
      var errorSpan = formulaContainer.querySelector('#sv-search-error');

      function runSearch() {
        var val = searchInput ? searchInput.value.trim() : '';
        if (errorSpan) errorSpan.textContent = '';
        if (!val) {
          if (errorSpan) errorSpan.textContent = 'Masukkan alamat sel';
          return;
        }
        // Populate search input with the queried address
        if (searchInput) searchInput.value = val;
        var ok = self.navigateTo(val);
        if (!ok && errorSpan) {
          errorSpan.textContent = 'Alamat tidak valid atau tidak ditemukan';
        }
      }

      if (searchBtn) {
        searchBtn.addEventListener('click', runSearch);
      }
      if (searchInput) {
        searchInput.addEventListener('keypress', function (e) {
          if (e.key === 'Enter') { runSearch(); }
        });
      }
      if (clearBtn) {
        clearBtn.addEventListener('click', function () {
          if (searchInput) searchInput.value = '';
          if (errorSpan) errorSpan.textContent = '';
          self.clearHighlight();
        });
      }

      // Attach cell click delegation
      this.gridTable.addEventListener('click', function (e) {
        var cellEl = e.target.closest('td.sv-cell');
        if (!cellEl) return;
        self._selectCellElement(cellEl, false);
      });

      this.isInitialized = true;
      return true;
    },

    setAnomalyCounts: function (counts) {
      this.anomalyCounts = counts || {};
      if (this.currentParsedWb && this.tabBar) {
        this._renderTabs(this.currentSheet);
      }
    },

    open: function (parsedWb, sheetName, sourceReference, findingContext) {
      if (!parsedWb) return;
      this.currentParsedWb = parsedWb;

      if (!this.isInitialized) {
        var defaultContainer = document.getElementById('audit-source-viewer');
        if (defaultContainer) this.init(defaultContainer);
      }
      if (!this.container) return;

      // Determine initial sheet
      var targetSheet = sheetName;
      if (!targetSheet || !parsedWb.sheets || !parsedWb.sheets[targetSheet]) {
        targetSheet = (parsedWb.inspection && parsedWb.inspection.primaryRabSheet) ||
                      parsedWb.primaryRabSheet ||
                      (parsedWb.sheetNames && parsedWb.sheetNames[0]);
      }

      this._renderTabs(targetSheet);

      // Check if target reference requires higher row count before rendering
      var neededRow = 0;
      if (sourceReference) {
        var refStr = typeof sourceReference === 'string'
          ? sourceReference
          : (sourceReference.cell || sourceReference.range || '');
        var match = refStr.match(/(\d+)/);
        if (match) neededRow = parseInt(match[1], 10);
      }

      this.renderSheet(targetSheet, neededRow);

      if (findingContext) {
        this.showFindingContext(findingContext);
      }

      if (sourceReference) {
        this.highlightProvenance(sourceReference);
      }
    },

    _renderTabs: function (activeSheet) {
      if (!this.tabBar || !this.currentParsedWb) return;
      this.tabBar.innerHTML = '';
      var sheetNames = this.currentParsedWb.sheetNames || Object.keys(this.currentParsedWb.sheets || {});
      var self = this;

      sheetNames.forEach(function (name) {
        var btn = document.createElement('button');
        btn.type = 'button';
        btn.className = 'sv-tab-btn' + (name === activeSheet ? ' active' : '');
        btn.dataset.sheet = name;

        var nameSpan = document.createElement('span');
        nameSpan.textContent = name;
        btn.appendChild(nameSpan);

        // Add anomaly notification bubble if sheet has findings
        var sheetAnomalies = self.anomalyCounts[name];
        if (sheetAnomalies && sheetAnomalies.total > 0) {
          var badge = document.createElement('span');
          badge.className = 'sv-tab-badge';
          if (sheetAnomalies.error > 0 && sheetAnomalies.warning > 0) {
            badge.classList.add('sv-tab-badge-mixed');
            badge.title = sheetAnomalies.error + ' error, ' + sheetAnomalies.warning + ' warning';
          } else if (sheetAnomalies.error > 0) {
            badge.classList.add('sv-tab-badge-error');
            badge.title = sheetAnomalies.error + ' error';
          } else {
            badge.classList.add('sv-tab-badge-warn');
            badge.title = sheetAnomalies.warning + ' warning';
          }
          badge.textContent = sheetAnomalies.total;
          btn.appendChild(badge);
        }

        btn.addEventListener('click', function () {
          self.renderSheet(name);
        });
        self.tabBar.appendChild(btn);
      });
    },

    renderSheet: function (sheetName, minRequiredRow) {
      if (!this.currentParsedWb || !this.currentParsedWb.sheets) return;
      var sheet = this.currentParsedWb.sheets[sheetName];
      if (!sheet) return;
      this.currentSheet = sheetName;

      // Update header titles
      var fname = this.currentParsedWb.filename || 'Workbook';
      if (this.filenameSpan) this.filenameSpan.textContent = fname;
      if (this.sheetnameSpan) this.sheetnameSpan.textContent = sheetName;

      // Update tab active classes
      if (this.tabBar) {
        var tabBtns = this.tabBar.querySelectorAll('.sv-tab-btn');
        tabBtns.forEach(function (btn) {
          if (btn.dataset.sheet === sheetName) btn.classList.add('active');
          else btn.classList.remove('active');
        });
      }

      // Clear existing grid
      this.gridTable.innerHTML = '';
      this.clearHighlight(false);

      // Determine dimensions
      var rangeStr = sheet['!ref'] || 'A1:A1';
      var range = window.XLSX && window.XLSX.utils
        ? window.XLSX.utils.decode_range(rangeStr)
        : { s: { r: 0, c: 0 }, e: { r: 50, c: 10 } };

      var startRow = range.s.r;
      var startCol = range.s.c;
      var endRow = Math.max(range.e.r, 25);
      var endCol = Math.max(range.e.c, 7);

      // If a specific row is needed (e.g. row 833), ensure endRow covers it
      if (minRequiredRow && minRequiredRow > endRow) {
        endRow = Math.min(minRequiredRow + 20, 3000);
      }

      // Sensible render boundary for large sheets (up to 2500 rows)
      if (endRow - startRow > 2500) {
        endRow = startRow + 2500;
      }
      if (endCol - startCol > 60) {
        endCol = startCol + 60;
      }

      if (this.metaSpan) {
        this.metaSpan.textContent = 'Rentang: ' + rangeStr +
          ' (' + (range.e.r - range.s.r + 1) + ' baris \xD7 ' +
          (range.e.c - range.s.c + 1) + ' kolom)';
      }

      // 1. Column Headers (A, B, C...)
      var thead = document.createElement('thead');
      var trHead = document.createElement('tr');
      var corner = document.createElement('th');
      corner.className = 'sv-corner';
      trHead.appendChild(corner);

      for (var c = startCol; c <= endCol; c++) {
        var th = document.createElement('th');
        th.className = 'sv-col-header';
        th.textContent = colIndexToLetter(c);
        th.dataset.col = c;
        trHead.appendChild(th);
      }
      thead.appendChild(trHead);
      this.gridTable.appendChild(thead);

      // 2. Data Rows
      var tbody = document.createElement('tbody');
      for (var r = startRow; r <= endRow; r++) {
        var tr = document.createElement('tr');
        var rowHeader = document.createElement('th');
        rowHeader.className = 'sv-row-header';
        rowHeader.textContent = String(r + 1);
        rowHeader.dataset.row = r;
        tr.appendChild(rowHeader);

        for (var col = startCol; col <= endCol; col++) {
          var td = document.createElement('td');
          td.className = 'sv-cell';
          td.dataset.row = r;
          td.dataset.col = col;

          var cellAddr = colIndexToLetter(col) + (r + 1);
          td.dataset.addr = cellAddr;

          var cell = sheet[cellAddr];
          if (cell) {
            var displayVal = cell.w !== undefined ? cell.w : (cell.v !== undefined ? cell.v : '');
            td.textContent = displayVal;

            if (typeof cell.v === 'number') {
              td.classList.add('num-val');
            } else if (typeof cell.v === 'string' && cell.v.length <= 6 &&
                       /^[A-Z0-9\-\.\/]+$/i.test(cell.v.trim())) {
              td.classList.add('center-val');
            }

            td._rawCell = {
              f: cell.f || '',
              v: cell.v,
              w: cell.w,
              t: cell.t
            };
          }

          tr.appendChild(td);
        }
        tbody.appendChild(tr);
      }
      this.gridTable.appendChild(tbody);
    },

    _selectCellElement: function (cellEl, scrollIntoView) {
      if (!cellEl) return;
      var prev = this.gridTable.querySelectorAll('.sv-selected');
      prev.forEach(function (el) { el.classList.remove('sv-selected'); });

      cellEl.classList.add('sv-selected');

      var addr = cellEl.dataset.addr ||
        (colIndexToLetter(parseInt(cellEl.dataset.col, 10)) + (parseInt(cellEl.dataset.row, 10) + 1));
      if (this.nameBox) this.nameBox.textContent = addr;

      var raw = cellEl._rawCell;
      var formulaText = '';
      if (raw && raw.f) {
        formulaText = '=' + raw.f;
      } else if (raw && raw.w !== undefined) {
        formulaText = String(raw.w);
      } else if (raw && raw.v !== undefined && raw.v !== null) {
        formulaText = String(raw.v);
      } else {
        formulaText = cellEl.textContent || '';
      }

      if (this.formulaBar) this.formulaBar.value = formulaText;

      if (scrollIntoView && cellEl.scrollIntoView) {
        cellEl.scrollIntoView({ block: 'center', inline: 'center' });
      }
    },

    showFindingContext: function (f) {
      if (!this.findingBanner || !f) return;

      var bannerId = this.findingBanner.querySelector('#sv-banner-id');
      var bannerSev = this.findingBanner.querySelector('#sv-banner-severity');
      var bannerRule = this.findingBanner.querySelector('#sv-banner-rule');
      var bannerLoc = this.findingBanner.querySelector('#sv-banner-loc');
      var bannerDesc = this.findingBanner.querySelector('#sv-banner-desc');
      var bannerValues = this.findingBanner.querySelector('#sv-banner-values');

      var isError = String(f.severity || '').toUpperCase() === 'ERROR';
      if (isError) {
        this.findingBanner.className = 'sv-finding-banner active banner-error';
        if (bannerSev) {
          bannerSev.className = 'badge-fail';
          bannerSev.textContent = 'ERROR';
        }
      } else {
        this.findingBanner.className = 'sv-finding-banner active';
        if (bannerSev) {
          bannerSev.className = 'badge-tinjau';
          bannerSev.textContent = f.severity || 'WARNING';
        }
      }

      if (bannerId) bannerId.textContent = f.id || 'UF';
      if (bannerRule) bannerRule.textContent = f.ruleId || 'Audit Anomaly';

      var locText = f.formattedLocation || f.sourceCell ||
        (f.sourceSheet ? f.sourceSheet + '!' + (f.cellReference || '') : '-');
      if (bannerLoc) bannerLoc.textContent = locText;

      if (bannerDesc) {
        bannerDesc.textContent = f.explanation || f.recommendedAction || 'Anomali terdeteksi pada sel sumber ini.';
      }

      if (bannerValues) {
        bannerValues.innerHTML = '';
        if (f.originalValue !== undefined && f.originalValue !== null) {
          var valItem1 = document.createElement('div');
          valItem1.className = 'sv-finding-banner-val-item';
          valItem1.innerHTML = '<strong>Nilai sumber:</strong> <span>' + f.originalValue + '</span>';
          bannerValues.appendChild(valItem1);
        }
        if (f.expectedValue !== undefined && f.expectedValue !== null) {
          var valItem2 = document.createElement('div');
          valItem2.className = 'sv-finding-banner-val-item';
          valItem2.innerHTML = '<strong>Ekspektasi:</strong> <span>' + f.expectedValue + '</span>';
          bannerValues.appendChild(valItem2);
        }
        if (f.difference !== undefined && f.difference !== null) {
          var valItem3 = document.createElement('div');
          valItem3.className = 'sv-finding-banner-val-item';
          valItem3.innerHTML = '<strong>Selisih:</strong> <span>' + f.difference + '</span>';
          bannerValues.appendChild(valItem3);
        }
      }
    },

    highlightProvenance: function (sourceRef) {
      if (!sourceRef) return false;
      this.clearHighlight(false);

      var sheetName = this.currentSheet;
      var cellOrRange = '';

      if (typeof sourceRef === 'string') {
        var str = cleanCoordinate(sourceRef);
        if (str.indexOf('!') !== -1) {
          var p = str.split('!');
          sheetName = p[0].trim();
          cellOrRange = p[1].trim();
        } else {
          cellOrRange = str;
        }
      } else if (typeof sourceRef === 'object') {
        sheetName = sourceRef.sheet || sourceRef.sourceSheet || this.currentSheet;
        cellOrRange = sourceRef.cell || sourceRef.range || sourceRef.ref || sourceRef.sourceCell || '';
        if (cellOrRange.indexOf('!') !== -1) {
          var parts2 = cellOrRange.split('!');
          sheetName = parts2[0];
          cellOrRange = parts2[1];
        }
      }

      cellOrRange = cleanCoordinate(cellOrRange);

      // Check for row references like "Baris 127"
      var rowMatch = cellOrRange.match(/^Baris\s*(\d+)$/i);
      if (rowMatch) {
        var rNum = parseInt(rowMatch[1], 10);
        cellOrRange = 'G' + rNum;
      }

      if (sheetName && sheetName !== this.currentSheet) {
        var rowNum = 0;
        var rMatch = cellOrRange.match(/(\d+)/);
        if (rMatch) rowNum = parseInt(rMatch[1], 10);
        this.renderSheet(sheetName, rowNum);
      }

      if (cellOrRange.indexOf(':') !== -1) {
        return this.focusRange(cellOrRange);
      } else {
        return this.focusCell(cellOrRange, false);
      }
    },

    // navigateTo: unified navigation used by manual search + location tags + finding clicks
    navigateTo: function (ref) {
      if (!ref) return false;

      var clean = cleanCoordinate(String(ref).trim());
      var sheetPart = null;
      var addrPart = clean;

      if (clean.indexOf('!') !== -1) {
        var parts = clean.split('!');
        sheetPart = parts[0].trim();
        addrPart = parts[1].trim();
        if (sheetPart && sheetPart !== this.currentSheet) {
          // Compute required row for the target address so we pre-render enough rows
          var rowNeeded = 0;
          var rowM = addrPart.match(/(\d+)/);
          if (rowM) rowNeeded = parseInt(rowM[1], 10);
          this.renderSheet(sheetPart, rowNeeded);
        }
      }

      if (addrPart.indexOf(':') !== -1) {
        return this.focusRange(addrPart);
      }
      return this.focusCell(addrPart, false);
    },

    focusCell: function (cellReference, isSupporting) {
      if (!cellReference || !this.gridTable) return false;

      var cleanRef = cleanCoordinate(cellReference);
      if (cleanRef.indexOf('!') !== -1) {
        var parts = cleanRef.split('!');
        var sName = parts[0];
        cleanRef = parts[1];
        if (sName && sName !== this.currentSheet) {
          var m = cleanRef.match(/(\d+)/);
          this.renderSheet(sName, m ? parseInt(m[1], 10) : 0);
        }
      }

      var cellEl = this.gridTable.querySelector('td.sv-cell[data-addr="' + cleanRef + '"]');

      // If not in DOM because row is beyond initial boundary, expand render
      if (!cellEl && window.XLSX && window.XLSX.utils) {
        try {
          var decoded = window.XLSX.utils.decode_cell(cleanRef);
          var requiredRow = decoded.r + 1;
          this.renderSheet(this.currentSheet, requiredRow);
          cellEl = this.gridTable.querySelector('td.sv-cell[data-addr="' + cleanRef + '"]');
        } catch (e) { /* invalid cell address */ }
      }

      if (cellEl) {
        // Select first (updates name box & formula bar), then apply colour highlight
        this._selectCellElement(cellEl, true);
        if (isSupporting) {
          cellEl.classList.add('sv-supporting-highlight');
        } else {
          cellEl.classList.add('sv-anomaly-highlight');
        }
        return true;
      }
      return false;
    },

    focusRange: function (rangeReference) {
      if (!rangeReference || !this.gridTable) return false;

      var cleanRange = cleanCoordinate(rangeReference);
      if (cleanRange.indexOf('!') !== -1) {
        var parts = cleanRange.split('!');
        var sName = parts[0];
        cleanRange = parts[1];
        if (sName && sName !== this.currentSheet) {
          this.renderSheet(sName);
        }
      }

      if (cleanRange.indexOf(':') === -1) {
        return this.focusCell(cleanRange, false);
      }

      if (!window.XLSX || !window.XLSX.utils) return false;
      try {
        var rng = window.XLSX.utils.decode_range(cleanRange);
        var primaryCellEl = null;

        // Ensure range end row is rendered
        var maxRowRequired = rng.e.r + 1;
        var existingRows = this.gridTable.querySelectorAll('tbody tr').length;
        if (maxRowRequired > existingRows) {
          this.renderSheet(this.currentSheet, maxRowRequired);
        }

        for (var r = rng.s.r; r <= rng.e.r; r++) {
          for (var c = rng.s.c; c <= rng.e.c; c++) {
            var el = this.gridTable.querySelector(
              'td.sv-cell[data-row="' + r + '"][data-col="' + c + '"]'
            );
            if (el) {
              if (r === rng.s.r && c === rng.s.c) {
                primaryCellEl = el;
                // Primary: select then mark as anomaly
                this._selectCellElement(el, true);
                el.classList.add('sv-anomaly-highlight');
              } else {
                el.classList.add('sv-supporting-highlight');
              }
            }
          }
        }

        if (primaryCellEl) {
          if (this.nameBox) this.nameBox.textContent = cleanRange;
          return true;
        }
      } catch (e) {
        console.warn('[SourceViewer] Failed to parse range:', rangeReference, e);
      }
      return false;
    },

    clearHighlight: function (hideBanner) {
      if (!this.gridTable) return;
      var highlighted = this.gridTable.querySelectorAll(
        '.sv-selected, .sv-range-selected, .sv-anomaly-highlight, .sv-supporting-highlight'
      );
      highlighted.forEach(function (el) {
        el.classList.remove('sv-selected', 'sv-range-selected', 'sv-anomaly-highlight', 'sv-supporting-highlight');
      });
      if (this.formulaBar) this.formulaBar.value = '';
      if (this.nameBox) this.nameBox.textContent = '-';
      if (hideBanner !== false && this.findingBanner) {
        this.findingBanner.classList.remove('active');
      }
    }
  };

  window.SourceViewer = SourceViewer;
})(typeof window !== 'undefined' ? window : this);
