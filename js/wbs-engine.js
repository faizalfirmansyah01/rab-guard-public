/*
 * RAB GUARD — WBS Engine
 * Module: wbs-engine.js
 *
 * Responsibilities:
 *   • Detect hierarchical WBS nodes (ROOT, DIVISION, SUBDIVISION, WORK_PACKAGE, WORK_ITEM, NON_WBS, UNKNOWN).
 *   • Infer parent/child relationships using code patterns, headings, row order, and indentation metadata when available.
 *   • Produce a tree of nodes with explicit properties required by the Phase 2 contract.
 *   • Export a simple UI helper to render the tree in the existing application (UniversalUI).
 */

(function (window) {
  'use strict';

  /** Utility: simple round‑away‑from‑zero as defined in the contract (used elsewhere). */
  function roundAwayFromZero(x) {
    return x >= 0 ? Math.floor(x + 0.5) : Math.ceil(x - 0.5);
  }

  /**
   * Detects the WBS node type for a given canonical work item.
   *   - Uses code patterns (e.g., "1", "1.1", "A", "I.A").
   *   - Considers rowType (DETAIL, SUBTOTAL, HEADER, NOTE, etc.).
   *   - Returns an object with type, confidence, and evidence description.
   */
  function detectNodeType(item) {
    var code = (item.code || '').trim();
    var desc = (item.description || '').trim().toLowerCase();
    var rowType = item.rowType || 'UNKNOWN';

    // Helper regexes for common hierarchical codes
    var numericHier = /^\d+(?:\.\d+)*$/;
    var alphaHier = /^[A-Z](?:\.\d+)*$/i;
    var romanHier = /^(?=[MDCLXVI])M{0,4}(CM|CD|D?C{0,3})(XC|XL|L?X{0,3})(IX|IV|V?I{0,3})(?:\.[A-Z0-9]+)*$/i;

    // ROOT is a synthetic node that will be added later; never inferred here.
    if (rowType === 'HEADER') {
      if (numericHier.test(code)) {
        var headerNumericParts = code.split('.');
        if (headerNumericParts.length === 1)
          return { type: 'DIVISION', confidence: 'CONFIRMED', evidence: 'Numeric header division' };
        if (headerNumericParts.length === 2)
          return { type: 'SUBDIVISION', confidence: 'INFERRED', evidence: 'Numeric header subdivision' };
        return { type: 'WORK_PACKAGE', confidence: 'INFERRED', evidence: 'Numeric header deeper hierarchy' };
      }

      if (romanHier.test(code)) {
        var headerRomanParts = code.split('.');
        if (headerRomanParts.length === 1)
          return { type: 'DIVISION', confidence: 'INFERRED', evidence: 'Roman header division' };
        if (headerRomanParts.length === 2)
          return { type: 'SUBDIVISION', confidence: 'INFERRED', evidence: 'Roman header subdivision' };
        return { type: 'WORK_PACKAGE', confidence: 'INFERRED', evidence: 'Roman header deeper hierarchy' };
      }

      if (alphaHier.test(code) && !romanHier.test(code)) {
        var headerAlphaParts = code.split('.');
        if (headerAlphaParts.length === 1)
          return { type: 'DIVISION', confidence: 'INFERRED', evidence: 'Alphabetic header division' };
        if (headerAlphaParts.length === 2)
          return { type: 'SUBDIVISION', confidence: 'INFERRED', evidence: 'Alphabetic header subdivision' };
        return { type: 'WORK_PACKAGE', confidence: 'INFERRED', evidence: 'Alphabetic header deeper hierarchy' };
      }

      return { type: 'DIVISION', confidence: 'INFERRED', evidence: 'Uncoded header row' };
    }
    if (rowType === 'NOTE') {
      return { type: 'NOTE', confidence: 'INFERRED', evidence: 'Note row detected' };
    }
    if (rowType === 'SUBTOTAL' || rowType === 'GRAND_TOTAL') {
      return { type: 'NON_WBS', confidence: 'CONFIRMED', evidence: 'Subtotal/Grand‑total row' };
    }
    if (rowType === 'EMPTY') {
      return { type: 'UNKNOWN', confidence: 'NEEDS_CONFIRMATION', evidence: 'Empty row' };
    }

    // Look at the code pattern first
    if (numericHier.test(code)) {
      var parts = code.split('.');
      if (parts.length === 1) return { type: 'DIVISION', confidence: 'CONFIRMED', evidence: 'Numeric top‑level code' };
      if (parts.length === 2) return { type: 'SUBDIVISION', confidence: 'INFERRED', evidence: 'Numeric second‑level code' };
      if (parts.length >= 3) return { type: 'WORK_PACKAGE', confidence: 'INFERRED', evidence: 'Numeric deeper hierarchy' };
    }
    if (alphaHier.test(code) && !romanHier.test(code)) {
      var partsA = code.split('.');
      if (partsA.length === 1) return { type: 'DIVISION', confidence: 'INFERRED', evidence: 'Alphabetic top‑level code' };
      if (partsA.length === 2) return { type: 'SUBDIVISION', confidence: 'INFERRED', evidence: 'Alphabetic second‑level code' };
      return { type: 'WORK_PACKAGE', confidence: 'INFERRED', evidence: 'Alphabetic deeper hierarchy' };
    }
    if (romanHier.test(code)) {
      // Treat roman numerals as divisions or subdivisions based on dot depth
      var partsR = code.split('.');
      if (partsR.length === 1) return { type: 'DIVISION', confidence: 'INFERRED', evidence: 'Roman numeral division' };
      if (partsR.length === 2) return { type: 'SUBDIVISION', confidence: 'INFERRED', evidence: 'Roman numeral subdivision' };
      return { type: 'WORK_PACKAGE', confidence: 'INFERRED', evidence: 'Roman numeral deeper level' };
    }

    // Fallback: if description contains typical WBS keywords, infer a WORK_ITEM
    if (desc && (desc.includes('pekerjaan') || desc.includes('item') || desc.includes('paket'))) {
      return { type: 'WORK_ITEM', confidence: 'INFERRED', evidence: 'Description keyword' };
    }

    // If none matches, mark as UNKNOWN needing confirmation
    return { type: 'UNKNOWN', confidence: 'NEEDS_CONFIRMATION', evidence: 'No pattern matched' };
  }

  /**
   * Build a WBS tree from the canonical data model.
   *   - Adds a synthetic ROOT node.
   *   - Assigns wbsNodeId (generated UUID‑like string), wbsParentId, wbsLevel, wbsCode, wbsPath.
   *   - Guarantees parent‑child consistency, no cycles.
   */
  function buildTree(canonicalData) {
    var nodes = [];
    var idCounter = 1;
    // Synthetic root
    var rootNode = {
      wbsNodeId: 'root',
      wbsParentId: null,
      wbsLevel: 0,
      wbsNodeType: 'ROOT',
      wbsCode: null,
      wbsPath: '/',
      wbsConfidence: 'CONFIRMED',
      wbsDetectionEvidence: 'Synthetic root',
      sourceSheet: null,
      sourceRow: null,
      children: []
    };
    nodes.push(rootNode);

    // First pass: detect type for each work item and create node objects
    canonicalData.workItems.forEach(function (item) {
      var detection = detectNodeType(item);
      var node = {
        wbsNodeId: 'node' + (idCounter++),
        wbsParentId: null, // to be filled later
        wbsLevel: null,
        wbsNodeType: detection.type,
        wbsCode: item.code || null,
        wbsPath: null,
        wbsConfidence: detection.confidence,
        wbsDetectionEvidence: detection.evidence,
        sourceSheet: item.raw ? item.raw.sheet : null,
        sourceRow: item.raw ? item.raw.row : null,
        // Preserve reference to original canonical item for later cost roll‑up
        canonicalItem: item,
        children: []
      };
      nodes.push(node);
    });

    // Simple hierarchical linking algorithm:
    // After linking, compute cost roll‑up aggregates for each node.
    function aggregateNode(node) {
      // Leaf: distinguish structural headers from actual cost items.
      if (!node.children || node.children.length === 0) {
        var item = node.canonicalItem || {};

        var sourceAmt =
          typeof item.amount === 'number' && Number.isFinite(item.amount)
            ? item.amount : null;

        var hasVolume =
          typeof item.volume === 'number' && Number.isFinite(item.volume);

        var hasPrice =
          typeof item.unitPrice === 'number' && Number.isFinite(item.unitPrice);

        var isStructuralHeader =
          String(item.rowType || '').trim().toUpperCase() === 'HEADER';

        var isSupportedType = [
          'DIVISION',
          'SUBDIVISION',
          'WORK_PACKAGE',
          'WORK_ITEM'
        ].includes(node.wbsNodeType);

        var isAggregatable = isSupportedType && !isStructuralHeader;

        var recomputedAmt = (hasVolume && hasPrice)
          ? item.volume * item.unitPrice
          : null;

        if (recomputedAmt !== null && !Number.isFinite(recomputedAmt)) {
          recomputedAmt = null;
        }

        node.totalSourceAmount = isAggregatable ? sourceAmt : null;
        node.totalRecomputedAmount = isAggregatable ? recomputedAmt : null;

        node.itemCount = isAggregatable ? 1 : 0;

        node.calculatedCount =
          (isAggregatable && recomputedAmt !== null) ? 1 : 0;

        node.problemCount =
          (isAggregatable && recomputedAmt === null) ? 1 : 0;

        if (isStructuralHeader) {
          node.calculationStatus = 'STRUCTURAL';
        } else if (!isAggregatable) {
          node.calculationStatus = 'EXCLUDED';
        } else {
          node.calculationStatus =
            recomputedAmt !== null ? 'COMPLETE' : 'PARTIAL';
        }

        return;
      }

      // Parent: aggregate all children after recursively processing them.
      var aggSource = 0;
      var aggRecalc = 0;
      var totalItems = 0;
      var calcItems = 0;
      var probItems = 0;

      var hasSourceAmount = false;
      var hasRecomputedAmount = false;
      var allComplete = true;

      node.children.forEach(function (child) {
        aggregateNode(child);

        if (child.totalSourceAmount !== null) {
          aggSource += child.totalSourceAmount;
          hasSourceAmount = true;
        }

        if (child.totalRecomputedAmount !== null) {
          aggRecalc += child.totalRecomputedAmount;
          hasRecomputedAmount = true;
        }

        totalItems += child.itemCount || 0;
        calcItems += child.calculatedCount || 0;
        probItems += child.problemCount || 0;

        if (
          child.calculationStatus !== 'COMPLETE' &&
          child.calculationStatus !== 'STRUCTURAL' &&
          child.calculationStatus !== 'EXCLUDED'
        ) {
          allComplete = false;
        }
      });

      node.totalSourceAmount = hasSourceAmount ? aggSource : null;
      node.totalRecomputedAmount = hasRecomputedAmount ? aggRecalc : null;

      node.itemCount = totalItems;
      node.calculatedCount = calcItems;
      node.problemCount = probItems;

      if (totalItems === 0) {
        node.calculationStatus = 'STRUCTURAL';
      } else {
        node.calculationStatus = allComplete ? 'COMPLETE' : 'PARTIAL';
      }
    }

    // Compute contribution percentages for each node (relative to project total)
    function computeContribution(node, projectTotal) {
      if (projectTotal && node.totalRecomputedAmount !== null && projectTotal > 0) {
        node.contributionPct = (node.totalRecomputedAmount / projectTotal) * 100;
      } else {
        node.contributionPct = 0;
      }
      if (node.children && node.children.length > 0) {
        node.children.forEach(function (c) { computeContribution(c, projectTotal); });
      }
    }

    // Parent-child linking: use strict hierarchical code prefixes.
    for (var i = 1; i < nodes.length; i++) {
      var cur = nodes[i];

      if (cur.wbsNodeType === 'NON_WBS' || cur.wbsNodeType === 'UNKNOWN') {
        cur.wbsParentId = rootNode.wbsNodeId;
        cur.wbsLevel = 1;
        cur.wbsPath = '/' + (cur.wbsCode || cur.wbsNodeId);
        rootNode.children.push(cur);
        continue;
      }

      var curCode = cur.wbsCode !== null && cur.wbsCode !== undefined
        ? String(cur.wbsCode).trim().toUpperCase()
        : '';

      var parent = null;

      // Find the nearest previous node whose code is a strict prefix.
      if (curCode) {
        for (var j = i - 1; j >= 1; j--) {
          var candidate = nodes[j];

          if (candidate.wbsNodeType === 'NON_WBS' ||
              candidate.wbsNodeType === 'UNKNOWN') {
            continue;
          }

          var candidateCode = candidate.wbsCode !== null &&
                              candidate.wbsCode !== undefined
            ? String(candidate.wbsCode).trim().toUpperCase()
            : '';

          if (candidateCode &&
              curCode.indexOf(candidateCode + '.') === 0) {
            parent = candidate;
            break;
          }
        }
      } else {
        // Fallback for items without codes: use semantic type hierarchy.
        var typeRank = {
          'ROOT': 0,
          'DIVISION': 1,
          'SUBDIVISION': 2,
          'WORK_PACKAGE': 3,
          'WORK_ITEM': 4,
          'NON_WBS': 5,
          'UNKNOWN': 6
        };

        var curRank = typeRank[cur.wbsNodeType];

        if (typeof curRank === 'number') {
          for (var k = i - 1; k >= 1; k--) {
            var previous = nodes[k];
            var previousRank = typeRank[previous.wbsNodeType];

            if (typeof previousRank === 'number' &&
                previousRank < curRank &&
                previous.wbsNodeType !== 'NON_WBS' &&
                previous.wbsNodeType !== 'UNKNOWN') {
              parent = previous;
              break;
            }
          }
        }
      }

      if (parent) {
        cur.wbsParentId = parent.wbsNodeId;
        cur.wbsLevel = (parent.wbsLevel || 0) + 1;
        cur.wbsPath = (parent.wbsPath || '/') +
          (parent.wbsPath === '/' ? '' : '/') +
          (cur.wbsCode || cur.wbsNodeId);

        parent.children.push(cur);
      } else {
        cur.wbsParentId = rootNode.wbsNodeId;
        cur.wbsLevel = 1;
        cur.wbsPath = '/' + (cur.wbsCode || cur.wbsNodeId);
        rootNode.children.push(cur);
      }
    }

    // Aggregate after hierarchy linking is complete.
    aggregateNode(rootNode);
    // Calculate contribution using final project total.
    computeContribution(rootNode, rootNode.totalRecomputedAmount);
    return rootNode;
  }

  /**
   * Simple recursive renderer that creates an indented HTML list of the tree.
   * Each node shows code, description (if available), confidence badge and source reference.
   */
  function renderNode(node) {
      var li = document.createElement('li');
      // store data attributes for search/filter
      if (node.wbsCode) li.setAttribute('data-code', node.wbsCode);
      if (node.canonicalItem && node.canonicalItem.description) li.setAttribute('data-description', node.canonicalItem.description);
      if (node.calculationStatus) li.setAttribute('data-status', node.calculationStatus);
      var span = document.createElement('span');
      var label = '';
      if (node.wbsNodeType !== 'ROOT') {
        label += '[' + node.wbsNodeType + '] ';
        if (node.wbsCode) label += node.wbsCode + ' – ';
        if (node.canonicalItem && node.canonicalItem.description) label += node.canonicalItem.description;
      } else {
        label = 'ROOT (Project)';
      }
      // Cost info badges
      if (node.totalRecomputedAmount !== undefined && node.totalRecomputedAmount !== null) {
        label += ' (Rp' + node.totalRecomputedAmount.toLocaleString('id-ID') + ')';
      }
      // Confidence badge
      var badge = document.createElement('strong');
      badge.textContent = ' (' + (node.wbsConfidence || node.calculationStatus) + ')';
      span.textContent = label;
      li.appendChild(span);
      li.appendChild(badge);
      if (node.sourceSheet && node.sourceRow) {
        var src = document.createElement('div');
        src.style.fontSize = '0.8em';
        src.style.color = '#555';
        src.textContent = 'Source: ' + node.sourceSheet + '! Row ' + node.sourceRow;
        li.appendChild(src);
      }
      // Simple edit controls (click to edit)
      var editBtn = document.createElement('button');
      editBtn.textContent = 'Edit';
      editBtn.style.marginLeft = '8px';
      editBtn.style.fontSize = '0.8em';
      editBtn.addEventListener('click', function (e) {
        e.stopPropagation();
        var newParent = prompt('Enter new parent nodeId (or leave empty to keep):', node.wbsParentId || '');
        var newType = prompt('Enter new node type (DIVISION, SUBDIVISION, WORK_PACKAGE, WORK_ITEM, NON_WBS, UNKNOWN):', node.wbsNodeType);
        if (newParent !== null) {
          window.WbsEngine.updateNodeParent(node.wbsNodeId, newParent.trim() || null);
        }
        if (newType !== null) {
          window.WbsEngine.setNodeType(node.wbsNodeId, newType.trim().toUpperCase());
        }
        // Refresh explorer
        window.WbsEngine.showExplorer(window.currentCanonicalData);
      });
      li.appendChild(editBtn);
      if (node.children && node.children.length > 0) {
        var ul = document.createElement('ul');
        node.children.forEach(function (child) { ul.appendChild(renderNode(child)); });
        li.appendChild(ul);
      }
      return li;
    }
  /**
   * Public API – build the tree and display it in a modal overlay.
   */
  function showExplorer(canonicalData) {
    // Build tree (may be expensive for large workbooks, but acceptable in browser)
    var treeRoot = buildTree(canonicalData);
    // Create overlay if not already present
    var overlay = document.getElementById('wbs-explorer-overlay');
    if (!overlay) {
      overlay = document.createElement('div');
      overlay.id = 'wbs-explorer-overlay';
      overlay.style.position = 'fixed';
      overlay.style.left = '0';
      overlay.style.top = '0';
      overlay.style.width = '100%';
      overlay.style.height = '100%';
      overlay.style.backgroundColor = 'rgba(0,0,0,0.8)';
      overlay.style.zIndex = '10000';
      overlay.style.overflow = 'auto';
      overlay.style.color = '#fff';
      overlay.style.padding = '20px';
      document.body.appendChild(overlay);
    }
    overlay.innerHTML = '<h2>WBS Explorer</h2><button id="wbs-close" style="margin-bottom:10px;">Close</button>'; // simple close button
    var closeBtn = document.getElementById('wbs-close');
    closeBtn.onclick = function () { overlay.style.display = 'none'; };
    var ul = document.createElement('ul');
    ul.appendChild(renderNode(treeRoot));
    overlay.appendChild(ul);
    overlay.style.display = 'block';
  }

  // Expose globally as WbsEngine
  window.WbsEngine = {
    detectNodeType: detectNodeType,
    buildTree: buildTree,
    showExplorer: showExplorer
  };

})(typeof window !== 'undefined' ? window : this);
