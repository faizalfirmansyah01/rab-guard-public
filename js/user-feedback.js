/**
 * RAB GUARD — User Feedback & Local Storage Manager
 * Module: UserFeedback (js/user-feedback.js)
 * 
 * Tanggung Jawab:
 * - Menyimpan evaluasi pengguna, laporan kesalahan pembacaan Excel,
 *   penandaan false positive, dan usulan aturan baru.
 * - Beroperasi 100% offline menggunakan localStorage (privasi terjamin, Rp0).
 * - Menyediakan ekspor sukarela dalam format JSON dan CSV.
 * - Tidak pernah mengunggah data atau berkas ke server tanpa persetujuan eksplisit.
 */

(function (window) {
  'use strict';

  var STORAGE_KEY = 'rab_guard_user_feedback_v1';

  function getStoredData() {
    try {
      var raw = localStorage.getItem(STORAGE_KEY);
      if (raw) {
        var parsed = JSON.parse(raw);
        if (parsed && typeof parsed === 'object') return parsed;
      }
    } catch (e) {
      console.warn('[UserFeedback] Gagal membaca localStorage:', e);
    }
    return {
      ratings: [],
      parsingIssues: [],
      falsePositives: {},
      featureProposals: [],
      createdAt: new Date().toISOString()
    };
  }

  function saveStoredData(data) {
    try {
      localStorage.setItem(STORAGE_KEY, JSON.stringify(data));
      return true;
    } catch (e) {
      console.error('[UserFeedback] Gagal menyimpan ke localStorage:', e);
      return false;
    }
  }

  var UserFeedback = {
    /**
     * Menyimpan penilaian kepuasan pengguna.
     * @param {number} stars - Skor 1-5
     * @param {string} comment - Catatan evaluasi
     */
    saveRating: function (stars, comment) {
      var data = getStoredData();
      var entry = {
        id: 'RAT-' + Date.now(),
        stars: Math.max(1, Math.min(5, parseInt(stars, 10) || 5)),
        comment: (comment || '').trim(),
        timestamp: new Date().toISOString()
      };
      data.ratings.push(entry);
      saveStoredData(data);
      return entry;
    },

    /**
     * Melaporkan pembacaan sel / struktur Excel yang keliru.
     * @param {string} sheetName - Nama lembar kerja
     * @param {string} cellRef - Sel (mis. F14)
     * @param {string} issueDescription - Deskripsi kesalahan
     * @param {string} expectedValue - Nilai atau perilaku yang diharapkan
     */
    reportParsingIssue: function (sheetName, cellRef, issueDescription, expectedValue) {
      var data = getStoredData();
      var entry = {
        id: 'ISSUE-' + Date.now(),
        sheetName: (sheetName || '').trim(),
        cellRef: (cellRef || '').trim(),
        issueDescription: (issueDescription || '').trim(),
        expectedValue: (expectedValue || '').trim(),
        timestamp: new Date().toISOString()
      };
      data.parsingIssues.push(entry);
      saveStoredData(data);
      return entry;
    },

    /**
     * Menandai temuan audit tertentu sebagai False Positive.
     * @param {string} findingId - ID temuan (mis. UF-1)
     * @param {string} ruleId - Aturan audit
     * @param {string} explanation - Uraian temuan
     * @param {string} userReason - Alasan pengguna menandai false positive
     */
    markFalsePositive: function (findingId, ruleId, explanation, userReason) {
      if (!findingId) return false;
      var data = getStoredData();
      data.falsePositives[findingId] = {
        findingId: findingId,
        ruleId: ruleId || '-',
        explanation: explanation || '-',
        userReason: (userReason || 'Ditinjau manual: tidak relevan / sengaja.').trim(),
        markedAt: new Date().toISOString()
      };
      saveStoredData(data);
      return true;
    },

    /**
     * Membatalkan penandaan False Positive.
     * @param {string} findingId
     */
    unmarkFalsePositive: function (findingId) {
      if (!findingId) return false;
      var data = getStoredData();
      if (data.falsePositives[findingId]) {
        delete data.falsePositives[findingId];
        saveStoredData(data);
        return true;
      }
      return false;
    },

    /**
     * Memeriksa apakah suatu temuan telah ditandai sebagai false positive.
     * @param {string} findingId
     * @returns {boolean}
     */
    isFalsePositive: function (findingId) {
      if (!findingId) return false;
      var data = getStoredData();
      return !!data.falsePositives[findingId];
    },

    /**
     * Mengambil daftar seluruh false positive yang ditandai.
     * @returns {Array}
     */
    getFalsePositivesList: function () {
      var data = getStoredData();
      var list = [];
      for (var k in data.falsePositives) {
        if (Object.prototype.hasOwnProperty.call(data.falsePositives, k)) {
          list.push(data.falsePositives[k]);
        }
      }
      return list;
    },

    /**
     * Mengajukan usulan aturan audit atau fitur baru.
     * @param {string} title
     * @param {string} description
     * @param {string} category - 'RULE' | 'PARSER' | 'SIMULATION' | 'UI'
     */
    submitProposal: function (title, description, category) {
      var data = getStoredData();
      var entry = {
        id: 'PROP-' + Date.now(),
        title: (title || '').trim(),
        description: (description || '').trim(),
        category: category || 'RULE',
        timestamp: new Date().toISOString()
      };
      data.featureProposals.push(entry);
      saveStoredData(data);
      return entry;
    },

    /**
     * Mengambil seluruh data feedback pengguna.
     * @returns {Object}
     */
    getAllFeedback: function () {
      return getStoredData();
    },

    /**
     * Mengunduh berkas feedback lengkap dalam format JSON.
     */
    exportJSON: function () {
      var data = getStoredData();
      var jsonStr = JSON.stringify(data, null, 2);
      var blob = new Blob([jsonStr], { type: 'application/json;charset=utf-8;' });
      var url = URL.createObjectURL(blob);
      var a = document.createElement('a');
      a.href = url;
      a.download = 'rab_guard_feedback_' + new Date().toISOString().slice(0, 10) + '.json';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    },

    /**
     * Mengunduh rekap masukan dalam format CSV.
     */
    exportCSV: function () {
      var data = getStoredData();
      var rows = [
        ['Kategori', 'ID', 'Timestamp', 'Detil 1', 'Detil 2', 'Keterangan / Alasan']
      ];

      (data.ratings || []).forEach(function (r) {
        rows.push(['Rating', r.id, r.timestamp, r.stars + ' Bintang', '', r.comment]);
      });

      (data.parsingIssues || []).forEach(function (p) {
        rows.push(['Parsing Issue', p.id, p.timestamp, p.sheetName, p.cellRef, p.issueDescription + ' (Harapan: ' + p.expectedValue + ')']);
      });

      var fpList = this.getFalsePositivesList();
      fpList.forEach(function (f) {
        rows.push(['False Positive', f.findingId, f.markedAt, f.ruleId, f.explanation, f.userReason]);
      });

      (data.featureProposals || []).forEach(function (pr) {
        rows.push(['Proposal', pr.id, pr.timestamp, pr.category, pr.title, pr.description]);
      });

      var csvContent = '\uFEFF' + rows.map(function (row) {
        return row.map(function (cell) {
          var str = String(cell || '').replace(/"/g, '""');
          return '"' + str + '"';
        }).join(',');
      }).join('\r\n');

      var blob = new Blob([csvContent], { type: 'text/csv;charset=utf-8;' });
      var url = URL.createObjectURL(blob);
      var a = document.createElement('a');
      a.href = url;
      a.download = 'rab_guard_feedback_' + new Date().toISOString().slice(0, 10) + '.csv';
      document.body.appendChild(a);
      a.click();
      document.body.removeChild(a);
      URL.revokeObjectURL(url);
    },

    /**
     * Membersihkan seluruh data masukan lokal.
     */
    clearAll: function () {
      localStorage.removeItem(STORAGE_KEY);
      return true;
    }
  };

  window.UserFeedback = UserFeedback;

})(typeof window !== 'undefined' ? window : this);
