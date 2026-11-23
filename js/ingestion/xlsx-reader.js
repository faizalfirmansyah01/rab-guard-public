/**
 * RAB GUARD — Universal XLSX Ingestion
 * Module: XlsxReader (js/ingestion/xlsx-reader.js)
 * 
 * Tanggung Jawab:
 * - Membaca berkas biner/ArrayBuffer dari File input atau Drag-and-Drop.
 * - Memanfaatkan SheetJS (XLSX) yang dibundel secara lokal (offline murni).
 * - Menangani validasi ukuran, format .xlsx, dan error parsing secara graceful.
 */

(function (window) {
  'use strict';

  var XlsxReader = {
    /**
     * Memeriksa ketersediaan pustaka SheetJS lokal.
     * @returns {boolean}
     */
    isAvailable: function () {
      return typeof window.XLSX !== 'undefined' && typeof window.XLSX.read === 'function';
    },

    /**
     * Memvalidasi berkas sebelum dibaca.
     * @param {File} file
     * @param {Object} [options]
     * @returns {{ valid: boolean, error?: string }}
     */
    validateFile: function (file, options) {
      options = options || {};
      var maxSizeBytes = options.maxSizeBytes || 50 * 1024 * 1024; // 50 MB default

      if (!file) {
        return { valid: false, error: 'Berkas tidak ditemukan atau belum dipilih.' };
      }

      var name = file.name || '';
      var ext = name.substring(name.lastIndexOf('.')).toLowerCase();
      var validExts = ['.xlsx', '.xlsm', '.xls', '.xlsb'];

      if (validExts.indexOf(ext) === -1) {
        return {
          valid: false,
          error: 'Format berkas tidak didukung (' + ext + '). Harap unggah berkas spreadsheet Excel (.xlsx).'
        };
      }

      if (file.size > maxSizeBytes) {
        var mbSize = (file.size / (1024 * 1024)).toFixed(2);
        var maxMb = (maxSizeBytes / (1024 * 1024)).toFixed(0);
        return {
          valid: false,
          error: 'Ukuran berkas (' + mbSize + ' MB) melebihi batas maksimum (' + maxMb + ' MB).'
        };
      }

      return { valid: true };
    },

    /**
     * Membaca berkas File (Async via FileReader).
     * @param {File} file
     * @param {Object} [readOptions]
     * @returns {Promise<Object>}
     */
    readFileAsync: function (file, readOptions) {
      var self = this;
      return new Promise(function (resolve, reject) {
        var val = self.validateFile(file);
        if (!val.valid) {
          return reject(new Error(val.error));
        }

        if (!self.isAvailable()) {
          return reject(new Error('Pustaka SheetJS lokal (xlsx.full.min.js) belum termuat. Pastikan berkas tersedia di folder js/.'));
        }

        var reader = new FileReader();

        reader.onload = function (e) {
          try {
            var data = new Uint8Array(e.target.result);
            var workbook = self.readUint8Array(data, file.name, file.size, readOptions);
            resolve(workbook);
          } catch (err) {
            reject(new Error('Gagal memproses struktur berkas Excel: ' + (err.message || err)));
          }
        };

        reader.onerror = function () {
          reject(new Error('Terjadi kesalahan peramban saat membaca berkas fisik.'));
        };

        reader.readAsArrayBuffer(file);
      });
    },

    /**
     * Memproses byte array menjadi struktur workbook universal.
     * @param {Uint8Array} data
     * @param {string} filename
     * @param {number} fileSize
     * @param {Object} [customOptions]
     * @returns {Object}
     */
    readUint8Array: function (data, filename, fileSize, customOptions) {
      if (!this.isAvailable()) {
        throw new Error('SheetJS XLSX parser tidak tersedia.');
      }

      var defaultOpts = {
        type: 'array',
        cellFormula: true,
        cellStyles: true,
        cellDates: true,
        cellNF: true,
        dense: false
      };

      var opts = Object.assign({}, defaultOpts, customOptions || {});
      var rawWb = window.XLSX.read(data, opts);

      return {
        filename: filename || 'workbook.xlsx',
        fileSize: fileSize || data.byteLength || 0,
        sheetNames: rawWb.SheetNames || [],
        sheets: rawWb.Sheets || {},
        rawWorkbook: rawWb,
        readAt: new Date().toISOString()
      };
    }
  };

  window.XlsxReader = XlsxReader;

})(typeof window !== 'undefined' ? window : this);
