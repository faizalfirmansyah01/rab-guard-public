/*
 * RAB GUARD - Project Storage (IndexedDB)
 * Database: RABGuardDB
 * Object stores:
 *   - projects
 *   - documents
 *   - workbookFiles (original Excel Blob)
 */

(function (window) {
  'use strict';

  const DB_NAME = 'RABGuardDB';
  const DB_VERSION = 4;

  const PROJECT_STORE = 'projects';
  const DOC_STORE = 'documents';
  const FILE_STORE = 'workbookFiles';

  let db = null;

  function upgradeSchema(database, transaction) {
    function objectStore(name, options) {
      return database.objectStoreNames.contains(name)
        ? transaction.objectStore(name)
        : database.createObjectStore(name, options);
    }
    const projects = objectStore(PROJECT_STORE, { keyPath: 'projectId' });
    if (!projects.indexNames.contains('projectName')) projects.createIndex('projectName', 'projectName', { unique: false });
    const documents = objectStore(DOC_STORE, { keyPath: 'documentId' });
    if (!documents.indexNames.contains('projectId')) documents.createIndex('projectId', 'projectId', { unique: false });
    objectStore(FILE_STORE, { keyPath: 'documentId' });
  }

  function initDB() {
    return new Promise((resolve, reject) => {
      if (db) return resolve(db);

      const request = indexedDB.open(DB_NAME, DB_VERSION);

      request.onupgradeneeded = function (e) {
        const database = e.target.result;
        upgradeSchema(database, e.target.transaction);
      };

      request.onsuccess = function (e) {
        db = e.target.result;

        db.onversionchange = function () {
          db.close();
          db = null;
        };

        resolve(db);
      };

      request.onerror = function (e) {
        reject(e.target.error);
      };

      request.onblocked = function () {
        reject(new Error(
          'Database upgrade blocked. Close other RAB Guard tabs and retry.'
        ));
      };
    });
  }

  function _tx(storeName, mode = 'readonly') {
    return initDB().then(database => {
      const tx = database.transaction(storeName, mode);
      return tx.objectStore(storeName);
    });
  }

  // --- Project operations ---

  function addProject(project) {
    return initDB().then(database => new Promise((resolve, reject) => {
      const now = new Date().toISOString();
      const record = Object.assign({}, project, { createdAt: now, updatedAt: now, projectVersion: 1, documents: [] });
      const tx = database.transaction(PROJECT_STORE, 'readwrite'), req = tx.objectStore(PROJECT_STORE).add(record);
      let requestError = null;
      req.onerror = event => { requestError = event.target.error; };
      tx.oncomplete = () => resolve(record);
      tx.onerror = () => reject(requestError || tx.error || new Error('Project create transaction failed.'));
      tx.onabort = () => reject(requestError || tx.error || new Error('Project create transaction aborted.'));
    }));
  }

  function getAllProjects() {
    return _tx(PROJECT_STORE).then(store => {
      const req = store.getAll();

      return new Promise((resolve, reject) => {
        req.onsuccess = () => resolve(req.result);
        req.onerror = e => reject(e.target.error);
      });
    });
  }

  function updateProject(project) {
    if (!project || !project.projectId) return Promise.reject(new Error('Project ID is required for update.'));
    return initDB().then(database => new Promise((resolve, reject) => {
      const tx = database.transaction(PROJECT_STORE, 'readwrite'), store = tx.objectStore(PROJECT_STORE), request = store.get(project.projectId);
      let updateError = null, updatedRecord = null;
      request.onsuccess = function () {
        const existing = request.result;
        if (!existing) { updateError = new Error('Project not found; update cannot create a project.'); tx.abort(); return; }
        const currentVersion = Number.isInteger(existing.projectVersion) ? existing.projectVersion : 1;
        const expectedVersion = Number.isInteger(project.projectVersion) ? project.projectVersion : 1;
        if (expectedVersion !== currentVersion) { updateError = new Error('Project version conflict; reload the project before updating.'); tx.abort(); return; }
        updatedRecord = Object.assign({}, existing, project, { projectVersion: currentVersion + 1, updatedAt: new Date().toISOString() });
        store.put(updatedRecord);
      };
      request.onerror = function () { updateError = request.error; };
      tx.oncomplete = () => resolve(updatedRecord);
      tx.onerror = () => reject(updateError || tx.error || new Error('Project update transaction failed.'));
      tx.onabort = () => reject(updateError || tx.error || new Error('Project update transaction aborted.'));
    }));
  }

  /*
   * Delete project, its documents, and original workbook Blobs
   * in ONE IndexedDB transaction.
   */
  function deleteProject(projectId) {
    return initDB().then(database => new Promise((resolve, reject) => {
      const tx = database.transaction(
        [PROJECT_STORE, DOC_STORE, FILE_STORE],
        'readwrite'
      );

      const projectStore = tx.objectStore(PROJECT_STORE);
      const docStore = tx.objectStore(DOC_STORE);
      const fileStore = tx.objectStore(FILE_STORE);

      projectStore.delete(projectId);

      const index = docStore.index('projectId');
      const cursorReq = index.openCursor(
        IDBKeyRange.only(projectId)
      );

      cursorReq.onsuccess = function (e) {
        const cursor = e.target.result;

        if (cursor) {
          const documentId = cursor.primaryKey;

          cursor.delete();
          fileStore.delete(documentId);

          cursor.continue();
        }
      };

      tx.oncomplete = () => resolve();

      tx.onerror = () => reject(
        tx.error || new Error('Project deletion failed.')
      );

      tx.onabort = () => reject(
        tx.error || new Error('Project deletion aborted.')
      );
    }));
  }

  // --- Document operations ---

  function addDocument(document) {
    return _tx(DOC_STORE, 'readwrite').then(store => {
      const now = new Date().toISOString();

      document.importedAt = now;
      document.updatedAt = now;

      const req = store.add(document);

      return new Promise((resolve, reject) => {
        req.onsuccess = () => resolve(document);
        req.onerror = e => reject(e.target.error);
      });
    });
  }

  /*
   * Atomically save document metadata and original workbook Blob.
   * Either both records are committed or neither is.
   */
  function addDocumentWithFile(document, file) {
    if (!file || typeof file.slice !== 'function') {
      return Promise.reject(
        new Error('Invalid workbook file.')
      );
    }

    return initDB().then(database => new Promise((resolve, reject) => {
      const tx = database.transaction(
        [DOC_STORE, FILE_STORE],
        'readwrite'
      );

      const docStore = tx.objectStore(DOC_STORE);
      const fileStore = tx.objectStore(FILE_STORE);

      const now = new Date().toISOString();

      document.importedAt = now;
      document.updatedAt = now;

      const blob = file.slice(
        0,
        file.size,
        file.type || document.fileType || ''
      );

      docStore.add(document);

      fileStore.add({
        documentId: document.documentId,
        projectId: document.projectId,
        fileName: document.fileName,
        fileType: document.fileType,
        fileSize: file.size,
        blob: blob,
        savedAt: now
      });

      tx.oncomplete = () => resolve(document);

      tx.onerror = () => reject(
        tx.error || new Error('Document and workbook save failed.')
      );

      tx.onabort = () => reject(
        tx.error || new Error('Document save transaction aborted.')
      );
    }));
  }

  /* Atomically add a set of role-bound workbook source documents and Blobs. */
  function addSourceDocumentsWithFiles(entries) {
    if (!Array.isArray(entries) || entries.some(entry => !entry || !entry.document || !entry.file || typeof entry.file.slice !== 'function')) {
      return Promise.reject(new Error('Invalid source document batch.'));
    }
    const ids = new Set(entries.map(entry => entry.document.documentId));
    if (ids.size !== entries.length) return Promise.reject(new Error('Duplicate source document ID in batch.'));
    if (!entries.length) return Promise.resolve([]);
    return initDB().then(database => new Promise((resolve, reject) => {
      const tx = database.transaction([DOC_STORE, FILE_STORE], 'readwrite');
      const docs = tx.objectStore(DOC_STORE), files = tx.objectStore(FILE_STORE), now = new Date().toISOString();
      entries.forEach(entry => {
        const doc = entry.document, file = entry.file;
        doc.importedAt = now; doc.updatedAt = now;
        docs.add(doc);
        files.add({ documentId: doc.documentId, projectId: doc.projectId, fileName: doc.fileName, fileType: doc.fileType,
          fileSize: file.size, blob: file.slice(0, file.size, file.type || doc.fileType || ''), savedAt: now });
      });
      tx.oncomplete = () => resolve(entries.map(entry => entry.document));
      tx.onerror = () => reject(tx.error || new Error('Source workbook batch save failed.'));
      tx.onabort = () => reject(tx.error || new Error('Source workbook batch save transaction aborted.'));
    }));
  }

  function getDocumentsByProject(projectId) {
    return _tx(DOC_STORE).then(store => {
      const index = store.index('projectId');
      const range = IDBKeyRange.only(projectId);
      const req = index.getAll(range);

      return new Promise((resolve, reject) => {
        req.onsuccess = () => resolve(req.result);
        req.onerror = e => reject(e.target.error);
      });
    });
  }

  function getDocument(documentId) {
    return _tx(DOC_STORE).then(store => {
      const req = store.get(documentId);

      return new Promise((resolve, reject) => {
        req.onsuccess = () => resolve(req.result);
        req.onerror = e => reject(e.target.error);
      });
    });
  }

  function updateDocument(document) {
    return _tx(DOC_STORE, 'readwrite').then(store => {
      document.updatedAt = new Date().toISOString();

      const req = store.put(document);

      return new Promise((resolve, reject) => {
        req.onsuccess = () => resolve(document);
        req.onerror = e => reject(e.target.error);
      });
    });
  }

  /*
   * Retrieve original workbook Blob.
   * Returns Blob or null if unavailable.
   */
  function getWorkbookFile(documentId) {
    return _tx(FILE_STORE).then(store => {
      const req = store.get(documentId);

      return new Promise((resolve, reject) => {
        req.onsuccess = () => {
          resolve(req.result ? req.result.blob : null);
        };

        req.onerror = e => reject(e.target.error);
      });
    });
  }

  function deleteWorkbookFile(documentId) {
    return _tx(FILE_STORE, 'readwrite').then(store => {
      const req = store.delete(documentId);

      return new Promise((resolve, reject) => {
        req.onsuccess = () => resolve();
        req.onerror = e => reject(e.target.error);
      });
    });
  }

  /*
   * Delete document metadata and original Blob atomically.
   */
  function deleteDocument(documentId) {
    return initDB().then(database => new Promise((resolve, reject) => {
      const tx = database.transaction(
        [DOC_STORE, FILE_STORE],
        'readwrite'
      );

      tx.objectStore(DOC_STORE).delete(documentId);
      tx.objectStore(FILE_STORE).delete(documentId);

      tx.oncomplete = () => resolve();

      tx.onerror = () => reject(
        tx.error || new Error('Document deletion failed.')
      );

      tx.onabort = () => reject(
        tx.error || new Error('Document deletion aborted.')
      );
    }));
  }

  // --- Public API ---

  window.ProjectStorage = {
    initDB,
    upgradeSchema,

    addProject,
    getAllProjects,
    updateProject,
    deleteProject,

    addDocument,
    addDocumentWithFile,
    addSourceDocumentsWithFiles,
    getDocumentsByProject,
    getDocument,
    updateDocument,
    deleteDocument,

    getWorkbookFile,
    deleteWorkbookFile
  };

})(typeof window !== 'undefined' ? window : this);
