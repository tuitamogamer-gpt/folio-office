import { chromium, expect } from '@playwright/test';

const browser = await chromium.launch({ headless: true, executablePath: '/usr/bin/chromium', args: ['--no-sandbox'] });
const origin = process.env.BASE_URL || 'http://localhost:5173';
const sample = (name = 'Saved work', content = '<p>Keep me</p>') => ({ id: 'test-file', name, kind: 'document', content, createdAt: 1, updatedAt: 2, starred: false, trashed: false });
async function blankPage(context) {
  const page = await context.newPage();
  await page.route('**/__workspace-storage-test', route => route.fulfill({ contentType: 'text/html', body: '<!doctype html><title>Workspace storage tests</title>' }));
  await page.goto(`${origin}/__workspace-storage-test`);
  return page;
}
async function dbRecord(page) {
  return page.evaluate(() => new Promise((resolve, reject) => {
    const request = indexedDB.open('folio-workspace-v1', 1);
    request.onerror = () => reject(request.error);
    request.onsuccess = () => { const db = request.result; const transaction = db.transaction('state', 'readonly'); const get = transaction.objectStore('state').get('files'); get.onsuccess = () => { resolve(get.result); db.close(); }; get.onerror = () => reject(get.error); };
  }));
}

try {
  const context = await browser.newContext();
  const page = await blankPage(context);
  const legacy = [sample('Legacy document')];
  await page.evaluate(files => localStorage.setItem('folio-files-v1', JSON.stringify(files)), legacy);
  expect(await page.evaluate(async () => (await import('/src/lib/workspaceStorage.ts')).loadWorkspace([]))).toEqual(legacy);
  expect((await dbRecord(page)).files).toEqual(legacy);
  console.log('PASS: legacy localStorage is migrated into a committed IndexedDB record.');

  const latest = [sample('Current document', '<p>Current draft</p>')];
  await page.evaluate(async files => { const storage = await import('/src/lib/workspaceStorage.ts'); await storage.saveWorkspace(files); }, latest);
  await page.evaluate(files => localStorage.setItem('folio-files-v1', JSON.stringify(files)), legacy);
  await page.reload();
  expect(await page.evaluate(async () => (await import('/src/lib/workspaceStorage.ts')).loadWorkspace([]))).toEqual(latest);
  console.log('PASS: save/reload prefers canonical storage over an older compatibility mirror.');

  const large = [sample('Photos and version history', `<p>${'A'.repeat(6 * 1024 * 1024)}</p>`)];
  large[0].versions = [{ id: 'history-1', name: 'Saved milestone', createdAt: 1, content: '<p>Original image caption</p>' }];
  await page.evaluate(async files => (await import('/src/lib/workspaceStorage.ts')).saveWorkspace(files), large);
  expect(await page.evaluate(() => localStorage.getItem('folio-files-v1'))).toBeNull();
  await page.reload();
  const loadedLarge = await page.evaluate(async () => { const files = await (await import('/src/lib/workspaceStorage.ts')).loadWorkspace([]); return { size: files[0].content.length, name: files[0].name, versions: files[0].versions }; });
  expect(loadedLarge.size).toBe(large[0].content.length); expect(loadedLarge.versions).toEqual(large[0].versions); expect(loadedLarge.name).toBe(large[0].name);
  console.log('PASS: a workspace larger than 5 MB, including history, survives reload without a stale mirror.');

  const queued = await page.evaluate(async () => {
    const storage = await import('/src/lib/workspaceStorage.ts');
    const first = { id: 'ordered', name: 'First', kind: 'document', content: 'first', createdAt: 1, updatedAt: 1, starred: false, trashed: false };
    const one = storage.saveWorkspace([first]); first.content = 'Mutated after call';
    const two = storage.saveWorkspace([{ ...first, name: 'Second', content: 'latest' }]);
    await Promise.all([one, two]); return storage.loadWorkspace([]);
  });
  expect(queued[0].name).toBe('Second'); expect(queued[0].content).toBe('latest');
  console.log('PASS: concurrent saves commit in call order and snapshot caller data.');

  const syncRecovery = await page.evaluate(async () => {
    const storage = await import('/src/lib/workspaceStorage.ts');
    const file = { id: 'ordered', name: 'Older queued snapshot', kind: 'document', content: 'old', createdAt: 1, updatedAt: 1, starred: false, trashed: false };
    const pending = storage.saveWorkspace([file]);
    const mirrored = storage.mirrorWorkspace([{ ...file, name: 'Last edit before closing', content: 'latest synchronous edit' }]);
    await pending;
    return { mirrored, pendingName: JSON.parse(localStorage.getItem('folio-workspace-recovery-v1')).files[0].name };
  });
  expect(syncRecovery).toEqual({ mirrored: true, pendingName: 'Last edit before closing' });
  await page.reload();
  const recovered = await page.evaluate(async () => (await import('/src/lib/workspaceStorage.ts')).loadWorkspace([]));
  expect(recovered[0].content).toBe('latest synchronous edit');
  expect((await dbRecord(page)).files[0].name).toBe('Last edit before closing');
  console.log('PASS: synchronous pagehide recovery survives an older queued commit and a full page reload.');

  const failedCanonical = await page.evaluate(async () => {
    const storage = await import('/src/lib/workspaceStorage.ts');
    const transaction = IDBDatabase.prototype.transaction;
    IDBDatabase.prototype.transaction = function (...args) { if (args[1] === 'readwrite') throw new DOMException('Simulated disk write failure', 'QuotaExceededError'); return transaction.apply(this, args); };
    try { await storage.saveWorkspace([{ id: 'fallback', name: 'Fallback after database failure', kind: 'document', content: 'recover me', createdAt: 1, updatedAt: 1, starred: false, trashed: false }]); }
    finally { IDBDatabase.prototype.transaction = transaction; }
    return JSON.parse(localStorage.getItem('folio-workspace-recovery-v1')).files[0].name;
  });
  expect(failedCanonical).toBe('Fallback after database failure');
  expect((await dbRecord(page)).files[0].name).toBe('Last edit before closing');
  await page.reload();
  expect((await page.evaluate(async () => (await import('/src/lib/workspaceStorage.ts')).loadWorkspace([])))[0].name).toBe('Fallback after database failure');
  expect((await dbRecord(page)).files[0].name).toBe('Fallback after database failure');
  console.log('PASS: newer local fallback supersedes stale canonical data after a database write failure.');
  await context.close();

  const fallbackContext = await browser.newContext();
  const fallbackPage = await blankPage(fallbackContext);
  await fallbackPage.evaluate(() => { window.__nativeIndexedDb = window.indexedDB; Object.defineProperty(window, 'indexedDB', { configurable: true, value: undefined }); });
  expect(await fallbackPage.evaluate(async files => { const storage = await import('/src/lib/workspaceStorage.ts'); await storage.saveWorkspace(files); return storage.loadWorkspace([]); }, latest)).toEqual(latest);
  const failure = await fallbackPage.evaluate(async () => {
    const storage = await import('/src/lib/workspaceStorage.ts'); const previous = localStorage.getItem('folio-files-v1');
    const originalSet = Storage.prototype.setItem;
    Storage.prototype.setItem = function () { throw new DOMException('Quota exceeded', 'QuotaExceededError'); };
    let rejected = false;
    try { await storage.saveWorkspace([{ id: 'new', name: 'Unsaved work', kind: 'document', content: 'new' }]); } catch { rejected = true; }
    const mirror = storage.mirrorWorkspace([]); Storage.prototype.setItem = originalSet;
    return { rejected, mirror, unchanged: localStorage.getItem('folio-files-v1') === previous };
  });
  expect(failure).toEqual({ rejected: true, mirror: false, unchanged: true });
  console.log('PASS: IndexedDB-unavailable fallback saves locally; dual write failures preserve the last good snapshot.');

  await fallbackPage.evaluate(() => Object.defineProperty(window, 'indexedDB', { configurable: true, value: window.__nativeIndexedDb }));
  await fallbackPage.evaluate(async files => (await import('/src/lib/workspaceStorage.ts')).saveWorkspace(files), [sample('Retry works')]);
  expect((await dbRecord(fallbackPage)).files[0].name).toBe('Retry works');
  console.log('PASS: a failed database open is retried after IndexedDB becomes available.');
  await fallbackContext.close();

  const readFailureContext = await browser.newContext();
  const readFailurePage = await blankPage(readFailureContext);
  const canonicalFiles = [sample('Canonical photos and history', `<p>${'B'.repeat(6 * 1024 * 1024)}</p>`)];
  canonicalFiles[0].versions = [{ id: 'retained-version', name: 'Original images', createdAt: 1, content: 'Saved milestone' }];
  await readFailurePage.evaluate(async files => (await import('/src/lib/workspaceStorage.ts')).saveWorkspace(files), canonicalFiles);
  await readFailurePage.evaluate(files => localStorage.setItem('folio-files-v1', JSON.stringify(files)), legacy);
  const readFailure = await readFailurePage.evaluate(async fallback => {
    const storage = await import('/src/lib/workspaceStorage.ts');
    const transaction = IDBDatabase.prototype.transaction;
    IDBDatabase.prototype.transaction = function (...args) { if (args[1] === 'readonly') throw new DOMException('Temporary read error', 'UnknownError'); return transaction.apply(this, args); };
    let loadRejected = false; let saveRejected = false;
    try { await storage.loadWorkspace(fallback); } catch { loadRejected = true; }
    finally { IDBDatabase.prototype.transaction = transaction; }
    try { await storage.saveWorkspace(fallback); } catch { saveRejected = true; }
    return { loadRejected, saveRejected, mirrorBlocked: !storage.mirrorWorkspace(fallback) };
  }, [sample('Default starter document')]);
  expect(readFailure).toEqual({ loadRejected: true, saveRejected: true, mirrorBlocked: true });
  const stillCanonical = await dbRecord(readFailurePage);
  expect(stillCanonical.files[0].name).toBe(canonicalFiles[0].name);
  expect(stillCanonical.files[0].content.length).toBe(canonicalFiles[0].content.length);
  const retriedLoad = await readFailurePage.evaluate(async () => { const files = await (await import('/src/lib/workspaceStorage.ts')).loadWorkspace([]); return { name: files[0].name, size: files[0].content.length, versions: files[0].versions }; });
  expect(retriedLoad).toEqual({ name: canonicalFiles[0].name, size: canonicalFiles[0].content.length, versions: canonicalFiles[0].versions });
  console.log('PASS: transient canonical read failure rejects startup and all writes; retry preserves large files over stale local/default data.');

  await readFailurePage.reload();
  const missingKnownDatabase = await readFailurePage.evaluate(async () => {
    const original = window.indexedDB; Object.defineProperty(window, 'indexedDB', { configurable: true, value: undefined });
    let rejected = false;
    try { await (await import('/src/lib/workspaceStorage.ts')).loadWorkspace([]); } catch { rejected = true; }
    finally { Object.defineProperty(window, 'indexedDB', { configurable: true, value: original }); }
    return rejected;
  });
  expect(missingKnownDatabase).toBe(true);
  expect((await dbRecord(readFailurePage)).files[0].name).toBe(canonicalFiles[0].name);
  await readFailureContext.close();

  const emptyContext = await browser.newContext();
  const emptyPage = await blankPage(emptyContext);
  expect(await emptyPage.evaluate(async files => (await import('/src/lib/workspaceStorage.ts')).loadWorkspace(files), latest)).toEqual(latest);
  await emptyPage.evaluate(async () => (await import('/src/lib/workspaceStorage.ts')).saveWorkspace([]));
  expect(await emptyPage.evaluate(async files => (await import('/src/lib/workspaceStorage.ts')).loadWorkspace(files), latest)).toEqual([]);
  console.log('PASS: confirmed missing storage initializes normally, and an intentionally empty workspace stays empty.');
  await emptyContext.close();

  const disabledContext = await browser.newContext();
  const disabledPage = await blankPage(disabledContext);
  const disabledFallback = await disabledPage.evaluate(async files => {
    IDBFactory.prototype.open = function () { throw new DOMException('IndexedDB disabled by browser settings', 'SecurityError'); };
    const storage = await import('/src/lib/workspaceStorage.ts');
    const loaded = await storage.loadWorkspace(files); await storage.saveWorkspace(loaded); return storage.loadWorkspace([]);
  }, latest);
  expect(disabledFallback).toEqual(latest);
  console.log('PASS: browsers explicitly disabling IndexedDB still support localStorage-only workspaces.');
  await disabledContext.close();
} finally { await browser.close(); }
