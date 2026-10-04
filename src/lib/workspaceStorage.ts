import type { OfficeFile } from '../types';

const DATABASE = 'folio-workspace-v1';
const STORE = 'state';
const FILES_KEY = 'files';
const LEGACY_KEY = 'folio-files-v1';
const RECOVERY_KEY = 'folio-workspace-recovery-v1';
const CANONICAL_MARKER_KEY = 'folio-workspace-canonical-v1';
const MAX_MIRROR_CHARACTERS = 2_000_000;
type SavedWorkspace = { files: OfficeFile[]; savedAt: number };
let databasePromise: Promise<IDBDatabase> | undefined;
let saveQueue: Promise<void> = Promise.resolve();
let lastRevision = 0;
let hydrationFailure: Error | undefined;
const revision = () => (lastRevision = Math.max(Date.now(), lastRevision + 1));
class IndexedDBUnavailableError extends Error {}
function databaseOpenError(error: unknown) {
  return error instanceof DOMException && ['SecurityError', 'NotSupportedError'].includes(error.name) ? new IndexedDBUnavailableError('This browser does not permit IndexedDB storage.', { cause: error }) : error;
}
function markCanonical() { try { globalThis.localStorage.setItem(CANONICAL_MARKER_KEY, '1'); } catch { /* The database itself remains canonical. */ } }
function hasCanonicalMarker() { try { return globalThis.localStorage.getItem(CANONICAL_MARKER_KEY) === '1'; } catch { return true; } }

function validFiles(value: unknown): value is OfficeFile[] {
  return Array.isArray(value) && value.every(file => file && typeof file === 'object' && typeof file.id === 'string' && typeof file.name === 'string' && ['document', 'spreadsheet', 'presentation'].includes(file.kind) && Object.hasOwn(file, 'content'));
}
function filesFrom(value: unknown): OfficeFile[] | undefined {
  if (validFiles(value)) return value;
  if (value && typeof value === 'object' && 'files' in value && validFiles(value.files)) return value.files;
}
function savedFrom(value: unknown): SavedWorkspace | undefined {
  const files = filesFrom(value);
  if (!files) return;
  const savedAt = value && typeof value === 'object' && 'savedAt' in value && typeof value.savedAt === 'number' && Number.isFinite(value.savedAt) ? value.savedAt : 0;
  lastRevision = Math.max(lastRevision, savedAt);
  return { files, savedAt };
}
function readLocal(key: string): SavedWorkspace | undefined {
  try { const value = globalThis.localStorage.getItem(key); return value === null ? undefined : savedFrom(JSON.parse(value)); } catch { return; }
}
function openDatabase(): Promise<IDBDatabase> {
  if (databasePromise) return databasePromise;
  const opening = new Promise<IDBDatabase>((resolve, reject) => {
    let settled = false;
    let request: IDBOpenDBRequest;
    try {
      if (!globalThis.indexedDB) throw new IndexedDBUnavailableError('IndexedDB is unavailable.');
      request = globalThis.indexedDB.open(DATABASE, 1);
    } catch (error) { reject(databaseOpenError(error)); return; }
    const timeout = setTimeout(() => finish(new Error('Opening workspace storage timed out.')), 4000);
    function finish(error: unknown) { if (settled) return; settled = true; clearTimeout(timeout); reject(error); }
    request.onupgradeneeded = () => { if (!request.result.objectStoreNames.contains(STORE)) request.result.createObjectStore(STORE); };
    request.onerror = () => finish(databaseOpenError(request.error || new Error('Workspace storage could not be opened.')));
    request.onblocked = () => finish(new Error('Another browser tab is blocking workspace storage.'));
    request.onsuccess = () => {
      if (settled) { request.result.close(); return; }
      settled = true; clearTimeout(timeout);
      const database = request.result;
      database.onversionchange = () => { database.close(); databasePromise = undefined; };
      database.onclose = () => { databasePromise = undefined; };
      resolve(database);
    };
  });
  databasePromise = opening;
  opening.catch(() => { if (databasePromise === opening) databasePromise = undefined; });
  return opening;
}
async function readDatabase(): Promise<unknown> {
  const database = await openDatabase();
  return new Promise((resolve, reject) => {
    try {
      const transaction = database.transaction(STORE, 'readonly');
      const request = transaction.objectStore(STORE).get(FILES_KEY);
      transaction.oncomplete = () => resolve(request.result);
      transaction.onabort = () => reject(transaction.error || request.error || new Error('Workspace storage could not be read.'));
      transaction.onerror = () => { /* Aborting the transaction reports the complete failure. */ };
    } catch (error) { databasePromise = undefined; reject(error); }
  });
}
async function writeDatabase(snapshot: SavedWorkspace): Promise<void> {
  const database = await openDatabase();
  return new Promise((resolve, reject) => {
    let transaction: IDBTransaction | undefined;
    try {
      try { transaction = database.transaction(STORE, 'readwrite', { durability: 'strict' }); }
      catch (error) { if (error instanceof TypeError) transaction = database.transaction(STORE, 'readwrite'); else throw error; }
      transaction.oncomplete = () => resolve();
      transaction.onabort = () => reject(transaction?.error || new Error('Workspace changes could not be saved.'));
      transaction.onerror = () => { /* Wait for abort: request success alone does not imply a committed write. */ };
      transaction.objectStore(STORE).put(snapshot, FILES_KEY);
    } catch (error) {
      try { transaction?.abort(); } catch { /* The last committed record is retained. */ }
      if (error instanceof DOMException && error.name === 'InvalidStateError') databasePromise = undefined;
      reject(error);
    }
  });
}
function writeLegacy(files: OfficeFile[], mirrorOnly = false): void {
  const serialized = JSON.stringify(files);
  if (mirrorOnly && serialized.length > MAX_MIRROR_CHARACTERS) throw new Error('Workspace is too large for the compatibility mirror.');
  globalThis.localStorage.setItem(LEGACY_KEY, serialized);
}

function writeRecovery(snapshot: SavedWorkspace, mirrorOnly: boolean): void {
  const existing = readLocal(RECOVERY_KEY);
  if (existing && existing.savedAt > snapshot.savedAt) return;
  const serialized = JSON.stringify(snapshot);
  if (mirrorOnly && serialized.length > MAX_MIRROR_CHARACTERS) throw new Error('Workspace is too large for synchronous recovery storage.');
  // One atomic storage operation keeps recovery data and its revision together.
  globalThis.localStorage.setItem(RECOVERY_KEY, serialized);
  try { writeLegacy(snapshot.files, true); } catch { /* The complete recovery snapshot is already durable. */ }
}

/** A small synchronous safety copy for pagehide. A failed mirror never deletes the last good copy. */
export function mirrorWorkspace(files: OfficeFile[]): boolean {
  if (hydrationFailure) return false;
  try { writeRecovery({ files, savedAt: revision() }, true); return true; } catch { return false; }
}

/** A read failure is not an empty workspace: callers must keep editing/autosave gated and offer retry. */
export async function loadWorkspace(fallback: OfficeFile[]): Promise<OfficeFile[]> {
  let canonical: SavedWorkspace | undefined;
  try {
    const stored = await readDatabase();
    canonical = savedFrom(stored);
    if (stored !== undefined && !canonical) throw new Error('The saved workspace could not be read.');
    if (canonical) markCanonical();
  } catch (error) {
    // Missing browser support is a supported localStorage-only mode. An error opening or
    // reading an available database may hide existing work, so never seed or migrate over it.
    if (!(error instanceof IndexedDBUnavailableError) || hasCanonicalMarker()) {
      hydrationFailure = new Error('Your saved workspace could not be opened. Retry to load it safely; your existing files have not been changed.', { cause: error });
      throw hydrationFailure;
    }
  }
  hydrationFailure = undefined;
  const recovery = readLocal(RECOVERY_KEY);
  if (canonical && (!recovery || recovery.savedAt <= canonical.savedAt)) return canonical.files;
  const local = recovery || readLocal(LEGACY_KEY);
  if (local) {
    try { await saveWorkspace(local.files); } catch { /* Reading existing work remains possible even when the device is full. */ }
    return local.files;
  }
  return structuredClone(fallback);
}

/** Serialize writes and resolve only when at least one persistent store has committed the snapshot. */
export function saveWorkspace(files: OfficeFile[]): Promise<void> {
  if (hydrationFailure) return Promise.reject(hydrationFailure);
  let snapshot: SavedWorkspace;
  try { snapshot = { files: structuredClone(files), savedAt: revision() }; } catch (error) { return Promise.reject(error); }
  const save = saveQueue.catch(() => {}).then(async () => {
    let databaseError: unknown;
    try {
      await writeDatabase(snapshot);
      markCanonical();
      const recovery = readLocal(RECOVERY_KEY);
      if (!recovery || recovery.savedAt <= snapshot.savedAt) {
        // Never clear a newer pagehide snapshot when an older queued database write completes.
        try { globalThis.localStorage.removeItem(RECOVERY_KEY); } catch { /* Canonical storage is already durable. */ }
        try { writeLegacy(snapshot.files, true); }
        catch {
          // Remove stale compatibility data only after the canonical replacement has committed.
          try { globalThis.localStorage.removeItem(LEGACY_KEY); } catch { /* Canonical storage is already durable. */ }
        }
      }
      markCanonical();
      return;
    } catch (error) { databaseError = error; }
    try { writeRecovery(snapshot, false); }
    catch (legacyError) { throw new AggregateError([databaseError, legacyError], 'Your browser could not save this workspace. Download a backup to keep your latest changes.'); }
  });
  saveQueue = save;
  return save;
}
