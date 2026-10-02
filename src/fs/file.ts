/**
 * File access, with a browser fallback.
 *
 * Under Tauri we get real open/save dialogs and can overwrite the file in
 * place. Running `pnpm dev` in a plain browser (handy for iterating on the
 * grid) falls back to a file input and a download, so the app still works —
 * it just cannot overwrite the original.
 */

export interface OpenedFile {
  /** Absolute path under Tauri; undefined in the browser (no overwrite possible). */
  path?: string;
  name: string;
  bytes: Uint8Array;
}

export interface RecentEntry {
  path: string;
  name: string;
  opened_at: number;
}

export const isTauri = typeof window !== 'undefined' && '__TAURI_INTERNALS__' in window;

/** Event the Rust side emits for files it was handed by the OS. */
const OPEN_FILES_EVENT = 'excel-editor://open-files';

const SPREADSHEET_FILTER = {
  name: '스프레드시트',
  extensions: ['xlsx', 'xlsm', 'csv', 'tsv'],
};
const XLSX_FILTER = { name: 'Excel 통합 문서', extensions: ['xlsx', 'xlsm'] };
const CSV_FILTER = { name: '구분자 텍스트', extensions: ['csv', 'tsv'] };

export function basename(path: string): string {
  const parts = path.split(/[/\\]/);
  return parts[parts.length - 1] || path;
}

/** True when the name looks like delimited text rather than a workbook. */
export function isDelimitedName(name: string): boolean {
  return /\.(csv|tsv|txt)$/i.test(name);
}

/* ----------------------------------------------------------------- browser */

function browserOpen(): Promise<OpenedFile | undefined> {
  return new Promise((resolve) => {
    const input = document.createElement('input');
    input.type = 'file';
    input.accept = '.xlsx,.xlsm,.csv,.tsv';
    input.onchange = async () => {
      const file = input.files?.[0];
      if (!file) return resolve(undefined);
      resolve({ name: file.name, bytes: new Uint8Array(await file.arrayBuffer()) });
    };
    // A cancelled picker fires no event in some browsers; `cancel` covers the rest.
    input.oncancel = () => resolve(undefined);
    input.click();
  });
}

function browserDownload(name: string, bytes: Uint8Array): void {
  const type = isDelimitedName(name)
    ? 'text/csv'
    : 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  const url = URL.createObjectURL(new Blob([bytes as BlobPart], { type }));
  const a = document.createElement('a');
  a.href = url;
  a.download = name;
  a.click();
  URL.revokeObjectURL(url);
}

/* -------------------------------------------------------------------- open */

export async function pickAndReadFile(): Promise<OpenedFile | undefined> {
  if (!isTauri) return browserOpen();

  const { open } = await import('@tauri-apps/plugin-dialog');
  const selected = await open({
    multiple: false,
    directory: false,
    filters: [SPREADSHEET_FILTER, XLSX_FILTER, CSV_FILTER],
  });
  if (typeof selected !== 'string') return undefined;

  const { readFile } = await import('@tauri-apps/plugin-fs');
  const file = { path: selected, name: basename(selected), bytes: await readFile(selected) };
  // The dialog put this path in scope; recording it lets it be reopened later.
  await recordRecent(selected);
  return file;
}

export async function readFileAt(path: string): Promise<Uint8Array> {
  if (!isTauri) throw new Error('브라우저에서는 경로로 파일을 읽을 수 없습니다.');
  const { readFile } = await import('@tauri-apps/plugin-fs');
  return readFile(path);
}

/* -------------------------------------------------------------------- save */

/** Writes to a known path. Returns false in the browser, where it downloads instead. */
export async function writeFileAt(path: string, bytes: Uint8Array): Promise<boolean> {
  if (!isTauri) {
    browserDownload(basename(path), bytes);
    return false;
  }
  const { writeFile } = await import('@tauri-apps/plugin-fs');
  await writeFile(path, bytes);
  return true;
}

/**
 * Asks where to save, without writing anything yet.
 *
 * The path has to be chosen before the bytes are built: picking `.csv` versus
 * `.xlsx` decides which serialiser runs, so building first would mean guessing.
 * Returns undefined when cancelled, or in the browser, which has no save dialog.
 */
export async function pickSavePath(suggestedName: string): Promise<string | undefined> {
  if (!isTauri) return undefined;

  const { save } = await import('@tauri-apps/plugin-dialog');
  const target = await save({
    defaultPath: suggestedName,
    filters: isDelimitedName(suggestedName)
      ? [CSV_FILTER, XLSX_FILTER]
      : [XLSX_FILTER, CSV_FILTER],
  });
  return target ?? undefined;
}

/** Records a path the user just saved to, so it shows up in the recent list. */
export async function noteSaved(path: string): Promise<void> {
  await recordRecent(path);
}

/** Browser-only: hands the bytes to the download machinery. */
export function downloadFile(name: string, bytes: Uint8Array): void {
  browserDownload(name, bytes);
}

/* ----------------------------------------------------------------- backups */

/**
 * Copies the file about to be overwritten to `name.bak.ext`.
 *
 * Runs in Rust because the backup lands on a path the user never picked, which
 * the filesystem scope would — correctly — refuse to the frontend.
 */
export async function backupFile(path: string): Promise<string | undefined> {
  if (!isTauri) return undefined;
  const { invoke } = await import('@tauri-apps/api/core');
  return invoke<string>('backup_file', { path });
}

/* ------------------------------------------------------------ recent files */

export async function recentList(): Promise<RecentEntry[]> {
  if (!isTauri) return [];
  const { invoke } = await import('@tauri-apps/api/core');
  return invoke<RecentEntry[]>('recent_list');
}

async function recordRecent(path: string): Promise<void> {
  if (!isTauri) return;
  const { invoke } = await import('@tauri-apps/api/core');
  // Rust refuses paths outside the current scope; a refusal is not worth
  // interrupting the open that just succeeded.
  await invoke('record_recent', { path }).catch(() => undefined);
}

/**
 * Re-grants filesystem access to a remembered path.
 *
 * Scope grants do not survive a restart, so reopening yesterday's file needs
 * this first. Rust only honours paths already in its own recent list.
 */
export async function grantRecent(path: string): Promise<void> {
  if (!isTauri) return;
  const { invoke } = await import('@tauri-apps/api/core');
  await invoke('grant_recent', { path });
}

export async function forgetRecent(path: string): Promise<void> {
  if (!isTauri) return;
  const { invoke } = await import('@tauri-apps/api/core');
  await invoke('forget_recent', { path });
}

export async function clearRecent(): Promise<void> {
  if (!isTauri) return;
  const { invoke } = await import('@tauri-apps/api/core');
  await invoke('clear_recent');
}

/* ----------------------------------------------- files handed over by the OS */

/**
 * Subscribes to open requests the app did not initiate: a Finder double-click,
 * or a file dropped onto the window. Rust has already put these paths in scope.
 */
export async function onOpenFiles(
  handler: (paths: string[]) => void,
): Promise<() => void> {
  if (!isTauri) return () => {};
  const { listen } = await import('@tauri-apps/api/event');
  const unlisten = await listen<string[]>(OPEN_FILES_EVENT, (event) => {
    if (Array.isArray(event.payload) && event.payload.length > 0) handler(event.payload);
  });
  return unlisten;
}

/**
 * Tells Rust the UI can handle open requests now, and collects any that arrived
 * before it could.
 *
 * Opening a file from Finder while the app is closed launches it, and the OS
 * request lands long before the webview is up — without this handover the file
 * is silently dropped and the app opens empty. Call only once the grid is
 * initialised and {@link onOpenFiles} is attached.
 */
export async function announceFrontendReady(): Promise<string[]> {
  if (!isTauri) return [];
  const { invoke } = await import('@tauri-apps/api/core');
  return invoke<string[]>('frontend_ready');
}

/* ------------------------------------------------------------------- window */

export async function confirmDialog(message: string, title: string): Promise<boolean> {
  if (!isTauri) return window.confirm(`${title}\n\n${message}`);
  const { confirm } = await import('@tauri-apps/plugin-dialog');
  return confirm(message, { title, kind: 'warning' });
}

export async function alertDialog(message: string, title: string): Promise<void> {
  if (!isTauri) {
    window.alert(`${title}\n\n${message}`);
    return;
  }
  const { message: show } = await import('@tauri-apps/plugin-dialog');
  await show(message, { title, kind: 'error' });
}

export async function setWindowTitle(title: string): Promise<void> {
  if (!isTauri) {
    document.title = title;
    return;
  }
  const { getCurrentWindow } = await import('@tauri-apps/api/window');
  await getCurrentWindow().setTitle(title);
}
