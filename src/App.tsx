import { useCallback, useEffect, useRef, useState } from 'react';
import { CommandType, LocaleType, createUniver, defaultTheme, merge } from '@univerjs/presets';
import type { FUniver, IWorkbookData } from '@univerjs/presets';
import { UniverSheetsCorePreset } from '@univerjs/preset-sheets-core';
import sheetsCoreKoKR from '@univerjs/preset-sheets-core/locales/ko-KR';
import '@univerjs/preset-sheets-core/lib/index.css';

import { UniverSheetsFindReplacePreset } from '@univerjs/preset-sheets-find-replace';
import findReplaceKoKR from '@univerjs/preset-sheets-find-replace/locales/ko-KR';
import '@univerjs/preset-sheets-find-replace/lib/index.css';

import { UniverSheetsFilterPreset } from '@univerjs/preset-sheets-filter';
import filterKoKR from '@univerjs/preset-sheets-filter/locales/ko-KR';
import '@univerjs/preset-sheets-filter/lib/index.css';

import { UniverSheetsSortPreset } from '@univerjs/preset-sheets-sort';
import sortKoKR from '@univerjs/preset-sheets-sort/locales/ko-KR';
import '@univerjs/preset-sheets-sort/lib/index.css';

import { createEmptySnapshot } from './xlsx/snapshot';
import {
  alertDialog,
  announceFrontendReady,
  backupFile,
  basename,
  confirmDialog,
  downloadFile,
  forgetRecent,
  grantRecent,
  isDelimitedName,
  isTauri,
  onOpenFiles,
  noteSaved,
  pickAndReadFile,
  pickSavePath,
  readFileAt,
  recentList,
  setWindowTitle,
  writeFileAt,
} from './fs/file';
import type { OpenedFile, RecentEntry } from './fs/file';
import './App.css';

const APP_NAME = 'Excel Editor';
const BACKUP_PREFERENCE_KEY = 'excel-editor.backup-on-save';

/**
 * The parsers are about a megabyte and are only needed once a file is opened or
 * saved, so they stay out of the startup bundle.
 */
const loadXlsx = () =>
  Promise.all([import('./xlsx/import'), import('./xlsx/export'), import('./xlsx/inspect')]);
const loadCsv = () => import('./csv/convert');


/**
 * Formula engine lifecycle mutations.
 *
 * Recalculation writes its results back as ordinary `set-range-values`
 * mutations, which are indistinguishable from a user edit by command id. Opening
 * a workbook with formulas therefore looks like an immediate edit and the
 * document would show as unsaved before the user touched anything. The engine
 * brackets every pass with a start and a stop, so dirty tracking ignores
 * whatever happens in between.
 */
const CALC_START = 'formula.mutation.set-formula-calculation-start';
const CALC_STOP = 'formula.mutation.set-formula-calculation-stop';
const CALC_LIFECYCLE = new Set([
  CALC_START,
  CALC_STOP,
  'formula.mutation.set-formula-calculation-result',
  'formula.mutation.set-formula-calculation-notification',
]);

/**
 * Grace period after loading a workbook, covering mutations that land before the
 * engine announces its first pass.
 */
const LOAD_SETTLE_MS = 1000;

type Format = 'xlsx' | 'csv';

interface DocState {
  /** Absolute path, when we can overwrite in place. */
  path?: string;
  name: string;
  format: Format;
  /** Delimiter the file was read with, so a CSV saves back the same shape. */
  delimiter?: string;
  /** Encoding detected on open, shown so a CP949 source is not a surprise. */
  encoding?: string;
  /** What a save of this document would discard. */
  warning?: string;
  /** Set once the user has acknowledged `warning` for this document. */
  warningAcknowledged: boolean;
}

const UNTITLED: DocState = { name: '제목 없음', format: 'xlsx', warningAcknowledged: true };

function readBackupPreference(): boolean {
  try {
    // Default on: saving is lossy for anything the converter does not model, so
    // a copy of the original is the cheap insurance.
    return localStorage.getItem(BACKUP_PREFERENCE_KEY) !== 'off';
  } catch {
    return true;
  }
}

export default function App() {
  const containerRef = useRef<HTMLDivElement>(null);
  const apiRef = useRef<FUniver | null>(null);
  /** Univer unit id of the workbook currently on screen. */
  const unitIdRef = useRef<string | null>(null);
  /** Disposes the dirty-tracking listener of the previous workbook. */
  const listenerRef = useRef<{ dispose: () => void } | null>(null);
  /** The formula worker, so it can be terminated with the Univer instance. */
  const workerRef = useRef<Worker | null>(null);
  /** True while the formula engine is mid-recalculation. */
  const calculatingRef = useRef(false);
  /** Timestamp of the last workbook load, for the settle window. */
  const loadedAtRef = useRef(0);

  const [doc, setDoc] = useState<DocState>(UNTITLED);
  const [dirty, setDirty] = useState(false);
  const [busy, setBusy] = useState<string | null>(null);
  const [ready, setReady] = useState(false);
  const [recent, setRecent] = useState<RecentEntry[]>([]);
  const [recentOpen, setRecentOpen] = useState(false);
  const [backupOnSave, setBackupOnSave] = useState(readBackupPreference);

  /* ----------------------------------------------------------- univer setup */

  useEffect(() => {
    const container = containerRef.current;
    if (!container) return;

    const worker = new Worker(new URL('./univer-worker.ts', import.meta.url), {
      type: 'module',
    });
    workerRef.current = worker;

    const { univer, univerAPI } = createUniver({
      locale: LocaleType.KO_KR,
      locales: {
        [LocaleType.KO_KR]: merge({}, sheetsCoreKoKR, findReplaceKoKR, filterKoKR, sortKoKR),
      },
      theme: defaultTheme,
      presets: [
        UniverSheetsCorePreset({ container, workerURL: worker }),
        UniverSheetsFindReplacePreset(),
        UniverSheetsFilterPreset(),
        UniverSheetsSortPreset(),
      ],
    });

    apiRef.current = univerAPI;
    setReady(true);

    return () => {
      listenerRef.current?.dispose();
      listenerRef.current = null;
      apiRef.current = null;
      unitIdRef.current = null;
      univer.dispose();
      // `univer.dispose()` does not own the worker we handed it.
      worker.terminate();
      workerRef.current = null;
    };
  }, []);

  /** Swaps the on-screen workbook for `snapshot` and restarts dirty tracking. */
  const loadSnapshot = useCallback((snapshot: IWorkbookData) => {
    const api = apiRef.current;
    // Throwing rather than returning quietly: a silent no-op here puts the file
    // name in the title bar with an empty grid underneath, which reads as data
    // loss. Callers are gated on `ready`, so this should be unreachable.
    if (!api) throw new Error('표가 아직 준비되지 않았습니다. 잠시 후 다시 시도해 주세요.');

    listenerRef.current?.dispose();
    listenerRef.current = null;

    if (unitIdRef.current) api.disposeUnit(unitIdRef.current);

    const workbook = api.createWorkbook(snapshot);
    unitIdRef.current = workbook.getId();
    calculatingRef.current = false;
    loadedAtRef.current = performance.now();
    setDirty(false);

    listenerRef.current = workbook.onCommandExecuted((command) => {
      if (command.id === CALC_START) {
        calculatingRef.current = true;
        return;
      }
      if (command.id === CALC_STOP) {
        calculatingRef.current = false;
        return;
      }
      if (CALC_LIFECYCLE.has(command.id)) return;

      // MUTATION is the only command type that changes the saved model; COMMAND
      // and OPERATION cover things like moving the selection, which must not mark
      // the document dirty.
      if (command.type !== CommandType.MUTATION) return;
      if (calculatingRef.current) return;
      if (performance.now() - loadedAtRef.current < LOAD_SETTLE_MS) return;

      setDirty(true);
    });
  }, []);

  // Start on an empty workbook so the grid is never blank.
  useEffect(() => {
    if (ready) loadSnapshot(createEmptySnapshot());
  }, [ready, loadSnapshot]);

  useEffect(() => {
    void setWindowTitle(`${dirty ? '• ' : ''}${doc.name} — ${APP_NAME}`);
  }, [doc.name, dirty]);

  const refreshRecent = useCallback(() => {
    void recentList().then(setRecent);
  }, []);

  useEffect(refreshRecent, [refreshRecent]);

  useEffect(() => {
    try {
      localStorage.setItem(BACKUP_PREFERENCE_KEY, backupOnSave ? 'on' : 'off');
    } catch {
      // A locked-down storage is not worth failing the toggle over.
    }
  }, [backupOnSave]);

  /* ----------------------------------------------------------------- open */

  const confirmDiscard = useCallback(async () => {
    if (!dirty) return true;
    return confirmDialog(
      '저장하지 않은 변경 사항이 사라집니다. 계속할까요?',
      '저장되지 않은 변경 사항',
    );
  }, [dirty]);

  /** The one place a file becomes the on-screen document, whatever opened it. */
  const openBytes = useCallback(
    async (file: OpenedFile) => {
      setBusy('파일 읽는 중…');
      try {
        if (isDelimitedName(file.name)) {
          const { csvToSnapshot } = await loadCsv();
          const result = csvToSnapshot(file.bytes, file.name);
          loadSnapshot(result.snapshot);
          setDoc({
            path: file.path,
            name: file.name,
            format: 'csv',
            delimiter: result.delimiter,
            encoding: result.encoding,
            warningAcknowledged: true,
          });
        } else {
          const [{ importXlsx }, , { describeLossyFeatures, detectLossyFeatures }] =
            await loadXlsx();
          const warning = describeLossyFeatures(detectLossyFeatures(file.bytes));
          const { snapshot } = await importXlsx(file.bytes, file.name);
          loadSnapshot(snapshot);
          setDoc({
            path: file.path,
            name: file.name,
            format: 'xlsx',
            warning,
            warningAcknowledged: !warning,
          });
        }
        refreshRecent();
      } catch (error) {
        await alertDialog(
          `파일을 열 수 없습니다.\n\n${error instanceof Error ? error.message : String(error)}`,
          '열기 실패',
        );
      } finally {
        setBusy(null);
      }
    },
    [loadSnapshot, refreshRecent],
  );

  const openFile = useCallback(async () => {
    if (!(await confirmDiscard())) return;
    const file = await pickAndReadFile();
    if (file) await openBytes(file);
  }, [confirmDiscard, openBytes]);

  /** Opens a path the app already knows about: a recent entry, or an OS request. */
  const openPath = useCallback(
    async (path: string, needsGrant: boolean) => {
      if (!(await confirmDiscard())) return;
      setRecentOpen(false);
      try {
        // Scope grants are lost on restart, so a remembered path needs one again.
        if (needsGrant) await grantRecent(path);
        await openBytes({ path, name: basename(path), bytes: await readFileAt(path) });
      } catch (error) {
        // Drop the entry before showing the dialog, not after: the dialog blocks
        // until it is dismissed, and until then the recent list would still be
        // offering a file that cannot be opened.
        await forgetRecent(path);
        refreshRecent();
        await alertDialog(
          `파일을 열 수 없습니다.\n\n${error instanceof Error ? error.message : String(error)}`,
          '열기 실패',
        );
      }
    },
    [confirmDiscard, openBytes, refreshRecent],
  );

  // Kept in a ref so the subscription below can depend on `ready` alone: it must
  // attach exactly once, and `openPath` changes identity whenever `dirty` does.
  const openPathRef = useRef(openPath);
  useEffect(() => {
    openPathRef.current = openPath;
  }, [openPath]);

  // Finder double-clicks and files dropped on the window arrive from Rust, which
  // has already put them in scope.
  useEffect(() => {
    // Waiting for `ready` matters: Univer must exist before a file can be shown,
    // and a request that arrives earlier would otherwise be swallowed.
    if (!ready) return;

    let cancelled = false;
    let dispose: (() => void) | undefined;

    void (async () => {
      dispose = await onOpenFiles((paths) => {
        void openPathRef.current(paths[0], false);
      });
      if (cancelled) {
        dispose();
        return;
      }

      // Only now can Rust stop queueing. Anything it held while the window was
      // starting up comes back here — this is the Finder-launch path.
      const queued = await announceFrontendReady();
      if (!cancelled && queued.length > 0) {
        void openPathRef.current(queued[0], false);
      }
    })().catch(async (error) => {
      // Without this the handover fails silently and the app just sits there
      // showing an empty grid — the exact failure this whole path exists to fix.
      await alertDialog(
        `파일 열기 요청을 처리하지 못했습니다.\n\n${
          error instanceof Error ? error.message : String(error)
        }`,
        '열기 실패',
      );
    });

    return () => {
      cancelled = true;
      dispose?.();
    };
  }, [ready]);

  /* ----------------------------------------------------------------- save */

  /** Serialises the live workbook in `format`, warning once about any loss. */
  const buildBytes = useCallback(
    async (format: Format): Promise<Uint8Array | undefined> => {
      const workbook = apiRef.current?.getActiveWorkbook();
      if (!workbook) return undefined;

      // A cell being typed into has not reached the model yet, so saving mid-edit
      // would write the old value and silently lose what the user just entered.
      if (workbook.isCellEditing()) await workbook.endEditingAsync(true);

      const snapshot = workbook.save();

      if (format === 'csv') {
        const { sheetsLostInCsv, snapshotToCsv } = await loadCsv();
        const sheetId = workbook.getActiveSheet().getSheetId();
        const lostSheets = sheetsLostInCsv(snapshot, sheetId);
        const extra = lostSheets.length > 0 ? ` 그리고 다른 시트(${lostSheets.join(', ')})` : '';
        const go = await confirmDialog(
          `CSV는 값만 담습니다. 서식·수식·병합셀${extra}은 저장되지 않습니다.\n\n계속할까요?`,
          'CSV로 저장',
        );
        if (!go) return undefined;
        return snapshotToCsv(snapshot, { sheetId, delimiter: doc.delimiter });
      }

      if (doc.warning && !doc.warningAcknowledged) {
        const go = await confirmDialog(
          `${doc.warning}\n\n그래도 저장할까요?`,
          '일부 기능이 저장되지 않습니다',
        );
        if (!go) return undefined;
        setDoc((prev) => ({ ...prev, warningAcknowledged: true }));
      }

      const [, { exportXlsx }] = await loadXlsx();
      return exportXlsx(snapshot);
    },
    [doc.delimiter, doc.warning, doc.warningAcknowledged],
  );

  const reportSaveFailure = useCallback(async (error: unknown) => {
    await alertDialog(
      `저장에 실패했습니다.\n\n${error instanceof Error ? error.message : String(error)}`,
      '저장 실패',
    );
  }, []);

  const saveAs = useCallback(async () => {
    try {
      if (!isTauri) {
        // No save dialog in the browser; keep the current format and download.
        const bytes = await buildBytes(doc.format);
        if (!bytes) return;
        downloadFile(doc.name, bytes);
        setDirty(false);
        return;
      }

      const suggested = /\.(xlsx|xlsm|csv|tsv)$/i.test(doc.name)
        ? doc.name
        : `${doc.name}.${doc.format === 'csv' ? 'csv' : 'xlsx'}`;
      const target = await pickSavePath(suggested);
      if (!target) return;

      // The chosen extension decides the format, which is why the path is picked
      // before the bytes are built.
      const format: Format = isDelimitedName(target) ? 'csv' : 'xlsx';

      setBusy('저장 중…');
      const bytes = await buildBytes(format);
      if (!bytes) return;

      await writeFileAt(target, bytes);
      await noteSaved(target);
      setDoc((prev) => ({ ...prev, path: target, name: basename(target), format }));
      setDirty(false);
      refreshRecent();
    } catch (error) {
      await reportSaveFailure(error);
    } finally {
      setBusy(null);
    }
  }, [buildBytes, doc.format, doc.name, refreshRecent, reportSaveFailure]);

  const save = useCallback(async () => {
    if (!doc.path) {
      await saveAs();
      return;
    }

    try {
      setBusy('저장 중…');
      const bytes = await buildBytes(doc.format);
      if (!bytes) return;

      if (backupOnSave) {
        // A failed backup must not block the save the user asked for, but they
        // should know the safety net was not there.
        try {
          await backupFile(doc.path);
        } catch (error) {
          const go = await confirmDialog(
            `백업을 만들지 못했습니다.\n\n${
              error instanceof Error ? error.message : String(error)
            }\n\n백업 없이 저장할까요?`,
            '백업 실패',
          );
          if (!go) return;
        }
      }

      await writeFileAt(doc.path, bytes);
      setDirty(false);
    } catch (error) {
      await reportSaveFailure(error);
    } finally {
      setBusy(null);
    }
  }, [backupOnSave, buildBytes, doc.format, doc.path, reportSaveFailure, saveAs]);

  const newFile = useCallback(async () => {
    if (!(await confirmDiscard())) return;
    loadSnapshot(createEmptySnapshot());
    setDoc(UNTITLED);
  }, [confirmDiscard, loadSnapshot]);

  /* ------------------------------------------------------------- shortcuts */

  useEffect(() => {
    const onKeyDown = (event: KeyboardEvent) => {
      if (event.key === 'Escape') {
        setRecentOpen(false);
        return;
      }
      if (!(event.metaKey || event.ctrlKey)) return;

      const key = event.key.toLowerCase();
      if (key === 's') {
        event.preventDefault();
        void (event.shiftKey ? saveAs() : save());
      } else if (key === 'o') {
        event.preventDefault();
        void openFile();
      } else if (key === 'n') {
        event.preventDefault();
        void newFile();
      }
    };
    // Capture phase: Univer binds its own handlers on the grid and would
    // otherwise swallow the event before it reaches the window. Find-and-replace
    // (⌘F) is deliberately left to Univer.
    window.addEventListener('keydown', onKeyDown, true);
    return () => window.removeEventListener('keydown', onKeyDown, true);
  }, [save, saveAs, openFile, newFile]);

  /* ----------------------------------------------- browser drag-drop fallback */

  useEffect(() => {
    // Under Tauri the OS handles the drop natively and Rust forwards the paths,
    // so DOM drag events never fire. This is only for `pnpm dev` in a browser.
    if (isTauri) return;

    const onDragOver = (event: DragEvent) => event.preventDefault();
    const onDrop = async (event: DragEvent) => {
      event.preventDefault();
      const file = event.dataTransfer?.files?.[0];
      if (!file) return;
      if (!(await confirmDiscard())) return;
      await openBytes({ name: file.name, bytes: new Uint8Array(await file.arrayBuffer()) });
    };

    window.addEventListener('dragover', onDragOver);
    window.addEventListener('drop', onDrop);
    return () => {
      window.removeEventListener('dragover', onDragOver);
      window.removeEventListener('drop', onDrop);
    };
  }, [confirmDiscard, openBytes]);

  /* ------------------------------------------------------------------- UI */

  const disabled = Boolean(busy);

  return (
    <div className="app">
      <header className="toolbar">
        <button type="button" onClick={() => void newFile()} disabled={disabled}>
          새 파일
        </button>
        <button type="button" onClick={() => void openFile()} disabled={disabled}>
          열기
        </button>

        {isTauri && (
          <div className="recent">
            <button
              type="button"
              onClick={() => setRecentOpen((open) => !open)}
              disabled={disabled || recent.length === 0}
              aria-expanded={recentOpen}
              title={recent.length === 0 ? '최근 파일 없음' : '최근 파일'}
            >
              최근 ▾
            </button>
            {recentOpen && recent.length > 0 && (
              <>
                {/* Click-away layer, so the menu closes without a document listener. */}
                <div className="scrim" onClick={() => setRecentOpen(false)} />
                <ul className="recent-menu">
                  {recent.map((entry) => (
                    <li key={entry.path}>
                      <button
                        type="button"
                        onClick={() => void openPath(entry.path, true)}
                        title={entry.path}
                      >
                        {entry.name}
                      </button>
                    </li>
                  ))}
                </ul>
              </>
            )}
          </div>
        )}

        <button type="button" onClick={() => void save()} disabled={disabled}>
          저장
        </button>
        <button type="button" onClick={() => void saveAs()} disabled={disabled}>
          다른 이름으로
        </button>

        <label className="backup" title="저장하기 전에 원본을 name.bak.xlsx 로 복사합니다">
          <input
            type="checkbox"
            checked={backupOnSave}
            onChange={(event) => setBackupOnSave(event.target.checked)}
          />
          백업
        </label>

        <span className="filename" title={doc.path ?? doc.name}>
          {dirty && (
            <span className="dot" aria-label="저장되지 않음">
              •
            </span>
          )}
          {doc.name}
          {doc.format === 'csv' && (
            <span className="badge">
              CSV{doc.encoding === 'cp949' ? ' · CP949' : ''}
              {doc.delimiter === '\t' ? ' · 탭' : doc.delimiter === ';' ? ' · 세미콜론' : ''}
            </span>
          )}
        </span>

        {busy && <span className="busy">{busy}</span>}
      </header>

      {doc.warning && (
        <div className="warning" role="status">
          <span>⚠️ {doc.warning}</span>
          <button type="button" onClick={() => setDoc((prev) => ({ ...prev, warning: undefined }))}>
            닫기
          </button>
        </div>
      )}

      <div className="grid" ref={containerRef} />
    </div>
  );
}
