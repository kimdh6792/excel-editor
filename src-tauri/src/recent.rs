//! Recently-opened files, and the filesystem scope that goes with them.
//!
//! The app grants the fs plugin access only to files the user actually picked,
//! so a path remembered in a previous session is no longer readable after a
//! restart. Rather than widen the scope to the whole disk, the recent list lives
//! here on the Rust side: `grant_recent` re-grants access, but only for a path
//! this module recorded, and `record_recent` refuses to record anything that is
//! not already in scope. The frontend therefore cannot use either command to
//! reach a file the user never chose.

use std::{
    fs,
    path::{Path, PathBuf},
    sync::Mutex,
    time::{SystemTime, UNIX_EPOCH},
};

use serde::{Deserialize, Serialize};
use tauri::{AppHandle, Manager, Runtime};
use tauri_plugin_fs::FsExt;

/// How many entries to keep. Enough to cover a working session, short enough to
/// show in one menu.
const MAX_ENTRIES: usize = 15;

#[derive(Clone, Debug, Serialize, Deserialize)]
pub struct RecentEntry {
    /// Absolute path, as the OS gave it to us.
    pub path: String,
    /// File name, so the frontend does not have to split paths.
    pub name: String,
    /// Unix milliseconds of the most recent open.
    pub opened_at: u64,
}

#[derive(Default)]
pub struct RecentStore {
    entries: Mutex<Vec<RecentEntry>>,
}

fn now_millis() -> u64 {
    SystemTime::now()
        .duration_since(UNIX_EPOCH)
        .map(|d| d.as_millis() as u64)
        .unwrap_or(0)
}

fn file_name_of(path: &Path) -> String {
    path.file_name()
        .map(|n| n.to_string_lossy().to_string())
        .unwrap_or_else(|| path.to_string_lossy().to_string())
}

fn store_path<R: Runtime>(app: &AppHandle<R>) -> Option<PathBuf> {
    let dir = app.path().app_config_dir().ok()?;
    Some(dir.join("recent.json"))
}

fn read_from_disk<R: Runtime>(app: &AppHandle<R>) -> Vec<RecentEntry> {
    let Some(path) = store_path(app) else {
        return Vec::new();
    };
    let Ok(text) = fs::read_to_string(path) else {
        return Vec::new();
    };
    // A corrupt list is not worth failing over; start fresh instead.
    serde_json::from_str(&text).unwrap_or_default()
}

fn write_to_disk<R: Runtime>(app: &AppHandle<R>, entries: &[RecentEntry]) {
    let Some(path) = store_path(app) else { return };
    if let Some(dir) = path.parent() {
        let _ = fs::create_dir_all(dir);
    }
    if let Ok(text) = serde_json::to_string_pretty(entries) {
        let _ = fs::write(path, text);
    }
}

impl RecentStore {
    /// Loads the persisted list into memory. Called once during setup.
    pub fn hydrate<R: Runtime>(&self, app: &AppHandle<R>) {
        let loaded = read_from_disk(app);
        if let Ok(mut entries) = self.entries.lock() {
            *entries = loaded;
        }
    }

    fn snapshot(&self) -> Vec<RecentEntry> {
        self.entries.lock().map(|e| e.clone()).unwrap_or_default()
    }

    fn contains(&self, path: &str) -> bool {
        self.entries
            .lock()
            .map(|e| e.iter().any(|entry| entry.path == path))
            .unwrap_or(false)
    }

    fn insert<R: Runtime>(&self, app: &AppHandle<R>, path: &Path) {
        let key = path.to_string_lossy().to_string();
        let entry = RecentEntry {
            path: key.clone(),
            name: file_name_of(path),
            opened_at: now_millis(),
        };

        let updated = {
            let Ok(mut entries) = self.entries.lock() else {
                return;
            };
            entries.retain(|e| e.path != key);
            entries.insert(0, entry);
            entries.truncate(MAX_ENTRIES);
            entries.clone()
        };
        write_to_disk(app, &updated);
    }

    fn remove<R: Runtime>(&self, app: &AppHandle<R>, path: &str) {
        let updated = {
            let Ok(mut entries) = self.entries.lock() else {
                return;
            };
            entries.retain(|e| e.path != path);
            entries.clone()
        };
        write_to_disk(app, &updated);
    }

    fn clear<R: Runtime>(&self, app: &AppHandle<R>) {
        if let Ok(mut entries) = self.entries.lock() {
            entries.clear();
        }
        write_to_disk(app, &[]);
    }
}

/// Grants fs access to a path and adds it to the recent list.
///
/// Used for paths the OS handed us — a Finder double-click or a drop onto the
/// window — which the dialog plugin never saw and so never put in scope.
pub fn admit_path<R: Runtime>(app: &AppHandle<R>, path: &Path) {
    if !path.is_file() {
        return;
    }
    if let Some(scope) = app.try_fs_scope() {
        let _ = scope.allow_file(path);
    }
    app.state::<RecentStore>().insert(app, path);
}

/// Converts the `file://` URLs of a system open request into local paths.
pub fn paths_from_urls(urls: &[tauri::Url]) -> Vec<PathBuf> {
    urls.iter()
        .filter(|u| u.scheme() == "file")
        .filter_map(|u| u.to_file_path().ok())
        .collect()
}

/* -------------------------------------------------------------- commands */

#[tauri::command]
pub fn recent_list(app: AppHandle, store: tauri::State<'_, RecentStore>) -> Vec<RecentEntry> {
    let entries = store.snapshot();
    let existing: Vec<RecentEntry> = entries
        .iter()
        .filter(|e| Path::new(&e.path).is_file())
        .cloned()
        .collect();

    // Prune vanished files so the list does not accumulate dead paths.
    if existing.len() != entries.len() {
        if let Ok(mut guard) = store.entries.lock() {
            *guard = existing.clone();
        }
        write_to_disk(&app, &existing);
    }
    existing
}

/// Records a path the user opened or saved through a dialog.
///
/// Refuses paths that are not already in the fs scope: the dialog plugin puts
/// the chosen file in scope, so a path the frontend invented will not be there.
#[tauri::command]
pub fn record_recent(
    app: AppHandle,
    store: tauri::State<'_, RecentStore>,
    path: String,
) -> Result<(), String> {
    let candidate = PathBuf::from(&path);
    if !candidate.is_file() {
        return Err("파일이 존재하지 않습니다.".into());
    }

    let allowed = app
        .try_fs_scope()
        .map(|scope| scope.is_allowed(&candidate))
        .unwrap_or(false);
    if !allowed {
        return Err("허용되지 않은 경로입니다.".into());
    }

    store.insert(&app, &candidate);
    Ok(())
}

/// Re-grants fs access to a recorded path, so a recent entry can be reopened
/// after a restart. Paths outside the recent list are rejected.
#[tauri::command]
pub fn grant_recent(
    app: AppHandle,
    store: tauri::State<'_, RecentStore>,
    path: String,
) -> Result<(), String> {
    if !store.contains(&path) {
        return Err("최근 파일 목록에 없는 경로입니다.".into());
    }

    let candidate = PathBuf::from(&path);
    if !candidate.is_file() {
        store.remove(&app, &path);
        return Err("파일이 더 이상 존재하지 않습니다.".into());
    }

    match app.try_fs_scope() {
        Some(scope) => scope.allow_file(&candidate).map_err(|e| e.to_string()),
        None => Err("파일 시스템 스코프를 사용할 수 없습니다.".into()),
    }
}

#[tauri::command]
pub fn forget_recent(app: AppHandle, store: tauri::State<'_, RecentStore>, path: String) {
    store.remove(&app, &path);
}

#[tauri::command]
pub fn clear_recent(app: AppHandle, store: tauri::State<'_, RecentStore>) {
    store.clear(&app);
}

/// Copies a file next to itself as `name.bak.ext` before it gets overwritten.
///
/// Done here rather than in the frontend because the backup lands on a path the
/// user never picked, which the fs scope rightly would not allow.
#[tauri::command]
pub fn backup_file(app: AppHandle, path: String) -> Result<String, String> {
    let source = PathBuf::from(&path);
    if !source.is_file() {
        return Err("원본 파일이 없습니다.".into());
    }

    let allowed = app
        .try_fs_scope()
        .map(|scope| scope.is_allowed(&source))
        .unwrap_or(false);
    if !allowed {
        return Err("허용되지 않은 경로입니다.".into());
    }

    let stem = source
        .file_stem()
        .map(|s| s.to_string_lossy().to_string())
        .unwrap_or_else(|| "backup".into());
    let backup_name = match source.extension() {
        Some(ext) => format!("{stem}.bak.{}", ext.to_string_lossy()),
        None => format!("{stem}.bak"),
    };
    let target = source.with_file_name(backup_name);

    fs::copy(&source, &target).map_err(|e| e.to_string())?;
    Ok(target.to_string_lossy().to_string())
}
