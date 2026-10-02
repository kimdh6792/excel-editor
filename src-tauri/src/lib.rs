mod pending;
mod recent;

use tauri::{Emitter, Manager};

use pending::PendingOpens;
use recent::{RecentStore, admit_path, paths_from_urls};

/// Event the frontend listens on to open files it did not pick itself.
const OPEN_FILES_EVENT: &str = "excel-editor://open-files";

/// Puts OS-supplied paths in scope, records them, and gets them to the frontend.
///
/// At launch the frontend does not exist yet, so the paths are queued rather
/// than emitted; `frontend_ready` hands them over once the UI can act on them.
fn announce_paths(app: &tauri::AppHandle, paths: Vec<std::path::PathBuf>) {
    let accepted: Vec<String> = paths
        .into_iter()
        .filter(|p| p.is_file())
        .map(|p| {
            admit_path(app, &p);
            p.to_string_lossy().to_string()
        })
        .collect();

    if accepted.is_empty() {
        return;
    }

    if let Some(live) = app.state::<PendingOpens>().queue_or_emit(accepted) {
        let _ = app.emit(OPEN_FILES_EVENT, live);
    }
}

#[cfg_attr(mobile, tauri::mobile_entry_point)]
pub fn run() {
    let app = tauri::Builder::default()
        .plugin(tauri_plugin_dialog::init())
        .plugin(tauri_plugin_fs::init())
        .manage(RecentStore::default())
        .manage(PendingOpens::default())
        .invoke_handler(tauri::generate_handler![
            pending::frontend_ready,
            recent::recent_list,
            recent::record_recent,
            recent::grant_recent,
            recent::forget_recent,
            recent::clear_recent,
            recent::backup_file,
        ])
        .setup(|app| {
            app.state::<RecentStore>().hydrate(app.handle());
            Ok(())
        })
        .on_window_event(|window, event| {
            // Files dropped on the window come from the OS, so they need the same
            // scope grant as a dialog pick before the frontend can read them.
            if let tauri::WindowEvent::DragDrop(tauri::DragDropEvent::Drop { paths, .. }) = event {
                announce_paths(window.app_handle(), paths.clone());
            }
        })
        .build(tauri::generate_context!())
        .expect("error while building tauri application");

    app.run(|app_handle, event| {
        // Fired when the user opens a file with the app from Finder, both at
        // launch and while it is already running.
        #[cfg(target_os = "macos")]
        if let tauri::RunEvent::Opened { urls } = &event {
            announce_paths(app_handle, paths_from_urls(urls));
        }

        // Silence unused-variable warnings on platforms without `Opened`.
        #[cfg(not(target_os = "macos"))]
        {
            let _ = (app_handle, &event);
        }
    });
}
