//! Files the OS asked us to open before the UI could listen for them.
//!
//! Double-clicking a file in Finder while the app is closed launches it and
//! fires `RunEvent::Opened` during startup — well before the webview has loaded
//! and registered its event listener. Emitting straight away would drop the
//! request on the floor and leave the user staring at an empty grid, which is
//! exactly what happened. So paths are queued until the frontend reports that it
//! is ready, and handed over in one go at that point.

use std::sync::Mutex;

#[derive(Default)]
struct State {
    /// Set once the frontend has registered its listener.
    ready: bool,
    /// Paths that arrived before that.
    paths: Vec<String>,
}

#[derive(Default)]
pub struct PendingOpens {
    state: Mutex<State>,
}

impl PendingOpens {
    /// Queues paths, or reports that they should be emitted live instead.
    ///
    /// The check and the queueing happen under one lock so a path cannot slip
    /// between "not ready yet" and the frontend draining the queue.
    pub fn queue_or_emit(&self, paths: Vec<String>) -> Option<Vec<String>> {
        let Ok(mut state) = self.state.lock() else {
            // A poisoned lock should not lose the user's file; emit and hope the
            // frontend is up.
            return Some(paths);
        };

        if state.ready {
            Some(paths)
        } else {
            state.paths.extend(paths);
            None
        }
    }

    fn mark_ready(&self) -> Vec<String> {
        let Ok(mut state) = self.state.lock() else {
            return Vec::new();
        };
        state.ready = true;
        std::mem::take(&mut state.paths)
    }
}

/// Called by the frontend once Univer is up and the open-files listener is
/// attached. Returns whatever the OS asked for in the meantime.
#[tauri::command]
pub fn frontend_ready(pending: tauri::State<'_, PendingOpens>) -> Vec<String> {
    pending.mark_ready()
}
