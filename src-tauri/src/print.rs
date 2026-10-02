//! Printing, which has to start in Rust.
//!
//! WKWebView does not implement JavaScript's `window.print()` — the call is a
//! silent no-op, so a web-only implementation would appear to work and print
//! nothing. Going through Tauri instead reaches wry, which builds an
//! `NSPrintOperation` from the webview and runs it as a sheet on the window.
//!
//! Paper size, orientation, scaling, copies and "Save as PDF" are then the
//! system dialog's job. The frontend's only responsibility is deciding what the
//! document contains before it calls this, which it does by swapping a
//! print-only rendering of the sheet in for the on-screen grid.

/// Opens the system print dialog for the calling window's webview.
///
/// Returns once the dialog has been put up, not once printing has finished: the
/// operation runs as a window-modal sheet. The caller must therefore leave the
/// printable DOM in place after this resolves.
#[tauri::command]
pub fn print_page(window: tauri::WebviewWindow) -> Result<(), String> {
    window.print().map_err(|e| e.to_string())
}
