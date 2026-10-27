//! The FFPLAY video pin's own road, written down when `RHP_PIN_TRACE` asks for it: the take-up,
//! the window and host it draws through, every decision to show or skip it, and what a hover's
//! `Hide` does to a pin that is already standing.
//!
//! The gate is one atomic read: a run without the variable writes nothing and builds no line.

use std::sync::atomic::{AtomicBool, Ordering};
use std::sync::Once;

static READ: Once = Once::new();
static ASKED: AtomicBool = AtomicBool::new(false);

/// Whether `RHP_PIN_TRACE` asks for the pin's road to be written out.
pub(super) fn asked() -> bool {
    READ.call_once(|| {
        ASKED.store(
            std::env::var_os("RHP_PIN_TRACE").is_some(),
            Ordering::Relaxed,
        );
    });
    ASKED.load(Ordering::Relaxed)
}

/// Write one timestamped line to `%TEMP%\rhp-pin-trace.log`, appended and flushed so a run that
/// ends badly leaves what it wrote.
pub(super) fn write(line: &str) {
    let stamp = std::time::SystemTime::now()
        .duration_since(std::time::UNIX_EPOCH)
        .map(|elapsed| format!("{}.{:03}", elapsed.as_secs(), elapsed.subsec_millis()))
        .unwrap_or_else(|_| "-".to_string());
    if let Ok(mut file) = std::fs::OpenOptions::new()
        .create(true)
        .append(true)
        .open(std::env::temp_dir().join("rhp-pin-trace.log"))
    {
        use std::io::Write;
        let _ = writeln!(file, "{stamp} {line}");
        let _ = file.flush();
    }
}

/// A path as the name of the file in it: the folder is the same on both sides of every question a
/// line asks, and a path long enough to be worth reading is a path too long to read.
pub(super) fn path_name(path: &std::path::Path) -> String {
    path.file_name()
        .map(|name| name.to_string_lossy().into_owned())
        .unwrap_or_else(|| "-".to_string())
}

/// The lines a test run wrote, so a test can hold the road to what it said rather than to what the
/// log happened to be configured to keep: the same seam `mpv_ipc` keeps its command lines with, and
/// for the same reason (see `note_pin_trace`).
#[cfg(test)]
static NOTED: std::sync::Mutex<Vec<String>> = std::sync::Mutex::new(Vec::new());

/// Write one line down for a test that reads it (see `NOTED`).
#[cfg(test)]
pub(super) fn note_a_line_for_a_test(line: String) {
    if let Ok(mut noted) = NOTED.lock() {
        noted.push(line);
    }
}

/// The lines this run has written (see `NOTED`).
#[cfg(test)]
pub(super) fn noted_lines() -> Vec<String> {
    NOTED
        .lock()
        .map(|noted| noted.clone())
        .unwrap_or_default()
}

/// Forget the lines written before a test (see `NOTED`).
#[cfg(test)]
pub(super) fn clear_noted() {
    if let Ok(mut noted) = NOTED.lock() {
        noted.clear();
    }
}

/// Write one line, and build nothing at all where the trace is off. A test run keeps every line it
/// wrote, whether the log is on or not (see `NOTED`).
macro_rules! note_pin_trace {
    ($($line:tt)*) => {{
        if super::pin_trace::asked() {
            super::pin_trace::write(&format!($($line)*));
        }
        #[cfg(test)]
        super::pin_trace::note_a_line_for_a_test(format!($($line)*));
    }};
}

pub(super) use note_pin_trace;
