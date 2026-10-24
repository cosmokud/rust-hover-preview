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

/// Write one line, and build nothing at all where the trace is off.
macro_rules! note_pin_trace {
    ($($line:tt)*) => {
        if super::pin_trace::asked() {
            super::pin_trace::write(&format!($($line)*));
        }
    };
}

pub(super) use note_pin_trace;
