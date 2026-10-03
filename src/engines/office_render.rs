//! The Office render tier: Word, Excel or PowerPoint draws a document's first
//! page, once, and the preview then draws from it.
//!
//! Nothing here is ever on the hover path. A page is asked for the moment a
//! hover needs one, it is drawn on a thread of its own, and the preview shows a
//! spinner in the meantime — so what a document costs is bounded by the render
//! tier even when it is an Office start, an export and a dialogs worth of
//! waiting, and what the hover waits for is that render rather than a timer in
//! front of it.
//!
//! Three rules shape the rest:
//!
//! * **The user's Office is never disturbed.** An automation instance may attach
//!   to a running Word or Excel — the applications are registered for multiple
//!   use — so nothing is hidden that is already visible, the settings that are
//!   changed are restored, and only an instance this app created is quit or
//!   ended.
//! * **A stuck engine must not wedge the app.** The calls below are COM calls
//!   into another process and cannot be cancelled or bounded; the thread that
//!   makes them may be lost to a modal dialog inside Office. What that costs is
//!   one render, never the tier: a worker that has been inside one piece of work
//!   for too long is given up on, the process it started is ended with it, and
//!   the next hover is answered by a fresh worker. Nothing waits on this thread.
//! * **A page outlives the hover that asked for it, and the run that drew it.** Office's
//!   export calls take a file name rather than a stream, so one render writes one scratch
//!   file under the temp folder and reads it back out. What it read is kept as the page both
//!   engines' documents are drawn from, bounded by `document_cache_mb` (see
//!   `document_cache`). What is kept *here* is the request slot, the documents that have
//!   refused a page, and which worker is current.

mod com;
mod engines;
mod renderers;
mod worker;

pub(crate) use worker::{
    enabled, held_page, hover_ended, page_is_narrower_than, page_is_workbook_picture, request,
    shutdown, stop_engines,
};

/// Kept at the tier's own path: a refusal may be remembered from outside the worker,
/// and the worker is not where a caller of the tier should have to look for it.
#[allow(unused_imports)]
pub(crate) use worker::remember_failure;

// What the tests below reach for, out of the four submodules this is split across.
#[cfg(test)]
use com::{last_failure, path_variant, Object};
#[cfg(test)]
use engines::{Engine, Engines};
#[cfg(test)]
use renderers::{alerts_off, has_zone_identifier, with_a_workbook, MAX_SLIDE_EXPORT_WIDTH};
#[cfg(test)]
use worker::{render_request, RenderOutcome, RenderRequest, WORKER_GENERATION};

#[cfg(test)]
use crate::app::engine_processes;
#[cfg(test)]
use crate::config::config::OfficeEngine;
#[cfg(test)]
use crate::engines::document_cache::{self, PageKind};
#[cfg(test)]
use crate::formats::office_formats::{app_for, container_kind, OfficeApp};
#[cfg(test)]
use std::path::{Path, PathBuf};
#[cfg(test)]
use std::sync::atomic::Ordering;
#[cfg(test)]
use std::time::{Duration, Instant};
#[cfg(test)]
use windows::core::VARIANT;
#[cfg(test)]
use windows::Win32::System::Com::{CoInitializeEx, COINIT_APARTMENTTHREADED};

#[cfg(test)]
mod tests;
