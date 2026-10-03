//! Drawing a document — or a font specimen — in the browser engine that is already on the
//! machine.
//!
//! Every document this app previews is drawn here — still, gzipped or animated: this app
//! rasterizes none of them, so what a document needs is the WebView2 runtime Windows 11
//! ships with, which is Chromium, and which is the only complete implementation of SVG
//! that is on the machine without installing anything. A machine without the runtime has
//! no SVG preview at all, which is the answer a file that will not decode gets.
//!
//! A font is drawn by the same engine for the same reason, and it is one of the two kinds
//! this app draws rather than rasterizes: the five formats a font goes by are one container
//! or another around the same outlines, and the browser reads all of them. What the page
//! carries for one is a specimen — the font's own name and the lines its character map
//! covers, worked out on this side by `font_preview` — and a collection is answered with the
//! face the setting names written out as a font of its own, because no page can name a face
//! inside one.
//!
//! What lives here is the engine and the window it draws in, not the preview loop: the
//! loop measures the document for the layout, hands it over with the box the layout came
//! out with, and this answers with a window of its own. That window is its own because
//! the preview window is a layered one, and a layered window has no window tree to put a
//! child in — the same shape as the video path, where the player's own window is the
//! preview.
//!
//! One engine is kept warm between documents and let go after `webview_idle`, ten
//! minutes by default: beginning one costs a browser start, and pointing a warm one at
//! another file costs a few milliseconds, so what a hover pays for a second document is
//! nothing worth measuring — and what a hover pays for the *first* one is a browser
//! start, which is what the waiting spinner is shown for. What is let go of is the
//! engine and not the thread that holds it, which stays parked on its channel for the
//! run and begins a new engine for the next document: a thread that ended with its
//! browser would leave every document after the first idle timeout with nobody to draw
//! it, which is no preview for the rest of the run. An app left alone has no browser
//! process and one thread asleep — and the settings it is given are the app's own rules
//! rather than a browser's: a document is drawn and not run, and nothing about it is a
//! way out of the preview. A page of HTML is the one exception to that, and the exception
//! is deliberate rather than a leak: it is handed to the browser as a page rather than as an
//! image, a page that draws itself is nothing without a run, and what a run is given stops
//! at everything that is a way out of the frame it is in (see `html_page`, `page_runs`).
//!
//! What is kept between documents is the browser, not the work it was doing, and
//! `Host::hide`/`Host::wake` are what stop and start it (see `suspend`).

mod api;
mod engine;
mod environment;
mod host;
mod pages;

pub(crate) use api::{
    can_draw, draws, is_available, is_behind, is_showing, page_runs, screen_rect, showing_hwnd,
    showing_path, take_failure_notice, Area,
};
pub(crate) use environment::{clear_stale_profiles, stale_profile_pids};
pub(crate) use pages::{hide, hide_html_preview, place, show, shutdown, wanted_here};

// What the tests below and the pinned-preview tests elsewhere reach for, out of the five
// submodules this is split across.
#[cfg(test)]
use api::{drop_placement, take_placement, ENGINE, PLACED, PLACE_ASKED};
#[cfg(test)]
use api::{last_timings, page_runs_under, runtime_version};
#[cfg(test)]
use environment::{ex_style_for, file_url, mouse_activate_answers};
#[cfg(test)]
use pages::{
    ask_place, box_change, escape_attribute, font_page, frame_html, frame_page, html_page,
    specimen_html, BoxChange,
};
#[cfg(test)]
pub(crate) use pages::{clear_want_for_test, publish_want_for_test};

#[cfg(test)]
use crate::config::config::TransparentBackground;
#[cfg(test)]
use std::path::{Path, PathBuf};
#[cfg(test)]
use std::sync::atomic::Ordering;
#[cfg(test)]
use std::time::{Duration, Instant};
#[cfg(test)]
use windows::Win32::Foundation::LRESULT;
#[cfg(test)]
use windows::Win32::UI::WindowsAndMessaging::{WS_EX_NOACTIVATE, WS_EX_TOOLWINDOW, WS_EX_TOPMOST};

#[cfg(test)]
mod tests;
