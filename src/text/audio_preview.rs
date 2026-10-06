//! The card a sound file is previewed as.
//!
//! A sound has no picture, so what a hover shows is what a person wants to know about one: a
//! mark, the file's name, what the file holds, and how far into it the sound is while it
//! plays. It is painted with the same GDI layer a text preview uses and colored by the same
//! theme, which is what keeps the three painted previews looking like one app — and it is
//! made of text and rules and nothing else, because that is what the layer draws.
//!
//! Two calls share the work, as they do for text and for an archive: `measure` answers how big
//! the card wants to be before the window is placed, and `render` fills the box the layout
//! settled on. Both build the same page and neither keeps it — what is cached is the track the
//! card is drawn from (`audio_track`), so a repaint of the clock is a layout rather than a
//! probe, and a second hover of the file is a layout rather than a probe as well.
//!
//! Two things change while the card is on screen, and both are the preview loop's: the clock,
//! with the bar under it, which is drawn from a player that is running, and a name the card
//! has no room for, which is scrolled across the card sideways rather than being cut short
//! (see [`NameScroll`] and `Card::name_offset`). A painted preview of this app's is otherwise
//! drawn once and held, so the preview loop is what asks for the card again while a sound is
//! playing (see `repaint_audio_card`).
//!
//! The card carries its own controls when it is the card a pinned window is showing: the three
//! buttons at the left of its bar row, the bar itself, the volume button at the bar's right, and
//! the pin's own two window buttons in the band above the name line — the minimize that shrinks
//! the pin into its bubble and the close that ends it — stood in a top margin grown to hold
//! buttons twice the side the margin alone would (see `window_button_band`).
//! Every question about where one of them is is answered against the card's own layout (see
//! [`control_at`], [`control_box`] and [`bar_share_at`]) rather than against anything kept beside
//! it, because a button laid out by one arithmetic and hit-tested by another answers a press in
//! the middle of the facts line. Only a pinned window asks: a hover's own window is a window
//! nobody is in, and a click on one of those lands on the file behind it, so the card a hover
//! shows is the card it has always been — no buttons, and a bar from margin to margin (see
//! [`Card::controls`]). The row the buttons stand in is the bar's own line either way, so what
//! those buttons cost is the width of the bar, not the height of the card (see `bar_row`). The
//! two window buttons cost height instead: the band they stand in is room the card grew to hold
//! them, and a hover's card grew with it, so a pin's card is still the size a hover's is.
//!
//! The card a sound is previewed as lives in two files below: `card` for what the card says and
//! the two calls that draw it, and `page` for the page it is laid out and painted from. What is
//! left in this file is the way in — every name the rest of the tree reaches as `audio_preview::`
//! — and the imports the tests beside it read the card and its page through.

mod card;
mod page;

pub(crate) use card::{
    bar_share_at, control_at, control_box, facts_of, measure, name_of, render, AudioPreviewOptions,
    Card, CardChrome, CardControl, NameScroll,
};

#[cfg(test)]
use crate::config::config::TextTheme;
#[cfg(test)]
use crate::readers::audio_track::Track;
#[cfg(test)]
use crate::text::text_paint::{rgb, readable, scaled, TextMetrics};
#[cfg(test)]
use crate::text::text_theme;
#[cfg(test)]
use card::{
    bitrate_label, channel_label, clock, rate_label, Fact, FactKind, BAR_GAP_PIXELS, BAR_PIXELS,
    BAR_REACH_PIXELS, CONTROL_GAP_PIXELS, CONTROL_SIDE_PIXELS, HEADER_LEVEL, NAME_HOLD,
    WINDOW_BUTTON_GAP_PIXELS,
};
#[cfg(test)]
use page::{
    bar_band, bar_row, build_page, clock_runs, fill_span, window_button_band, CardBoxes, Page,
};
#[cfg(test)]
use std::path::Path;
#[cfg(test)]
use std::time::{Duration, Instant};
#[cfg(test)]
use windows::Win32::Foundation::RECT;
#[cfg(test)]
use windows::Win32::Graphics::Gdi::{CreateCompatibleDC, DeleteDC};

#[cfg(test)]
mod tests;
