//! Text previews: one screenful of a text file, colored by the syntax
//! definition its extension names.
//!
//! The file is read once, decoded, turned into styled lines, laid out for the
//! box the caller asks for, and painted with GDI into a BGRA frame — the same
//! frame shape a decoded image arrives in, so everything downstream (the layered
//! surface, the spinner, the hover generation check) is unchanged.
//!
//! Two calls share that work. `measure` answers "how big is this preview?"
//! before anything is painted, the way a PDF page's size is resolved before the
//! window is placed; `render` then lays the same document out for the box the
//! layout settled on and fills it. The parsed document is cached per file, mode
//! and theme, so a second hover, a theme switch or a repaint costs a layout
//! instead of a parse.

mod document;
mod frame;
mod layout;
mod markdown;
mod source;

pub use document::text_of;
pub use frame::{
    measure, position_in, render_scrolled, word_in, FrameLine, ScrollBar, Selection,
    TextPreviewOptions,
};
// A run is what a line is made of, and it is named only by the tests that build lines of their
// own rather than painting a document.
#[cfg(test)]
pub use frame::FrameRun;
pub use layout::scroll_line_at_track_y;

#[cfg(test)]
mod tests;
