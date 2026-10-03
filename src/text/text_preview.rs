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

pub use frame::{
    frame_text, measure, point_is_on_text, position_in, render_scrolled, text_in, FrameLine,
    ScrollBar, Selection, TextPreviewOptions,
};
pub use layout::scroll_line_at_track_y;

#[cfg(test)]
mod tests;
