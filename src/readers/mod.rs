pub(crate) mod archive_listing;
pub(crate) mod audio_seek;
pub(crate) mod audio_track;
pub(crate) mod bcn;
pub(crate) mod comic_preview;
pub(crate) mod dds_image;
pub(crate) mod eps_image;
pub(crate) mod font_preview;
pub(crate) mod heif_sequence;
pub(crate) mod jxl_image;
pub(crate) mod metafile_image;
pub(crate) mod office_preview;
pub(crate) mod pdf_preview;
pub(crate) mod project_image;
pub(crate) mod psd_image;
// The spike, which exists to be run once by hand and answered, and which the application itself
// never calls — so it is not compiled outside a test build at all. See the module's own
// documentation for what it measured.
#[cfg(test)]
pub(crate) mod source_reader_spike;
pub(crate) mod svg_preview;
pub(crate) mod tone_map;
pub(crate) mod video_player;
pub(crate) mod webp_image;
pub(crate) mod wic_image;
