//! The scale a preview is drawn at: the display's own scale, what the kind of file allows,
//! and the reductions a box too small for the full one takes.

use super::*;

/// Every scale a hover is laid out by, read from the configuration together so that the
/// measure of a file and the render that follows it cannot disagree about the size.
#[derive(Debug, Clone, Copy)]
pub(super) struct HoverScales {
    /// The share of its own size a picture is drawn at, which is also the scale every
    /// format that is none of the others below keeps.
    pub(super) picture: PreviewScale,
    /// The share of its own size a video is drawn at.
    pub(super) video: PreviewScale,
    /// The share of its own size an animated picture is drawn at.
    ///
    /// It is read apart from the picture scale beside it — and read for an animated
    /// file whether it moves or not, since a single-frame GIF or WebP is a picture —
    /// so that what moves is drawn at the size one wants it at rather than at the size
    /// one wants a photograph at.
    pub(super) animated: PreviewScale,
    /// The share of the display a PDF — the `Ebook` kind — page is drawn at.
    pub(super) ebook: PreviewScale,
    /// The share of the display a page of the `Document` kind is shown at: what an installed
    /// engine hands back is a page, so the share is of the room the display has rather than of
    /// a size the file asks for, exactly as a PDF page's is.
    pub(super) document: PreviewScale,
    /// The share of the display a font specimen is drawn at.
    pub(super) font: PreviewScale,
    /// The share of the display a design document is drawn at.
    ///
    /// What a design document is previewed from is the picture its own format keeps of
    /// the whole thing, so what the share is of is the room the display has rather than
    /// the size that picture happens to be — the same question the document scales above
    /// answer, and a setting of its own because a drawing wants a different share of the
    /// screen from a page or a specimen.
    pub(super) design: PreviewScale,
    /// The share of the display a vector drawing is replayed over.
    ///
    /// A drawing is not a bitmap: the records are played again at whatever size the box
    /// asks for, so the room the display has is free quality the way a document's is, and
    /// the share is of that room.
    pub(super) vector: PreviewScale,
}

impl HoverScales {
    /// Every scale one hover is laid out by, read out of a configuration the caller already
    /// holds — which is what `HoverFacts` does, so that installing a hover reads the
    /// configuration once for its scales and for the answers below rather than once each.
    pub(super) fn of(config: &crate::config::config::AppConfig) -> Self {
        Self {
            picture: config.preview_scale,
            video: config.video_scale,
            animated: config.animated_scale,
            ebook: config.ebook_scale,
            document: config.document_scale,
            font: config.font_scale,
            design: config.design_scale,
            vector: config.vector_scale,
        }
    }
}

/// The scale a preview is laid out and rendered with.
///
/// A PDF page is a vector, so the engine draws it at whatever size it is asked
/// for and a larger preview is sharper text rather than an enlarged raster. The
/// room the display has is therefore the page's size, and a configured percentage
/// below `100%` reduces that size rather than being ignored — the page is no
/// longer enlarged by it either, since enlarging a page is what fit-to-screen
/// already does (see `fit_reduced`). What share of that room a page is drawn at is
/// `ebook_scale`'s to say, and it says the whole of it unless it is asked for less.
///
/// Text is the opposite case: it is drawn at a fixed, display-scaled font size,
/// so enlarging it would only stretch the window around text that stays the same
/// size. `100%` is exactly the rule text wants — never enlarged, reduced only
/// when the space beside the cursor cannot hold it — and the text renderer reads
/// the size it is given as "as many lines and columns as fit".
///
/// An SVG is the same case as a page: it is drawn at whatever size it is asked for,
/// so the room the display has is free quality and the whole of it is what a document
/// is drawn at. What it is asked for is a share of that room rather than a share of
/// the size the file asks for, which is the one thing a document and a picture do not
/// agree on: a picture at `50%` is half of its own size, a document at `50%` is half
/// of the screen. The engine's window is the size that comes out of this and its page
/// fills it, so the setting is the document's size and nothing else: see
/// `webview_preview::frame_page`.
///
/// A page of the `Document` kind is the PDF rule again: whether the application that owns the
/// format exported it or an installed render engine drew it, it is drawn at whatever size it is
/// asked for, at the share of the room `document_scale` names — one setting for both, since it
/// is one question about one shape of preview. The one source that is not a page is the bitmap a
/// workbook is answered with where no page can be exported, and it follows the share the way
/// `bitmap_at_display_scale` reads it.
///
/// A font is the same rule once more, at the share `font_scale` names: the specimen is a
/// page of this app's own — the box `font_preview` measures a font at — and the glyphs are
/// sized from the window the engine draws it in, so a share of the room is a share of the
/// type. A file that will not parse as a font is not measured at all, so a hover onto one
/// never reaches this.
///
/// A design document is the document rule rather than the picture's, at the share
/// `design_scale` names: what its preview is made of is the picture the file keeps of the
/// whole document rather than a picture the file *is*, so the room the display has is what
/// the share is of — the same question a page answers, and a setting of its own because a
/// drawing and a page want different shares of it. See `load_design_preview`.
///
/// A vector drawing is the document rule once more, at the share `vector_scale` names: the
/// records are replayed at whatever size they are asked for, so the room the display has is
/// free quality and a share of it is what the setting means — there is nothing to enlarge
/// and nothing to lose by it. See `load_vector_preview`.
///
/// A video keeps the share of its own size `video_scale` names, which is the picture's
/// rule: what a video's preview is, until the player's window is over it, is its first
/// frame — a bitmap measured the way a picture is — so the share is of the file's own
/// size, and it is a setting of its own because a size that suits a photograph is not
/// always the size one wants to watch a file at. The one thing read beside it is whether
/// the probe has answered yet (see `video_probe_due`).
///
/// An animated picture keeps the share of its own size `animated_scale` names, which is
/// the picture's rule once more — its frames are bitmaps — with the file's own content
/// asked which of the two settings it is under: a `.gif`, `.webp` or `.png` that holds
/// more than one frame is animated and follows `animated_scale`, while one that holds a
/// single frame is a still picture and keeps `preview_scale` like any other. Which one
/// it is is the probe below's answer, and it is asked as one question so that the size
/// a hover is placed at and the size its frames are decoded at are the same answer.
///
/// Every other format keeps the picture scale.
pub(super) fn effective_preview_scale(path: &Path, scales: HoverScales) -> PreviewScale {
    effective_preview_scale_of(&HoverFacts::read(path), scales)
}

/// The share one file is placed at, asked of the answer that hover already has.
///
/// Everything this asks is a field of `hover` or something `hover` holds: the file's kind, the
/// one probe that runs off the thread and is remembered per file and version, and whether the
/// file's own head says it moves. That is the whole of what a preview's scale is — which is why
/// the path is not beside it, and why the scale and the box a preview is laid out at cannot
/// disagree: they are asked of one answer.
///
/// A hover does not come here directly. It goes through `hover_preview_scale_of`, which is this
/// answer with the bitmap rule applied to a video, because a fit enlarges a pinned window's
/// media and must not enlarge a hover's.
pub(super) fn effective_preview_scale_of(hover: &HoverFacts, scales: HoverScales) -> PreviewScale {
    // What the file's own bytes say it is comes first, as it does for the loader that draws
    // it and for the box the layout places it at: a picture under a video's name is laid out
    // at the picture's share, and one under a document's name at the picture's share too.
    // Where the bytes have nothing to say the name decides below, which is every file that
    // is called what it is.
    if let crate::formats::content_type::Content::Kind(kind) = hover.route.content {
        return scale_of_kind(kind, hover, scales);
    }

    // Two questions that are about the run rather than about the file's kind, and both come
    // before it: a page painted to the frame it is given is not scaled within it, and a video
    // whose probe has not answered yet is the spinner rather than a video. Neither can be asked
    // of the kind, which knows nothing about what the run has done so far.
    if hover.is_painted_page() {
        // A listing is a page of text painted to the box it is given, whether this app read the
        // archive itself or an engine listed it, so both are the text rule.
        return scale_of_kind(PreviewType::Text, hover, scales);
    }

    if video_probe_due(hover) {
        // A video that has not been probed yet is a hover that is waiting, and what is on
        // screen for one is the waiting spinner: a wait is placed at the size it is rather
        // than fitted to the display, and what the probe answers is what the replay that
        // follows it is laid out at (see `video_probe_due`).
        return PreviewScale::Percent(100);
    }

    // And the kind decides the rest, asked of the one table every side asks it in: what a hover
    // is measured at is the answer the hook admitted it under and the loader draws it by (see
    // `formats::routing`). A name no list claims is measured as the picture it ends up being
    // decoded as, which is where the loader's own chain sends one.
    scale_of_kind(
        hover.route.named.unwrap_or(PreviewType::Images),
        hover,
        scales,
    )
}

/// The share a *hover* of one file is placed at: the file's own share, with the one thing a
/// video has and a pinned window does not taken back to its own size.
///
/// A hover is a bitmap. What a hover puts on screen is a picture with pixels of its own, and a
/// `Fit to Screen` read for a bitmap is the picture at the size it is rather than the whole of
/// the display — the rule every other picture in this app is laid out by, and the one
/// `bitmap_at_display_scale` exists to state. A video is no different here: until the player's
/// window is over it, what a video's preview *is* is its first frame, and that frame is
/// decoded, stored and handed to the compositor at the box the layout chose, sixty times a
/// second for a film that is playing. Enlarging a 1080p file to a 4K display's whole work area
/// therefore costs a 3840x2160 surface — thirty-three megabytes copied into the layered DIB
/// and thirty-three more handed to the compositor, every frame — to draw a picture that has
/// 1920x1080 of detail in it. Nothing here is ever enlarged to fill the room.
///
/// A configured percentage is left exactly as it was configured, above `100%` included: a user
/// who has asked for `150%` of a video is asking to watch it larger than it is, which is the
/// one enlargement in this app that is a person saying so rather than a rule guessing.
///
/// It is here, and not in `scale_of_kind`'s own video arm, because a pinned window is the
/// exception rather than an oversight: a pin is fitted to the media it shows, and at a fit the
/// media is scaled up to the room as well as down to it — that is what makes a maximize a
/// maximize (see `pinned_media_box`). Clamping in the table would take that away along with the
/// hover's waste, so the two roads are kept apart here, at the two places a hover's scale is
/// asked for.
///
/// And not in `scale_in_room`, which is the narrowest function of all and is the wrong one: it
/// is one function for the hover's box and the pin's box both, so a rule written in it is a rule
/// about both, and it cannot tell a video from the document beside it.
pub(super) fn hover_preview_scale_of(hover: &HoverFacts, scales: HoverScales) -> PreviewScale {
    let share = effective_preview_scale_of(hover, scales);

    // Asked of the answer rather than of the kind, so a video the bytes name under a name the
    // video list does not carry is caught by the same predicate the loader plays it by.
    if hover.is_video() {
        bitmap_at_display_scale(share)
    } else {
        share
    }
}

/// The share a preview of one kind is drawn at.
///
/// It is one place per kind rather than a share written into each arm of the chain above,
/// because two questions arrive here now — what the name says a file is, and what its bytes
/// say — and both have to come out at the same share for the same kind: a picture is drawn
/// at the picture's share whether it is called `tomcat.png` or `tomcat.mp4`, or the same
/// bytes would be two sizes depending on the name they were left under.
pub(super) fn scale_of_kind(
    kind: PreviewType,
    hover: &HoverFacts,
    scales: HoverScales,
) -> PreviewScale {
    let path = hover.probe.path();
    match kind {
        // A page is a vector, so the room the display has is free quality: the setting is
        // the whole of that room unless it asks for less (see `fit_reduced`).
        PreviewType::Ebook => fit_reduced(scales.ebook),

        // A page of HTML the engine draws is the one exception to the arm below: the engine
        // draws it at whatever box it is given rather than painting it at a fixed size, so
        // the share is of the room — the rule a document follows (see `html_page_box`).
        PreviewType::Text if html_is_engine_drawn(path) => fit_reduced(scales.document),

        // Text is drawn at a fixed, display-scaled font size and a listing is painted to
        // the frame it is given, so neither is enlarged or reduced by a setting: the size
        // the box came out at is the size they are drawn at. A text page is measured
        // against the share its own `text_scale` setting asks for rather than at the share
        // written here, so what is left for placement to do is only to shrink a box that
        // the room cannot take. An archive an engine listed is the second of those: the
        // same page, painted the same way, from a listing that came back from somewhere
        // else.
        PreviewType::Text | PreviewType::Archives | PreviewType::Peazip | PreviewType::Audio => {
            PreviewScale::Percent(100)
        }

        // A page of the `Document` kind is the Ebook rule at the share `document_scale` names,
        // and both halves of the kind answer to it: the page an Office document's own
        // application exported, and the page an installed render engine drew. The raster
        // picture a workbook is answered with where no printer can export a page is the
        // exception: it is only as good as the pixels it holds, so it follows the configured
        // share the way an image does rather than being enlarged to fit. And a document with
        // nothing drawn for it yet is placed at the spinner's own size, since a page on the way
        // has no shape to fit.
        PreviewType::Document => match office_preview::source_kind(path) {
            office_preview::SourceKind::None => PreviewScale::Percent(100),
            source if source.may_be_enlarged() => fit_reduced(scales.document),
            _ => bitmap_at_display_scale(scales.document),
        },

        // A document an engine draws is the other half of that kind, and it answers to the same
        // setting: what the engine hands back is a page, not a picture with a size of its own to
        // be scaled from, so the share is of the display the way a PDF page's is.
        PreviewType::Libre => fit_reduced(scales.document),

        // A picture an engine developed is the picture's rule: what comes back is a PNG,
        // which is a bitmap with a size of its own — the size the engine wrote it at — so
        // the share is of that size rather than of the display. It is asked here rather than
        // left to the arm below so that a `.nef` the content named and one the name named
        // come out at the same size (see `effective_preview_scale`).
        PreviewType::Magick => scales.picture,

        // And a book an engine converted keeps the book rule, at the share `ebook_scale` names:
        // what the engine hands back is a PDF, which is a page rather than a picture with a size of
        // its own to be scaled from — so the share is of the display, the same question a PDF page
        // and a page an engine drew answer. It is asked here rather than left to an arm of its own
        // because the setting is the same one: a user who wants their books smaller wants them
        // smaller whichever reader drew one.
        PreviewType::Calibre => fit_reduced(scales.ebook),

        // A design document is a document for this question rather than a picture: what is
        // previewed is the picture the file keeps of the whole of itself, at whatever size
        // that is, so the share is of the display the way a page's or a specimen's is.
        PreviewType::Design => fit_reduced(scales.design),

        // Both halves of the drawing kind: what a document costs to draw and what a
        // drawing costs to replay are the same question, and the setting is the same one.
        PreviewType::Vector => fit_reduced(scales.vector),

        // And the specimen once more: the glyphs are sized from the window the engine draws
        // it in, so a share of the room is a share of the type.
        PreviewType::Fonts => fit_reduced(scales.font),

        // A video that is still being probed is a wait rather than a video, and a wait is
        // placed at the size it is; one that has been probed keeps the share of its own
        // size `video_scale` names, which is the picture's rule — what a video's preview is
        // until the player's window is over it is its first frame, a bitmap.
        PreviewType::Videos => {
            if video_probe_due(hover) {
                PreviewScale::Percent(100)
            } else {
                scales.video
            }
        }

        // A picture keeps the share of its own size, and an animation one of its own. The
        // animated arm is asked last because asking it is the one thing here that reads the
        // file, and a file that has already answered as another kind never pays for it (see
        // `image_is_animated`).
        PreviewType::Images => animated_scale_for(path, scales).unwrap_or(scales.picture),
    }
}

/// The share of a bitmap's own size an animated picture is drawn at, for a file that
/// holds an animation — or `None` for every file this question is not about: one that
/// is not an animated format, one whose animation scale is the same as its picture
/// scale, and one that turned out to hold a single frame after all.
///
/// The two scales being equal is the cheap early exit, and it is the whole reason this
/// can be asked on every hover: where the answer cannot change what is drawn, the file
/// is not read at all, which is the state a fresh install is in because both settings
/// start at `100%`. Only a user who has taken the trouble to give animations a size of
/// their own pays for the probe, and what that probe costs is the file's own structure
/// rather than its pixels (see `image_is_animated`).
pub(super) fn animated_scale_for(path: &Path, scales: HoverScales) -> Option<PreviewScale> {
    if scales.animated == scales.picture {
        return None;
    }

    image_is_animated(path).then_some(scales.animated)
}

/// Whether a picture file holds a sequence rather than a single frame: a GIF with more than one
/// frame, a WebP with an animation chunk, a PNG with an animation control chunk, or one of the
/// sequences this app has no reader for.
///
/// Nothing is decoded, and the question is asked of the file's own bytes rather than of its
/// name: a still `.gif` and an animated `.png` are both files a name cannot settle, which is the
/// same reason the loader asks the file and not its extension what it is (see
/// `head::PictureNature`). It is asked of the head, which the router has read already by the
/// time a picture reaches this question, so the answer costs a lookup rather than a read.
pub(super) fn image_is_animated(path: &Path) -> bool {
    crate::formats::head::picture_nature(path).is_some_and(|nature| nature.moves)
}

/// The scale a bitmap is drawn at, for a share of the display it is asked to follow.
///
/// The room the display has is free quality for a source that is drawn at the size it
/// is asked for, and it is not for a bitmap: enlarging one only stretches the pixels it
/// holds, and a worksheet's corner is a few hundred pixels across rather than a
/// screenful. So a share of the display is read for a bitmap as the same share of its
/// own size, and the whole of the display — a fit — as the bitmap at the size it is,
/// which is what `100%` means for a picture. Nothing here is ever enlarged to fill the
/// room.
pub(super) fn bitmap_at_display_scale(display_scale: PreviewScale) -> PreviewScale {
    match display_scale {
        PreviewScale::Percent(percent) => PreviewScale::Percent(percent),
        PreviewScale::FitToScreen | PreviewScale::FitToScreenReduced(_) => {
            PreviewScale::Percent(100)
        }
    }
}

/// The room the display has, reduced to the configured share of it where the
/// configuration asks for less than the whole of it.
///
/// A source drawn at any size it is asked for — a PDF page, an SVG document, a page
/// Office rendered — is laid out at fit-to-screen, because the display's room is free
/// quality there. A configured percentage at or above `100%` asks for at least
/// that room and is answered with it, so those settings are one setting for such
/// a source; one below `100%` is a size the user picked, and is answered by
/// reducing the fitted size — `50%` halves it — rather than being ignored.
pub(super) fn fit_reduced(preview_scale: PreviewScale) -> PreviewScale {
    match preview_scale {
        PreviewScale::Percent(percent) if percent < 100 => {
            PreviewScale::FitToScreenReduced(percent)
        }
        _ => PreviewScale::FitToScreen,
    }
}
