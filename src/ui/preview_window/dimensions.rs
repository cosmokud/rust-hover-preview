//! How big a file wants to be: every box and dimension question, the room a preview has to
//! fit into, and the size a renderer is asked for at the scale of the display.

use super::*;
use crate::text::text_paint::TextMetrics;

/// The size of the picture a design document is previewed from: the page an installed
/// render engine drew for it, the document's own size for a Photoshop file, and the size of
/// the picture a project container, a CorelDRAW document or a PostScript one holds.
///
/// What is neither of those is either a CorelDRAW document of the older shape — a RIFF
/// container holding a bitmap rather than a zip holding a file — or the encapsulated
/// PostScript a document was saved as before Illustrator wrote PDFs, and each reader
/// answers for what a file is rather than for what it is called.
///
/// A file none of them will answer for reports no size, which is how a design document
/// this app has no reader for comes to show nothing at all rather than a picture of some
/// other format's making.
pub(super) fn design_dimensions(path: &Path) -> Option<(u32, u32)> {
    // The page the engine drew is measured first, where there is one: what it holds is the
    // document drawn — a page, sharp at whatever size the preview is shown at — rather than
    // a picture of the document that its application kept at some smaller size. Nothing is
    // asked of the engine here: a page it has not drawn is the preview loop's to ask for.
    if let Some(page) = libreoffice_render::rendered_page(path) {
        return pdf_preview::page_dimensions(&page);
    }

    if psd_image::is_psd_file(path) {
        return psd_image::dimensions(path);
    }

    project_image::dimensions(path).or_else(|| eps_image::dimensions(path))
}

/// The box a comic is placed at: the first plate's own size, and nothing at all for a container
/// with no plate in it.
///
/// It is the one box of the book kind that is read out of the file rather than out of a page
/// something drew: the plate is inside the container and this side is the reader, so the two
/// answers are the size of that plate and nothing at all — and nothing is a hover that shows no
/// preview and starts no engine, which is what a box of text under a comic's name gets (see
/// `comic_preview`).
///
/// The size is read once per version of the comic, because it costs a walk of the container's
/// own table of contents and a read of the plate it names rather than a header read of a file.
/// Both of those are felt on a comic of any size, so the read is taken off the preview thread
/// and the hover is laid out as the wait for it (see `measured_off_the_tick`).
pub(super) fn comic_box(path: &Path) -> Option<(u32, u32)> {
    let source = path.to_path_buf();

    measured_off_the_tick(
        path,
        MeasureScope::File,
        (office_preview::WAITING_BOX, office_preview::WAITING_BOX),
        move || comic_preview::dimensions(&source),
    )
}

/// The size a drawing asks to be shown at: what the preview inside an `.eps` is of, or what
/// a metafile's own header declares.
pub(super) fn vector_dimensions(path: &Path) -> Option<(u32, u32)> {
    if metafile_image::is_metafile_name(path) {
        return metafile_image::dimensions(path).or_else(|| eps_image::dimensions(path));
    }

    eps_image::dimensions(path).or_else(|| metafile_image::dimensions(path))
}

/// The box a video hover is placed at.
///
/// A video is measured by a probe — `ffprobe` and a cropdetect pass, two external
/// processes — and the probe is what a hover waits for when its file has not been
/// measured yet: the answer then is the waiting box, which is the box every wait is shown
/// in, and the layout that follows the probe's own replay reads the size here instead
/// (see `video_probe_due`). What the probe answered when it does answer is one of two
/// things, and the file's own lack of an answer is neither: a shape is the shape, and a
/// file with nothing to measure is placed at the 16:9 box FFmpeg's player is handed a
/// file it could not measure at — while a file the media engine would not open is answered
/// with no size at all, which is how the layout drops it rather than putting up a box
/// nothing would be drawn into. That is the only case the engine's answer is read for here,
/// and it is a case only a machine with no FFmpeg on it reaches.
pub(super) fn video_box(path: &Path) -> Option<(u32, u32)> {
    match cached_video_geometry(path) {
        Some(ProbedGeometry::Measured(geometry)) => Some((geometry.width, geometry.height)),
        Some(ProbedGeometry::Unmeasurable) => match video_route(path) {
            VideoRoute::Ffplay => Some((1920, 1080)),
            // No player will take the file: the engine has turned it down on a machine where
            // it was the only player there is, or the name is one only the `[ffmpeg]` list
            // carries on such a machine. That is a file with no preview at all, and the layout
            // is answered with no size so the hover is dropped rather than laid out into a box
            // nothing would be drawn into (see `video_route`).
            VideoRoute::MediaEngine | VideoRoute::NoPreview => None,
        },
        // Not probed yet: the wait for the probe, which is the box the hover is placed
        // in until the answer lands and the hover is replayed.
        None => Some((office_preview::WAITING_BOX, office_preview::WAITING_BOX)),
    }
}

/// The picture a video's preview is drawn from: the size the file is shown at, and the part of
/// the frame that picture is cut from where the probe settled on a crop, in the frame's own
/// pixels.
///
/// It is the one answer both players are given, each in its own terms — FFmpeg's player as a
/// filter on its command line, and the media engine as the source rectangle of its frame
/// transfer, which is what it is normalized over here (see `video_player::Picture`) — and it is
/// also what settles who scales a preview shown above 100%: the picture is what the engine is
/// asked for at its own size, and the box is what this side scales it into (see
/// `video_player::play`). The box is placed at the picture's shape, so a whole frame drawn into
/// one is scaled down to fit and padded with the engine's border colour on the sides the file's
/// own bars leave over — which is the black bar a preview of a file with a crop was growing
/// down two of its edges.
///
/// It is read from the cache and never probed for, for the reason `video_box` reads a shape
/// from there: this is asked on the preview thread, where the probe's two processes are the one
/// wait that must not happen. A file the cache has no answer for is handed the box itself as its
/// picture, which is a picture the size of the preview: nothing for either side to scale, and the
/// arrangement every video had before this side could scale one.
pub(super) fn probed_picture(path: &Path, width: u32, height: u32) -> video_player::Picture {
    match cached_video_geometry(path) {
        Some(ProbedGeometry::Measured(geometry)) => video_player::Picture {
            width: geometry.width,
            height: geometry.height,
            crop: geometry.crop.map(|crop| video_player::Crop {
                x: crop.x,
                y: crop.y,
                width: crop.width,
                height: crop.height,
                frame_width: geometry.frame_width,
                frame_height: geometry.frame_height,
            }),
        },
        _ => video_player::Picture {
            width,
            height,
            crop: None,
        },
    }
}

/// How long a video plays, as the probe that measured it read. A container that does not say is
/// answered with nothing, which is a transport bar with no length to draw a playhead against.
pub(super) fn video_duration(path: &Path) -> Option<f64> {
    match cached_video_geometry(path) {
        Some(ProbedGeometry::Measured(geometry)) => geometry.duration,
        _ => None,
    }
}

/// The subtitle file lying beside this file, as the probe that measured the film found it, or
/// nothing where the folder holds none.
///
/// It is read from the cache and never looked for here, which is the whole of what it is for:
/// finding it is a `read_dir` of the film's own folder (`video_launch::sidecar_for`), and this is
/// asked on the preview thread — inside the launch, so a seek, a resize, a volume change and a
/// track change would each have paid the walk again. The probe resolves it once per file and
/// version, beside the two processes it already runs, and this is where that answer is read
/// (see `probe_video_geometry`).
///
/// It is asked of by the tests rather than by the launch, because the launch already holds the
/// geometry it would name — it read the same entry a few lines above for the crop — so it takes
/// the field out of that answer rather than paying a second lookup for it (see
/// `start_video_playback`). This is the named way to ask, so that the reader is a thing with a
/// name and not a field read spelled out at each use.
#[cfg(test)]
pub(super) fn video_sidecar(path: &Path) -> Option<PathBuf> {
    match cached_video_geometry(path) {
        Some(ProbedGeometry::Measured(geometry)) => geometry.sidecar,
        _ => None,
    }
}

/// A file's subtitle streams, as the probe that measured it read them, or none at all for a file
/// the probe has no answer for.
///
/// "No answer" is answered as *no subtitle streams* rather than as nothing, because the two
/// callers want the same thing out of it and only one of them can tell them apart: a relaunch
/// naming `-sst s:0` at a file the probe never read is a player asked for a track in a file it
/// has not been given, and a track this app has no idea exists is the one thing a specifier must
/// never be guessed at. So an unprobed file plays with the player's own choice, which is what it
/// has always done, and is re-seeded the moment the probe answers (see `PinTransport::subtitle`).
pub(super) fn video_subtitles(path: &Path) -> SubtitleStreams {
    match cached_video_geometry(path) {
        Some(ProbedGeometry::Measured(geometry)) => geometry.subtitles,
        _ => SubtitleStreams { count: 0, first: 0 },
    }
}

/// Whether the small files this app copied out of this film's own subtitle tracks are ready to
/// draw, as the probe's answer for the film holds them.
///
/// What asks is the pin's take-up: the `ffplay` a hover began is adopted by the pin as it stands
/// (see `take_up_pinned_window`), so a player begun while the copy was still coming is one a
/// pinned window would draw on without the subtitles — and the take-up begins it again where
/// this answers yes and the player was one of those (see `reload_adopted_subtitles`). A film no
/// probe has answered for is answered `false` here on the same terms as `video_subtitles`: an
/// answer nothing read is not a copy, and a player begun for such a film has nothing to correct.
pub(super) fn video_copy_ready(path: &Path) -> bool {
    match cached_video_geometry(path) {
        Some(ProbedGeometry::Measured(geometry)) => geometry.derived.is_some(),
        _ => false,
    }
}

/// Whether this file is a video the probe has not answered for yet.
///
/// A hover for one cannot be laid out as a video — the layout has no shape to place — so
/// it is laid out as the wait for the probe and replayed when the answer lands (see
/// `video_probe` in the preview loop). A file the probe has already answered for is not a
/// wait, whatever the answer was: an unmeasurable video is a video with a fallback box,
/// not one to be probed again on every hover.
pub(super) fn video_probe_due(hover: &HoverFacts) -> bool {
    hover.is_video() && cached_video_geometry(hover.probe.path()).is_none()
}

/// The box a PDF page asks for, measured off the preview thread.
///
/// A page's size is read out of the document, and the PDF engine opens it whole to read it —
/// which on a book of a thousand pages is a read that can be felt — so it is measured the way a
/// video's shape is: the hover is laid out as the wait and replayed when the answer lands (see
/// `measured_off_the_tick`).
pub(super) fn pdf_page_box(path: &Path) -> Option<(u32, u32)> {
    let source = path.to_path_buf();

    measured_off_the_tick(
        path,
        MeasureScope::File,
        (office_preview::WAITING_BOX, office_preview::WAITING_BOX),
        move || pdf_preview::page_dimensions(&source),
    )
}

/// The box a document the engine draws asks for: the size its own markup declares, read the
/// same way and for the same reason — a document is read and parsed whole to be measured.
pub(super) fn svg_box(path: &Path) -> Option<(u32, u32)> {
    let source = path.to_path_buf();

    measured_off_the_tick(
        path,
        MeasureScope::File,
        (office_preview::WAITING_BOX, office_preview::WAITING_BOX),
        move || svg_preview::measure(&source),
    )
}

/// The box a specimen is drawn in, held for a file that parses as a font.
///
/// The box is this app's own — a font has no size it asks to be drawn at — so what the measure
/// answers is whether the file is a font at all, and that costs a read of it: a collection is
/// read whole to reach the face the specimen shows (see `font_preview::probe`).
pub(super) fn font_box(path: &Path) -> Option<(u32, u32)> {
    let source = path.to_path_buf();

    measured_off_the_tick(
        path,
        MeasureScope::File,
        (office_preview::WAITING_BOX, office_preview::WAITING_BOX),
        move || {
            font_preview::probe(&source)
                .map(|_| (font_preview::SPECIMEN_WIDTH, font_preview::SPECIMEN_HEIGHT))
        },
    )
}

/// The box a vector drawing asks for, measured the same way: what a metafile declares is read
/// out of the whole file, which a drawing of any size takes with it (see
/// `metafile_image::dimensions`).
pub(super) fn vector_box(path: &Path) -> Option<(u32, u32)> {
    let source = path.to_path_buf();

    measured_off_the_tick(
        path,
        MeasureScope::File,
        (office_preview::WAITING_BOX, office_preview::WAITING_BOX),
        move || vector_dimensions(&source),
    )
}

/// The box a listing asks for, measured off the preview thread: an archive's table of contents
/// is a read that can be felt — every entry of a zip walked, a `.tar.gz` inflated to reach one
/// — and the wait for it is the same kind of wait (see `measured_off_the_tick`).
///
/// It is asked of the archives this app reads itself. An archive an engine lists is measured by
/// `archive_box` instead: what that side waits for is the engine's own listing, and a listing
/// that has been remembered is a page measured out of memory rather than out of a file (see
/// `peazip_box`).
pub(super) fn archive_box_off_the_tick(
    path: &Path,
    bounds: ScreenBounds,
    dpi: u32,
) -> Option<(u32, u32)> {
    let source = path.to_path_buf();
    let cap_width = (bounds.right - bounds.left).max(1) as u32;
    let cap_height = bounds.height().max(1) as u32;
    let options = current_archive_options();

    measured_off_the_tick(
        path,
        MeasureScope::Room {
            cap_width,
            cap_height,
            dpi,
            theme: options.theme,
            font_scale_percent: options.font_scale_percent,
        },
        (office_preview::WAITING_BOX, office_preview::WAITING_BOX),
        move || archive_preview::measure(&source, cap_width, cap_height, dpi, options),
    )
}

/// The box a sound's card asks for, with the probe that fills it beside it on the same thread.
///
/// Two things a hover on a sound waits for, and both of them are here. The first is the probe:
/// whether this machine has anything that plays the file at all, and what the file says about
/// itself — a source reader for the engine's own decoders, an `ffprobe` pass for FFmpeg's —
/// and a file neither of them can play is a hover answered with nothing rather than with a card
/// of facts nothing will ever play. The second is the card's own layout, which is a page of
/// text wrapped to the room the Audio Scaling setting gives it (see `audio_box_room`).
///
/// Both are off the preview thread, and what the hover waits in meanwhile is the spinner: a
/// probe is a process, and the box it answers with is what the replayed hover is laid out at
/// (see `measured_off_the_tick`). The wait is laid out at the room itself, so that the
/// spinner stands at the box the card will take and does not jump size when the answer
/// lands.
pub(super) fn audio_box(path: &Path, bounds: ScreenBounds, dpi: u32) -> Option<(u32, u32)> {
    let source = path.to_path_buf();
    let room = audio_box_room(bounds, current_audio_scale(), dpi);
    let options = current_audio_options();

    measured_off_the_tick(
        path,
        MeasureScope::Room {
            cap_width: room.width,
            cap_height: room.height,
            dpi,
            theme: options.theme,
            font_scale_percent: options.font_scale_percent,
        },
        (room.width, room.height),
        move || {
            // What the machine has for the file, asked once per file and version and held for
            // the hovers that follow — a file the engine will not play costs one probe rather
            // than one per hover.
            if matches!(audio_track::probed(&source), Probed::NotAsked) {
                let probed = probe_audio_track(&source);
                audio_track::remember(
                    &source,
                    match &probed {
                        Some(track) => Probed::Track(track.clone()),
                        None => Probed::Nothing,
                    },
                );
            }

            // Nothing: this is the card a *hover* is measured with, and a hover's own window is a window
            // nobody is in (see `audio_card`).
            let card = audio_card(&source, None, None, 0, None)?;
            audio_preview::measure(&card, room.width, room.height, dpi, options)
        },
    )
}

/// Get original dimensions of media for positioning calculations, of a file whose answer the
/// caller already has.
///
/// It is asked of the answer rather than of the path because the caller is `media_dimensions`,
/// which was handed the same answer a few lines above and used to read the file's directory
/// entry again to get it (see `HoverFacts`).
pub(super) fn get_media_dimensions_of(hover: &HoverFacts, path: &PathBuf) -> Option<(u32, u32)> {
    // What the file's content says it is comes ahead of what its name does, where the two
    // disagree: the box a file is placed at is the box of the kind its content belongs to,
    // and a format no kind previews is placed nowhere at all — see `content_type`.
    match hover.route.content {
        crate::formats::content_type::Content::Kind(kind) => {
            return media_dimensions_of_kind(kind, path)
        }
        crate::formats::content_type::Content::Foreign => return None,
        crate::formats::content_type::Content::Unknown => {}
    }

    if hover.is_video() {
        return video_box(path);
    }

    // A PDF is measured from its own first page; one that cannot be read as a
    // PDF reports no dimensions, which drops the preview instead of guessing.
    if pdf_preview::is_pdf_preview(path) {
        return pdf_page_box(path);
    }

    // And a comic, whose page is a picture inside the container: it is measured where the hook asks
    // it, beside the PDF, because the two are the same kind of preview — a page of a book — and are
    // told apart by the reader rather than by the user. A box with no plate in it is measured as
    // nothing, which is the hover that shows no preview at all (see `comic_box`).
    if previewed_as(path, PreviewType::Ebook) {
        return comic_box(path);
    }

    // An Office document is measured from the page that has been drawn for it — by Office
    // where its own application is installed, and by the render engine beside it where it is
    // not. A document with neither is measured as the page it is about to get — while a page
    // is coming, which is the only case where one is.
    if previewed_as(path, PreviewType::Document) {
        return office_preview::measure(path);
    }

    // A document this app hands to a render engine is measured from the page that engine
    // drew, and a page not drawn yet is the wait for one: the layout places the spinner's
    // own box, the preview loop asks the engine for the document, and the hover is replayed
    // when the page lands (see `libre_render_is_due`). Nothing is converted here, and
    // nothing is waited on — a launch on this thread is a preview, a tray and a pointer
    // held for as long as the engine takes, which is what a document the engine cannot draw
    // never ends.
    //
    // It is asked where the hook asks it — after the office list, ahead of the design list
    // — because a name can sit in two lists: CorelDRAW is a design document to this app and
    // a drawing to the engine, and it is the engine that draws it (see `libre_formats`).
    if previewed_as(path, PreviewType::Libre) {
        return libre_box(path);
    }

    // And a picture the ImageMagick engine develops, measured where the hook asks it: after
    // the documents an engine draws, ahead of the design, vector, font and image lists, none
    // of which would have claimed a `.nef` anyway. What is measured is the picture the engine
    // wrote, and one it has not written yet is the wait for it (see `magick_box`).
    if previewed_as(path, PreviewType::Magick) {
        return magick_box(path);
    }

    // And a book the ebook engine converts, measured where the hook asks it: beside the listing
    // engine and the document engines, none of whose lists would have claimed a `.mobi` anyway.
    // What is measured is the page the engine wrote, and one it has not written yet is the wait
    // for it (see `calibre_box`).
    if previewed_as(path, PreviewType::Calibre) {
        return calibre_box(path);
    }

    // A design document is measured from the picture it is previewed from — the merged
    // image at the end of a Photoshop file, or the picture a project container holds —
    // and a file neither reader will answer for reports no size, which is how it comes
    // to show nothing at all rather than a box nothing would be drawn into.
    //
    // It is asked ahead of the two kinds below it because the gate asks it ahead of them:
    // a name written into the design list as well as into the vector or font list is a
    // design document, and the two questions further down would report a size for another
    // kind than the one the tray was asked to switch.
    if previewed_as(path, PreviewType::Design) {
        return design_dimensions(path);
    }

    // An SVG is measured from the document rather than from a header: the size it asks
    // to be drawn at is the size the layout places, and the engine draws it at whatever
    // box comes out of that. It is asked ahead of the readers of the other half of its
    // kind — the vector list names a document beside the metafiles, and neither of those
    // readers would take one — and ahead of the `Images` gate, which is the other list
    // that may have claimed it. A document is its own kind either way: what draws one is
    // not a decoder, and the switch for it is not the switch for pictures. A file of a
    // kind that is switched off reports no size, which is how the layout drops its
    // preview.
    if svg_preview::is_svg_file(path) {
        if !PreviewType::Vector.enabled() {
            return None;
        }

        // The engine is what draws a document, so a machine without one — or a spell
        // the engine has stood down for — has no document preview at all. Reporting no
        // size is what keeps a hover from opening a box nothing would be drawn into,
        // and it costs no read of the file.
        if !webview_preview::can_draw() {
            return None;
        }

        return svg_box(path);
    }

    // A vector drawing is measured from the records it holds: what an `.eps` keeps a
    // preview of, or what a metafile's own header declares its drawing to be.
    if previewed_as(path, PreviewType::Vector) {
        return vector_box(path);
    }

    // A font is measured at a box of this app's own rather than by anything the file says:
    // a font has no size it asks to be drawn at — what it holds is outlines — so the box is
    // the shape a specimen wants, and the share of the display `font_scale` names decides
    // how large that box is shown. `Fonts` is the gate, and a file that will not parse as a
    // font reports no size at all: that is how a `.ttf` that is something else comes to show
    // nothing rather than a page of another font's glyphs.
    if previewed_as(path, PreviewType::Fonts) {
        if !webview_preview::can_draw() {
            return None;
        }

        return font_box(path);
    }

    // Whatever is left is a picture, so the `Images` gate is what decides it.
    if !PreviewType::Images.enabled() {
        return None;
    }

    picture_dimensions(path)
}

/// The box for one of the kinds a file's content answered with — see `content_type`.
///
/// It is the arm the chain above would have taken had the file been named what its content
/// says it is, and each arm is what that arm measures: the same reader, asked of the file
/// itself rather than of the name it is under. The gate is asked here as the chain asks it
/// per kind, so a kind switched off has no box and comes down the way it always does.
///
/// Two kinds have no size of their own: an archive listing and a text preview are both laid
/// out from the frame the loader paints rather than from anything their file says.
pub(super) fn media_dimensions_of_kind(kind: PreviewType, path: &PathBuf) -> Option<(u32, u32)> {
    if !kind.enabled() {
        return None;
    }

    match kind {
        PreviewType::Videos => video_box(path),
        // A sound is measured against the room it is drawn in — the card is a page of text
        // wrapped to the box it is given — so it is asked where the bounds and the DPI are,
        // beside the drawn kinds and not here (see `media_dimensions`).
        PreviewType::Audio => None,
        PreviewType::Ebook => pdf_page_box(path),
        PreviewType::Archives | PreviewType::Peazip => None,
        // A page of HTML the engine draws is the exception among the kinds that have no size
        // of their own: it is drawn in a box of this app's rather than painted into one (see
        // `html_page_box`).
        PreviewType::Text => html_is_engine_drawn(path).then(html_page_box),
        PreviewType::Document => office_preview::measure(path),
        PreviewType::Libre => libre_box(path),
        PreviewType::Magick => magick_box(path),
        PreviewType::Calibre => calibre_box(path),
        PreviewType::Design => design_dimensions(path),
        PreviewType::Vector => {
            if svg_preview::is_svg_file(path) {
                webview_preview::can_draw().then(|| svg_box(path)).flatten()
            } else {
                vector_box(path)
            }
        }
        PreviewType::Fonts => webview_preview::can_draw()
            .then(|| font_box(path))
            .flatten(),
        PreviewType::Images => picture_dimensions(path),
    }
}

/// The box a document the render engine draws is placed at: the page it has already drawn
/// for this version of the document, the wait for one that is on its way, and nothing at
/// all for a document the engine has turned down or for a machine with no engine to draw
/// one with.
pub(super) fn libre_box(path: &Path) -> Option<(u32, u32)> {
    if !libreoffice_render::available() {
        // Nothing to draw it with, so there is nothing to show: a machine without the
        // engine shows no preview for these names rather than the thumbnail the file
        // carries, which is the whole reason the name is in this list.
        return None;
    }

    if let Some(page) = libreoffice_render::rendered_page(path) {
        return pdf_preview::page_dimensions(&page);
    }

    // A document the engine has already turned down is not one to wait for.
    if libreoffice_render::refused(path) {
        return None;
    }

    Some((office_preview::WAITING_BOX, office_preview::WAITING_BOX))
}

/// The box a book the ebook engine converts is placed at: the page it has already converted for
/// this version of the book, the wait for one that is on its way, and nothing at all for a book the
/// engine has turned down or for a machine with no engine to convert one with.
///
/// It is the shape `libre_box` has, and it is the same question: what is placed is a page the
/// engine wrote rather than a size the file asks for — a book holds no page of its own, which is
/// the whole reason it is handed to an engine — so the box is the page's own and, until there is
/// one, the spinner's. Nothing is converted here: a book the engine has answered nothing for yet is
/// the wait, and the loop asks for the page the moment there is a hover to ask for it (see
/// `request_calibre_render`).
pub(super) fn calibre_box(path: &Path) -> Option<(u32, u32)> {
    if !calibre_render::available() {
        // Nothing to convert it with, so there is nothing to show: a machine without the engine
        // shows no preview for these names rather than the first page of the markup a `.fb2` is,
        // which is the whole reason the name is in this list.
        return None;
    }

    if let Some(page) = calibre_render::rendered_page(path) {
        return pdf_preview::book_page(&page).map(|book| book.size);
    }

    // A book the engine has already turned down is not one to wait for.
    if calibre_render::refused(path) {
        return None;
    }

    Some((office_preview::WAITING_BOX, office_preview::WAITING_BOX))
}

/// The box a picture the ImageMagick engine develops is placed at: the size the engine
/// developed it at, the wait for one that is on its way, and nothing at all for a file the
/// engine has turned down or for a machine with no engine to develop one with.
///
/// The size is the picture's own — read from the header of the bytes the engine wrote, which
/// is the size the preview is drawn at under the picture scale — and it is what the layout
/// places and what the frame is keyed by. Nothing is converted here: a file the engine has
/// developed nothing for yet is the wait, and the loop asks for the picture the moment there
/// is a hover to ask for it (see `request_magick_render`).
///
/// The wait is a spinner's box and says nothing about how large the picture will be drawn,
/// which is why it is not the room the engine is asked for: a picture is only this size
/// *after* the conversion, so the room it is developed in is a ceiling on the size its
/// preview can ever be drawn at (see `PendingLoad::room`).
pub(super) fn magick_box(path: &Path) -> Option<(u32, u32)> {
    if !imagemagick_render::available() {
        // Nothing to develop it with, so there is nothing to show: a machine without the
        // engine shows no preview for these names rather than the picture the camera left
        // inside the file, which is what the shell's own thumbnail is for.
        return None;
    }

    if let Some(size) = imagemagick_render::dimensions(path) {
        return Some(size);
    }

    // A file the engine has already turned down is not one to wait for.
    if imagemagick_render::refused(path) {
        return None;
    }

    // A raw sample dump is the one file the engine cannot measure for itself: what shape it is
    // comes from its own length rather than from a conversion, which is arithmetic this side can
    // do — and a dump whose length does not settle a shape is a file with no preview rather than
    // one to wait for (see `raw_geometry`).
    if imagemagick_render::is_raw_sample(path) {
        return imagemagick_render::raw_geometry(path);
    }

    Some((office_preview::WAITING_BOX, office_preview::WAITING_BOX))
}

/// A picture's own size, read the way this app reads one: its own header, and the codec
/// Windows keeps for the picture formats this app's decoder has no reader for.
pub(super) fn picture_dimensions(path: &PathBuf) -> Option<(u32, u32)> {
    image_dimensions_with_header_check(path)
}

/// Whether this hover is the wait for a page rather than a preview of one: an Office
/// document, a document the render engine draws or a book the ebook engine converts, with
/// nothing drawn for it yet and a page on the way.
///
/// A hover like that is placed by the spinner's own box, a pointer gap off the hand,
/// rather than by the size a preview would take: it is the wait for the file under the
/// hand, which belongs at the hand. The page is laid out again by the replay that arrives
/// with it, so nothing here has to guess how large it will be.
pub(super) fn page_is_on_the_way(path: &Path) -> bool {
    // A file whose bytes are another kind is not a document a page is coming for, which is
    // the same question the render tier is asked before it is asked for one: a picture left
    // under a document's name is drawn here, and placing it at the pointer as the wait for a
    // page would be a preview waiting for nothing (see `content_names_another_kind`).
    let office = previewed_as(path, PreviewType::Document)
        && !content_names_another_kind(path, PreviewType::Document)
        && matches!(
            office_preview::source_kind(path),
            office_preview::SourceKind::None
        );

    office
        || libre_render_is_due(path)
        || magick_render_is_due(path)
        || peazip_render_is_due(path)
        || calibre_render_is_due(path)
}

/// The size the layout should place and scale a preview from.
///
/// A text file has no size of its own, so the box its first screenful wants is
/// measured here, bounded by the display it will be shown on. That makes the
/// measurement an intrinsic size in the same sense a PDF page's is: the layout
/// can fit it into the space beside the cursor, and the text renderer is handed
/// the box that comes out of that.
pub(super) fn media_dimensions(
    path: &PathBuf,
    bounds: ScreenBounds,
    dpi: u32,
) -> Option<(u32, u32)> {
    media_dimensions_of(&HoverFacts::read(path), path, bounds, dpi)
}

/// The same, of a file whose answer the caller already has — which is the form the `Show` arm
/// asks it in, so that installing a hover reads one file's directory entry rather than the two
/// this and `get_media_dimensions` used to read between them (see `HoverFacts`).
pub(super) fn media_dimensions_of(
    hover: &HoverFacts,
    path: &PathBuf,
    bounds: ScreenBounds,
    dpi: u32,
) -> Option<(u32, u32)> {
    // What the file's own bytes say it is comes first, as it does for the loader that draws
    // it and for the share it is laid out at: a `.txt` whose bytes are a picture is measured
    // as the picture it is rather than read as a page of text it is not — which for a file
    // whose bytes are not text is no measurement at all, and a preview that never appears
    // for a file that would otherwise be drawn.
    if let crate::formats::content_type::Content::Kind(kind) = hover.route.content {
        return match kind {
            // The three kinds measured against the room they are drawn in, which is a question
            // this side has the answer to and `media_dimensions_of_kind` does not.
            PreviewType::Text if html_is_engine_drawn(path) && PreviewType::Text.enabled() => {
                Some(html_page_box())
            }
            // And a text file that is not one, which is the arm this group is about: the
            // room it is drawn in is a question `media_dimensions_of_kind` cannot answer.
            PreviewType::Text => text_box(path, bounds, dpi),
            PreviewType::Archives => archive_box_off_the_tick(path, bounds, dpi),
            PreviewType::Peazip => peazip_box(path, bounds, dpi),
            // And the fourth kind that is wrapped to its room: a sound's card is painted at a
            // fixed font size and cut to the box it is given, and what stands in for it until
            // the probe beside it has answered is the waiting spinner.
            PreviewType::Audio => audio_box(path, bounds, dpi),
            _ => media_dimensions_of_kind(kind, path),
        };
    }

    // A page of HTML the engine draws is measured at a box of this app's own, for the reason
    // a specimen is: the page asks for no size, and the box is the one the layout scales by
    // `document_scale` (see `html_page_box`).
    if html_is_engine_drawn(path) && PreviewType::Text.enabled() {
        return Some(html_page_box());
    }

    if hover.is_text() {
        return text_box(path, bounds, dpi);
    }

    if previewed_as(path, PreviewType::Archives) {
        return archive_box_off_the_tick(path, bounds, dpi);
    }

    // And an archive an engine lists, measured where the hook asks it: beside the archive list
    // above it, which is where the two are told apart — a name in that list is read by this app
    // itself, and one in this list is read by an engine. What is measured is the page the
    // engine's listing makes, and a listing that has not come back yet is the wait for one (see
    // `peazip_box`).
    if peazip_formats::is_peazip_file(path) && PreviewType::Peazip.enabled() {
        return peazip_box(path, bounds, dpi);
    }

    // A sound is the fourth kind measured against its room, and the last: what a hover on one
    // asks is a card whose facts a probe has to bring back first (see `audio_box`).
    if hover.is_audio() {
        return audio_box(path, bounds, dpi);
    }

    get_media_dimensions_of(hover, path)
}

/// The room a page of text is measured in: the display's work area cut down to the share
/// `text_scale` asks for.
///
/// A page of text has no size of its own to be drawn at — it is measured at the font size the
/// DPI beside it gives, and the answer is however many rows and columns that font takes in the
/// room it is measured in — so what the setting names is that room rather than a share of
/// anything the file holds. `Fit to Screen` is the whole of the work area, which is the room
/// this kind was given before the setting existed, and a share past the whole is the whole: a
/// page is never measured in more display than the display has.
///
/// It is a ceiling and not a zoom, which is the whole of the difference between this and the
/// scale of a picture: a page that takes less room than the share allows keeps the room it
/// takes, and a `Text Size` large enough to fill the share cannot grow the page past it (see
/// `text_box`).
pub(super) fn text_box_room(bounds: ScreenBounds) -> (u32, u32) {
    let share = CONFIG
        .lock()
        .map(|config| config.text_scale.target_scale().unwrap_or(1.0).min(1.0))
        .unwrap_or(1.0);

    // A column is a whole character and a row a whole line, so each side is rounded to the
    // nearest one rather than cut: the share is a room to measure in, and a room a pixel
    // narrower than it asks for is a page with a column missing from it.
    let width = ((bounds.right - bounds.left).max(1) as f32 * share).round();
    let height = (bounds.height().max(1) as f32 * share).round();

    (width.max(1.0) as u32, height.max(1.0) as u32)
}

/// The font size a sound's card is built at for a share: the default
/// text size at the 10% anchor, scaled by the share's fraction of the
/// anchor and rounded to the nearest whole percent — 63%, 125%, 188%,
/// 250% and 313% for the five shares the menu offers. A share a hand
/// edits past the menu's top end is past the menu but not past the
/// metrics, which honor 1% to 1000% (see `TextMetrics::new`).
pub(super) fn audio_font_scale_percent(scale: PreviewScale) -> u32 {
    let share = scale.target_scale().unwrap_or(1.0).min(1.0);
    let anchor = DEFAULT_AUDIO_SCALE_PERCENT as f32 / 100.0;

    (DEFAULT_TEXT_FONT_SCALE_PERCENT as f32 * share / anchor).round() as u32
}

/// What the Audio Scaling setting answers for a work area: the room a
/// sound's card is laid out over, and the font size the card is built
/// at (see `audio_font_scale_percent`).
pub(super) struct AudioRoom {
    pub(super) width: u32,
    pub(super) height: u32,
    #[cfg_attr(not(test), allow(dead_code))]
    pub(super) font_scale_percent: u32,
}

/// The room a sound's card is laid out over, and the font size the card
/// is built at: the work area of the display the pointer is on times the
/// share the Audio Scaling setting names (see `audio_box`), and the
/// default text size scaled by the share's fraction of the 10% anchor.
///
/// A share past the whole display is the whole display, so a hand-edited
/// value above `100%` cannot put a card wider than the screen. The width
/// room is that share of the width plus twice the card's own margin — the
/// margin being the padding the card is drawn inside, read at the font
/// size the card is built at for the share, so the margin scales with the
/// card the way Windows scaling scales a window's chrome — and the height
/// room is the plain share of the height: the card's height is its own,
/// content-driven at that font, and the share of the height is only the
/// guard that answers a room too short to draw a card in with nothing.
pub(super) fn audio_box_room(bounds: ScreenBounds, scale: PreviewScale, dpi: u32) -> AudioRoom {
    let share = scale.target_scale().unwrap_or(1.0).min(1.0);
    let font_scale_percent = audio_font_scale_percent(scale);

    // The card's own margin, at the font size the card is built at for
    // this share: a compatible DC to measure with, created and dropped
    // here, the way `window_button_band` takes one.
    let dc = unsafe { CreateCompatibleDC(None) };
    let padding = TextMetrics::new(dc, dpi, font_scale_percent)
        .map(|metrics| metrics.padding)
        .unwrap_or(0);
    unsafe {
        let _ = DeleteDC(dc);
    }

    // A column is a whole character and a row a whole line, so each side is
    // rounded to the nearest one rather than cut (see `text_box_room`).
    let width = ((bounds.right - bounds.left).max(1) as f32 * share).round() + (padding * 2) as f32;
    let height = (bounds.height().max(1) as f32 * share).round();

    AudioRoom {
        width: width.max(1.0) as u32,
        height: height.max(1.0) as u32,
        font_scale_percent,
    }
}

/// The box a page of text asks for: as many lines and columns as the room its own setting
/// gives it holds, at the font size its DPI gives them.
pub(super) fn text_box(path: &Path, bounds: ScreenBounds, dpi: u32) -> Option<(u32, u32)> {
    let (cap_width, cap_height) = text_box_room(bounds);

    text_preview::measure(path, cap_width, cap_height, dpi, current_text_options())
}

/// And the box a listing asks for, which is the same shape of question asked of an archive's
/// own table of contents.
pub(super) fn archive_box(path: &Path, bounds: ScreenBounds, dpi: u32) -> Option<(u32, u32)> {
    let cap_width = (bounds.right - bounds.left).max(1) as u32;
    let cap_height = bounds.height().max(1) as u32;

    archive_preview::measure(path, cap_width, cap_height, dpi, current_archive_options())
}

/// The box an archive the PeaZip engine lists is placed at: the box the same listing would be
/// measured at had this app read the archive itself, the spinner's own box while the engine has
/// not answered yet, and nothing at all for a file the engine has turned down or for a machine
/// with no engine to list it.
///
/// The measurement is the archive page's own, and it is the same one either way: a listing is a
/// listing, and where it came from is not something the layout is told (see
/// `archive_listing::listing_for`). What is waited for is a page that does not exist yet — the box
/// an engine's answer has not arrived for says nothing about how large it will be — so it is
/// placed as the spinner, at the pointer's own corner, and laid out again by the replay that
/// arrives with the answer.
pub(super) fn peazip_box(path: &Path, bounds: ScreenBounds, dpi: u32) -> Option<(u32, u32)> {
    if !peazip_render::available_for(path) {
        // Nothing to list it with, so there is nothing to show: a machine without the engine — or
        // with an installation that does not carry the one tool this name is read by — shows no
        // preview for it rather than a page of something else.
        return None;
    }

    if let Some(size) = archive_box(path, bounds, dpi) {
        return Some(size);
    }

    // A file the engine has already turned down is not one to wait for.
    if peazip_render::refused(path) {
        return None;
    }

    Some((office_preview::WAITING_BOX, office_preview::WAITING_BOX))
}

/// A text preview's box, placed for the width it came out with.
///
/// The box a text preview is placed from is measured at the width the *display* can
/// give, and the layout fits that box beside the cursor or the focused item by
/// shrinking it whole. Text does not shrink with it: the lines wrap at the width the
/// box ends up with, and a narrower box needs *more* rows than the proportional
/// height leaves — a long line that took two rows at the display's width takes three
/// in the box that came back — so the frame could show the first of them and the
/// rest of the line was cut off below it. Measuring the document again at the width
/// the box actually has is what gives it the height the text really takes there, and
/// placing the result again is the same rule that put it there the first time.
///
/// The size that measure answered with is handed back with the layout, because the
/// wait that follows the pointer re-places from the size it was given as well: left
/// at the display's own measurement it would step around the name for a height the
/// wrapped frame does not have (see `HoverPlacement`).
///
/// The re-measure is held to the room `text_scale` named rather than to the work area
/// (see `text_box_room`): wrapping at a narrower width takes *more* rows, so a page
/// measured at a share of the display and again at the width the layout gave it would
/// otherwise come back taller than the share allows — and at `Fit to Screen` the room is
/// the whole work area, which is wider and taller than the layout is, so nothing is
/// clamped at the setting's own start.
pub(super) fn text_preview_layout(
    hover: &HoverFacts,
    path: &Path,
    layout: PreviewLayout,
    bounds: ScreenBounds,
    dpi: u32,
    place: impl FnOnce((u32, u32)) -> Option<PreviewLayout>,
) -> (PreviewLayout, Option<(u32, u32)>) {
    if !hover.is_text() {
        return (layout, None);
    }

    let room = text_box_room(bounds);
    let Some(size) = text_preview::measure(
        path,
        layout.preview_w.min(room.0),
        layout.max_height.min(room.1),
        dpi,
        current_text_options(),
    ) else {
        return (layout, None);
    };

    match place(size) {
        Some(placed) => (placed, Some(size)),
        None => (layout, None),
    }
}

/// The effective DPI of the display nearest `(x, y)`, which is what a text preview's
/// font size is scaled by — and what every margin a layout is written around is
/// scaled by (see `logical_px`).
///
/// The display is asked and not the window the point happens to be over, and the display's
/// answer is kept for one display and asked about again for the next: both of those are the
/// adapter's, and the question of what to do where the machine cannot name a display is the
/// same one a placement asks and so is asked in the same place (see `displays::display_at`).
pub(crate) fn monitor_dpi_from_point(x: i32, y: i32) -> u32 {
    dpi_at(x, y)
}
