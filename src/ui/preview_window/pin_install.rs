//! Putting a loaded media into a pin: the answer off the worker thread, the relayout it asks
//! for, and the install that gives it the window.

use super::*;

/// The wait a pinned window is in, in the form a repaint can draw: the millisecond its arc
/// began turning at, counted from [`PIN_ARC_BASE`], and `0` for a window that is waiting for
/// nothing.
///
/// It is a global rather than a field of the wait because a repaint is not the loop: the arc
/// is drawn by whatever paint is asked for — a tick of the loop's own, or a window message
/// from a hand that moved across the caption — and all a paint needs to know is whether one is
/// up and how far round it has come. The wait itself belongs to the loop, which is what
/// started it and what takes the answer (see `PinLoad` and `paint_pin_spinner`).
pub(super) static PIN_ARC: AtomicU64 = AtomicU64::new(0);

/// What [`PIN_ARC`] counts from: this module's own first moment, so that a moment of a wait
/// fits in the `u64` an atomic can hold rather than in an `Instant` it cannot.
pub(super) static PIN_ARC_BASE: Lazy<Instant> = Lazy::new(Instant::now);

/// Show a pinned window another file: what is on screen is taken down — the frame this app holds,
/// the player it started, the browser another engine draws in — and the file that replaces it is
/// put up in the box the pin's media occupies.
///
/// What this is handed is the whole of a load's answer, already loaded (`pin_media_load`): what is
/// left here is everything that has to be asked of this thread, which is the browser a document is
/// handed to, the media engine a video is played through and FFmpeg's player a sound and a film
/// are played by. A load that has run for `spinner_delay_ms` puts an arc in the middle of the
/// pin's media while it runs, so a file whose read or decode is slow is a wait the pin is painted
/// for rather than a window frozen for as long as the disk takes (see `PinLoad` and
/// `paint_pin_spinner`).
///
/// What is installed is put up by the take-up that follows, so a swap reaches the window by the
/// same path a first pin does. Nothing is answered where the file cannot be shown: the pin keeps
/// the file it is showing, which is the only other thing a window that is already up can do for a
/// file that has no preview — and a step of the pin's own walk that lands on one is stepped over
/// rather than stopped at (see `PinStep`).
///
/// A video the media engine plays is the one kind this does not finish: the engine's first frame
/// is a tick of its own away, and installing the placeholder in the meantime is a flash of the
/// backdrop the tray keeps for pictures every time a caption's **Next** steps onto a film. So
/// that one is started here and held, and the file the pin is already showing keeps drawing
/// until the engine has actually drawn something (see `PinSwapHold`).
pub(super) fn swap_pinned_media(answer: PinAnswer) -> PinSwap {
    let PinAnswer {
        path,
        update,
        media,
        walk,
        arc,
    } = answer;

    let content = update.content;
    let volume = update.volume;
    let width = (content.2 - content.0).max(1);
    let height = (content.3 - content.1).max(1);

    // A file nothing could read — not there, still in the cloud, a decode that came back with
    // nothing at all — is a file the walk steps over, and there is no frame of it for any of
    // what follows to be asked about (see `refuse_pinned_media`).
    let Some(media) = media else {
        return PinSwap::Refused { path, walk };
    };

    let mut file = PinInstallable {
        path,
        update,
        media,
        audio: None,
        walk,
    };

    // The pin is being shown another file, so the copy of the film it was showing is dropped:
    // the pass reads the whole film, and the film the window is actually showing is the one
    // worth that read (see `subtitle_files`). A copy for this same file is kept, which is what
    // a pin taken up over a hover's own extraction needs.
    subtitle_files::keep_extraction_for(Some(&file.path));

    // What was showing goes before what replaces it: the player a video of this app's was started
    // in, and the frame this side holds.
    //
    // A video the media engine plays is the exception, and the exception is the whole of what
    // the hold below is for. This take-down is what takes the pin's frame off the screen, and the
    // engine's first frame is a tick or more away — so a swap that did it here would leave the
    // window standing on nothing for the length of that wait, compositing the placeholder over
    // the backdrop, which is the flash the hold exists to answer. Every other kind still goes
    // first, because for those the frame in hand is the frame the pin is shown the moment this
    // returns, and a player that would still be running behind it is a player playing over the
    // file that replaced it.
    //
    // The hold keeps the frame and nothing else: the player behind the standing file goes all the
    // same, and at once (see `stop_pinned_player`).
    // The cover goes up before either take-down and not after it, because the take-down is what makes
    // the hole: the player is killed on both branches below, the band's pixels are transparent ones
    // for a video, and the file that replaces this one has no window of its own for the length of a
    // decode — or, for the kind the engine plays, for the length of the wait for its first frame. So
    // the outgoing frame is read off the screen, painted over the band and the player's window put
    // away first, and the record is armed so the settle hands the band to the window that replaces
    // it rather than to whatever merely exists (see `cover_step_swap_for_video`).
    //
    // **A file no player of this app's is coming for is not covered at all**: a picture paints
    // itself, so a cover here would only have to be taken down again over the picture that replaces
    // the film (see `give_up_pinned_park`) — and neither is a film on a machine FFmpeg is not on,
    // which is answered the same way for the same reason (see `pinned_player_is_coming`).
    let player_coming = pinned_player_is_coming(file.media.media_type);
    cover_step_swap_for_video(player_coming);
    if file.media.media_type.is_native_video() {
        stop_pinned_player();
    } else {
        take_down_pinned_media();
    }

    if file.media.media_type.is_engine() {
        // A document or a specimen is drawn by the browser in a window of its own, and the engine
        // the old document was in is the same engine: what it is told is another file and the box
        // it goes in, which is a navigation rather than a browser started again (see
        // `webview_preview` and `restore_pin`, which asks it the same way).
        webview_preview::show(
            &file.path,
            webview_preview::Area {
                x: content.0,
                y: content.1,
                width,
                height,
            },
            engine_background(&file.path),
        );
    } else {
        // Every other kind is drawn here, so a browser still up — a document the pin was showing
        // before this file — comes down with the frame that replaces it.
        webview_preview::hide();
    }

    if file.media.media_type.is_native_video() {
        // A video the media engine plays is started before it is shown, for the reason a hover
        // starts one before its preview goes up: an engine that will not play the file is a file
        // with no preview rather than a box of the placeholder pixels a video is loaded with.
        let (video_width, video_height) = (file.media.current_width(), file.media.current_height());
        video_player::play(
            &file.path,
            video_width,
            video_height,
            volume,
            probed_picture(&file.path, video_width, video_height),
        );

        if !video_player::is_playing() {
            // A refusal rather than a wait, and it is the refusal this has always given: the
            // engine would not take the file at all, which is a file the walk steps over rather
            // than one a frame is waited for. The standing file is still taken down first, exactly
            // as it was before the hold existed — nothing was drawn in the meantime, so there is
            // nothing to keep (see `refuse_pinned_media`).
            take_down_pinned_media();
            return PinSwap::Refused {
                path: file.path,
                walk: file.walk,
            };
        }

        // And the file is not installed. What the load came back with is the buffer a video is
        // always loaded with — of the right size, and holding nothing, which every frame the
        // engine draws afterwards is written into, and out of which nothing is read until the
        // engine has reported a frame of its own (see `take_native_video_frame`) — and the
        // first of those arrives on a tick of its own. So the pin keeps the file it is
        // showing, at the frame it had stopped on, and the loop's own per-tick take is kept off
        // the standing file until the wait is over: a frame pulled into the media the pin is
        // still showing is a frame of the *new* film in the *old* file's buffer at the old
        // file's size, which is a worse thing to see than the flash this answers (see
        // `PinSwapHold`).
        //
        // The arc is the load's own, carried into the hold rather than dropped with the load: the
        // pin is still waiting for a file, and what the user is shown while it does is the file
        // it already had with the arc turning over it.
        return PinSwap::Holding(PinSwapHold { file, arc });
    }

    if player_coming {
        // FFmpeg's player draws it in a window of its own, which the take-up puts in the pin's
        // media band — and a player that would not start is the same answer an engine that will not
        // play gets: there is nothing of the file to show, so the pin keeps the file it has. The
        // player of the file the pin was showing may have survived its own stop (a dropped handle,
        // a kill nobody confirmed), and one of those playing over the file that replaces it is what
        // the sweep before every other player start is for (see `start_audio_playback`).
        kill_stray_video_process();

        // The track is the new file's own choice, which is the one the take-up that
        // installs this file writes down beside it — so the player and the bar
        // agree from the first frame rather than the bar being corrected by the
        // first seek (see `PinTransport::subtitle`).
        let Some(process) = start_video_playback(
            &file.path,
            content.0,
            content.1,
            width,
            height,
            0.0,
            volume,
            video_subtitles(&file.path).chosen(),
        ) else {
            // A cover standing over a player that was killed for a replacement that never came is
            // a frozen frame nothing will ever end, so it is given up here rather than left for
            // the walk's next file to inherit (see `give_up_pinned_park`).
            give_up_pinned_park();
            return PinSwap::Refused {
                path: file.path,
                walk: file.walk,
            };
        };

        file.media.video_process = Some(process);

        return PinSwap::Ready(file);
    }

    // A sound is started here, the way it is started for a hover, and where it is started *from* is
    // read the way the hover reads it: the tray's `Volume → Audio Seek` decides, and a share of a
    // length nothing has read yet is kept for the tick that can ask for it (see
    // `audio_seek::start_position`). A sound no player will take is a card with no clock behind it,
    // which is the same answer the load path gives one.
    //
    // At the level the plan was made with rather than at the tray's, which is the level a film two
    // arms above is started at and the level the card is about to be drawn at: a knob moved on this
    // window's own bar is a level the next sound is played at, and a player begun at the tray's
    // under a card that says otherwise is a sound the card and the user's ear disagree about (see
    // `PinVolume`).
    if file.media.media_type.is_audio() {
        let seek = current_audio_seek();
        let length = audio_track::playable(&file.path).and_then(|track| track.duration);
        let start = audio_seek::start_position(&file.path, seek, length);

        if !start_audio_playback_at(&file.path, &mut file.media, start, volume) {
            return PinSwap::Refused {
                path: file.path,
                walk: file.walk,
            };
        }

        // What the loop's clock is written down from, read before the media is handed on: a sound
        // is timed from the player this app started, and there is one to time from exactly where
        // one was started — a file nothing would play, at any level, gets none, and a card whose
        // clock ran anyway would be a sound it says is playing that is not.
        let started = file.media.video_process.is_some().then(Instant::now);

        file.audio = Some(SwappedAudio {
            started,
            from: start,
            share: (start == 0.0
                && matches!(seek, AudioSeek::Middle | AudioSeek::Random)
                && length.is_none())
            .then_some(seek),
        });

        return PinSwap::Ready(file);
    }

    PinSwap::Ready(file)
}

/// Whether a file of this kind is one a player of this app's is going to be put up for.
///
/// **A video, and only where FFmpeg is on this machine.** FFmpeg's player draws it in a window of
/// its own, so a file it would not play has no player at all — which is why every road that depends
/// on the answer asks this rather than the kind on its own: the cover a step raises over the
/// outgoing frame, the start a step makes, and the give-up a take-up owes a cover already standing.
///
/// It is asked once per road rather than carried across the ticks between them, so the two readings
/// that must agree are the two sides of one decision — the cover and the start — made together in
/// `swap_pinned_media`. The drift the reader of this is warned about is not two reads racing inside
/// one swap; it is a later tick asking again (`install_pinned_media`) and getting a different answer,
/// which is a cover standing over the film being left behind for the life of a pin whose player is
/// never begun: nothing is behind the band to hand it back to, so the cover never comes down, and
/// `pin_media_is_alive` reads its own record as a player still on its way (see `give_up_pinned_park`
/// and `reconcile_swap_take_up`). The later tick therefore answers the give-up, not the cover.
pub(super) fn pinned_player_is_coming(media_type: MediaType) -> bool {
    media_type == MediaType::Video && ffplay_is_here()
}

/// Whether FFmpeg's player is on this machine, which is the half of [`pinned_player_is_coming`]
/// that is about the machine rather than about the file.
///
/// It is asked through here rather than inline so that a test can stand an answer in for it, and it
/// has to be: the arm this exists for is one a machine with no FFmpeg takes every time and a machine
/// with one never takes at all, so nothing on a machine this is developed on can reach it (see
/// `stand_ffplay_in`).
pub(super) fn ffplay_is_here() -> bool {
    #[cfg(test)]
    if let Some(answer) = ffplay_stand_in() {
        return answer;
    }

    codecs::ffplay_available()
}

/// The answer [`ffplay_is_here`] is given while a test stands one in, read under the slot's lock.
#[cfg(test)]
fn ffplay_stand_in() -> Option<bool> {
    FFPLAY_STAND_IN.lock().ok().and_then(|held| *held)
}

/// Stand an answer for [`ffplay_is_here`] in place of the machine's own, or give the machine's own
/// answer back with `None`.
///
/// One slot for the whole process, so it is put back by the tests that set it rather than left to
/// whichever test the runner reaches next — the rule a park's own one slot keeps (see
/// `clear_park_swap_arm`).
#[cfg(test)]
pub(super) fn stand_ffplay_in(answer: Option<bool>) {
    if let Ok(mut held) = FFPLAY_STAND_IN.lock() {
        *held = answer;
    }
}

/// The answer [`ffplay_is_here`] is given while a test stands one in (see `stand_ffplay_in`).
#[cfg(test)]
static FFPLAY_STAND_IN: Mutex<Option<bool>> = Mutex::new(None);

/// Where the player a sound's card is drawn against was started, read by the caller into the clock
/// the loop draws that card from — the instant a player this app started began, the second of the
/// file it was begun at, and a start that was a share of a length nothing has read yet, which is
/// asked for on the first tick a player can answer (see `audio_clock`).
pub(super) struct SwappedAudio {
    pub(super) started: Option<Instant>,
    pub(super) from: f64,
    pub(super) share: Option<AudioSeek>,
}

/// The loop's own bookkeeping about what a pinned window is showing, borrowed for the one tick
/// that changes it.
///
/// It is a struct of borrows rather than a list of `&mut` parameters because there are fifteen of
/// them and a call that lists fifteen arguments is a call nobody can read. What an install writes
/// is scattered across the loop's locals because each of them is a fact some *other* part of the
/// tick reads — where a player's window belongs, the hover a take-down ends, the clock a sound's
/// card is drawn against, the request that puts the window up — and none of them belongs to the
/// swap. Bundling the borrows names the whole of what a swap reaches out of the loop, which is
/// the one thing a reader of `install_pinned_media` cannot see from its own arguments and wants
/// to know: a swap that looks like it returns a frame and nothing else is in fact rewriting a
/// good deal of the loop.
///
/// It is put together by `pin_install` rather than written out as a literal at either call site,
/// so that the list of what a swap touches is one list and not one per caller. There are two
/// callers, and they are two because of the only interesting thing about a swap: one of them
/// runs the tick a load answers on and the other runs the tick a held video's first frame lands
/// on, and the install has to be the *same* install both times or a film reached by pressing
/// **Next** is not the film a listing pick would have shown (see `PinSwapHold`).
pub(super) struct PinInstall<'a> {
    /// The walk a file that cannot be shown is stepped on from, and the walk the file being
    /// installed came from is put back into as the walk behind what is on screen (see
    /// `pin_step_off`).
    pub(super) walk: &'a mut Option<PinStep>,
    /// The walk that made the file now on screen, kept rather than spent with the load: an engine
    /// that takes a file and then never draws a frame of it says so a tick or three later, long
    /// after the walk that reached it is gone, and without this a corrupted film ends the walk
    /// rather than the walk stepping past it.
    pub(super) walk_of_current: &'a mut Option<PinStep>,
    /// The load the mark for a file that could not be shown is queued as, which a refusal queues
    /// one of where the walk has nothing left to offer (see `show_pin_failure`).
    pub(super) load: &'a mut Option<PinLoad>,
    /// When the player behind what is on screen now was started, which is what tells a player
    /// that is gone from a file that refused it rather than from a film that ended or a window
    /// the user closed — and is `None` for every kind this app's own player does not draw (see
    /// `pin_media_failed_before_a_frame`).
    pub(super) player_started: &'a mut Option<Instant>,
    /// Where the player a sound's card is drawn against was started.
    pub(super) audio_started: &'a mut Option<Instant>,
    /// The second of the file that player was begun at.
    pub(super) audio_start_offset: &'a mut f64,
    /// A seek asked for at a share of a length nothing has read yet, kept for the tick that can
    /// ask for it (see `audio_seek::start_position`).
    pub(super) audio_share_seek: &'a mut Option<AudioSeek>,
    /// Where a sound the user held was let go of, and for how long: a file swapped into the
    /// window is not held, whatever the file it replaced was doing.
    pub(super) audio_paused: &'a mut Option<f64>,
    /// When the card was last repainted, which a swap restarts so that the new card is drawn at
    /// once rather than at the end of the old one's cadence.
    pub(super) audio_repaint_at: &'a mut Instant,
    /// The display the card is laid out for.
    pub(super) audio_card_dpi: &'a mut u32,
    /// The marquee a name too long for its card scrolls across, which is the card's own box —
    /// the frame that has just been loaded, rather than the one it replaced.
    pub(super) audio_name_scroll: &'a mut Option<audio_preview::NameScroll>,
    /// Where a player's window belongs while a pin is up: the pin's media band, which is what the
    /// tick's own re-assertion of the order reads.
    pub(super) video_pos: &'a mut (i32, i32, i32, i32),
    /// The file a player this app started is playing, which is the one whose window is put on top
    /// again every tick.
    pub(super) current_video_path: &'a mut Option<PathBuf>,
    /// The hover the media behind the pin is the answer to, which a take-down ends and a walk is
    /// matched against.
    pub(super) current_show: &'a mut Option<PreviewMessage>,
    /// The pin's own request for the next tick, which is the take-up of what is installed here.
    pub(super) request: &'a mut Option<PreviewMessage>,
}

/// Borrow the loop's own bookkeeping about what the pin is showing, for one install or one
/// refusal.
///
/// The fifteen arguments are the fifteen the struct carries and the order is the struct's, so
/// that reading the two side by side is reading one list rather than two that have to be matched
/// up by hand — which is the only way a list this long can be kept right when a swap grows
/// another thing to write. Nothing is copied and nothing is read: every one of them is a mutable
/// borrow of a local that outlives the call, and the borrow ends with it (see `PinInstall`).
#[allow(clippy::too_many_arguments)]
pub(super) fn pin_install<'a>(
    walk: &'a mut Option<PinStep>,
    walk_of_current: &'a mut Option<PinStep>,
    load: &'a mut Option<PinLoad>,
    player_started: &'a mut Option<Instant>,
    audio_started: &'a mut Option<Instant>,
    audio_start_offset: &'a mut f64,
    audio_share_seek: &'a mut Option<AudioSeek>,
    audio_paused: &'a mut Option<f64>,
    audio_repaint_at: &'a mut Instant,
    audio_card_dpi: &'a mut u32,
    audio_name_scroll: &'a mut Option<audio_preview::NameScroll>,
    video_pos: &'a mut (i32, i32, i32, i32),
    current_video_path: &'a mut Option<PathBuf>,
    current_show: &'a mut Option<PreviewMessage>,
    request: &'a mut Option<PreviewMessage>,
) -> PinInstall<'a> {
    PinInstall {
        walk,
        walk_of_current,
        load,
        player_started,
        audio_started,
        audio_start_offset,
        audio_share_seek,
        audio_paused,
        audio_repaint_at,
        audio_card_dpi,
        audio_name_scroll,
        video_pos,
        current_video_path,
        current_show,
        request,
    }
}

/// Put a file a pinned window is being shown another one into the window.
///
/// It is the one install there is, and both roads a swap can take run through it: the tick a load
/// answers on for every kind but a video the media engine plays, and the tick that video's first
/// frame lands on (see `PinSwapHold`). Anything the swap did to what it is not itself is done
/// here — the arc comes down, the walk behind what is on screen is rewritten, a sound's card
/// gets its clock — because the whole of what "the pin is showing this file now" means to the
/// rest of the loop is written here and nowhere else, and two copies of it would be two files
/// that drift apart over a run (see `PinInstall`).
///
/// The media is installed last of the loop's own facts and before the request, so that the tick
/// which takes the request up reads a loop that already agrees with the window: the request
/// carries only the file and the box, and everything else about the file being on screen is
/// state the take-up reads rather than state it is handed.
pub(super) fn install_pinned_media(install: PinInstall<'_>, file: PinInstallable) {
    let PinInstallable {
        path,
        update,
        media,
        audio,
        walk,
    } = file;
    let PinUpdate { content, dpi, .. } = update;

    // Everything a swap is taking up with, before this file's own facts are written over the
    // loop's: a cover for this file's player is one of the things being carried, and the player
    // this is asked about is the one that is actually going to be begun — a film nothing will play
    // is given the cover up rather than handed it (see `reconcile_swap_take_up` and
    // `pinned_player_is_coming`).
    reconcile_swap_take_up(pinned_player_is_coming(media.media_type));

    // The pin is not waiting for a file any more, whatever the file it was waiting for turned
    // out to be: an arc left up over a window showing a video is a spinner for a question nobody
    // is asking, and one of them is on screen for as long as it is left there (see `pin_arc_set`).
    pin_arc_set(None);

    // The walk that made the file now on screen is kept rather than spent with the load: an
    // engine that takes this file and then never draws a frame of it says so a tick or three
    // from here, with nothing left to step on. A pick that was not a walk's arms no walk, which
    // is the same answer (see `pin_step_off`).
    *install.walk_of_current = walk;

    // A volume popup floating over the media belongs to the box it was opened over, and that box
    // has just been given another file: it is put away rather than left where the hand left it
    // (see the box change below, which does the same).
    close_pin_volume();

    // A sound's card is drawn against the clock of the player this app started, and the clock is
    // the loop's rather than the media's: where that player was started and from which second is
    // read here the way the load path reads it, so a card swapped into a pin ticks like a card a
    // hover put up (see `audio_clock`). The marquee its name needs is the card's own box, which
    // is the frame that has just been loaded.
    if let Some(start) = audio {
        let card_width = media.current_width();

        *install.audio_started = start.started;
        *install.audio_start_offset = start.from;
        *install.audio_share_seek = start.share;
        // A file swapped into the window is not held, whatever the file it replaced was doing:
        // the sound a card is drawn at belongs to the sound behind it.
        *install.audio_paused = None;
        *install.audio_repaint_at = Instant::now();
        *install.audio_card_dpi = dpi;
        *install.audio_name_scroll = Some(audio_preview::NameScroll::of(
            &audio_preview::name_of(&path),
            card_width,
            dpi,
            current_audio_options(),
        ));
    }

    // Where a player's window belongs while a pin is up is the pin's media band, which is what
    // the tick's own re-assertion reads.
    *install.video_pos = (
        content.0,
        content.1,
        content.2 - content.0,
        content.3 - content.1,
    );
    *install.current_video_path = match media.media_type {
        MediaType::Video => Some(path.clone()),
        _ => None,
    };
    // When the player behind what is on screen now was started, which is what tells a player
    // that is gone from a file that refused it rather than from a film that ended or a window
    // the user closed (see `pin_media_failed_before_a_frame`).
    *install.player_started = (media.media_type == MediaType::Video).then(Instant::now);

    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        *current = Some(media);
    }

    // What the loop's own bookkeeping is about is the file on screen, and the pin is showing
    // another one now: the card a sound is drawn from, the render tier's own record, and the
    // hover a take-down ends are all read of this.
    *install.current_show = Some(PreviewMessage::Show(
        path.clone(),
        content.0,
        content.1,
        None,
    ));

    *install.request = Some(PreviewMessage::Pin {
        path,
        rect: content,
    });
}

/// The answer a swap of a pinned window's file gives where there is nothing to show for it: the
/// walk is asked for the next file rather than left standing on one that cannot be shown, and
/// where the walk has nothing left to ask for, the mark for the file is what the window is left
/// over.
///
/// It is a function rather than an `if let` at the call site for the same reason the install is
/// one: a swap can be refused on the tick its load answers *or* on the tick a held video's wait
/// is given up on, and the second is a road that did not exist before the hold did. A refusal on
/// it is not a new answer — the file is installed the way it would have been, and the machinery
/// that watches a video for a player that never drew a frame runs on afterwards exactly as it
/// does for a swap that was never held (see `pin_media_failed_before_a_frame`).
///
/// The mark is not queued where something else is already on its way, which is the one
/// condition: a load in hand is a read and a decode that have not been paid for yet, and putting
/// the mark over it would throw that work away for a window that is about to be shown the file it
/// was for (see `show_pin_failure`).
pub(super) fn refuse_pinned_media(install: PinInstall<'_>, path: PathBuf, walk: Option<PinStep>) {
    pin_arc_set(None);

    pin_step_off(walk, install.walk);

    if install.walk.is_none() && install.load.is_none() {
        show_pin_failure(&path, install.load);
    }
}

/// A file a pinned window is loading, and the wait it is.
///
/// A load a hover waits for has a `PendingLoad`: the media is a thread's work, the window shows
/// the spinner's own box at the hand that asked, and the answer is matched against a generation
/// so that a load the pointer has left is dropped rather than shown. A pin has none of that — the
/// file was picked rather than hovered, so there is no hand to put a spinner at and no generation
/// to match against — but it has what the whole of that is for: the file on screen is a window
/// with a caption, and a window that freezes for as long as a large decode takes says nothing at
/// all. So the load is a thread's work here too, and what the wait shows is the arc in the middle
/// of the pin's own media (see `paint_pin_spinner`).
///
/// The file the load is for, the box it was laid out for and the walk it is a step of are all
/// kept, because the answer is taken up a tick or more after the load was started: what the answer
/// is installed with is the plan made when the file was picked, and a file that could not be shown
/// is a step the walk is asked to carry on from (see `PinStep`).
pub(super) struct PinLoad {
    pub(super) path: PathBuf,
    /// The box the file's media is loaded for, and the level the pin plays it at — the plan the
    /// pick was made with, kept rather than asked for again so that the answer and the plan can
    /// never disagree.
    pub(super) update: PinUpdate,
    /// The arc painted while this load runs.
    pub(super) arc: PinArc,
    /// The thread's answer, taken where it lands (see `take_pin_load`).
    pub(super) answer: Receiver<Option<MediaData>>,
    /// The walk this load is a step of, which is stepped on where the file cannot be shown.
    pub(super) walk: Option<PinStep>,
}

impl PinLoad {
    /// Start loading the file a pick was made for, and answer the wait it is.
    ///
    /// The thread is started here and the answer is left on a channel rather than handed over,
    /// because the loop is what owns the window and everything installed into it: a frame read
    /// on a thread of its own is installed by the thread that draws (see `take_pin_load`).
    pub(super) fn start(path: &Path, update: PinUpdate, walk: Option<PinStep>) -> Self {
        let width = (update.content.2 - update.content.0).max(1) as u32;
        let height = (update.content.3 - update.content.1).max(1) as u32;
        let answer = pin_media_load(path, width, height, update.dpi);

        PinLoad {
            path: path.to_path_buf(),
            arc: PinArc::new(),
            update,
            answer,
            walk,
        }
    }

    /// A load that has already answered, for the mark a file that could not be shown is drawn as.
    ///
    /// The channel is the one every load answers on, and it is answered before the load is
    /// handed over rather than by a thread: there is no read and no decode behind a cross, so a
    /// thread would be a wait for nothing, and `take_pin_load` reads the answer with `try_recv` —
    /// which is what makes this one taken up on the tick it is queued on (see `show_pin_failure`).
    pub(super) fn answered(path: &Path, update: PinUpdate) -> Self {
        let (answer, answers) = channel();
        let _ = answer.send(Some(unplayable_media((
            (update.content.2 - update.content.0).max(1) as u32,
            (update.content.3 - update.content.1).max(1) as u32,
        ))));

        PinLoad {
            path: path.to_path_buf(),
            arc: PinArc::new(),
            update,
            answer: answers,
            walk: None,
        }
    }
}

/// The arc a wait is painted on: when it began, how long before it is put up, and when it was
/// last turned.
///
/// Whether a wait is due a paint: once it has run for the delay `spinner_delay_ms` names — the
/// same moment a hover's own wait is given — and then once per turn of the arc for as long as it
/// stands. The turn is the spinner's own cadence, the one `MediaType::Loading` advances at,
/// because it is the same arc (see `MediaData::update_loading_frame`). It is measured from the
/// last turn rather than from the start of the wait, so the moment the delay runs out is one
/// paint and not one paint a tick until the cadence catches up with it.
pub(super) struct PinArc {
    pub(super) started: Instant,
    pub(super) spinner_delay: Duration,
    /// When the arc was last turned, and nothing while it has not been put up at all.
    pub(super) turned: Option<Instant>,
}

impl PinArc {
    pub(super) fn new() -> Self {
        PinArc {
            started: Instant::now(),
            spinner_delay: load_spinner_delay(),
            turned: None,
        }
    }

    pub(super) fn due(&self) -> bool {
        match self.turned {
            None => self.started.elapsed() >= self.spinner_delay,
            Some(turned) => turned.elapsed() >= Duration::from_millis(u64::from(SPINNER_TURN_MS)),
        }
    }

    /// Note that the arc has been turned, so that the next turn is a cadence away rather than
    /// due at once.
    pub(super) fn spun(&mut self) {
        self.turned = Some(Instant::now());
    }
}

/// Load the frame a pinned window is about to show, on a thread of its own.
///
/// This is the whole of what a swap can be waited for: everything the file is drawn *by* — the
/// browser a document hands to, the media engine a video is played through, FFmpeg's player — is
/// this thread's and cannot be moved, and all of it is asked for once the frame is in hand
/// (see `swap_pinned_media`). The read and the decode are not, and those are what a slow file
/// spends its time on.
///
/// The call is the one a hover's own load makes and the one a pin's own relayout makes, at the
/// same box, so what is drawn is what any other preview of the file would be. Nothing of the pin
/// is touched while it runs, which is what lets a wait be shown over the file the pin is still
/// showing: the old file is the pin's until the new one has answered.
pub(super) fn pin_media_load(
    path: &Path,
    width: u32,
    height: u32,
    dpi: u32,
) -> Receiver<Option<MediaData>> {
    let (answer, answers) = channel();
    let path = path.to_path_buf();

    std::thread::spawn(move || {
        // The same apartment a hover's loader takes before its first load: a PDF page is
        // rendered through Windows.Data.Pdf and a picture of a format this app has no decoder
        // for is decoded by the codec Windows has, and this is a thread of its own either way.
        pdf_preview::initialize_apartment();
        wic_image::initialize_apartment();

        // A panic in a decoder is a file with no preview rather than a thread this app's own
        // that takes the window down with it, which is what the hover's loader answers the same
        // way (see `spawn_load_worker`).
        let media = std::panic::catch_unwind(std::panic::AssertUnwindSafe(|| {
            load_media(
                &path,
                width,
                height,
                PreviewScale::FitToScreen,
                dpi,
                Arc::new(AtomicBool::new(false)),
            )
        }))
        .unwrap_or(None);

        let _ = answer.send(media);
    });

    answers
}

/// A pinned window's load, answered: the file, the plan it was loaded for, the frame, and the
/// walk the load is a step of.
pub(super) struct PinAnswer {
    pub(super) path: PathBuf,
    pub(super) update: PinUpdate,
    /// The frame the load came back with, or nothing at all for a file this app cannot read — a
    /// file that is not there, a file still in the cloud, a file whose decode is nothing. It is
    /// the walk that answers that, and not the pin: the file it is showing stays on screen
    /// (see `PinStep`).
    pub(super) media: Option<MediaData>,
    pub(super) walk: Option<PinStep>,
    /// The arc the load painted while it ran, carried out of the load with the answer rather than
    /// dropped with it. It is the one part of the wait that is not over when the answer is: a
    /// swap that is then held for a video's first frame is still a wait to the person looking at
    /// the pin, and taking the arc down the moment the load answered is what left a frozen file
    /// with nothing over it saying the window was still asking (see `PinSwapHold`).
    pub(super) arc: PinArc,
}

/// What a decode thread has answered, or `None` where it has not answered yet — which is the
/// ordinary case for the first ticks after a load is asked for, and what the pin keeps showing
/// its own frame over until the answer lands.
///
/// A thread that is gone with nothing to say — a panic on the way out, a sender dropped — is
/// answered as a file with no frame, which is what it is.
pub(super) fn answered(answer: Option<&Receiver<Option<MediaData>>>) -> Option<Option<MediaData>> {
    match answer?.try_recv() {
        Ok(media) => Some(media),
        Err(TryRecvError::Empty) => None,
        Err(TryRecvError::Disconnected) => Some(None),
    }
}

/// Take a pinned window's load, where it has answered.
///
/// `None` is a load still running: what the pin is showing is the file it already had, and the
/// arc is due for the wait or is up (see `PinLoad`).
pub(super) fn take_pin_load(pin_load: &mut Option<PinLoad>) -> Option<PinAnswer> {
    let media = answered(pin_load.as_ref().map(|load| &load.answer))?;
    let load = pin_load.take()?;

    Some(PinAnswer {
        path: load.path,
        update: load.update,
        media,
        walk: load.walk,
        arc: load.arc,
    })
}

/// The pin a window is taken up with: the whole of what a `PreviewMessage::Pin` builds, from the
/// media's own box to the window that stands around it.
///
/// It is a function rather than an arm of the loop's match for the reason `install_pinned_media` is
/// one: it is the *only* place a pinned window's state is written from whole, and both roads that
/// reach it — a pick out of the listing and a swap's own install — must build the same pin, or a
/// window that has just been shown another file is a window with different geometry from one that
/// has just been shown its first. A test that could not ask this question of it could only assert
/// the arithmetic, which is why the state dump of a taken-up window is taken here (see
/// `ws_m_frame_truth`).
///
/// `rect` is the media's own box: the hover's preview box for a first pin, and the box the swap
/// laid the incoming file out in for the other road. Everything else is read off the machine and
/// off the pin being replaced, and the caller publishes the result.
pub(super) fn take_up_pinned_window(path: &Path, rect: ScreenRegion) -> PinnedPreview {
    let kind = CURRENT_MEDIA
        .lock()
        .ok()
        .and_then(|media| media.as_ref().map(|media| media.media_type));
    let dpi = dpi_at(rect.0, rect.1);
    let transport_bar = pin_transport_kind(kind);
    let overlay = pin_overlay_chrome(kind);
    let hides_chrome = pin_hides_chrome(kind);
    let caption = pinned_caption_height(dpi, kind);

    // A sound's card is not the box its hover was: a hover's card carries no
    // controls — a hover's own window is a window nobody is in — and the row of
    // controls a pinned one carries is taller than the bar it stands in for, so
    // a pin that took the hover's box would be a window with the bottom of the
    // card cut off it (see `pinned_audio_card_box`).
    let rect = match kind {
        Some(MediaType::Audio) => pinned_audio_card_box(rect, path, dpi),
        _ => rect,
    };

    // The window is the media's box with the chrome around it, and it is the
    // *window* that is held to the display rather than the media: a hover can
    // sit flush against the top of the screen — the placement above it has
    // nowhere else to go — and a caption drawn above that would be a caption
    // off the top of the screen, with the buttons that close the pin on it.
    // So the box the media is given is the media's box shifted back into the
    // display by however much of the chrome fell off it.
    let window = clamp_pinned_box(
        pinned_window_box_of(rect, dpi, transport_bar, overlay, caption),
        dpi,
        &DESKTOPS,
    );
    let content = content_box_of(window, dpi, transport_bar, overlay, caption);

    // Whether the file this pin is coming up on is a sound, which is what says
    // which of the tray's two level settings the level it plays at is read of:
    // `Volume → Audio` for a sound, because that is the setting its preview was
    // playing at, and `Volume → Video` for every other kind. From the moment the
    // pin is up the level is this window's own, so it is asked of the pin where
    // it is holding one and of the tray where it is not (see `pin_volume_taken_up`
    // and `PinVolume`).
    let audio_kind = matches!(kind, Some(MediaType::Audio));

    // A pin taken up over another one — the file it was showing was picked by the
    // pointer or the keyboard while it was up (see `PinUpdate`) — is the same
    // window showing another file, so what belongs to the window rather than to
    // the file is carried over: a maximized pin stays maximized and restores to
    // the box it would have restored to, a level moved on its own bar stays where it
    // was moved to, and chrome that is showing over a picture is not brought back as
    // if the window had just arrived. A sound's card is the one kind that shows no
    // maximize of its own — a card is its own size, offers no maximize to stay in
    // (`PinFrame::None`) — but the state the window was in when the card arrived is
    // the state the file after the card is laid out under, which is the box that
    // file is fitted to (see `pin_update_content`); a walk through a sound that
    // un-maximized the window would be a window the user maximized once and every
    // file after the sound came at its own bound. There is nothing to carry for a
    // first pin, which is why the take-up below reads exactly as it always did.

    // The length the probe read is asked before the pin's own lock is taken:
    // the answer comes from the geometry cache, which is a lock and the
    // file's own metadata besides, and a take-up is no place to hold the pin
    // across either (see `cached_video_geometry`).
    let duration = video_duration(path);

    let carried = pin_state().and_then(|pinned| {
        pinned.pin().map(|pin| {
            (
                pin.restore,
                pin.chrome,
                pin.volume,
                pin.overlay,
                pin.hides_chrome,
                pin.bound,
            )
        })
    });

    // The two facts that have to be read before the pin's own lock is
    // taken, because each of them is a read of the machine rather than of
    // the state in hand: `pin_keeps_its_box` is
    // answered by a content probe that takes `CONFIG` over a file read,
    // and the Shell's own answer to "which program opens this" is two
    // `AssocQueryStringW` calls. None of them belongs inside a guard
    // that the window procedure takes for every mouse move, press,
    // release and paint — a take-up holding `PINNED` across any of them
    // is a window whose buttons stop answering for as long as the disk
    // or the Shell takes. One that pumps or re-enters while the lock is
    // held is worse still: the window procedure would be asking for a
    // lock this thread already owns, which is not a wait but a stop (see
    // `pin_media_is_alive` for the same rule, written down for the media).
    //
    // A take-up runs for every file a walk lands on rather than once per
    // pin, so it is the arrow keys, the two step buttons and a pick in the
    // listing alike that pay it.
    let keeps_its_box = pin_keeps_its_box(path);

    // The name the hand-off button says is asked of the planner rather
    // than here, for the same reason and by the same rule: it is a
    // question about the machine's own associations, and the pin is
    // answered about the file it holds until it holds another one (see
    // `PinTooltip`). The button is drawn without a name until the
    // answer lands rather than the take-up waiting for it.
    ask_pin_open_with(path.to_path_buf());

    let now = Instant::now();
    PinnedPreview {
        path: path.to_path_buf(),
        content,
        // A pin taken up over another one keeps the bound the window has:
        // what a swap is measured against is the size the window was given
        // rather than the size the file it is showing came out at. A pin
        // without one keeps none while the file on screen is drawn to its
        // own box, and takes the longest side of the box the first file
        // with a shape of its own came out at — which is where a pin taken
        // up on a picture gets its bound too (see `pin_bound_after`).
        bound: pin_bound_after(
            carried.and_then(|(.., bound)| bound),
            keeps_its_box,
            content,
        ),
        restore: carried.and_then(|(restore, ..)| restore),
        dpi,
        transport_bar,
        // Both kinds of video carry a bar that does something, and for
        // opposite reasons: the engine answers every question the bar asks,
        // and FFmpeg's player answers the two that are keys.
        transport_live: pin_transport_live(kind),
        frame: pin_frame(kind),
        overlay,
        hides_chrome,
        caption,
        chrome: match carried {
            // Chrome belongs to the kind it was drawn over: one kind's
            // strip has nothing to say about another's, so a swap that
            // changes it arrives as the new kind's own does. What is
            // compared is whether the kind draws its chrome over its
            // media and whether it can hide it at all — a picture's
            // arrived-and-gone title bar is a strip of chrome, and a kind
            // that keeps its chrome in bands has never shown one (see
            // `pin_hides_chrome`).
            Some((_, chrome, _, was_overlay, was_hiding, _))
                if was_overlay == overlay && was_hiding == hides_chrome =>
            {
                chrome
            }
            _ => {
                if hides_chrome {
                    PinChrome::on_arrival(now)
                } else {
                    PinChrome::always()
                }
            }
        },
        collapsed: false,
        bubble_pause: None,
        hovered: None,
        pressed: None,
        // The name the hand-off button says. It is left empty here
        // and filled in when the planner's answer lands, because
        // the Shell is asked on a thread of its own and this
        // take-up must not wait for it (see `ask_pin_open_with`).
        // The name is still asked once per format per run, as it
        // was below the lock before (see `default_app_name`).
        tooltip: PinTooltip::default(),
        dragging: None,
        // A cover standing for the file this pin is being taken up with
        // comes with it: the swap armed it over the outgoing frame and
        // the settle is what takes it down, on the tick that finds the
        // incoming player's window. Written false here it would be given
        // up by the first tick that finds it, which is the hole the cover
        // was raised to close (see `pin_park_carried_forward`).
        parked: pin_park_carried_forward(),
        transport: PinTransport {
            // The length the probe read, and where a player this app
            // started has got to: a video FFmpeg plays has been running
            // since before the pin existed, and its own clock starts
            // here — which is the best that can be said about a player
            // that reports nothing at all (see `PinTransport`).
            duration,
            started: (kind == Some(MediaType::Video)).then_some((Instant::now(), 0.0)),
            // Which subtitle track this pin opens on, which is the
            // player's own choice for the file until a key says
            // otherwise — written down rather than left unnamed, so that
            // the first seek does not quietly replace it (see
            // `video_subtitles`).
            subtitle: video_subtitles(path).chosen(),
            ..Default::default()
        },
        // A level carried onto a file of the same kind is that file's
        // own and is carried; one carried onto a file of another kind
        // is that other kind's bar having been turned, and is not
        // carried — the file arrives at the level the tray names for
        // it, and the level walked off is kept so that walking back is
        // answered by the pin rather than by the setting (see
        // `pin_volume_taken_up`).
        volume: match carried {
            Some((_, _, volume, ..)) if volume.audio == audio_kind => volume,
            carried => pin_volume_taken_up(carried.map(|(_, _, volume, ..)| volume), audio_kind),
        },
        audio_hovered: None,
        audio_pressed: None,
    }
}

/// Lay the pinned media out again for the box its window has been given — one it was maximized
/// to, restored from, or resized to.
///
/// What it costs is what a hover of the same file costs: the media is laid out at the size it
/// is now drawn at, and a frame the image cache already holds is handed back rather than
/// decoded, so a box that changed once is paid for once. For the kinds something else draws it
/// is cheaper still: a player's window is resized and its picture scales with it, and a browser
/// is told the new bounds of the page it is already showing.
///
/// It runs on the preview thread, and that is deliberate for everything a box change asks of
/// *another* window — a player's window to move, a browser's to travel with the band, a card to be
/// laid out again — because all of that is a message to a window this thread owns and nothing else
/// can send it.
///
/// What it asks of this thread's own decoding is not, and is not asked here any more. A decode is
/// the one thing in this function with no bound on it, and it was being run inline on the thread
/// that pumps the pinned window's messages: a large picture or a long document resized on the
/// loop is a tick that takes seconds, and a tick is a stretch of the loop in which no window
/// message is dispatched at all — so the drag under the hand stops answering, the caption's
/// buttons stop answering, and the window is the frozen one this function's own comment used to
/// call a deliberate trade. It is the same wait a swap already waits on, and it is waited on the
/// same way: the read and the decode go to a thread of their own and the loop installs what comes
/// back (see `PinRelayout`). Nothing is lost by the frame not being there the instant the box
/// changes, because the band is already filled from the frame that *is* there and scaled into the
/// new box — the same answer a resized video is given for as long as the engine's next frame is
/// on its way (see `compose_media_into_band`).
///
/// **`road` is what tells the one kind whose answer differs which of the two roads is asking**
/// (see `PinRelayoutRoad`): every other kind lays its media out again wherever the box came from,
/// but a video's box change is a player being ended and begun again, and that is the one thing a
/// take-up is not doing.
pub(super) fn relayout_pinned_media(
    path: &Path,
    content: ScreenRegion,
    dpi: u32,
    card: Option<AudioCardClock>,
    road: PinRelayoutRoad,
) -> Option<PinRelayout> {
    let width = (content.2 - content.0).max(1) as u32;
    let height = (content.3 - content.1).max(1) as u32;
    let kind = CURRENT_MEDIA
        .lock()
        .ok()
        .and_then(|media| media.as_ref().map(|media| media.media_type));

    match kind {
        // FFmpeg's player draws in a window of its own and scales the picture to it, so the box
        // that changed is a window that has to be moved — and, because the player rendered at
        // whatever size it was begun at, one that has to be *begun again* to be rendered sharply
        // at the size it has settled on.
        //
        // The two happen at different times and that is the whole of the two-phase resize. While
        // the hand is still moving the edge, what is on screen is asked for on every pointer move
        // (see `place_pinned_siblings`, called from `apply_pin_drag`): the window is put to the
        // new size with `SetWindowPos` and the frame already rendered is scaled into it. It is
        // responsive, it follows the pointer at the pointer's own pace, and it is soft — a 1440p
        // frame shown in a box half its size and scaled back up is not the picture the user is
        // choosing a size for.
        //
        // This is the settle, and it is the only place a relaunch happens for a resize. The
        // player is ended and begun again at the box the drag ended up with, cropped from the
        // same geometry probe the first launch cropped from, so the file that comes up is the
        // file that was already there at full resolution — and it arrives over the old window
        // rather than after it (see `retire_replaced_player`), so the soft frame under the hand
        // is replaced by a sharp one without the band ever showing the desktop.
        //
        // The alternative — holding the old size until the relaunch lands — was rejected, and
        // the reason is what a resize *is* while it is happening. A window that refused to follow
        // the hand and sprang to a new size at the end is not a window being resized; it is a
        // window that was closed and opened again, and the thing the user is doing — judging a
        // size by how the film looks in it — cannot be done against a picture that does not
        // change. Softness for the duration of a drag is a cost paid while the hand is moving
        // and nothing else is happening; a band that stays at the old size until a relaunch
        // completes is a visible discontinuity in the middle of a gesture, and it is the
        // discontinuity the user would notice rather than the softness.
        //
        // The position is kept, so a resize does not throw away where the film was: the bar is
        // read for the second the player had got to and the new one is begun there. A resize that
        // restarted the film at the beginning would be the one thing about a resize that has to be
        // undone by hand afterwards — and a hold is kept for the same reason, because a pin that
        // was holding the film and was resized is a pin holding the film at a different size, not a
        // pin that has started playing it (see `restart_pinned_player`).
        Some(MediaType::Video) => {
            // **A box change is a player being ended and another begun, so it takes the press road
            // first**: the cover stands over the outgoing frame and the old player is killed before
            // the replacement exists, which is what keeps the replacement's window — published within
            // milliseconds and empty until the file is open and the first frame is decoded — from
            // being raised over a transparent band (see `begin_covered_box_change`).
            //
            // **And nothing else does.** A take-up has pressed nothing: it builds the pin whole
            // for a file this window has not stood on before, with no hand in it and no player
            // behind it to kill. Asking the press road anyway is not harmless even so — it is silent
            // about the absence of a press, not deaf to it. A press snapshots the film it found and
            // freezes the clock (`gesture_press_freeze`), and the relaunch below reads its hold out
            // of that snapshot, so a road with no press in it has no snapshot to read: what it gets
            // is either no relaunch at all (the end finds nothing to take) or the second and the
            // hold of a press belonging to some other road — with a cover left standing over the
            // frame and nothing in the pin's own state to end it (see `relaunch_gesture_end_at` and
            // `pin_park_carried_forward`).
            //
            // A resize's own press has already run this road and left its cover standing, so on that
            // box change the call answers false too, and the relaunch below is the one the resize's
            // release armed: one press, one player.
            let covered =
                matches!(road, PinRelayoutRoad::BoxChange) && unsafe { begin_covered_box_change() };
            ensure_pinned_sibling_box(content);
            match pinned_media_owner() {
                // The press's end: it takes the snapshot, so this is the one relaunch the box
                // change owes — at the second the film had got to, and not twice across the
                // release that re-enters behind the settle (see `relaunch_gesture_end_at`).
                Some(_) if covered => {
                    relaunch_gesture_end_at(content, None);
                }
                Some((path, _)) => {
                    // And with no press behind it the hold is the pin's own: what the transport says
                    // the film is doing now, which for a take-up is the transport that take-up built
                    // rather than one a press has frozen (see `pinned_is_held`).
                    restart_pinned_player(
                        &path,
                        content,
                        pinned_playhead().unwrap_or(0.0),
                        pinned_is_held(),
                    );
                }
                None => {}
            }
            None
        }
        // The media engine draws into a surface of the size it was started at, and this window
        // draws the frames it hands back: both are resized, and the picture follows.
        Some(MediaType::NativeVideo) => {
            // What is deliberately *not* done here is redeclaring the frame to be the size of the
            // box: the frame holds the pixels of the size it was drawn at, and a buffer of one
            // width read as a buffer of another is a picture sheared a row at a time — the
            // diagonal striping a resize used to flash for the moment before the engine handed
            // the next frame over. The band is filled from the frame that is there, scaled, which
            // is exactly what it is filled with while the edge is still under the hand, and the
            // frame that lands on the next tick is the one at the size the window now is (see
            // `compose_media_into_band`).
            video_player::resize(width, height);
            None
        }
        // The browser draws in a window of its own, so the band that changed is a window that has
        // to travel with the box: it is moved rather than only asked again — a document still on
        // its way is moved too, by the want it will land on (see `webview_preview::place`).
        Some(MediaType::EngineSvg) | Some(MediaType::EngineFont) => {
            webview_preview::place(
                path,
                webview_preview::Area {
                    x: content.0,
                    y: content.1,
                    width: width as i32,
                    height: height as i32,
                },
                engine_background(path),
            );
            None
        }
        // A sound's card, which is a page of text laid out again rather than a file decoded —
        // and the one kind whose layout needs something the loop holds rather than the media:
        // where the player it started is, how far a name it had no room for has been scrolled,
        // and where a key put it on hold (see `AudioCardClock`).
        Some(MediaType::Audio) => {
            let clock = card?;
            let (elapsed, duration) = audio_clock(path, clock.started, clock.from, clock.paused);

            if let Ok(mut media) = CURRENT_MEDIA.lock() {
                if let Some(media) = media.as_mut() {
                    media.relayout_audio_card(path, clock, elapsed, duration, (width, height));
                }
            }
            None
        }
        // A frame this app draws for itself: the media is read and decoded again at the box it
        // is now shown in, which is the same call a hover's own load makes — and the same
        // wait one is, on a thread of its own. The decode used to run here, inline, which
        // put an unbounded stretch of unpumped loop in front of a window the hand was
        // still dragging; it is now answered a tick or more later by `take_pin_relayout`,
        // and the band is filled from the frame that is there until it lands.
        Some(_) => Some(PinRelayout {
            path: path.to_path_buf(),
            answer: pin_media_load(path, width, height, dpi),
        }),
        None => None,
    }
}

/// Which of the two roads asked for a relayout, for the one kind of media whose answer differs.
///
/// Every kind lays its media out again in whatever box it is given, so the road is a fact only a
/// video needs — and a video needs it because a box change for one is a player being ended and
/// another begun, which is the WS-J press road: freeze the clock, cover the outgoing frame, kill
/// the old player before the replacement's window is on screen to be seen empty over a transparent
/// band (see `begin_covered_box_change`).
#[derive(Clone, Copy, PartialEq, Eq, Debug)]
pub(super) enum PinRelayoutRoad {
    /// A take-up: the pin is being built whole for a file this window has not stood on before. No
    /// hand is in it, no player was killed, and nothing is owed a resume — so this road relaunches
    /// the film that is there, at the second it is at, in whatever state the transport the take-up
    /// just built says it is in (see `pinned_is_held`).
    TakeUp,
    /// A box change: the window the player is standing in has been given another box, so the
    /// player is ended and begun again in it. The press road, and the relaunch the press's own end
    /// makes (see `begin_covered_box_change` and `relaunch_gesture_end_at`).
    BoxChange,
}

/// A relayout of a pinned window's own media, waited for on a thread of its own.
///
/// The wait is held by the loop rather than inside the tick that started it, for the reason a
/// `PinLoad` is: the read and the decode are a thread's work, so the question is still out when
/// the tick that asked it ends and the answer to it arrives on one of the next.
///
/// There is no arc for this wait, and no spinner, and that is deliberate rather than an
/// oversight. A relayout answers a box the window has *already* been given: the file on screen is
/// the right file at a size it is being taken to, so there is nothing to wait *for* as far as the
/// person looking at it is concerned — the picture is there, scaled into the new box, and the frame
/// that is drawn for that box arrives when it arrives. An arc over it would say a thing that is
/// not true: that the window has nothing to show yet. It is the same answer a resized video is
/// given while the media engine's next frame is on its way.
pub(super) struct PinRelayout {
    /// The file being decoded, which is what the answer is installed over: a relayout left in
    /// flight across a swap would otherwise install a frame of the file the pin has left behind,
    /// at a box belonging to the one it is now showing.
    pub(super) path: PathBuf,
    pub(super) answer: Receiver<Option<MediaData>>,
}

/// Take a relayout that has answered, and install it.
///
/// One still running is the ordinary case for the first few ticks after a window is dragged to a
/// new size: what is on screen is the frame already decoded, and this is where it is replaced by
/// one drawn for the box the window is now at.
///
/// The frame is installed here, on the loop's own thread, because `CURRENT_MEDIA` is the loop's
/// (see `swap_pinned_media` for the same rule on a swap's answer) — and `keep_text_place` is asked
/// of the media being replaced *before* that lock is taken, for the reason it takes the lock
/// itself: a lock this thread already holds is not a lock it can wait for.
pub(super) fn take_pin_relayout(relayout: &mut Option<PinRelayout>) {
    let Some(media) = answered(relayout.as_ref().map(|pending| &pending.answer)) else {
        return;
    };
    let Some(pending) = relayout.take() else {
        return;
    };

    // A relayout is only an answer while the pin is still standing on the file it was asked for.
    // A swap that landed first, or a pin that has been closed, leaves it asking a question about
    // a window that is no longer there, and its frame must not be installed over whatever is
    // there now.
    if pinned_path().as_deref() != Some(pending.path.as_path()) {
        return;
    }

    let Some(media) = media else { return };
    let media = keep_text_place(media);
    if let Ok(mut current) = CURRENT_MEDIA.lock() {
        if let Some(ref mut existing) = *current {
            existing.cancel_background_work();
        }
        *current = Some(media);
    }
}

/// Carry where a page of text was up to — the line its frame starts at and what was selected —
/// from the media being replaced onto the one that replaces it.
///
/// A pinned text preview is laid out again whenever its window is given another box, and a box is
/// what the document is wrapped for: without this, resizing the window of a document someone is
/// reading would put them back at the top of it. The two positions survive being carried because
/// they are places in the document rather than points on the screen (see `Selection`), and the
/// rest of the state — the lines, the scrollbar, what can be reached — belongs to the box the new
/// media was laid out in.
///
/// It takes the lock itself and is called *before* the one above is taken, for the reason the
/// deadlock in `pin_media_is_alive` is written down as: a lock this thread already holds is not a
/// lock it can wait for.
pub(super) fn keep_text_place(media: MediaData) -> MediaData {
    let Ok(current) = CURRENT_MEDIA.lock() else {
        return media;
    };
    let Some(old) = current
        .as_ref()
        .and_then(|existing| existing.text_state.as_ref())
    else {
        return media;
    };

    let mut media = media;
    if let Some(state) = media.text_state.as_mut() {
        state.first_line = old.first_line;
        state.selection = old.selection;
    }
    media
}
