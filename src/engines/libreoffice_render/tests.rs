use super::*;

/// The engine is never started for a name its own filters do not read. A Photoshop
/// document, a Krita project, a Sketch file and a Procreate document are this app's
/// own readers' business, and asking an office suite about one would cost a launch and
/// answer nothing. A Flash animation is the case that made the rule worth stating: the
/// engine spins on one rather than failing, so a name like that must not reach it at
/// all (see `libre_formats`).
#[test]
fn asks_the_engine_only_about_the_names_it_reads() {
    for name in [
        "poster.psd",
        "painting.kra",
        "design.sketch",
        "art.procreate",
        "icon.svg",
        "animation.swf",
    ] {
        assert!(
            !crate::formats::libre_formats::engine_page_kind(Path::new(name)).is_some(),
            "`{name}` is not one of its formats"
        );
    }
}

/// And a name it does not read is not queued either: nothing about a document is asked
/// of the engine that its list has not claimed.
#[test]
fn queues_nothing_for_a_name_it_does_not_read() {
    assert!(
        !WORKER.take_queued(),
        "the slot is empty before anything is asked of the engine"
    );

    request(Path::new("animation.swf"));

    assert!(
        !WORKER.take_queued(),
        "the engine is not asked about a name no list of its own holds"
    );
}

/// And the give-up is not a decision on its own: what it names is ended, which is what
/// frees the seat and the core the engine was holding and lets the document behind it be
/// drawn. A process that stays up stands in for the engine — a test is not going to make
/// LibreOffice spin on a file — recorded the way the engine is, by image name, which is
/// the check that keeps an id from being acted on by itself.
///
/// The bound is the one this engine gives itself rather than the one the give-up is
/// decided by (`supervisor`): what is being asked here is that the two ends of the
/// decision reach the process, and the number that decides is that module's business and
/// tested there.
#[test]
fn ends_the_engine_only_once_its_conversion_has_outrun_the_give_up() {
    let _stand_in = crate::app::engine_processes::STAND_IN
        .lock()
        .unwrap_or_else(|poisoned| poisoned.into_inner());
    let _in_flight = crate::engines::supervisor::IN_FLIGHT_TAKEN
        .lock()
        .unwrap_or_else(|poisoned| poisoned.into_inner());
    let mut engine = crate::app::engine_processes::hidden_command("ping")
        .args(["-n", "30", "127.0.0.1"])
        .stdout(std::process::Stdio::null())
        .spawn()
        .expect("a process to stand in for the engine");
    let pid = engine.id();

    assert!(
        crate::app::engine_processes::processes_named("ping.exe").contains(&pid),
        "the stand-in runs the image the record below names it by"
    );
    crate::app::engine_processes::record("ping.exe", pid);

    // A conversion that has just started is a document being drawn, and is left to it.
    supervisor::begin(Adapter::LibreOffice, Path::new("drawing.cdr"), pid);
    end_hung_engine();
    assert!(
        crate::app::engine_processes::is_running(pid),
        "a conversion that has just started is not an engine to end"
    );

    // One that has outrun the give-up is an engine that has stopped answering. A bound of
    // nothing is the shortest a run can be outrun by, which says the same thing about
    // the decision without waiting half a minute to say it.
    supervisor::end_hung(Adapter::LibreOffice, Duration::ZERO);
    end_hung_engine();
    assert!(
        !crate::app::engine_processes::is_running(pid),
        "the engine a conversion has outrun is ended"
    );

    supervisor::stop(Adapter::LibreOffice);
    let _ = engine.wait();
}

/// And the names it does read are the CorelDRAW family and the formats of the same
/// libraries, whatever case they are written in.
#[test]
fn reads_coreldraw_and_the_formats_beside_it() {
    for name in [
        "logo.cdr",
        "drawing.CDR",
        "artwork.cmx",
        "poster.pub",
        "plan.vsd",
    ] {
        assert!(
            crate::formats::libre_formats::engine_page_kind(Path::new(name)).is_some(),
            "`{name}` is one of its formats"
        );
    }
}

/// What a kept engine is waited for is the document it holds being *open* — the lock file
/// LibreOffice writes beside it — and a wait answers as soon as that is there.
#[test]
fn waits_for_a_kept_engine_until_its_document_is_open() {
    let folder = std::env::temp_dir()
        .join("rust-hover-preview-libre-tests")
        .join("ready");
    std::fs::create_dir_all(&folder).expect("a test folder");

    let holder = folder.join(HOLDER_NAME);
    write_holder(&holder).expect("a written stub");
    assert_eq!(
        std::fs::read_to_string(&holder).expect("the stub"),
        HOLDER_DOCUMENT,
        "and what it holds is the document the engine is given"
    );

    // The lock file is what says the engine has its document open: with one there the
    // wait is answered at once rather than at its bound. The wait is on this process,
    // which is running, so what is measured is the lock file and nothing else.
    std::fs::write(lock_file(&holder), b"").expect("a written lock file");
    let started = Instant::now();
    assert!(wait_until_ready_for(
        std::process::id(),
        &holder,
        Duration::from_secs(30)
    ));
    assert!(
        started.elapsed() < Duration::from_secs(1),
        "the wait ends when the engine is ready rather than at the bound"
    );

    // Without it, the wait is the bound and no more — an engine that has not opened its
    // document yet is waited for, and one that never does is not waited for forever.
    std::fs::remove_file(lock_file(&holder)).expect("the lock file removed");
    let started = Instant::now();
    assert!(!wait_until_ready_for(
        std::process::id(),
        &holder,
        Duration::from_millis(200)
    ));
    assert!(started.elapsed() >= Duration::from_millis(200));

    // And a process that is gone is not waited for at all: what that costs is a launch
    // for the document at hand, which is what a machine without an engine pays anyway.
    let mut gone = crate::app::engine_processes::hidden_command("ping")
        .args(["-n", "30", "127.0.0.1"])
        .stdout(std::process::Stdio::null())
        .spawn()
        .expect("a process to stand in for an engine that is gone");
    let pid = gone.id();
    let _ = gone.kill();
    let _ = gone.wait();

    let started = Instant::now();
    assert!(!wait_until_ready_for(pid, &holder, Duration::from_secs(30)));
    assert!(
        started.elapsed() < Duration::from_secs(2),
        "an engine whose process is gone is not waited for, and the bound is not paid"
    );

    let _ = std::fs::remove_dir_all(&folder);
}

/// The idle setting is what says whether an engine is kept at all, and `0 seconds` — the
/// bottom of the tray's list — is the setting that keeps none: every document is the
/// launch it has always been. It is read through the app's own configuration, so what is
/// asserted here is the shape of the answer rather than the value a machine holds.
#[test]
fn a_setting_of_no_seconds_keeps_no_engine() {
    let keep = kept_idle().is_some();
    let configured = CONFIG
        .lock()
        .map(|config| config.libreoffice_idle)
        .expect("the configuration");

    assert_eq!(
        keep,
        configured != EngineIdle::Seconds(0),
        "an engine is kept exactly where the idle setting is not `0 seconds`"
    );
}

/// What the idle setting buys, measured: the same document converted with nothing kept,
/// and then again with the engine the first one started. Ignored because it starts the
/// installed LibreOffice, and run when that setting is being looked at:
/// `cargo test -- --ignored --nocapture engine_warmth_probe`.
#[test]
#[ignore = "starts the installed LibreOffice"]
fn engine_warmth_probe() {
    let Some(program) = soffice() else {
        println!("no LibreOffice installed: nothing to measure");
        return;
    };
    let Some(holder) = holder_document() else {
        println!("no folder for the stub document");
        return;
    };
    let Some(folder) = workspace() else {
        println!("no folder for the engine to work in");
        return;
    };
    std::fs::create_dir_all(&folder).ok();

    let source = folder.join("probe-source.fodt");
    std::fs::write(&source, HOLDER_DOCUMENT).ok();

    // Nothing is kept to begin with, so the first row is the launch every document paid
    // for on its own before there was a setting.
    document_cache::forget(&source, OfficeEngine::LibreOffice.as_str());
    let_go();
    let started = Instant::now();
    let first = convert(&program, &source);
    println!(
        "first document: {} in {:?} — the engine started and handed the page",
        first.is_some(),
        started.elapsed()
    );

    // And the same document again, with the engine the first one started.
    document_cache::forget(&source, OfficeEngine::LibreOffice.as_str());
    let started = Instant::now();
    let next = convert(&program, &source);
    println!(
        "next document: {} in {:?} — the engine kept, the page handed to it",
        next.is_some(),
        started.elapsed()
    );

    let launcher = kept_pid();
    let child = launcher.and_then(engine_child);
    println!("kept launcher {launcher:?}, engine behind it {child:?}");

    let_go();

    // A process that has been terminated is still in the process table for a moment, so
    // what is reported is the check that matters: whether it is still running.
    let gone = |pid: u32| !crate::app::engine_processes::is_running(pid);
    println!(
        "left after letting go: launcher gone {}, engine gone {}",
        launcher.map(gone).unwrap_or(true),
        child.map(gone).unwrap_or(true),
    );

    // And the bottom row of the setting: `0 seconds` keeps no engine, so one that is
    // running when it is chosen is let go of at the next look — which is this call, made
    // from the engine thread once a second while it waits for documents.
    let Some(pid) = keep_engine(&program) else {
        println!("no engine started for the last check");
        return;
    };
    wait_until_ready(pid);
    if let Ok(mut config) = CONFIG.lock() {
        config.libreoffice_idle = EngineIdle::Seconds(0);
    }
    let_go_if_expired();
    println!("0 seconds: engine gone {}", gone(pid));
    if let Ok(mut config) = CONFIG.lock() {
        config.libreoffice_idle = EngineIdle::Seconds(DEFAULT_LIBREOFFICE_IDLE_SECS);
    }

    document_cache::forget(&source, OfficeEngine::LibreOffice.as_str());
    std::fs::remove_file(&source).ok();
    std::fs::remove_file(lock_file(&holder)).ok();
}
