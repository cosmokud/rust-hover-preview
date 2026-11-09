use super::*;
use crate::readers::audio_track::{ask, probed, remember, Probed, ONE_AT_A_TIME};

/// A sound the machine has read once, in a folder of its own so that two tests are not each
/// other's file: `len` bytes of anything at all, because what is under test is the stamp read off
/// the file's own entry rather than what is inside it.
fn a_sound(case: &str, name: &str, len: usize) -> PathBuf {
    let folder = std::env::temp_dir()
        .join("rust-hover-preview-audio-track-disk")
        .join(case);
    let _ = std::fs::remove_dir_all(&folder);
    std::fs::create_dir_all(&folder).expect("a test folder");

    let path = folder.join(name);
    std::fs::write(&path, vec![0u8; len]).expect("a file a test can write");

    path
}

fn a_track(duration: f64) -> Track {
    Track {
        player: Player::Ffmpeg,
        codec: Some("FLAC".to_string()),
        rate: Some(44_100),
        channels: Some(2),
        bitrate: Some(900_000),
        duration: Some(duration),
    }
}

/// The name this version of this file's answer is kept under, with what a test before this one left
/// there given up first rather than counted against the folder's one budget.
fn entry_of(path: &Path) -> PathBuf {
    let entry = entry_path(&head::key(path));
    let _ = std::fs::remove_file(&entry);

    entry
}

/// What a run that starts after this one sees: the memo is empty, and the disk is everything the
/// run before it kept.
fn a_restart() {
    super::super::PROBED
        .lock()
        .expect("the memo's own lock")
        .clear();
}

fn set_modified_at(path: &Path, seconds_from_now: u64) {
    let moved = SystemTime::now() + std::time::Duration::from_secs(seconds_from_now);
    std::fs::File::options()
        .write(true)
        .open(path)
        .expect("the file just written")
        .set_modified(moved)
        .expect("a file whose time a test can move");
}

/// What the whole of this exists for: a folder opened after a restart is a folder of rows filled
/// with no probe at all, because the answer the probe gave is still on the disk.
#[test]
fn a_file_a_run_before_read_is_answered_from_the_disk_and_no_probe_is_entered() {
    let _one = ONE_AT_A_TIME.lock();
    let path = a_sound("answered-from-the-disk", "song.flac", 8);
    let _entry = entry_of(&path);

    remember(&path, a_track(180.0));
    a_restart();

    assert_eq!(
        probed(&path),
        Probed::NotAsked,
        "a cleared memo is the memo a run after this one starts with"
    );

    assert_eq!(
        ask(&path, || panic!(
            "a file the disk answers for was probed again"
        )),
        Some(a_track(180.0)),
        "the disk's answer is the ask's"
    );
    assert_eq!(
        probed(&path),
        Probed::Track(a_track(180.0)),
        "and a hit is held for the rest of this run"
    );
}

/// A hit is only a hit for the version of the file it was written for: a longer file, and one of the
/// same length saved again, are both files no read has answered for — the path is only the identity.
#[test]
fn a_file_saved_again_is_a_file_the_disk_does_not_answer_for() {
    let _one = ONE_AT_A_TIME.lock();
    let path = a_sound("saved-again", "song.flac", 8);
    let _entry = entry_of(&path);

    remember(&path, a_track(180.0));
    a_restart();

    // Same path, another length.
    std::fs::write(&path, vec![0u8; 16]).expect("a file a test can write");

    let mut reads = 0;
    assert_eq!(
        ask(&path, || {
            reads += 1;
            Some(a_track(200.0))
        }),
        Some(a_track(200.0))
    );
    assert_eq!(reads, 1, "a file of another length is probed again");

    // Same length, another time: what a file saved over itself is.
    a_restart();
    set_modified_at(&path, 60);

    assert_eq!(
        ask(&path, || {
            reads += 1;
            Some(a_track(240.0))
        }),
        Some(a_track(240.0))
    );
    assert_eq!(
        reads, 2,
        "and a file of the same length saved again is probed again too"
    );

    let _ = std::fs::remove_file(entry_path(&head::key(&path)));
}

/// A read that found nothing is not written anywhere: an emptiness on the disk outlives the run that
/// found it, and nothing would ever read the file again to find that a player had since been
/// installed (see `remember`).
#[test]
fn a_read_that_found_nothing_leaves_nothing_on_the_disk() {
    let _one = ONE_AT_A_TIME.lock();
    let path = a_sound("nothing-on-the-disk", "song.flac", 8);
    let entry = entry_of(&path);

    assert_eq!(ask(&path, || None), None);

    assert!(
        !entry.exists(),
        "a read that found nothing is not a fact about the file, on the disk least of all"
    );
}

/// What a run whose bytes are not this version's, cut short inside a write, or of no bytes at all
/// wrote: none of them is an answer, none of them is a panic, and each is given up rather than read
/// again for the rest of the machine's life.
#[test]
fn an_entry_of_another_version_or_one_cut_short_is_a_miss_and_is_given_up() {
    let _one = ONE_AT_A_TIME.lock();
    let path = a_sound("another-version", "song.flac", 8);
    let _ = entry_of(&path);

    remember(&path, a_track(180.0));
    let entry = entry_path(&head::key(&path));
    let whole = std::fs::read(&entry).expect("the entry just written");

    let mut other = whole.clone();
    other[0] = VERSION + 1;

    for (case, bytes) in [
        ("an entry of another version", other),
        ("an entry cut short", whole[..whole.len() / 2].to_vec()),
        ("an entry of no bytes at all", Vec::new()),
    ] {
        std::fs::write(&entry, &bytes).expect("an entry a test can write");
        a_restart();

        let mut reads = 0;
        assert_eq!(
            ask(&path, || {
                reads += 1;
                None
            }),
            None,
            "{case} is not an answer"
        );
        assert_eq!(reads, 1, "{case} is a miss, so the file is probed once");
        assert!(
            !entry.exists(),
            "{case} is given up rather than left on the disk"
        );
    }
}

/// The stamp is in the entry and not only in its name: bytes copied under another file's name are an
/// answer about the file they were written for, and a miss for this one.
#[test]
fn an_entry_copied_under_another_files_name_is_a_miss() {
    let _one = ONE_AT_A_TIME.lock();
    let written_for = a_sound("copied-under-a-name", "a.flac", 8);
    let copied_to = written_for.with_file_name("b.flac");
    std::fs::write(&copied_to, vec![0u8; 8]).expect("a file a test can write");
    let _written = entry_of(&written_for);
    let entry = entry_of(&copied_to);

    remember(&written_for, a_track(180.0));
    std::fs::copy(entry_path(&head::key(&written_for)), &entry).expect("an entry a test can copy");
    a_restart();

    let mut reads = 0;
    assert_eq!(
        ask(&copied_to, || {
            reads += 1;
            None
        }),
        None,
        "another file's stamp is not this file's answer"
    );
    assert_eq!(reads, 1, "so the file of that name is probed once");
    assert!(!entry.exists(), "and the entry is given up");

    let _ = std::fs::remove_file(entry_path(&head::key(&written_for)));
}
