use super::persistence::ConfigFileAsItWas;
use super::*;

use std::fs;

/// A file as a build before this one wrote it: every setting this app writes, less the numbers
/// this build writes about the file itself, which no build before it wrote.
fn file_before_this_build() -> Ini {
    let mut ini = Ini::new();
    for (section, keys) in AppConfig::default().to_ini().get_map_ref() {
        for (key, value) in keys {
            if key == migrate::CURRENT_VERSION_KEY || key == migrate::MIN_READER_KEY {
                continue;
            }
            ini.set(section, key, value.clone());
        }
    }

    ini
}

/// A value of that file set to what a build before this one wrote it as.
fn with(ini: &mut Ini, key: &str, value: &str) {
    ini.set(CONFIG_SECTION, key, Some(value.to_string()));
}

/// A file that states both numbers is read as it states them, and a file that states neither is
/// read as the oldest shape there is and the reader of now — which is what every file written
/// before either key existed says by being silent.
#[test]
fn a_file_that_says_nothing_about_itself_is_the_oldest_shape() {
    let ini = file_before_this_build();
    let stated = migrate::Stated::of(&ini);

    assert_eq!(stated.version, migrate::BASE_VERSION);
    assert_eq!(stated.min_reader, migrate::SELF_MIN_READER);
    assert!(
        !stated.is_ahead_of_this_build(),
        "a file of the oldest shape is not one a newer build wrote"
    );

    let mut ini = file_before_this_build();
    with(&mut ini, migrate::CURRENT_VERSION_KEY, "9");
    with(&mut ini, migrate::MIN_READER_KEY, "9");

    let stated = migrate::Stated::of(&ini);
    assert_eq!(stated.version, 9);
    assert_eq!(stated.min_reader, 9);
    assert!(
        stated.is_ahead_of_this_build(),
        "a file a build after this one wrote is one this build may read and must not write"
    );
}

/// A file at the shape of this build is a file there is nothing to bring up: the steps below it are
/// not run again, and the second read of a file this app has just written is a read that writes
/// nothing. It is the property the watcher depends on — a write of this app's own that asked for
/// another write would ask for one every second, forever.
#[test]
fn a_file_at_the_shape_of_this_build_has_nothing_to_bring_up() {
    let mut ini = AppConfig::default().to_ini();
    assert_eq!(
        ini.get(CONFIG_SECTION, migrate::CURRENT_VERSION_KEY),
        Some(migrate::CURRENT_VERSION.to_string()),
        "a file this app writes states the shape it is in"
    );

    let stated = migrate::Stated::of(&ini);
    let applied = migrate::apply_pending(&mut ini, stated).expect("a file of this build's own");

    assert!(
        !applied.touched_anything(),
        "there is nothing in a file of this build's own to write"
    );
    assert!(
        applied.keys.is_empty(),
        "and no setting of it was brought up to another value"
    );
}

/// The file a build before this one wrote has no numbers of its own, so a start reads it, brings up
/// what the steps below this shape say, and stamps it — which is what makes the file of a fresh
/// installation and the file of an old one the same file afterwards.
#[test]
fn an_older_file_is_brought_up_and_stamped() {
    let mut ini = file_before_this_build();
    let stated = migrate::Stated::of(&ini);

    let applied = migrate::apply_pending(&mut ini, stated).expect("a file of an older shape");

    assert!(
        applied.stamped,
        "a file that never carried the numbers of this build is stamped with them"
    );
    assert_eq!(
        ini.get(CONFIG_SECTION, migrate::CURRENT_VERSION_KEY),
        Some(migrate::CURRENT_VERSION.to_string())
    );
    assert_eq!(
        ini.get(CONFIG_SECTION, migrate::MIN_READER_KEY),
        Some(migrate::SELF_MIN_READER.to_string())
    );

    // And read again, it is a file at the shape of this build: the numbers are the ones it carries
    // and there is nothing below it left to run.
    let stated = migrate::Stated::of(&ini);
    let again = migrate::apply_pending(&mut ini, stated).expect("a file of this build's own");
    assert!(
        !again.touched_anything(),
        "one pass brings the file all the way up, so the next read has nothing to do"
    );
}

/// A `pin_key` that this app wrote when it shipped one key gains the middle button beside it.
///
/// The whole reason the step exists: a person whose file says `space` because the build they
/// installed shipped `space` ends up with the middle button AND the `space`, without being told to
/// open the file and without a reset row that would take every other setting with it.
#[test]
fn a_pin_key_this_app_shipped_gains_the_middle_button() {
    for shipped in ["space", "mouse3"] {
        let mut ini = file_before_this_build();
        with(&mut ini, "pin_key", shipped);

        let stated = migrate::Stated::of(&ini);
        let applied = migrate::apply_pending(&mut ini, stated).expect("a file of an older shape");

        assert_eq!(
            ini.get(CONFIG_SECTION, "pin_key"),
            Some("mouse3,space".to_string()),
            "`{shipped}` was a value this app wrote, so the setting is brought up to the pair"
        );
        assert!(
            applied.keys.contains(&"pin_key"),
            "and the change is named, so the file is written and the person is told"
        );
    }
}

/// And a `pin_key` anybody wrote is kept exactly as it stands.
///
/// The list the step is held to is the values this app shipped and no others: `space,page_down` and
/// `space,mouse3` are both spellings this app's own documentation offers, and a rule that read
/// "any binding naming `space`" as its own would take both of them from the person who wrote them.
#[test]
fn a_pin_key_a_person_wrote_is_left_exactly_as_it_stands() {
    for written in [
        "space,page_down",
        "space,mouse3",
        "mouse3,space",
        "ctrl+space",
        "f7",
        "page_down",
    ] {
        let mut ini = file_before_this_build();
        with(&mut ini, "pin_key", written);

        let stated = migrate::Stated::of(&ini);
        let applied = migrate::apply_pending(&mut ini, stated).expect("a file of an older shape");

        assert_eq!(
            ini.get(CONFIG_SECTION, "pin_key"),
            Some(written.to_string()),
            "`{written}` is nobody else's, and is kept"
        );
        assert!(
            !applied.keys.contains(&"pin_key"),
            "`{written}` is not a setting this build moved"
        );
    }
}

/// A file that names a shape newer than the build reading it is refused rather than read as the
/// newest shape there is: what it holds was written by a build this one has no steps for, and a
/// guess at it is how a file loses what it holds without anything saying so.
#[test]
fn a_file_from_a_newer_build_is_refused_rather_than_read() {
    let mut ini = file_before_this_build();
    with(
        &mut ini,
        migrate::CURRENT_VERSION_KEY,
        &(migrate::CURRENT_VERSION + 1).to_string(),
    );

    let stated = migrate::Stated::of(&ini);
    let refused = match migrate::apply_pending(&mut ini, stated) {
        Err(ahead) => ahead,
        Ok(_) => panic!("a shape this build has no steps for is not guessed at"),
    };

    assert_eq!(
        refused,
        migrate::Ahead::NewerThanThisBuild {
            file: migrate::CURRENT_VERSION + 1,
            build: migrate::CURRENT_VERSION,
        }
    );

    let says = refused.says_itself();
    assert!(
        says.contains(&(migrate::CURRENT_VERSION + 1).to_string())
            && says.contains(&migrate::CURRENT_VERSION.to_string()),
        "the refusal names both versions, since the alternative is settings that revert for a \
         reason nothing on the screen gives: `{says}`"
    );
}

/// The two numbers are keys this app writes, so they are named by a heading like every other setting
/// rather than showing up under `; Ungrouped`.
#[test]
fn the_numbers_of_the_file_are_written_and_named_like_every_other_setting() {
    let written = AppConfig::default().to_ini();
    let keys = written
        .get_map_ref()
        .get(CONFIG_SECTION)
        .expect("a settings section");

    for key in [migrate::CURRENT_VERSION_KEY, migrate::MIN_READER_KEY] {
        assert!(keys.contains_key(key), "`{key}` is a key this app writes");
        assert!(
            SETTING_GROUPS.iter().any(|(_, group)| group.contains(&key)),
            "`{key}` is named by a heading, or it is written under `; Ungrouped`"
        );
    }
}

/// Nothing this process writes goes into a file a newer build left behind, whatever asked for the
/// write: a tray row, a reset and a save of the app's own all end at the same door, and the door
/// answers the file rather than the caller.
#[test]
fn a_process_that_has_read_no_file_keeps_its_hands_off_nothing() {
    assert!(
        !settings_io::file_is_ahead_of_us(),
        "a process that has read no file is one with nothing to keep its hands off"
    );
}

/// The pairs the step names are the ones this app wrote, and the one it writes is not among them: a
/// step that read its own value of today as a value of the past would replace it with itself on
/// every read, which is a write every second for a file that never changed.
#[test]
fn the_step_of_a_pin_key_does_not_name_the_value_of_today() {
    let mut ini = file_before_this_build();
    with(&mut ini, "pin_key", "mouse3,space");

    let stated = migrate::Stated::of(&ini);
    let applied = migrate::apply_pending(&mut ini, stated).expect("a file of an older shape");

    assert!(
        !applied.keys.contains(&"pin_key"),
        "the value this build writes is not a value of the past, so a file holding it is left alone"
    );
}

/// A file a newer build left behind is read and never written, and its settings are still the ones
/// this build goes on using: what the process does is its own reading of the file, and what the file
/// holds is the newer build's.
///
/// It is the whole of the downgrade guard — a build that wrote its own view over such a file would
/// be the build that dropped the `pin_key` a user set, with nothing in the file saying so.
#[test]
fn a_start_on_a_file_a_newer_build_wrote_reads_it_and_writes_nothing() {
    let file = ConfigFileAsItWas::take();

    // A file the way a newer build leaves one: the settings of today's shape, the numbers of a build
    // after this one, and a value of that build's own this one has no name for.
    let mut ini = AppConfig::default().to_ini();
    with(
        &mut ini,
        migrate::CURRENT_VERSION_KEY,
        &(migrate::CURRENT_VERSION + 1).to_string(),
    );
    with(
        &mut ini,
        migrate::MIN_READER_KEY,
        &(migrate::SELF_MIN_READER + 1).to_string(),
    );
    with(&mut ini, "pin_key", "mouse3,space,page_down");
    let text = ordered_text(&ini);
    fs::write(&file.path, &text).expect("the file a newer build wrote");

    let loaded = AppConfig::load();

    assert_eq!(
        fs::read_to_string(&file.path).expect("the file where it was"),
        text,
        "the start wrote nothing over a file that says it is newer than this build"
    );
    assert_eq!(
        loaded.pin_key, "mouse3,space,page_down",
        "and what the file says is what the app uses, even where it is a spelling of a build this \
         one has no name for"
    );
}

/// And the same file seen by the watcher: the read happens, and what is on the disk afterwards is
/// what was on the disk before it.
#[test]
fn a_watch_on_a_file_a_newer_build_wrote_reads_it_and_writes_nothing() {
    let file = ConfigFileAsItWas::take();

    let mut ini = AppConfig::default().to_ini();
    with(
        &mut ini,
        migrate::CURRENT_VERSION_KEY,
        &(migrate::CURRENT_VERSION + 1).to_string(),
    );
    let text = ordered_text(&ini);
    fs::write(&file.path, &text).expect("the file a newer build wrote");

    let mut watched = AppConfig::default();
    watched.reload_from_disk();

    assert_eq!(
        fs::read_to_string(&file.path).expect("the file where it was"),
        text,
        "the watcher's own read wrote nothing of its own over the file"
    );
}

/// A file of this build's own is one the start writes nothing over: what it holds is what `save`
/// would write, and a write of a file that already says the same thing is a write that moves its
/// mtime for nothing — which the watcher would read back, and write again.
#[test]
fn a_start_on_a_file_of_this_builds_own_leaves_it_exactly_as_it_stands() {
    let file = ConfigFileAsItWas::take();

    let text = ordered_text(&AppConfig::default().to_ini());
    fs::write(&file.path, &text).expect("a file this app wrote");

    let _loaded = AppConfig::load();

    assert_eq!(
        fs::read_to_string(&file.path).expect("the file where it was"),
        text,
        "a file holding what this build writes is one there is nothing to write"
    );
}

/// A file with a `pin_key` a build of the past shipped is brought up on the start that reads it, and
/// the file on disk is what it was brought up to: the middle button beside the `space` that build
/// shipped, and the numbers of this build written down beside it.
#[test]
fn a_start_brings_up_a_pin_key_an_older_build_shipped() {
    let file = ConfigFileAsItWas::take();

    let mut ini = file_before_this_build();
    with(&mut ini, "pin_key", "space");
    fs::write(&file.path, ordered_text(&ini)).expect("a file an older build wrote");

    let loaded = AppConfig::load();

    assert_eq!(
        loaded.pin_key, "mouse3,space",
        "the setting is the one this build ships, with the middle button the older one did not"
    );

    let written = fs::read_to_string(&file.path).expect("the file where it was");
    assert!(
        written.contains("pin_key=mouse3,space"),
        "and the file was written with it: {written}"
    );
    assert!(
        written.contains(&format!(
            "{}={}",
            migrate::CURRENT_VERSION_KEY,
            migrate::CURRENT_VERSION
        )),
        "beside the shape this build writes, so the next start has nothing to bring up"
    );
}
